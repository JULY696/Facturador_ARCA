"""Envío de comprobantes por mail (SMTP) desde la casilla del estudio."""
from __future__ import annotations

import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, make_msgid

PROVEEDORES = {
    "Hostinger": {"servidor": "smtp.hostinger.com", "puerto": 465, "seguridad": "SSL"},
    "Gmail / Google Workspace": {"servidor": "smtp.gmail.com", "puerto": 465, "seguridad": "SSL"},
    "Microsoft 365 / Outlook": {"servidor": "smtp.office365.com", "puerto": 587, "seguridad": "STARTTLS"},
    "Otro": {"servidor": "", "puerto": 465, "seguridad": "SSL"},
}

ASUNTO_DEF = "{emisor} - {comprobante} {numero}"
CUERPO_DEF = """Estimado/a {receptor}:

Adjuntamos {comprobante} N° {numero} de fecha {fecha} emitida por {emisor}, por un total de {moneda} {total}.

CAE: {cae}

Ante cualquier consulta, quedamos a disposición.

Saludos cordiales,
ESTUDIO MARSAN
info@estudiomarsan.com.ar"""

DEFECTO = {
    "proveedor": "Hostinger", "servidor": "smtp.hostinger.com", "puerto": 465, "seguridad": "SSL",
    "usuario": "info@estudiomarsan.com.ar", "password": "",
    "remitente_email": "info@estudiomarsan.com.ar", "remitente_nombre": "ESTUDIO MARSAN",
    "copia_cliente": True, "copia_estudio": False,
    "asunto": ASUNTO_DEF, "cuerpo": CUERPO_DEF,
}


def _p(n: float) -> str:
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _nombre(nombre: str, letra: str) -> str:
    n = nombre[0] + nombre[1:].lower()
    return f"{n.replace('mipymes (fce)', 'MiPyMEs (FCE)')} {letra}"


def datos_plantilla(c: dict) -> dict:
    cli = c["cliente"]
    simbolo = {"PES": "$", "DOL": "USD", "060": "EUR"}.get(c["moneda"], c["moneda"])
    return {
        "emisor": cli.get("razon_social", ""), "receptor": c.get("rec_nombre") or "cliente",
        "comprobante": _nombre(c["nombre_tipo"], c["letra"]),
        "numero": f"{c['pto_vta']:05d}-{c['nro']:08d}", "fecha": c["fecha"].strftime("%d/%m/%Y"),
        "total": _p(c["total"]), "moneda": simbolo, "cae": c.get("cae", ""),
    }


def datos_desde_historial(h: dict) -> dict:
    from . import tablas as T
    cod, letra, nombre = T.TIPOS_CBTE.get(h["tipo"], (0, "", h["tipo"]))
    simbolo = {"PES": "$", "DOL": "USD", "060": "EUR"}.get(h.get("moneda"), h.get("moneda", "$"))
    return {
        "emisor": h.get("razon_social", ""), "receptor": h.get("receptor") or "cliente",
        "comprobante": _nombre(nombre, letra),
        "numero": f"{int(h['pto_vta']):05d}-{int(h['numero']):08d}", "fecha": h.get("fecha", ""),
        "total": _p(float(h.get("total") or 0)), "moneda": simbolo, "cae": h.get("cae", ""),
    }


def _rellenar(txt: str, datos: dict) -> str:
    for k, v in datos.items():
        txt = txt.replace("{" + k + "}", str(v))
    return txt


def armar_mensaje(cfg: dict, datos: dict, cli: dict, pdf: bytes, nombre_pdf: str,
                  destinos: list[str]) -> EmailMessage:
    cfg = {**DEFECTO, **cfg}
    msg = EmailMessage()
    msg["From"] = formataddr((cfg["remitente_nombre"], cfg["remitente_email"]))
    msg["To"] = ", ".join(destinos)
    cc = []
    if cfg.get("copia_cliente") and cli.get("email"):
        cc.append(cli["email"])
    if cfg.get("copia_estudio"):
        cc.append(cfg["remitente_email"])
    if cc:
        msg["Cc"] = ", ".join(cc)
    if cli.get("email"):
        msg["Reply-To"] = cli["email"]  # si el receptor responde, le llega al cliente del estudio
    msg["Subject"] = _rellenar(cfg["asunto"] or ASUNTO_DEF, datos)
    msg["Message-ID"] = make_msgid(domain=cfg["remitente_email"].split("@")[-1])
    msg.set_content(_rellenar(cfg["cuerpo"] or CUERPO_DEF, datos))
    msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=nombre_pdf)
    return msg


class Mailer:
    def __init__(self, cfg: dict):
        self.cfg = {**DEFECTO, **cfg}
        self.smtp = None

    def __enter__(self):
        c = self.cfg
        ctx = ssl.create_default_context()
        if c["seguridad"] == "SSL":
            self.smtp = smtplib.SMTP_SSL(c["servidor"], int(c["puerto"]), context=ctx, timeout=30)
        else:
            self.smtp = smtplib.SMTP(c["servidor"], int(c["puerto"]), timeout=30)
            if c["seguridad"] == "STARTTLS":
                self.smtp.starttls(context=ctx)
        if c.get("usuario"):
            self.smtp.login(c["usuario"], c["password"])
        return self

    def enviar(self, msg: EmailMessage):
        self.smtp.send_message(msg)

    def __exit__(self, *_):
        try:
            self.smtp.quit()
        except Exception:
            pass


def enviar_lote(cfg: dict, resultados: list[dict], forzar_todos: bool = False) -> list[dict]:
    """Envía los autorizados que tienen 'Enviar por mail' = S (o todos con email si forzar_todos).
    Devuelve los resultados con el campo 'mail' actualizado."""
    pendientes = [r for r in resultados if r.get("cbte") and r.get("emails")
                  and (r.get("enviar_mail") or forzar_todos) and not r.get("mail", "").startswith("Enviado")]
    if not pendientes:
        return resultados
    try:
        with Mailer(cfg) as m:
            for r in pendientes:
                msg = armar_mensaje(cfg, datos_plantilla(r["cbte"]), r["cbte"]["cliente"],
                                    r["pdf"], r["pdf_nombre"], r["emails"])
                r["mail_info"] = {"para": msg["To"], "cc": msg.get("Cc", "") or "", "asunto": str(msg["Subject"])}
                try:
                    m.enviar(msg)
                    r["mail"] = "Enviado a " + ", ".join(r["emails"])
                except Exception as ex:
                    r["mail"] = f"Error: {ex}"
    except smtplib.SMTPAuthenticationError:
        for r in pendientes:
            r["mail"] = "Error: usuario o contraseña del mail incorrectos"
    except Exception as ex:
        for r in pendientes:
            r["mail"] = f"Error de conexión con el servidor de mail: {ex}"
    return resultados


def probar(cfg: dict, destino: str):
    cfg = {**DEFECTO, **cfg}
    msg = EmailMessage()
    msg["From"] = formataddr((cfg["remitente_nombre"], cfg["remitente_email"]))
    msg["To"] = destino
    msg["Subject"] = "Prueba del Facturador ARCA - Estudio Marsan"
    msg.set_content("Si recibiste este mail, el envío de facturas está bien configurado.")
    with Mailer(cfg) as m:
        m.enviar(msg)
