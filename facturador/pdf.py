"""PDF del comprobante con el formato habitual de ARCA y el código QR (RG 4892)."""
from __future__ import annotations

import base64
import io
import json

import qrcode
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas

from . import tablas as T

ANCHO, ALTO = A4
M = 10 * mm


def _p(n: float) -> str:
    s = f"{n:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _cuit(c: str) -> str:
    c = str(c)
    return f"{c[:2]}-{c[2:10]}-{c[10:]}" if len(c) == 11 else c


def url_qr(c: dict, cae: str) -> str:
    datos = {
        "ver": 1, "fecha": c["fecha"].isoformat(), "cuit": int(c["cuit"]), "ptoVta": c["pto_vta"],
        "tipoCmp": c["tipo_cod"], "nroCmp": c["nro"], "importe": round(c["total"], 2),
        "moneda": c["moneda"], "ctz": c["cotizacion"], "tipoDocRec": c["doc_tipo"],
        "nroDocRec": int(c["doc_nro"] or 0), "tipoCodAut": "E", "codAut": int(cae),
    }
    return "https://www.afip.gob.ar/fe/qr/?p=" + base64.b64encode(json.dumps(datos).encode()).decode()


def _texto(cv, x, y, txt, tam=8, negrita=False, ancho=None, alin="izq"):
    fuente = "Helvetica-Bold" if negrita else "Helvetica"
    cv.setFont(fuente, tam)
    lineas = simpleSplit(str(txt), fuente, tam, ancho) if ancho else [str(txt)]
    for i, l in enumerate(lineas):
        yy = y - i * (tam + 2)
        if alin == "der":
            cv.drawRightString(x, yy, l)
        elif alin == "cen":
            cv.drawCentredString(x, yy, l)
        else:
            cv.drawString(x, yy, l)
    return y - len(lineas) * (tam + 2)


def generar_pdf(c: dict, cae: str, cae_vto: str, simulado: bool = False) -> bytes:
    buf = io.BytesIO()
    cv = canvas.Canvas(buf, pagesize=A4)
    cv.setTitle(f"{c['nombre_tipo']} {c['letra']} {c['pto_vta']:05d}-{c['nro']:08d}")
    for copia in ("ORIGINAL", "DUPLICADO"):
        _pagina(cv, c, cae, cae_vto, copia, simulado)
        cv.showPage()
    cv.save()
    return buf.getvalue()


def _pagina(cv, c, cae, cae_vto, copia, simulado):
    cli = c["cliente"]
    letra = c["letra"]
    top = ALTO - M
    cv.setLineWidth(0.8)
    # Copia
    cv.rect(M, top - 8 * mm, ANCHO - 2 * M, 8 * mm)
    _texto(cv, ANCHO / 2, top - 5.5 * mm, copia, 11, True, alin="cen")

    # Cabecera
    y0 = top - 8 * mm
    alto_cab = 42 * mm
    cv.rect(M, y0 - alto_cab, ANCHO - 2 * M, alto_cab)
    cv.line(ANCHO / 2, y0 - alto_cab, ANCHO / 2, y0 - 15 * mm)
    # Letra
    lado = 15 * mm
    cv.setFillColor(colors.white)
    cv.rect(ANCHO / 2 - lado / 2, y0 - lado, lado, lado, fill=1)
    cv.setFillColor(colors.black)
    _texto(cv, ANCHO / 2, y0 - 9 * mm, letra, 24, True, alin="cen")
    _texto(cv, ANCHO / 2, y0 - 13.5 * mm, f"COD. {c['tipo_cod']:03d}", 6.5, True, alin="cen")

    # Emisor (izquierda)
    xi = M + 4 * mm
    y = y0 - 9 * mm
    y = _texto(cv, xi, y, cli.get("nombre_fantasia") or cli.get("razon_social", ""), 14, True, ANCHO / 2 - M - 22 * mm)
    y -= 3 * mm
    for et, val in (("Razón Social:", cli.get("razon_social", "")),
                    ("Domicilio Comercial:", cli.get("domicilio", "")),
                    ("Condición frente al IVA:", cli.get("cond_iva", ""))):
        cv.setFont("Helvetica-Bold", 8); cv.drawString(xi, y, et)
        w = cv.stringWidth(et, "Helvetica-Bold", 8) + 2
        y = _texto(cv, xi + w, y, val, 8, ancho=ANCHO / 2 - M - 8 * mm - w) - 1.5

    # Comprobante (derecha)
    xd = ANCHO / 2 + 10 * mm
    y = y0 - 9 * mm
    if len(c["nombre_tipo"]) > 20:
        y_fin = _texto(cv, xd, y + 1 * mm, c["nombre_tipo"], 10.5, True, ancho=ANCHO - M - xd - 2 * mm)
        y = min(y - 8 * mm, y_fin - 1.5 * mm)
    else:
        _texto(cv, xd, y, c["nombre_tipo"], 16, True)
        y -= 8 * mm
    cv.setFont("Helvetica-Bold", 9)
    cv.drawString(xd, y, f"Punto de Venta: {c['pto_vta']:05d}     Comp. Nro: {c['nro']:08d}")
    y -= 5 * mm
    cv.drawString(xd, y, f"Fecha de Emisión: {c['fecha']:%d/%m/%Y}")
    y -= 6 * mm
    for et, val in (("CUIT:", _cuit(c["cuit"])), ("Ingresos Brutos:", cli.get("iibb", "")),
                    ("Fecha de Inicio de Actividades:", cli.get("inicio_act", ""))):
        cv.setFont("Helvetica-Bold", 8); cv.drawString(xd, y, et)
        cv.setFont("Helvetica", 8); cv.drawString(xd + cv.stringWidth(et, "Helvetica-Bold", 8) + 2, y, str(val))
        y -= 4.2 * mm

    # Período
    y = y0 - alto_cab
    if c["concepto"] in (2, 3):
        cv.rect(M, y - 7 * mm, ANCHO - 2 * M, 7 * mm)
        cv.setFont("Helvetica-Bold", 8)
        cv.drawString(M + 4 * mm, y - 4.6 * mm,
                      f"Período Facturado Desde: {c['serv_desde']:%d/%m/%Y}     Hasta: {c['serv_hasta']:%d/%m/%Y}"
                      f"     Fecha de Vto. para el pago: {c['vto_pago']:%d/%m/%Y}")
        y -= 7 * mm

    # Receptor
    alto_rec = 18 * mm
    cv.rect(M, y - alto_rec - 1 * mm, ANCHO - 2 * M, alto_rec)
    yy = y - 5 * mm
    doc_et = f"{c['doc_tipo_txt']}:" if c["doc_tipo_txt"] != "Sin identificar" else "Doc.:"
    pares = [
        ((doc_et, _cuit(c["doc_nro"]) if c["doc_tipo"] == 80 else (c["doc_nro"] if c["doc_nro"] != "0" else "")),
         ("Apellido y Nombre / Razón Social:", c["rec_nombre"])),
        (("Condición frente al IVA:", c["rec_cond_iva"]), ("Domicilio:", c["rec_domicilio"])),
        (("Condición de venta:", c["cond_venta"]), ("", "")),
    ]
    for (e1, v1), (e2, v2) in pares:
        cv.setFont("Helvetica-Bold", 8); cv.drawString(M + 3 * mm, yy, e1)
        cv.setFont("Helvetica", 8); cv.drawString(M + 3 * mm + cv.stringWidth(e1, "Helvetica-Bold", 8) + 2, yy, str(v1))
        if e2:
            x2 = M + 75 * mm
            cv.setFont("Helvetica-Bold", 8); cv.drawString(x2, yy, e2)
            w = cv.stringWidth(e2, "Helvetica-Bold", 8) + 2
            _texto(cv, x2 + w, yy, v2, 8, ancho=ANCHO - M - x2 - w - 3 * mm)
        yy -= 5 * mm
    y = y - alto_rec - 1 * mm

    if c["asociados"]:
        a = c["asociados"][0]
        y -= 4.5 * mm
        cv.setFont("Helvetica", 8)
        cv.drawString(M + 3 * mm, y, f"Comprobante asociado: {a['_txt']} {a['PtoVta']:05d}-{a['Nro']:08d}")
        y -= 1.5 * mm

    # Tabla de ítems
    y -= 3 * mm
    discrimina = letra in ("A", "M")
    if discrimina:
        cols = [("Código", 16), ("Producto / Servicio", 64), ("Cantidad", 16), ("U. Medida", 16),
                ("Precio Unit.", 20), ("% Bonif", 12), ("Subtotal", 20), ("Alícuota IVA", 14), ("Subtotal c/IVA", 22)]
    else:
        cols = [("Código", 16), ("Producto / Servicio", 80), ("Cantidad", 16), ("U. Medida", 18),
                ("Precio Unit.", 22), ("% Bonif", 12), ("Imp. Bonif.", 16), ("Subtotal", 20)]
    escala = (ANCHO - 2 * M) / (sum(w for _, w in cols) * mm)
    xs = [M]
    for _, w in cols:
        xs.append(xs[-1] + w * mm * escala)
    cv.setFillColor(colors.HexColor("#D9D9D9"))
    cv.rect(M, y - 6 * mm, ANCHO - 2 * M, 6 * mm, fill=1, stroke=1)
    cv.setFillColor(colors.black)
    for (t, _), x0, x1 in zip(cols, xs, xs[1:]):
        _texto(cv, (x0 + x1) / 2, y - 4 * mm, t, 7, True, alin="cen")
    y -= 6 * mm + 4 * mm
    for l in c["lineas"]:
        ancho_desc = xs[2] - xs[1] - 2 * mm
        if discrimina:
            pu = l["neto"] / l["cantidad"] / (1 - l["bonif"] / 100) if l["cantidad"] and l["bonif"] < 100 else 0
            alic_txt = (l["alicuota"].replace(".", ",") + "%") if l["alicuota"] not in ("Exento", "No gravado") else l["alicuota"]
            valores = [l["codigo"], None, f"{l['cantidad']:g}".replace(".", ","), l["unidad"], _p(pu),
                       _p(l["bonif"]), _p(l["neto"]), alic_txt, _p(round(l["neto"] * (1 + l["tasa"]), 2))]
        else:
            # B y C: precio final al consumidor
            if letra == "B" and not c["precios_con_iva"]:
                pu = l["precio"] * (1 + l["tasa"])
                sub = round(l["importe"] * (1 + l["tasa"]), 2)
            else:
                pu, sub = l["precio"], l["importe"]
            bonif_imp = round(l["cantidad"] * pu * l["bonif"] / 100, 2)
            valores = [l["codigo"], None, f"{l['cantidad']:g}".replace(".", ","), l["unidad"], _p(pu),
                       _p(l["bonif"]), _p(bonif_imp), _p(sub)]
        cv.setFont("Helvetica", 7.5)
        for i, (v, x0, x1) in enumerate(zip(valores, xs, xs[1:])):
            if i == 1:
                continue
            if i in (0, 3):
                cv.drawString(x0 + 1 * mm, y, str(v))
            else:
                cv.drawRightString(x1 - 1 * mm, y, str(v))
        y_fin = _texto(cv, xs[1] + 1 * mm, y, l["descripcion"], 7.5, ancho=ancho_desc)
        y = y_fin - 2 * mm
        if y < 95 * mm:
            break  # (para facturas muy largas se corta la lista; caso raro)

    # Pie: totales
    base_pie = 58 * mm
    alto_tot = 36 * mm
    cv.rect(M, base_pie, ANCHO - 2 * M, alto_tot)
    xt = ANCHO - M - 4 * mm
    yt = base_pie + alto_tot - 6 * mm
    sim = "USD " if c["moneda"] == "DOL" else "EUR " if c["moneda"] == "060" else "$"
    renglones = []
    if discrimina:
        renglones.append(("Importe Neto Gravado:", c["imp_neto"]))
        por_alic = {i["_alic"]: i["Importe"] for i in c["iva"]}
        for a in ("27", "21", "10.5", "5", "2.5", "0"):
            if a in por_alic:
                renglones.append((f"IVA {a.replace('.', ',')}%:", por_alic[a]))
        if c["imp_opex"]:
            renglones.append(("Importe Exento:", c["imp_opex"]))
        if c["imp_no_gravado"]:
            renglones.append(("Importe No Gravado:", c["imp_no_gravado"]))
    else:
        renglones.append(("Subtotal:", c["total"] - c["imp_trib"]))
    renglones.append(("Importe Otros Tributos:", c["imp_trib"]))
    for et, val in renglones:
        cv.setFont("Helvetica-Bold", 8.5)
        cv.drawRightString(xt - 32 * mm, yt, et)
        cv.drawRightString(xt, yt, sim + _p(val))
        yt -= 4.5 * mm
    cv.setFont("Helvetica-Bold", 10)
    cv.drawRightString(xt - 32 * mm, yt - 1 * mm, "Importe Total:")
    cv.drawRightString(xt, yt - 1 * mm, sim + _p(c["total"]))

    # Izquierda del pie: tributos + leyendas
    xl = M + 3 * mm
    yl = base_pie + alto_tot - 5 * mm
    if c["tributos"]:
        cv.setFont("Helvetica-Bold", 7.5); cv.drawString(xl, yl, "Otros tributos"); yl -= 3.8 * mm
        for t in c["tributos"]:
            cv.setFont("Helvetica", 7.5)
            cv.drawString(xl, yl, f"{t['Desc'][:40]}  {_p(t['Alic'])}%  {sim}{_p(t['Importe'])}")
            yl -= 3.6 * mm
    leyendas = []
    fce = c.get("fce") or {}
    if fce.get("cbu"):
        leyendas.append(f"CBU del emisor: {fce['cbu']}" + (f"   Alias: {fce['alias']}" if fce.get("alias") else ""))
        leyendas.append(f"Opción de transferencia: {fce['transferencia']} - {T.FCE_TRANSFERENCIA[fce['transferencia']]}")
    if fce.get("anulacion"):
        leyendas.append("Anulación de FCE: " + ("Sí" if fce["anulacion"] == "S" else "No"))
    leyendas += list(c["leyendas"])
    if c["moneda"] != "PES":
        leyendas.append(f"Moneda: {c['moneda_txt']} — Tipo de cambio: $ {_p(c['cotizacion'])}"
                        f" — Equivalente en pesos: $ {_p(c['total'] * c['cotizacion'])}"
                        + (" — Cancela en moneda extranjera" if c.get("cancela_mon_ext") else ""))
    if letra == "M":
        leyendas.append("OPERACIÓN SUJETA A RETENCIÓN")
    for ley in leyendas:
        yl = _texto(cv, xl, yl, ley, 7.5, ancho=95 * mm) - 1

    # Régimen de Transparencia Fiscal al Consumidor (Ley 27.743) en comprobantes B
    if letra == "B":
        yr = base_pie + 9 * mm
        cv.setFont("Helvetica-Oblique", 7.5)
        cv.drawString(xl, yr, "Régimen de Transparencia Fiscal al Consumidor (Ley 27.743)")
        cv.setFont("Helvetica", 7.5)
        cv.drawString(xl, yr - 3.8 * mm, f"IVA Contenido: {sim}{_p(c['imp_iva'])}")
        otros = sum(t["Importe"] for t in c["tributos"] if t["Id"] in (1, 4))
        cv.drawString(xl + 48 * mm, yr - 3.8 * mm, f"Otros Impuestos Nacionales Indirectos: {sim}{_p(otros)}")

    # QR + CAE
    img = qrcode.make(url_qr(c, cae), box_size=4, border=1)
    b = io.BytesIO(); img.save(b, format="PNG"); b.seek(0)
    cv.drawImage(ImageReader(b), M, M + 2 * mm, 36 * mm, 36 * mm)
    _texto(cv, M + 40 * mm, M + 28 * mm, "Comprobante Autorizado", 10, True)
    _texto(cv, M + 40 * mm, M + 22 * mm,
           "Esta Agencia no se responsabiliza por los datos ingresados en el detalle de la operación", 6.5)
    cv.setFont("Helvetica-Bold", 9)
    cv.drawRightString(ANCHO - M - 4 * mm, M + 28 * mm, f"CAE N°: {cae}")
    vto = f"{cae_vto[6:8]}/{cae_vto[4:6]}/{cae_vto[:4]}" if len(cae_vto) == 8 else cae_vto
    cv.drawRightString(ANCHO - M - 4 * mm, M + 22 * mm, f"Fecha de Vto. de CAE: {vto}")
    cv.setFont("Helvetica", 7)
    cv.drawRightString(ANCHO - M - 4 * mm, M + 4 * mm, "Generado con el Facturador de Estudio Marsan")

    if simulado:
        cv.saveState()
        cv.setFillColor(colors.Color(1, 0, 0, alpha=0.18))
        cv.setFont("Helvetica-Bold", 60)
        cv.translate(ANCHO / 2, ALTO / 2); cv.rotate(40)
        cv.drawCentredString(0, 20, "SIMULACIÓN")
        cv.setFont("Helvetica-Bold", 30)
        cv.drawCentredString(0, -25, "SIN VALIDEZ FISCAL")
        cv.restoreState()
