"""Orquesta la emisión: numeración, pedido de CAE, PDF e historial."""
from __future__ import annotations

import datetime as dt
from typing import Callable

from .arca import WSAA, WSFE, WSFESimulado, ErrorARCA
from .clientes import Almacen
from .pdf import generar_pdf
from .plantilla import armar_para_arca


def nombre_pdf(c: dict) -> str:
    return f"{c['cuit']}_{c['tipo']}_{c['pto_vta']:05d}-{c['nro']:08d}.pdf"


def emitir_lote(cbtes: list[dict], almacen: Almacen, simulacion: bool, archivo_excel: str,
                progreso: Callable[[int, int, str], None] | None = None) -> list[dict]:
    servicios: dict[str, object] = {}
    simulador = WSFESimulado()
    proximos: dict[tuple, int] = {}
    resultados = []

    for n, c in enumerate(cbtes, start=1):
        cli = c["cliente"]
        etiqueta = f"{cli.get('razon_social', c['cuit'])} — {c['tipo']} ID {c['id']}"
        if progreso:
            progreso(n, len(cbtes), etiqueta)
        res = {"id": c["id"], "emails": c.get("emails", []), "enviar_mail": c.get("enviar_mail", False),
               "cbte": None, "mail": "", "cuit": c["cuit"], "cliente": cli.get("razon_social", ""), "tipo": c["tipo"],
               "pto_vta": c["pto_vta"], "numero": None, "total": c["total"], "estado": "", "cae": "",
               "cae_vto": "", "mensaje": "", "pdf": None, "pdf_nombre": ""}
        try:
            if simulacion:
                ws = simulador
            else:
                if c["cuit"] not in servicios:
                    cert, clave = almacen.credenciales(c["cuit"])
                    if not cert or not clave:
                        raise ErrorARCA("El cliente no tiene certificado y clave cargados")
                    servicios[c["cuit"]] = WSFE(WSAA(c["cuit"], cert, clave, cli.get("entorno", "homologacion"),
                                                     almacen))
                ws = servicios[c["cuit"]]

            k = (c["cuit"], c["pto_vta"], c["tipo_cod"])
            if k not in proximos:
                proximos[k] = ws.ultimo_autorizado(c["pto_vta"], c["tipo_cod"]) + 1
            nro = proximos[k]
            if c["cotizacion"] is None:
                c = {**c, "cotizacion": ws.cotizacion(c["moneda"])}
            pedido = armar_para_arca(c, nro)
            try:
                r = ws.autorizar(pedido)
            except Exception as ex_red:
                # Si se cortó la conexión, verificamos si ARCA llegó a autorizarlo antes de reintentar
                try:
                    cons = ws.consultar(c["pto_vta"], c["tipo_cod"], nro)
                except Exception:
                    raise ErrorARCA(
                        f"⚠️ Se cortó la conexión con ARCA justo al pedir el CAE y no se pudo confirmar si quedó "
                        f"emitido. ANTES DE REINTENTAR, tocá 'Probar conexión' en el cliente: si el último "
                        f"autorizado de este tipo es {nro}, la factura SÍ se emitió (buscala en Mis Comprobantes). "
                        f"Detalle: {ex_red}")
                if cons and cons.get("CAE") and abs(cons["ImpTotal"] - c["total"]) < 0.01:
                    r = {"resultado": "A", "CAE": cons["CAE"], "CAEFchVto": cons["CAEFchVto"],
                         "errores": [], "observaciones": ["Recuperado por consulta tras corte de conexión"]}
                else:
                    raise ErrorARCA(f"Error de comunicación: {ex_red}")

            if r["resultado"] == "A" and r["CAE"]:
                proximos[k] = nro + 1
                c = {**c, "nro": nro}
                pdf = generar_pdf(c, r["CAE"], r["CAEFchVto"], simulado=simulacion)
                nom = nombre_pdf(c)
                if not simulacion:
                    try:
                        almacen.guardar_pdf(c["cuit"], nom, pdf)
                    except Exception as ex_pdf:  # el CAE ya está: no se pierde, se avisa
                        r["observaciones"].append(f"No se pudo guardar el PDF en el Drive: {ex_pdf}")
                    fila_hist = {
                        "fecha_emision": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "entorno": cli.get("entorno"), "cuit_emisor": c["cuit"],
                        "razon_social": cli.get("razon_social"), "tipo": c["tipo"], "pto_vta": c["pto_vta"],
                        "numero": nro, "fecha": c["fecha"].strftime("%d/%m/%Y"),
                        "receptor_doc": c["doc_nro"], "receptor": c["rec_nombre"],
                        "neto": c["imp_neto"], "iva": c["imp_iva"], "tributos": c["imp_trib"],
                        "total": c["total"], "moneda": c["moneda"], "cae": r["CAE"], "cae_vto": r["CAEFchVto"],
                        "id_excel": c["id"], "archivo_excel": archivo_excel, "pdf": nom,
                        "email": ";".join(c.get("emails", [])),
                    }
                    try:
                        almacen.registrar(fila_hist)
                    except Exception as ex_h:
                        r["observaciones"].append(f"No se pudo anotar en el historial: {ex_h}")
                res["cbte"] = {**c, "cae": r["CAE"], "cae_vto": r["CAEFchVto"]}
                res.update(estado="Autorizado" + (" (simulación)" if simulacion else ""), numero=nro,
                           cae=r["CAE"], cae_vto=r["CAEFchVto"], pdf=pdf, pdf_nombre=nom,
                           mensaje=" | ".join(r["observaciones"]))
            else:
                res.update(estado="Rechazado", mensaje=" | ".join(r["errores"] + r["observaciones"]) or "Rechazado por ARCA")
        except Exception as ex:
            res.update(estado="Error", mensaje=str(ex))
        resultados.append(res)
    return resultados
