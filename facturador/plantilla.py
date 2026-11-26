"""Plantilla Excel de carga, lectura, validación y cálculo de importes."""
from __future__ import annotations

import calendar
import datetime as dt
import io
import re
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

from . import tablas as T

# ------------------------------------------------------------------ columnas
COLS_CBTE = [
    ("id", "ID", "Identificador tuyo de la factura (ej. 1, 2, 3). Une esta fila con sus ítems y tributos."),
    ("cuit_emisor", "CUIT emisor", "CUIT del cliente del estudio que factura (sin guiones). Tiene que estar dado de alta en la app."),
    ("tipo", "Tipo comprobante", "FA, FB, FC, NCA, NCB, NCC, NDA, NDB, NDC, FM..."),
    ("pto_vta", "Punto de venta", "Vacío = el punto de venta por defecto del cliente."),
    ("fecha", "Fecha", "Vacío = hoy. Productos: hasta 5 días antes/después. Servicios: hasta 10."),
    ("concepto", "Concepto", "Productos / Servicios / Productos y Servicios. Vacío = el del cliente."),
    ("serv_desde", "Servicio desde", "Solo servicios. Vacío = primer día del mes de la fecha."),
    ("serv_hasta", "Servicio hasta", "Solo servicios. Vacío = último día del mes de la fecha."),
    ("vto_pago", "Vto. pago", "Solo servicios. Vacío = la fecha del comprobante."),
    ("doc_tipo", "Receptor tipo doc", "CUIT, CUIL, DNI, CDI, Pasaporte, Sin identificar."),
    ("doc_nro", "Receptor nro doc", "Sin guiones. 'Sin identificar' = dejar vacío."),
    ("rec_nombre", "Receptor nombre / razón social", ""),
    ("rec_domicilio", "Receptor domicilio", ""),
    ("rec_cond_iva", "Receptor condición IVA", "Responsable Inscripto, Monotributo, Consumidor Final, Exento..."),
    ("cond_venta", "Condición de venta", "Contado, Transferencia bancaria, Cuenta corriente..."),
    ("moneda", "Moneda", "Pesos / Dólares / Euros. Vacío = Pesos."),
    ("cotizacion", "Cotización", "Solo moneda extranjera."),
    ("precios_con_iva", "Precios con IVA incluido", "S = los precios de los ítems ya incluyen IVA (típico Factura B). Vacío = N."),
    ("asoc_tipo", "Cbte asociado tipo", "Obligatorio en notas de crédito/débito (ej. FA)."),
    ("asoc_pto_vta", "Cbte asociado pto vta", ""),
    ("asoc_nro", "Cbte asociado número", ""),
    ("asoc_fecha", "Cbte asociado fecha", ""),
    ("leyenda", "Leyenda / observaciones", "Texto libre que sale en la factura. Se suma a la leyenda fija del cliente."),
    ("email", "Email receptor", "Para mandarle el PDF. Podés poner varios separados por ;"),
    ("enviar_mail", "Enviar por mail", "S = mandar el PDF por mail desde info@estudiomarsan.com.ar. Vacío = no."),
    ("cancela_mon_ext", "Cancela en moneda extranjera", "Solo dólares/euros: S si el cliente paga en esa misma moneda. Vacío = N."),
    ("fce_cbu", "FCE CBU", "Solo FCE: CBU del emisor (22 dígitos). Vacío = el CBU cargado en el cliente."),
    ("fce_alias", "FCE alias", "Solo FCE: alias del CBU (opcional). Vacío = el del cliente."),
    ("fce_transf", "FCE transferencia", "Solo FCE: SCA (circulación abierta) o ADC (agente de depósito colectivo). Vacío = el del cliente."),
    ("fce_anulacion", "FCE anulación", "Solo notas de crédito/débito FCE: S si es por rechazo/anulación de la FCE. Vacío = N."),
]
COLS_ITEM = [
    ("id", "ID comprobante", "El mismo ID de la hoja Comprobantes."),
    ("codigo", "Código", "Opcional."),
    ("descripcion", "Descripción", "Producto o servicio, como querés que salga en la factura."),
    ("cantidad", "Cantidad", "Vacío = 1."),
    ("unidad", "Unidad de medida", "unidades, horas, kg, m2, servicio... Vacío = unidades."),
    ("precio", "Precio unitario", "Sin IVA en A/B (salvo 'Precios con IVA incluido' = S). Precio final en C."),
    ("bonif", "Bonificación %", "Opcional."),
    ("alicuota", "Alícuota IVA", "21, 10.5, 27, 5, 2.5, 0, Exento, No gravado. Se ignora en Factura C."),
]
COLS_TRIB = [
    ("id", "ID comprobante", ""),
    ("tipo", "Tipo", "Percepción IIBB, Percepción IVA, Impuestos internos, Otros..."),
    ("descripcion", "Descripción", "Ej. Percepción IIBB CABA"),
    ("base", "Base imponible", ""),
    ("alicuota", "Alícuota %", ""),
    ("importe", "Importe", "Vacío = base x alícuota."),
]

ALIAS_COND_IVA = {
    "ri": "IVA Responsable Inscripto", "responsable inscripto": "IVA Responsable Inscripto",
    "iva responsable inscripto": "IVA Responsable Inscripto",
    "monotributo": "Responsable Monotributo", "monotributista": "Responsable Monotributo",
    "responsable monotributo": "Responsable Monotributo",
    "cf": "Consumidor Final", "consumidor final": "Consumidor Final",
    "exento": "IVA Sujeto Exento", "iva sujeto exento": "IVA Sujeto Exento", "iva exento": "IVA Sujeto Exento",
}

AZUL = PatternFill("solid", fgColor="1F3864")


def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


# ------------------------------------------------------------------ plantilla
def generar_plantilla(clientes: list[dict] | None = None) -> bytes:
    wb = Workbook()
    hojas = [("Comprobantes", COLS_CBTE), ("Items", COLS_ITEM), ("Tributos", COLS_TRIB)]
    ws_listas = None
    for i, (nombre, cols) in enumerate(hojas):
        ws = wb.active if i == 0 else wb.create_sheet(nombre)
        ws.title = nombre
        for j, (_k, titulo, ayuda) in enumerate(cols, start=1):
            c = ws.cell(row=1, column=j, value=titulo)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = AZUL
            c.alignment = Alignment(wrap_text=True, vertical="center")
            if ayuda:
                c.comment = Comment(ayuda, "Facturador")
            ws.column_dimensions[c.column_letter].width = max(14, len(titulo) + 2)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 32

    # Listas para desplegables
    ws_listas = wb.create_sheet("Listas")
    listas = {
        "A": list(T.TIPOS_CBTE), "B": list(T.CONCEPTOS), "C": list(T.TIPOS_DOC),
        "D": list(T.COND_IVA_RECEPTOR), "E": T.CONDICIONES_VENTA, "F": list(T.MONEDAS),
        "G": ["S", "N"], "H": list(T.ALICUOTAS), "I": list(T.TRIBUTOS),
        "J": [c["cuit"] for c in (clientes or [])],
        "L": list(T.FCE_TRANSFERENCIA),
    }
    titulos = ["Tipos", "Conceptos", "Tipo doc", "Cond. IVA", "Cond. venta", "Moneda", "S/N", "Alícuotas", "Tributos", "CUIT clientes", "FCE transf."]
    for (col, valores), tit in zip(listas.items(), titulos):
        ws_listas[f"{col}1"] = tit
        ws_listas[f"{col}1"].font = Font(bold=True)
        for r, v in enumerate(valores, start=2):
            ws_listas[f"{col}{r}"] = v
        ws_listas.column_dimensions[col].width = 28
    if clientes:
        ws_listas["K1"] = "Cliente"
        ws_listas["K1"].font = Font(bold=True)
        for r, c in enumerate(clientes, start=2):
            ws_listas[f"K{r}"] = c.get("razon_social", "")
        ws_listas.column_dimensions["K"].width = 36

    def dv(ws, col_listas, n, rango):
        v = DataValidation(type="list", formula1=f"=Listas!${col_listas}$2:${col_listas}${n + 1}", allow_blank=True)
        v.error, v.showErrorMessage = "Elegí un valor de la lista", False
        ws.add_data_validation(v)
        v.add(rango)

    wc, wi, wt = wb["Comprobantes"], wb["Items"], wb["Tributos"]
    letra = {k: wc.cell(row=1, column=j).column_letter for j, (k, *_r) in enumerate(COLS_CBTE, start=1)}
    R = "2:1000"
    def rg(k):
        a, b = R.split(":")
        return f"{letra[k]}{a}:{letra[k]}{b}"
    dv(wc, "A", len(listas["A"]), rg("tipo"))
    dv(wc, "A", len(listas["A"]), rg("asoc_tipo"))
    dv(wc, "B", len(listas["B"]), rg("concepto"))
    dv(wc, "C", len(listas["C"]), rg("doc_tipo"))
    dv(wc, "D", len(listas["D"]), rg("rec_cond_iva"))
    dv(wc, "E", len(listas["E"]), rg("cond_venta"))
    dv(wc, "F", len(listas["F"]), rg("moneda"))
    dv(wc, "G", 2, rg("precios_con_iva"))
    for k in ("enviar_mail", "cancela_mon_ext", "fce_anulacion"):
        dv(wc, "G", 2, rg(k))
    dv(wc, "L", 2, rg("fce_transf"))
    if clientes:
        dv(wc, "J", len(clientes), rg("cuit_emisor"))
    dv(wi, "H", len(listas["H"]), "H2:H5000")
    dv(wt, "I", len(listas["I"]), "B2:B2000")
    for k in ("fecha", "serv_desde", "serv_hasta", "vto_pago", "asoc_fecha"):
        for r in range(2, 301):
            wc[f"{letra[k]}{r}"].number_format = "DD/MM/YYYY"
    for r in range(2, 1001):
        wi[f"F{r}"].number_format = "#,##0.00"

    # Ejemplos
    cuit_ej = clientes[0]["cuit"] if clientes else "20111111112"
    es_ri = bool(clientes) and clientes[0].get("cond_iva") == "IVA Responsable Inscripto"
    hoy = dt.date.today()
    ej_c = [
        [1, cuit_ej, "FA" if es_ri else "FC", None, hoy, "Servicios", None, None, None, "CUIT", "30712345671",
         "Empresa Ejemplo S.A.", "Av. Corrientes 1234, CABA", "IVA Responsable Inscripto", "Transferencia bancaria",
         "Pesos", None, "N", None, None, None, None, "Honorarios del mes", ""],
        [2, cuit_ej, "FB" if es_ri else "FC", None, hoy, "Productos", None, None, None, "DNI", "30111222",
         "Juan Pérez", "", "Consumidor Final", "Contado", "Pesos", None, "S", None, None, None, None, "", ""],
    ]
    for fila in ej_c:
        wc.append(fila)
    for fila in ([1, "", "Asesoramiento contable mensual", 1, "servicio", 150000, None, "21"],
                 [2, "P-001", "Producto de ejemplo", 2, "unidades", 12100, 10, "21"]):
        wi.append(fila)
    for ws in (wc, wi):
        for row in ws.iter_rows(min_row=2, max_row=3):
            for cell in row:
                cell.font = Font(italic=True, color="808080")
    ws_ayuda = wb.create_sheet("Ayuda", 0)
    texto = [
        "FACTURADOR ARCA — ESTUDIO MARSAN",
        "",
        "1. Hoja Comprobantes: una fila por factura / nota de crédito / nota de débito.",
        "2. Hoja Items: una fila por cada producto o servicio. Se une con el comprobante por el ID.",
        "3. Hoja Tributos: percepciones u otros tributos (opcional).",
        "4. Las filas grises son ejemplos: borralas o reemplazalas.",
        "5. Pasá el mouse sobre cada título para ver qué va en la columna.",
        "",
        "Precios: en Factura A/B/M el precio va SIN IVA, salvo que pongas 'S' en 'Precios con IVA incluido'.",
        "En Factura C el precio es el final (no se discrimina IVA).",
        "Notas de crédito/débito: completar las columnas 'Cbte asociado'.",
        "Servicios: si dejás vacías las fechas de servicio, se toma el mes de la fecha del comprobante.",
        "",
        "DÓLARES / EUROS: elegí la moneda. Si dejás vacía la cotización, se toma la oficial de ARCA al emitir.",
        "   Los precios de los ítems van en esa moneda. 'Cancela en moneda extranjera' = S si te pagan en dólares.",
        "",
        "FACTURA DE CRÉDITO ELECTRÓNICA MiPyME: tipos FCEA / FCEB / FCEC (notas: NCFCEA, NDFCEA, etc.).",
        "   Requiere CUIT del receptor, CBU del emisor (de la fila o del cliente) y SCA/ADC.",
        "   Si no ponés Vto. pago se toma a 30 días. Las notas FCE llevan comprobante asociado con fecha y 'FCE anulación'.",
        "",
        "MAIL: poné el email del receptor y 'S' en 'Enviar por mail'. Sale desde info@estudiomarsan.com.ar como ESTUDIO MARSAN.",
    ]
    for r, t in enumerate(texto, start=1):
        ws_ayuda[f"A{r}"] = t
    ws_ayuda["A1"].font = Font(bold=True, size=14)
    ws_ayuda.column_dimensions["A"].width = 110
    wb.active = 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------ lectura
def _leer_hoja(wb, nombre: str, cols) -> list[dict]:
    if nombre not in wb.sheetnames:
        return []
    ws = wb[nombre]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    por_titulo = {_norm(t): k for k, t, _ in cols}
    por_titulo.update({_norm(k): k for k, _t, _ in cols})
    claves = [por_titulo.get(_norm(h)) for h in filas[0]]
    out = []
    for n, fila in enumerate(filas[1:], start=2):
        if fila is None or all(v is None or str(v).strip() == "" for v in fila):
            continue
        reg = {"_fila": n}
        for k, v in zip(claves, fila):
            if k:
                reg[k] = v.strip() if isinstance(v, str) else v
        out.append(reg)
    return out


def _fecha(v, campo, errores) -> dt.date | None:
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%Y%m%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    errores.append(f"{campo}: fecha inválida '{v}' (usar DD/MM/AAAA)")
    return None


def _num(v, campo, errores, defecto=0.0) -> float:
    if v in (None, ""):
        return defecto
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if "," in s:  # formato argentino 1.234,56
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        errores.append(f"{campo}: número inválido '{v}'")
        return defecto


def _solo_digitos(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v))


def cuit_valido(cuit: str) -> bool:
    if not re.fullmatch(r"\d{11}", cuit or ""):
        return False
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    s = sum(int(a) * b for a, b in zip(cuit[:10], pesos))
    dv = 11 - s % 11
    dv = 0 if dv == 11 else 9 if dv == 10 else dv
    return dv == int(cuit[10])


def _buscar(dic: dict, valor, alias: dict | None = None):
    if valor in (None, ""):
        return None
    n = _norm(valor)
    if alias and n in alias:
        return alias[n]
    for k in dic:
        if _norm(k) == n:
            return k
    for k in dic:  # coincidencia parcial
        if n and n in _norm(k):
            return k
    return None


def leer_excel(contenido: bytes, clientes: dict[str, dict], hoy: dt.date | None = None) -> list[dict]:
    """Devuelve una lista de comprobantes ya validados y calculados.
    Cada uno trae 'errores' (bloquean la emisión) y 'avisos' (no bloquean)."""
    hoy = hoy or dt.date.today()
    wb = load_workbook(io.BytesIO(contenido), data_only=True)
    cbtes = _leer_hoja(wb, "Comprobantes", COLS_CBTE)
    items = _leer_hoja(wb, "Items", COLS_ITEM)
    tribs = _leer_hoja(wb, "Tributos", COLS_TRIB)
    if not cbtes:
        raise ValueError("No encontré la hoja 'Comprobantes' o está vacía. Usá la plantilla de la app.")

    def clave_id(v):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).strip()

    items_por_id, tribs_por_id = {}, {}
    for it in items:
        items_por_id.setdefault(clave_id(it.get("id")), []).append(it)
    for tr in tribs:
        tribs_por_id.setdefault(clave_id(tr.get("id")), []).append(tr)

    vistos = set()
    resultado = []
    for r in cbtes:
        errores, avisos = [], []
        cid = clave_id(r.get("id"))
        if not cid or cid == "None":
            errores.append("Falta el ID")
        elif cid in vistos:
            errores.append(f"ID {cid} repetido")
        vistos.add(cid)

        cuit = _solo_digitos(r.get("cuit_emisor"))
        cli = clientes.get(cuit)
        if not cli:
            errores.append(f"El CUIT emisor {cuit or '(vacío)'} no está dado de alta como cliente")
            cli = {}

        tipo = str(r.get("tipo") or "").strip().upper()
        if tipo not in T.TIPOS_CBTE:
            errores.append(f"Tipo de comprobante '{r.get('tipo')}' no válido")
            tipo_cod, letra, nombre_tipo = 0, "?", ""
        else:
            tipo_cod, letra, nombre_tipo = T.TIPOS_CBTE[tipo]
        if cli and letra != "?" and letra not in T.LETRAS_POR_EMISOR.get(cli.get("cond_iva"), set()):
            errores.append(f"{cli.get('razon_social')} es {cli.get('cond_iva')}: no puede emitir comprobantes {letra}")

        pto_vta = int(_num(r.get("pto_vta"), "Punto de venta", errores, cli.get("pto_vta") or 0))
        if pto_vta <= 0:
            errores.append("Falta el punto de venta (ni en la fila ni en el cliente)")

        fecha = _fecha(r.get("fecha"), "Fecha", errores) or hoy
        concepto_txt = _buscar(T.CONCEPTOS, r.get("concepto")) or cli.get("concepto") or "Servicios"
        if r.get("concepto") and not _buscar(T.CONCEPTOS, r.get("concepto")):
            errores.append(f"Concepto '{r.get('concepto')}' no válido")
        concepto = T.CONCEPTOS.get(concepto_txt, 2)

        lim = 5 if concepto == 1 else 10
        if abs((fecha - hoy).days) > lim:
            errores.append(f"La fecha {fecha:%d/%m/%Y} está fuera del rango que acepta ARCA (±{lim} días de hoy)")

        es_fce_fact, es_fce_nota = tipo in T.FCE_FACTURAS, tipo in T.FCE_NOTAS
        serv_desde = serv_hasta = vto_pago = None
        if es_fce_fact and concepto == 1:
            vto_pago = _fecha(r.get("vto_pago"), "Vto. pago", errores)
            if not vto_pago:
                vto_pago = fecha + dt.timedelta(days=30)
                avisos.append("FCE: vencimiento de pago a 30 días (no estaba completo)")
            if vto_pago < fecha:
                errores.append("El vencimiento de pago no puede ser anterior a la fecha del comprobante")
        if concepto in (2, 3):
            serv_desde = _fecha(r.get("serv_desde"), "Servicio desde", errores)
            serv_hasta = _fecha(r.get("serv_hasta"), "Servicio hasta", errores)
            vto_pago = _fecha(r.get("vto_pago"), "Vto. pago", errores)
            if not serv_desde:
                serv_desde = fecha.replace(day=1)
                avisos.append("Servicio desde: se usó el 1° del mes")
            if not serv_hasta:
                serv_hasta = fecha.replace(day=calendar.monthrange(fecha.year, fecha.month)[1])
                avisos.append("Servicio hasta: se usó el último día del mes")
            if not vto_pago:
                vto_pago = fecha + dt.timedelta(days=30) if es_fce_fact else fecha
                if es_fce_fact:
                    avisos.append("FCE: vencimiento de pago a 30 días (no estaba completo)")
            if es_fce_nota:
                vto_pago = None  # ARCA no lo admite en notas de crédito/débito FCE
            if serv_hasta < serv_desde:
                errores.append("Servicio hasta es anterior a Servicio desde")
            if vto_pago and vto_pago < fecha:
                errores.append("El vencimiento de pago no puede ser anterior a la fecha del comprobante")

        cond_iva = _buscar(T.COND_IVA_RECEPTOR, r.get("rec_cond_iva"), ALIAS_COND_IVA)
        if not cond_iva:
            errores.append(f"Condición IVA del receptor '{r.get('rec_cond_iva') or ''}' no válida (es obligatoria)")
        doc_tipo_txt = _buscar(T.TIPOS_DOC, r.get("doc_tipo")) or ("Sin identificar" if not r.get("doc_nro") else None)
        if not doc_tipo_txt:
            errores.append(f"Tipo de documento '{r.get('doc_tipo')}' no válido")
            doc_tipo_txt = "CUIT"
        doc_nro = _solo_digitos(r.get("doc_nro")) or "0"
        if doc_tipo_txt in ("CUIT", "CUIL") and not cuit_valido(doc_nro):
            errores.append(f"El {doc_tipo_txt} del receptor {doc_nro} no es válido (dígito verificador)")
        if doc_tipo_txt == "Sin identificar":
            doc_nro = "0"
            if cond_iva and cond_iva != "Consumidor Final":
                errores.append("Solo se puede facturar sin identificar a Consumidor Final")
        if letra == "A":
            if doc_tipo_txt != "CUIT":
                errores.append("La Factura A requiere CUIT del receptor")
            if cond_iva not in ("IVA Responsable Inscripto", "Responsable Monotributo",
                                "Monotributista Social", "Monotributo Trabajador Independiente Promovido", None):
                errores.append(f"No corresponde comprobante A a un receptor {cond_iva}")
        if letra == "B" and cond_iva == "IVA Responsable Inscripto":
            errores.append("A un Responsable Inscripto corresponde comprobante A, no B")
        if not r.get("rec_nombre") and doc_tipo_txt != "Sin identificar":
            avisos.append("Falta el nombre del receptor (sale vacío en el PDF)")

        moneda_txt = _buscar(T.MONEDAS, r.get("moneda")) or "Pesos"
        mon_id = T.MONEDAS[moneda_txt]
        cotiz = 1.0
        cancela_ext = False
        if mon_id != "PES":
            cotiz = _num(r.get("cotizacion"), "Cotización", errores, 0.0)
            if cotiz < 0:
                errores.append("La cotización no puede ser negativa")
            if not cotiz:
                cotiz = None
                avisos.append(f"{moneda_txt}: la cotización se toma de ARCA al momento de emitir")
            cancela_ext = str(r.get("cancela_mon_ext") or "N").strip().upper().startswith("S")

        con_iva = str(r.get("precios_con_iva") or "N").strip().upper().startswith("S")

        asociados = []
        if tipo.startswith(("NC", "ND")):
            at = str(r.get("asoc_tipo") or "").strip().upper()
            if at not in T.TIPOS_CBTE or not r.get("asoc_nro"):
                errores.append("Nota de crédito/débito: completá tipo y número del comprobante asociado")
            else:
                asociados.append({
                    "Tipo": T.TIPOS_CBTE[at][0],
                    "PtoVta": int(_num(r.get("asoc_pto_vta"), "Asoc pto vta", errores, pto_vta)),
                    "Nro": int(_num(r.get("asoc_nro"), "Asoc número", errores)),
                    "Cuit": cuit,
                    "CbteFch": (_fecha(r.get("asoc_fecha"), "Asoc fecha", errores) or None),
                    "_txt": at,
                })
                if es_fce_nota and at not in T.FCE_FACTURAS | T.FCE_NOTAS:
                    errores.append("Las notas FCE tienen que asociarse a un comprobante FCE (FCEA, FCEB, FCEC)")
                if es_fce_nota and not asociados[0]["CbteFch"]:
                    errores.append("Notas FCE: falta la fecha del comprobante asociado")

        # ------------------------------------------- FCE MiPyME
        opcionales, fce = [], {}
        if es_fce_fact or es_fce_nota:
            if doc_tipo_txt != "CUIT":
                errores.append("La FCE requiere CUIT del receptor")
        if es_fce_fact:
            cbu = _solo_digitos(r.get("fce_cbu")) or _solo_digitos(cli.get("cbu"))
            alias_cbu = str(r.get("fce_alias") or cli.get("alias_cbu") or "").strip()
            transf = str(r.get("fce_transf") or cli.get("fce_transferencia") or "SCA").strip().upper()
            if len(cbu) != 22:
                errores.append("FCE: falta el CBU del emisor (22 dígitos), en la fila o en el cliente")
            if transf not in T.FCE_TRANSFERENCIA:
                errores.append("FCE transferencia: tiene que ser SCA o ADC")
            opcionales.append({"Id": T.OPC_CBU, "Valor": cbu})
            if alias_cbu:
                opcionales.append({"Id": T.OPC_ALIAS, "Valor": alias_cbu})
            opcionales.append({"Id": T.OPC_TRANSFERENCIA, "Valor": transf})
            fce = {"cbu": cbu, "alias": alias_cbu, "transferencia": transf}
        if es_fce_nota:
            anul = "S" if str(r.get("fce_anulacion") or "N").strip().upper().startswith("S") else "N"
            opcionales.append({"Id": T.OPC_ANULACION, "Valor": anul})
            fce = {"anulacion": anul}

        # ------------------------------------------- ítems y cálculo
        lineas = []
        for it in items_por_id.get(cid, []):
            e_it = []
            cant = _num(it.get("cantidad"), "Cantidad", e_it, 1.0)
            precio = _num(it.get("precio"), "Precio", e_it)
            bonif = _num(it.get("bonif"), "Bonificación", e_it)
            try:
                alic = T.normalizar_alicuota(it.get("alicuota")) if letra != "C" else "C"
            except ValueError as ex:
                e_it.append(str(ex)); alic = "21"
            if not it.get("descripcion"):
                e_it.append("ítem sin descripción")
            if cant <= 0:
                e_it.append("cantidad debe ser mayor a 0")
            if precio < 0:
                e_it.append("precio negativo")
            errores += [f"Ítem fila {it['_fila']}: {x}" for x in e_it]
            bruto = round(cant * precio * (1 - bonif / 100), 2)
            tasa = T.ALICUOTAS[alic][1] if alic in T.ALICUOTAS else 0.0
            lineas.append({
                "codigo": "" if it.get("codigo") is None else str(it.get("codigo")),
                "descripcion": str(it.get("descripcion") or ""),
                "cantidad": cant, "unidad": str(it.get("unidad") or "unidades"),
                "precio": precio, "bonif": bonif, "alicuota": alic, "tasa": tasa, "importe": bruto,
            })
        if not lineas:
            errores.append(f"El comprobante ID {cid} no tiene ítems en la hoja Items")

        imp_neto = imp_iva = imp_opex = imp_noGrav = 0.0
        iva_arr = []
        if letra == "C":
            imp_neto = round(sum(l["importe"] for l in lineas), 2)
        else:
            grupos: dict[str, float] = {}
            for l in lineas:
                grupos[l["alicuota"]] = grupos.get(l["alicuota"], 0.0) + l["importe"]
            for alic, monto in grupos.items():
                monto = round(monto, 2)
                id_iva, tasa = T.ALICUOTAS[alic]
                if alic == "Exento":
                    imp_opex += monto
                elif alic == "No gravado":
                    imp_noGrav += monto
                else:
                    if con_iva:
                        base = round(monto / (1 + tasa), 2)
                        iva = round(monto - base, 2)
                    else:
                        base, iva = monto, round(monto * tasa, 2)
                    imp_neto += base
                    imp_iva += iva
                    iva_arr.append({"Id": id_iva, "BaseImp": base, "Importe": iva, "_alic": alic})
            # precio neto por línea (para mostrar en A)
            for l in lineas:
                if con_iva and l["tasa"]:
                    l["neto"] = round(l["importe"] / (1 + l["tasa"]), 2)
                else:
                    l["neto"] = l["importe"]

        trib_arr = []
        for tr in tribs_por_id.get(cid, []):
            e_tr = []
            tipo_tr = _buscar(T.TRIBUTOS, tr.get("tipo"))
            if not tipo_tr:
                e_tr.append(f"tipo de tributo '{tr.get('tipo')}' no válido")
                tipo_tr = "Otros"
            base = _num(tr.get("base"), "Base", e_tr)
            alic = _num(tr.get("alicuota"), "Alícuota", e_tr)
            imp = _num(tr.get("importe"), "Importe", e_tr, round(base * alic / 100, 2))
            errores += [f"Tributo fila {tr['_fila']}: {x}" for x in e_tr]
            trib_arr.append({"Id": T.TRIBUTOS[tipo_tr], "Desc": str(tr.get("descripcion") or tipo_tr),
                             "BaseImp": base, "Alic": alic, "Importe": round(imp, 2)})
        imp_trib = round(sum(t["Importe"] for t in trib_arr), 2)

        imp_neto, imp_iva = round(imp_neto, 2), round(imp_iva, 2)
        imp_opex, imp_noGrav = round(imp_opex, 2), round(imp_noGrav, 2)
        total = round(imp_neto + imp_iva + imp_opex + imp_noGrav + imp_trib, 2)
        if lineas and total <= 0:
            errores.append("El total del comprobante es 0")

        leyendas = [x for x in (cli.get("leyenda_fija"), r.get("leyenda")) if x]
        if (letra == "A" and not es_fce_fact and not tipo.startswith(("NC", "ND")) and mon_id == "PES"
                and total >= T.MONTO_MINIMO_FCE):
            avisos.append(f"Total igual o mayor a $ {T.MONTO_MINIMO_FCE:,.2f}: si el receptor es gran empresa "
                          "corresponde Factura de Crédito Electrónica (FCEA)")
        mails = [m.strip() for m in re.split(r"[;,\s]+", str(r.get("email") or "")) if m.strip()]
        enviar = str(r.get("enviar_mail") or "N").strip().upper().startswith("S")
        if enviar and not mails:
            errores.append("Marcaste 'Enviar por mail' pero falta el email del receptor")
        for m in mails:
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", m):
                errores.append(f"Email inválido: {m}")

        resultado.append({
            "id": cid, "fila": r["_fila"], "cuit": cuit, "cliente": cli,
            "tipo": tipo, "tipo_cod": tipo_cod, "letra": letra, "nombre_tipo": nombre_tipo,
            "pto_vta": pto_vta, "fecha": fecha, "concepto": concepto, "concepto_txt": concepto_txt,
            "serv_desde": serv_desde, "serv_hasta": serv_hasta, "vto_pago": vto_pago,
            "doc_tipo": T.TIPOS_DOC[doc_tipo_txt], "doc_tipo_txt": doc_tipo_txt, "doc_nro": doc_nro,
            "rec_nombre": r.get("rec_nombre") or ("Consumidor Final" if doc_tipo_txt == "Sin identificar" else ""),
            "rec_domicilio": r.get("rec_domicilio") or "",
            "rec_cond_iva": cond_iva or "", "rec_cond_iva_id": T.COND_IVA_RECEPTOR.get(cond_iva or "", 0),
            "cond_venta": r.get("cond_venta") or "Contado", "moneda": mon_id, "moneda_txt": moneda_txt,
            "cotizacion": cotiz, "precios_con_iva": con_iva, "asociados": asociados,
            "leyendas": leyendas, "emails": mails, "enviar_mail": enviar,
            "cancela_mon_ext": cancela_ext, "opcionales": opcionales, "fce": fce,
            "es_fce": es_fce_fact or es_fce_nota,
            "lineas": lineas, "iva": iva_arr, "tributos": trib_arr,
            "imp_neto": imp_neto, "imp_iva": imp_iva, "imp_opex": imp_opex,
            "imp_no_gravado": imp_noGrav, "imp_trib": imp_trib, "total": total,
            "errores": errores, "avisos": avisos,
        })
    return resultado


def armar_para_arca(c: dict, nro: int) -> dict:
    f = lambda d: d.strftime("%Y%m%d") if d else None
    return {
        "PtoVta": c["pto_vta"], "CbteTipo": c["tipo_cod"], "CbteNro": nro,
        "Concepto": c["concepto"], "DocTipo": c["doc_tipo"], "DocNro": c["doc_nro"],
        "CbteFch": f(c["fecha"]), "ImpTotal": c["total"], "ImpTotConc": c["imp_no_gravado"],
        "ImpNeto": c["imp_neto"], "ImpOpEx": c["imp_opex"], "ImpTrib": c["imp_trib"], "ImpIVA": c["imp_iva"],
        "FchServDesde": f(c["serv_desde"]), "FchServHasta": f(c["serv_hasta"]), "FchVtoPago": f(c["vto_pago"]),
        "MonId": c["moneda"], "MonCotiz": c["cotizacion"], "CanMisMonExt": "S" if c["cancela_mon_ext"] else "N",
        "Opcionales": c["opcionales"], "CondicionIVAReceptorId": c["rec_cond_iva_id"],
        "CbtesAsoc": [{**a, "CbteFch": f(a.get("CbteFch"))} for a in c["asociados"]],
        "Tributos": c["tributos"], "Iva": [{k: v for k, v in i.items() if not k.startswith("_")} for i in c["iva"]],
    }
