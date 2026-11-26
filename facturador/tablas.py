"""Tablas de códigos de ARCA (WSFEv1) usadas por el facturador."""

# Tipos de comprobante: código corto que se usa en el Excel -> (código ARCA, letra, nombre)
TIPOS_CBTE = {
    "FA":  (1,  "A", "FACTURA"),
    "NDA": (2,  "A", "NOTA DE DÉBITO"),
    "NCA": (3,  "A", "NOTA DE CRÉDITO"),
    "FB":  (6,  "B", "FACTURA"),
    "NDB": (7,  "B", "NOTA DE DÉBITO"),
    "NCB": (8,  "B", "NOTA DE CRÉDITO"),
    "FC":  (11, "C", "FACTURA"),
    "NDC": (12, "C", "NOTA DE DÉBITO"),
    "NCC": (13, "C", "NOTA DE CRÉDITO"),
    "FM":  (51, "M", "FACTURA"),
    "NDM": (52, "M", "NOTA DE DÉBITO"),
    "NCM": (53, "M", "NOTA DE CRÉDITO"),
    # Factura de Crédito Electrónica MiPyME (FCE)
    "FCEA":   (201, "A", "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NDFCEA": (202, "A", "NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NCFCEA": (203, "A", "NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
    "FCEB":   (206, "B", "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NDFCEB": (207, "B", "NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NCFCEB": (208, "B", "NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
    "FCEC":   (211, "C", "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NDFCEC": (212, "C", "NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE)"),
    "NCFCEC": (213, "C", "NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"),
}
FCE_FACTURAS = {"FCEA", "FCEB", "FCEC"}
FCE_NOTAS = {"NDFCEA", "NCFCEA", "NDFCEB", "NCFCEB", "NDFCEC", "NCFCEC"}
FCE_TODOS = FCE_FACTURAS | FCE_NOTAS
FCE_TRANSFERENCIA = {"SCA": "Sistema de Circulación Abierta", "ADC": "Agente de Depósito Colectivo"}
# Monto mínimo a partir del cual es obligatoria la FCE con grandes empresas (Res. 1/2026, desde 14/04/2026).
# Actualizar cuando cambie.
MONTO_MINIMO_FCE = 5_549_862.00

# Ids de "Opcionales" de WSFEv1 para FCE
OPC_CBU, OPC_ALIAS, OPC_TRANSFERENCIA, OPC_ANULACION = "2101", "2102", "27", "22"
CODIGO_A_TIPO = {v[0]: k for k, v in TIPOS_CBTE.items()}

CONCEPTOS = {
    "Productos": 1,
    "Servicios": 2,
    "Productos y Servicios": 3,
}

TIPOS_DOC = {
    "CUIT": 80,
    "CUIL": 86,
    "CDI": 87,
    "DNI": 96,
    "Pasaporte": 94,
    "Sin identificar": 99,
}

# Condición frente al IVA del receptor (obligatoria en WSFEv1)
COND_IVA_RECEPTOR = {
    "IVA Responsable Inscripto": 1,
    "IVA Sujeto Exento": 4,
    "Consumidor Final": 5,
    "Responsable Monotributo": 6,
    "Sujeto No Categorizado": 7,
    "Proveedor del Exterior": 8,
    "Cliente del Exterior": 9,
    "IVA Liberado - Ley 19.640": 10,
    "Monotributista Social": 13,
    "IVA No Alcanzado": 15,
    "Monotributo Trabajador Independiente Promovido": 16,
}

# Alícuotas de IVA: texto del Excel -> (Id ARCA, tasa). Exento / No gravado no van al array Iva.
ALICUOTAS = {
    "21":   (5, 0.21),
    "10.5": (4, 0.105),
    "27":   (6, 0.27),
    "5":    (8, 0.05),
    "2.5":  (9, 0.025),
    "0":    (3, 0.0),
    "Exento": (None, 0.0),
    "No gravado": (None, 0.0),
}

TRIBUTOS = {
    "Percepción IIBB": 7,
    "Percepción IVA": 6,
    "Impuestos nacionales": 1,
    "Impuestos provinciales": 2,
    "Impuestos municipales": 3,
    "Impuestos internos": 4,
    "Ingresos Brutos": 5,
    "Percepción municipal": 8,
    "Otras percepciones": 9,
    "Otros": 99,
}

MONEDAS = {"Pesos": "PES", "Dólares": "DOL", "Euros": "060"}

CONDICIONES_VENTA = [
    "Contado", "Transferencia bancaria", "Cuenta corriente", "Tarjeta de débito",
    "Tarjeta de crédito", "Cheque", "Otra",
]

COND_IVA_EMISOR = ["IVA Responsable Inscripto", "Responsable Monotributo", "IVA Sujeto Exento"]

# Letras que puede emitir cada condición del emisor
LETRAS_POR_EMISOR = {
    "IVA Responsable Inscripto": {"A", "B", "M"},
    "Responsable Monotributo": {"C"},
    "IVA Sujeto Exento": {"C"},
}


def normalizar_alicuota(valor) -> str:
    """Acepta 21, '21', '21%', 0.21, '10,5', 'exento'... y devuelve la clave de ALICUOTAS."""
    if valor is None or str(valor).strip() == "":
        return "21"
    s = str(valor).strip().replace("%", "").replace(",", ".").lower()
    if s.startswith("exen"):
        return "Exento"
    if s.startswith("no"):
        return "No gravado"
    try:
        f = float(s)
    except ValueError:
        raise ValueError(f"Alícuota de IVA no reconocida: {valor!r}")
    if 0 < f < 1:  # vino como 0.21
        f = round(f * 100, 2)
    for k in ("21", "10.5", "27", "5", "2.5", "0"):
        if abs(float(k) - f) < 1e-6:
            return k
    raise ValueError(f"Alícuota de IVA no válida: {valor!r}")
