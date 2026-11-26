"""Guardado de clientes (datos fiscales + certificado + clave), historial, PDF y tickets de ARCA.

Almacen      -> en la PC: carpeta datos/ (clientes/<CUIT>/, comprobantes/, historial.csv)
AlmacenNube  -> en tu Google Drive, a través del Apps Script (sirve igual en la PC y en la web)
"""
from __future__ import annotations

import base64
import csv
import json
from pathlib import Path

import requests

CAMPOS = ["cuit", "razon_social", "nombre_fantasia", "domicilio", "cond_iva", "iibb", "inicio_act",
          "pto_vta", "concepto", "leyenda_fija", "email", "entorno", "cbu", "alias_cbu", "fce_transferencia"]

COLS_HIST = ["fecha_emision", "entorno", "cuit_emisor", "razon_social", "tipo", "pto_vta", "numero", "fecha",
             "receptor_doc", "receptor", "neto", "iva", "tributos", "total", "moneda", "cae", "cae_vto",
             "id_excel", "archivo_excel", "pdf", "email", "mail_enviado"]

# Tablas extra (registro de mails y cobranzas)
TABLAS = {
    "Mails": ["id", "fecha", "cuit_emisor", "razon_social", "comprobante", "pdf", "para", "cc", "asunto",
              "estado", "detalle"],
    "Cobros": ["id", "fecha", "cuit_emisor", "receptor_doc", "receptor", "importe", "moneda", "medio",
               "referencia", "aplica_a", "observaciones", "cargado"],
}


def num(v) -> float:
    """Convierte números que pueden venir como '1234.5', '1.234,50' o 1234.5."""
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("$", "").replace(" ", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


class Almacen:
    def __init__(self, base: Path, secretos: dict | None = None):
        self.base = Path(base)
        self.dir_clientes = self.base / "clientes"
        self.dir_pdf = self.base / "comprobantes"
        self.dir_cache = self.base / "cache_tokens"
        for d in (self.dir_clientes, self.dir_pdf, self.dir_cache):
            d.mkdir(parents=True, exist_ok=True)
        self.secretos = secretos or {}

    # ---------------------------------------------------------- clientes
    def listar(self) -> dict[str, dict]:
        out: dict[str, dict] = {}
        for cuit, datos in (self.secretos or {}).items():
            d = dict(datos)
            d["cuit"] = str(d.get("cuit") or cuit)
            d["_origen"] = "secrets"
            out[d["cuit"]] = d
        for carpeta in sorted(self.dir_clientes.iterdir()):
            cfg = carpeta / "config.json"
            if cfg.exists():
                d = json.loads(cfg.read_text(encoding="utf-8"))
                d["_origen"] = "local"
                out[d["cuit"]] = d
        for d in out.values():
            d["pto_vta"] = int(d.get("pto_vta") or 0)
            d.setdefault("entorno", "homologacion")
        return out

    def guardar(self, datos: dict, cert: bytes | None = None, clave: bytes | None = None):
        carpeta = self.dir_clientes / datos["cuit"]
        carpeta.mkdir(parents=True, exist_ok=True)
        limpio = {k: datos.get(k, "") for k in CAMPOS}
        (carpeta / "config.json").write_text(json.dumps(limpio, ensure_ascii=False, indent=2), encoding="utf-8")
        if cert:
            (carpeta / "certificado.crt").write_bytes(cert)
        if clave:
            (carpeta / "clave.key").write_bytes(clave)

    def eliminar(self, cuit: str):
        carpeta = self.dir_clientes / cuit
        if carpeta.exists():
            for f in carpeta.iterdir():
                f.unlink()
            carpeta.rmdir()

    def credenciales(self, cuit: str) -> tuple[bytes | None, bytes | None]:
        s = (self.secretos or {}).get(cuit)
        if s and s.get("certificado") and s.get("clave"):
            return s["certificado"].encode(), s["clave"].encode()
        carpeta = self.dir_clientes / cuit
        cert = (carpeta / "certificado.crt").read_bytes() if (carpeta / "certificado.crt").exists() else None
        clave = (carpeta / "clave.key").read_bytes() if (carpeta / "clave.key").exists() else None
        return cert, clave

    def guardar_clave_pendiente(self, cuit: str, clave: bytes):
        """Clave privada generada con el CSR, a la espera de que ARCA devuelva el certificado."""
        carpeta = self.dir_clientes / cuit
        carpeta.mkdir(parents=True, exist_ok=True)
        (carpeta / "clave.key").write_bytes(clave)
        if (carpeta / "certificado.crt").exists():
            (carpeta / "certificado.crt").unlink()  # el certificado viejo ya no corresponde a la clave nueva

    # ---------------------------------------------------------- PDF y tickets de ARCA
    def guardar_pdf(self, cuit: str, nombre: str, pdf: bytes):
        carpeta = self.dir_pdf / cuit
        carpeta.mkdir(parents=True, exist_ok=True)
        (carpeta / nombre).write_bytes(pdf)

    def leer_pdf(self, cuit: str, nombre: str) -> bytes | None:
        ruta = self.dir_pdf / cuit / nombre
        return ruta.read_bytes() if ruta.exists() else None

    def ta_leer(self, clave: str) -> dict | None:
        ruta = self.dir_cache / f"{clave}.json"
        return json.loads(ruta.read_text()) if ruta.exists() else None

    def ta_guardar(self, clave: str, ta: dict):
        (self.dir_cache / f"{clave}.json").write_text(json.dumps(ta))

    # ---------------------------------------------------------- historial
    @property
    def ruta_hist(self) -> Path:
        return self.base / "historial.csv"

    def registrar(self, fila: dict):
        nuevo = not self.ruta_hist.exists()
        with self.ruta_hist.open("a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=COLS_HIST, delimiter=";")
            if nuevo:
                w.writeheader()
            w.writerow({k: fila.get(k, "") for k in COLS_HIST})

    def historial(self) -> list[dict]:
        if not self.ruta_hist.exists():
            return []
        with self.ruta_hist.open(encoding="utf-8-sig") as f:
            return list(csv.DictReader(f, delimiter=";"))

    def marcar_mail(self, pdf: str, cuando: str):
        filas = self.historial()
        for f in filas:
            if f["pdf"] == pdf:
                f["mail_enviado"] = cuando
        with self.ruta_hist.open("w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=COLS_HIST, delimiter=";")
            w.writeheader()
            for f in filas:
                w.writerow({k: f.get(k, "") for k in COLS_HIST})

    # ---------------------------------------------------------- tablas (mails, cobros)
    def _ruta_tabla(self, tabla: str) -> Path:
        return self.base / f"{tabla.lower()}.csv"

    def tabla_listar(self, tabla: str) -> list[dict]:
        ruta = self._ruta_tabla(tabla)
        if not ruta.exists():
            return []
        with ruta.open(encoding="utf-8-sig") as f:
            return list(csv.DictReader(f, delimiter=";"))

    def _tabla_escribir(self, tabla: str, filas: list[dict]):
        with self._ruta_tabla(tabla).open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=TABLAS[tabla], delimiter=";")
            w.writeheader()
            for fila in filas:
                w.writerow({k: fila.get(k, "") for k in TABLAS[tabla]})

    def tabla_agregar(self, tabla: str, fila: dict):
        self._tabla_escribir(tabla, self.tabla_listar(tabla) + [fila])

    def tabla_eliminar(self, tabla: str, id_: str):
        self._tabla_escribir(tabla, [f for f in self.tabla_listar(tabla) if str(f.get("id")) != str(id_)])

    # ---------------------------------------------------------- mail
    @property
    def ruta_mail(self) -> Path:
        return self.base / "config_mail.json"

    def config_mail(self, secretos_mail: dict | None = None) -> dict:
        cfg = dict(secretos_mail or {})
        if self.ruta_mail.exists():
            cfg.update(json.loads(self.ruta_mail.read_text(encoding="utf-8")))
        return cfg

    def guardar_config_mail(self, cfg: dict):
        self.ruta_mail.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    def ya_emitido(self, archivo: str, id_excel: str, cuit: str) -> dict | None:
        for h in self.historial():
            if h["archivo_excel"] == archivo and h["id_excel"] == id_excel and h["cuit_emisor"] == cuit:
                return h
        return None


# ====================================================================== Google Drive
class ErrorNube(Exception):
    pass


class AlmacenNube(Almacen):
    """Mismo funcionamiento que Almacen, pero todo queda en tu Google Drive vía Apps Script.
    Los tickets de ARCA también quedan ahí, así la PC y la web los comparten y ARCA no bloquea."""

    def __init__(self, base: Path, url: str, clave: str, secretos: dict | None = None):
        super().__init__(base, secretos)
        self.url, self.clave = url, clave
        self._clientes_cache: dict | None = None
        self._hist_cache: list | None = None

    def _llamar(self, accion: str, **datos):
        try:
            r = requests.post(self.url, json={"clave": self.clave, "accion": accion, **datos}, timeout=60)
            res = r.json()
        except Exception as ex:
            raise ErrorNube(f"No me pude comunicar con el Drive (Apps Script): {ex}")
        if not res.get("ok"):
            raise ErrorNube(f"Drive respondió: {res.get('error')}")
        return res.get("datos")

    def probar(self) -> bool:
        self._llamar("kv_get", k="prueba")
        return True

    # ---------------------------------------------------------- clientes
    def _clientes_nube(self) -> dict[str, dict]:
        if self._clientes_cache is None:
            self._clientes_cache = {str(d["cuit"]): d for d in (self._llamar("clientes_listar") or [])}
        return self._clientes_cache

    def listar(self) -> dict[str, dict]:
        out = {}
        for cuit, d in self._clientes_nube().items():
            d = {k: v for k, v in d.items() if k not in ("certificado", "clave")}
            d["_origen"] = "drive"
            d["pto_vta"] = int(d.get("pto_vta") or 0)
            d.setdefault("entorno", "homologacion")
            out[cuit] = d
        return out

    def _guardar_nube(self, cuit: str, datos: dict):
        self._llamar("cliente_guardar", cuit=cuit, datos=datos)
        self._clientes_nube()[cuit] = datos

    def guardar(self, datos: dict, cert: bytes | None = None, clave: bytes | None = None):
        cuit = datos["cuit"]
        previo = self._clientes_nube().get(cuit, {})
        nuevo = {k: datos.get(k, "") for k in CAMPOS}
        nuevo["certificado"] = cert.decode() if cert else previo.get("certificado", "")
        nuevo["clave"] = clave.decode() if clave else previo.get("clave", "")
        self._guardar_nube(cuit, nuevo)

    def eliminar(self, cuit: str):
        self._llamar("cliente_eliminar", cuit=cuit)
        self._clientes_nube().pop(cuit, None)

    def credenciales(self, cuit: str):
        d = self._clientes_nube().get(cuit, {})
        cert, clave = d.get("certificado"), d.get("clave")
        return (cert.encode() if cert else None), (clave.encode() if clave else None)

    def guardar_clave_pendiente(self, cuit: str, clave: bytes):
        d = dict(self._clientes_nube().get(cuit, {"cuit": cuit}))
        d["clave"], d["certificado"] = clave.decode(), ""
        self._guardar_nube(cuit, d)

    def subir_locales(self, local: Almacen) -> list[str]:
        """Copia al Drive los clientes que estaban guardados en la PC."""
        subidos = []
        for cuit, d in local.listar().items():
            if d.get("_origen") != "local" or cuit in self._clientes_nube():
                continue
            cert, clave = local.credenciales(cuit)
            self.guardar(d, cert, clave)
            subidos.append(d.get("razon_social") or cuit)
        return subidos

    # ---------------------------------------------------------- PDF y tickets
    def guardar_pdf(self, cuit: str, nombre: str, pdf: bytes):
        self._llamar("pdf_guardar", cuit=cuit, nombre=nombre, b64=base64.b64encode(pdf).decode())

    def leer_pdf(self, cuit: str, nombre: str) -> bytes | None:
        b64 = self._llamar("pdf_obtener", cuit=cuit, nombre=nombre)
        return base64.b64decode(b64) if b64 else None

    def ta_leer(self, clave: str) -> dict | None:
        v = self._llamar("kv_get", k="ta_" + clave)
        return json.loads(v) if v else None

    def ta_guardar(self, clave: str, ta: dict):
        self._llamar("kv_set", k="ta_" + clave, v=json.dumps(ta))

    # ---------------------------------------------------------- historial
    def registrar(self, fila: dict):
        self._llamar("hist_agregar", fila={k: fila.get(k, "") for k in COLS_HIST})
        self._hist_cache = None

    def historial(self) -> list[dict]:
        if self._hist_cache is None:
            self._hist_cache = self._llamar("hist_listar") or []
        return self._hist_cache

    def marcar_mail(self, pdf: str, cuando: str):
        self._llamar("hist_marcar_mail", pdf=pdf, cuando=cuando)
        self._hist_cache = None

    # ---------------------------------------------------------- tablas (mails, cobros)
    def _tabla_llamar(self, accion: str, **datos):
        try:
            return self._llamar(accion, **datos)
        except ErrorNube as ex:
            if "is not a function" in str(ex) or "not a function" in str(ex):
                raise ErrorNube("Hay que actualizar el Apps Script con la versión nueva (apps_script/Codigo.gs) "
                                "para usar mails enviados y cuentas corrientes.")
            raise

    def tabla_listar(self, tabla: str) -> list[dict]:
        cache = self.__dict__.setdefault("_tablas_cache", {})
        if tabla not in cache:
            cache[tabla] = self._tabla_llamar("tabla_listar", tabla=tabla, columnas=TABLAS[tabla]) or []
        return cache[tabla]

    def tabla_agregar(self, tabla: str, fila: dict):
        self._tabla_llamar("tabla_agregar", tabla=tabla, columnas=TABLAS[tabla],
                           fila={k: fila.get(k, "") for k in TABLAS[tabla]})
        self.__dict__.setdefault("_tablas_cache", {}).pop(tabla, None)

    def tabla_eliminar(self, tabla: str, id_: str):
        self._tabla_llamar("tabla_eliminar", tabla=tabla, columnas=TABLAS[tabla], id=str(id_))
        self.__dict__.setdefault("_tablas_cache", {}).pop(tabla, None)

    # ---------------------------------------------------------- mail
    def config_mail(self, secretos_mail: dict | None = None) -> dict:
        if getattr(self, "_mail_cache", None) is None:
            v = self._llamar("kv_get", k="config_mail")
            self._mail_cache = json.loads(v) if v else {}
        return {**dict(secretos_mail or {}), **self._mail_cache}

    def guardar_config_mail(self, cfg: dict):
        self._llamar("kv_set", k="config_mail", v=json.dumps(cfg, ensure_ascii=False))
        self._mail_cache = dict(cfg)
