"""Conexión con los web services de ARCA: WSAA (autenticación) y WSFEv1 (factura electrónica).

Cada cliente usa SU certificado y SU clave privada. El token (TA) se guarda en caché
por CUIT y entorno hasta que vence (ARCA rechaza pedir otro mientras el anterior sigue vigente).
"""
from __future__ import annotations

import base64
import datetime as dt
import json
import ssl
import time
from pathlib import Path
from xml.sax.saxutils import escape

import requests
from requests.adapters import HTTPAdapter
from lxml import etree
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives.serialization import pkcs7
from cryptography.x509.oid import NameOID

TZ_AR = dt.timezone(dt.timedelta(hours=-3))

URLS = {
    "homologacion": {
        "wsaa": "https://wsaahomo.afip.gov.ar/ws/services/LoginCms",
        "wsfe": "https://wswhomo.afip.gov.ar/wsfev1/service.asmx",
    },
    "produccion": {
        "wsaa": "https://wsaa.afip.gov.ar/ws/services/LoginCms",
        "wsfe": "https://servicios1.afip.gov.ar/wsfev1/service.asmx",
    },
}
NS_FE = "http://ar.gov.afip.dif.FEV1/"


class ErrorARCA(Exception):
    pass


# ---------------------------------------------------------------- HTTP
class _AdaptadorARCA(HTTPAdapter):
    """Los servidores de ARCA usan claves DH cortas: hay que bajar el nivel de seguridad de OpenSSL."""

    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)


def _sesion() -> requests.Session:
    s = requests.Session()
    s.mount("https://", _AdaptadorARCA())
    return s


def _post_soap(url: str, cuerpo: str, accion: str = "") -> etree._Element:
    headers = {"Content-Type": "text/xml; charset=utf-8", "SOAPAction": accion}
    r = _sesion().post(url, data=cuerpo.encode("utf-8"), headers=headers, timeout=60)
    try:
        raiz = etree.fromstring(r.content)
    except etree.XMLSyntaxError:
        raise ErrorARCA(f"Respuesta inválida de ARCA (HTTP {r.status_code}): {r.text[:300]}")
    falla = raiz.find(".//{http://schemas.xmlsoap.org/soap/envelope/}Fault")
    if falla is not None:
        msg = falla.findtext("faultstring") or etree.tostring(falla, encoding="unicode")
        raise ErrorARCA(msg)
    return raiz


# ---------------------------------------------------------------- certificados
def generar_clave_y_csr(cuit: str, razon_social: str, alias: str) -> tuple[bytes, bytes]:
    """Genera la clave privada y el pedido de certificado (CSR) para subir a ARCA."""
    clave = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    sujeto = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "AR"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, razon_social[:64]),
        x509.NameAttribute(NameOID.COMMON_NAME, alias[:64]),
        x509.NameAttribute(NameOID.SERIAL_NUMBER, f"CUIT {cuit}"),
    ])
    csr = x509.CertificateSigningRequestBuilder().subject_name(sujeto).sign(clave, hashes.SHA256())
    clave_pem = clave.private_bytes(
        serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
        serialization.NoEncryption())
    return clave_pem, csr.public_bytes(serialization.Encoding.PEM)


def info_certificado(cert_pem: bytes) -> dict:
    cert = x509.load_pem_x509_certificate(cert_pem)
    return {
        "sujeto": cert.subject.rfc4514_string(),
        "emisor": cert.issuer.rfc4514_string(),
        "vence": cert.not_valid_after_utc.astimezone(TZ_AR).strftime("%d/%m/%Y"),
        "vencido": cert.not_valid_after_utc < dt.datetime.now(dt.timezone.utc),
    }


def clave_coincide(cert_pem: bytes, clave_pem: bytes) -> bool:
    cert = x509.load_pem_x509_certificate(cert_pem)
    clave = serialization.load_pem_private_key(clave_pem, password=None)
    return cert.public_key().public_numbers() == clave.public_key().public_numbers()


# ---------------------------------------------------------------- WSAA
class WSAA:
    def __init__(self, cuit: str, cert_pem: bytes, clave_pem: bytes, entorno: str, almacen=None):
        """almacen: objeto con ta_leer(clave) / ta_guardar(clave, ta) — PC (archivo) o Google Drive."""
        self.cuit, self.cert_pem, self.clave_pem, self.entorno = str(cuit), cert_pem, clave_pem, entorno
        self.almacen = almacen
        self._memoria: dict | None = None
        # el ticket se asocia al certificado: si se cambia el certificado, se pide uno nuevo
        huella = x509.load_pem_x509_certificate(cert_pem).serial_number
        self._clave_cache = f"{self.cuit}_{entorno}_wsfe_{huella:x}"[:80]

    def _leer_cache(self) -> dict | None:
        ta = self._memoria
        if ta is None and self.almacen:
            ta = self.almacen.ta_leer(self._clave_cache)
        if ta and dt.datetime.fromisoformat(ta["vence"]) > dt.datetime.now(TZ_AR) + dt.timedelta(minutes=5):
            return ta
        return None

    def _tra(self) -> bytes:
        ahora = dt.datetime.now(TZ_AR)
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<loginTicketRequest version="1.0"><header>'
            f"<uniqueId>{int(time.time())}</uniqueId>"
            f"<generationTime>{(ahora - dt.timedelta(minutes=10)).isoformat(timespec='seconds')}</generationTime>"
            f"<expirationTime>{(ahora + dt.timedelta(hours=10)).isoformat(timespec='seconds')}</expirationTime>"
            "</header><service>wsfe</service></loginTicketRequest>"
        ).encode()

    def _firmar(self, datos: bytes) -> str:
        cert = x509.load_pem_x509_certificate(self.cert_pem)
        clave = serialization.load_pem_private_key(self.clave_pem, password=None)
        cms = (pkcs7.PKCS7SignatureBuilder().set_data(datos)
               .add_signer(cert, clave, hashes.SHA256()).sign(serialization.Encoding.DER, []))
        return base64.b64encode(cms).decode()

    def token(self) -> tuple[str, str]:
        ta = self._leer_cache()
        if ta:
            return ta["token"], ta["sign"]
        cuerpo = (
            '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
            'xmlns:wsaa="http://wsaa.view.sua.dvadac.desein.afip.gov"><soapenv:Header/><soapenv:Body>'
            f"<wsaa:loginCms><wsaa:in0>{self._firmar(self._tra())}</wsaa:in0></wsaa:loginCms>"
            "</soapenv:Body></soapenv:Envelope>"
        )
        try:
            raiz = _post_soap(URLS[self.entorno]["wsaa"], cuerpo)
        except ErrorARCA as e:
            if "TA valido" in str(e):
                raise ErrorARCA("ARCA dice que ya hay un ticket vigente para este certificado y no está guardado "
                                "(¿se usó desde otra PC sin conexión al Drive?). Se libera solo en unas horas "
                                "(máx. 12). Si es urgente, generá otro certificado para este cliente.")
            raise
        ret = raiz.find(".//{*}loginCmsReturn")
        ticket = etree.fromstring(ret.text.encode())
        ta = {
            "token": ticket.findtext(".//token"),
            "sign": ticket.findtext(".//sign"),
            "vence": ticket.findtext(".//expirationTime"),
        }
        self._memoria = ta
        if self.almacen:
            self.almacen.ta_guardar(self._clave_cache, ta)
        return ta["token"], ta["sign"]


# ---------------------------------------------------------------- WSFEv1
def _fmt(n: float) -> str:
    return f"{round(n + 0.0, 2):.2f}"


class WSFE:
    def __init__(self, wsaa: WSAA):
        self.wsaa = wsaa
        self.url = URLS[wsaa.entorno]["wsfe"]

    def _auth(self) -> str:
        token, sign = self.wsaa.token()
        return (f"<ar:Auth><ar:Token>{token}</ar:Token><ar:Sign>{sign}</ar:Sign>"
                f"<ar:Cuit>{self.wsaa.cuit}</ar:Cuit></ar:Auth>")

    def _llamar(self, metodo: str, interior: str) -> etree._Element:
        cuerpo = (
            f'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ar="{NS_FE}">'
            f"<soapenv:Header/><soapenv:Body><ar:{metodo}>{interior}</ar:{metodo}></soapenv:Body></soapenv:Envelope>"
        )
        raiz = _post_soap(self.url, cuerpo, NS_FE + metodo)
        return raiz.find(f".//{{{NS_FE}}}{metodo}Result")

    @staticmethod
    def _errores(res) -> list[str]:
        return [f"{e.findtext('{*}Code')}: {e.findtext('{*}Msg')}" for e in res.findall(".//{*}Errors/{*}Err")]

    def ultimo_autorizado(self, pto_vta: int, cbte_tipo: int) -> int:
        res = self._llamar("FECompUltimoAutorizado",
                           f"{self._auth()}<ar:PtoVta>{pto_vta}</ar:PtoVta><ar:CbteTipo>{cbte_tipo}</ar:CbteTipo>")
        errs = self._errores(res)
        if errs:
            raise ErrorARCA("; ".join(errs))
        return int(res.findtext("{*}CbteNro") or 0)

    def cotizacion(self, mon_id: str) -> float:
        res = self._llamar("FEParamGetCotizacion", f"{self._auth()}<ar:MonId>{mon_id}</ar:MonId>")
        errs = self._errores(res)
        if errs:
            raise ErrorARCA("; ".join(errs))
        return float(res.findtext(".//{*}MonCotiz"))

    def consultar(self, pto_vta: int, cbte_tipo: int, nro: int) -> dict | None:
        res = self._llamar("FECompConsultar",
                           f"{self._auth()}<ar:FeCompConsReq><ar:CbteTipo>{cbte_tipo}</ar:CbteTipo>"
                           f"<ar:CbteNro>{nro}</ar:CbteNro><ar:PtoVta>{pto_vta}</ar:PtoVta></ar:FeCompConsReq>")
        get = res.find("{*}ResultGet")
        if get is None:
            return None
        return {"CAE": get.findtext("{*}CodAutorizacion"), "CAEFchVto": get.findtext("{*}FchVto"),
                "ImpTotal": float(get.findtext("{*}ImpTotal") or 0), "DocNro": get.findtext("{*}DocNro")}

    def autorizar(self, c: dict) -> dict:
        """c: dict con los campos calculados del comprobante (ver calculo.armar_para_arca)."""
        d = []
        add = lambda tag, val: d.append(f"<ar:{tag}>{val}</ar:{tag}>")
        add("Concepto", c["Concepto"]); add("DocTipo", c["DocTipo"]); add("DocNro", c["DocNro"])
        add("CbteDesde", c["CbteNro"]); add("CbteHasta", c["CbteNro"]); add("CbteFch", c["CbteFch"])
        add("ImpTotal", _fmt(c["ImpTotal"])); add("ImpTotConc", _fmt(c["ImpTotConc"]))
        add("ImpNeto", _fmt(c["ImpNeto"])); add("ImpOpEx", _fmt(c["ImpOpEx"]))
        add("ImpTrib", _fmt(c["ImpTrib"])); add("ImpIVA", _fmt(c["ImpIVA"]))
        if c["Concepto"] in (2, 3):
            add("FchServDesde", c["FchServDesde"]); add("FchServHasta", c["FchServHasta"])
        if c.get("FchVtoPago"):
            add("FchVtoPago", c["FchVtoPago"])
        add("MonId", c["MonId"]); add("MonCotiz", f"{c['MonCotiz']:.6f}")
        if c["MonId"] != "PES":
            add("CanMisMonExt", c.get("CanMisMonExt") or "N")
        add("CondicionIVAReceptorId", c["CondicionIVAReceptorId"])
        if c.get("CbtesAsoc"):
            d.append("<ar:CbtesAsoc>" + "".join(
                f"<ar:CbteAsoc><ar:Tipo>{a['Tipo']}</ar:Tipo><ar:PtoVta>{a['PtoVta']}</ar:PtoVta>"
                f"<ar:Nro>{a['Nro']}</ar:Nro><ar:Cuit>{a['Cuit']}</ar:Cuit>"
                + (f"<ar:CbteFch>{a['CbteFch']}</ar:CbteFch>" if a.get("CbteFch") else "")
                + "</ar:CbteAsoc>" for a in c["CbtesAsoc"]) + "</ar:CbtesAsoc>")
        if c.get("Tributos"):
            d.append("<ar:Tributos>" + "".join(
                f"<ar:Tributo><ar:Id>{t['Id']}</ar:Id><ar:Desc>{escape(t['Desc'][:80])}</ar:Desc>"
                f"<ar:BaseImp>{_fmt(t['BaseImp'])}</ar:BaseImp><ar:Alic>{_fmt(t['Alic'])}</ar:Alic>"
                f"<ar:Importe>{_fmt(t['Importe'])}</ar:Importe></ar:Tributo>" for t in c["Tributos"]) + "</ar:Tributos>")
        if c.get("Iva"):
            d.append("<ar:Iva>" + "".join(
                f"<ar:AlicIva><ar:Id>{i['Id']}</ar:Id><ar:BaseImp>{_fmt(i['BaseImp'])}</ar:BaseImp>"
                f"<ar:Importe>{_fmt(i['Importe'])}</ar:Importe></ar:AlicIva>" for i in c["Iva"]) + "</ar:Iva>")
        if c.get("Opcionales"):
            d.append("<ar:Opcionales>" + "".join(
                f"<ar:Opcional><ar:Id>{o['Id']}</ar:Id><ar:Valor>{escape(str(o['Valor']))}</ar:Valor></ar:Opcional>"
                for o in c["Opcionales"]) + "</ar:Opcionales>")

        interior = (f"{self._auth()}<ar:FeCAEReq><ar:FeCabReq><ar:CantReg>1</ar:CantReg>"
                    f"<ar:PtoVta>{c['PtoVta']}</ar:PtoVta><ar:CbteTipo>{c['CbteTipo']}</ar:CbteTipo></ar:FeCabReq>"
                    f"<ar:FeDetReq><ar:FECAEDetRequest>{''.join(d)}</ar:FECAEDetRequest></ar:FeDetReq></ar:FeCAEReq>")
        res = self._llamar("FECAESolicitar", interior)
        det = res.find(".//{*}FECAEDetResponse")
        obs = [f"{o.findtext('{*}Code')}: {o.findtext('{*}Msg')}" for o in res.findall(".//{*}Obs")]
        out = {
            "resultado": det.findtext("{*}Resultado") if det is not None else "R",
            "CAE": (det.findtext("{*}CAE") or "").strip() if det is not None else "",
            "CAEFchVto": (det.findtext("{*}CAEFchVto") or "").strip() if det is not None else "",
            "errores": self._errores(res),
            "observaciones": obs,
        }
        return out


class WSFESimulado:
    """Imita a ARCA para probar el circuito completo sin emitir nada real."""

    def __init__(self, *_a, **_k):
        self._ultimos: dict = {}

    def ultimo_autorizado(self, pto_vta, cbte_tipo):
        return self._ultimos.get((pto_vta, cbte_tipo), 0)

    def consultar(self, *_a):
        return None

    def cotizacion(self, mon_id):
        return {"DOL": 1000.0, "060": 1100.0}.get(mon_id, 1.0)  # valor ficticio de simulación

    def autorizar(self, c):
        self._ultimos[(c["PtoVta"], c["CbteTipo"])] = c["CbteNro"]
        vto = (dt.date.today() + dt.timedelta(days=10)).strftime("%Y%m%d")
        return {"resultado": "A", "CAE": "7" + str(int(time.time() * 1000))[-13:], "CAEFchVto": vto,
                "errores": [], "observaciones": ["SIMULACIÓN: comprobante sin validez fiscal"]}
