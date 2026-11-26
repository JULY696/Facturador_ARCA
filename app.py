"""Facturador ARCA — Estudio Marsan.
Subís un Excel con las facturas de tus clientes y la app pide el CAE a ARCA con el certificado de cada uno.
"""
from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from facturador import tablas as T
from facturador.arca import (WSAA, WSFE, ErrorARCA, clave_coincide, generar_clave_y_csr, info_certificado)
from facturador.clientes import Almacen, AlmacenNube, ErrorNube
from facturador.emitir import emitir_lote
from facturador import mailer
from facturador.plantilla import cuit_valido, generar_plantilla, generar_plantilla_simple, leer_excel

st.set_page_config(page_title="Facturador ARCA · Estudio Marsan", page_icon="🧾", layout="wide")

BASE = Path(os.environ.get("FACTURADOR_DATOS", Path(__file__).parent / "datos"))


def _secretos(clave, defecto=None):
    try:
        return st.secrets.get(clave, defecto)
    except Exception:
        return defecto


# ------------------------------------------------------------------ acceso
clave_app = _secretos("APP_PASSWORD")
if clave_app and not st.session_state.get("ok"):
    st.title("🧾 Facturador ARCA")
    pw = st.text_input("Clave de acceso", type="password")
    if st.button("Ingresar"):
        if pw == clave_app:
            st.session_state.ok = True
            st.rerun()
        st.error("Clave incorrecta")
    st.stop()

sec_clientes = {k: dict(v) for k, v in (_secretos("clientes", {}) or {}).items()}
NUBE_URL, NUBE_CLAVE = _secretos("NUBE_URL"), _secretos("NUBE_CLAVE")
if "alm" not in st.session_state:
    if NUBE_URL and NUBE_CLAVE:
        st.session_state.alm = AlmacenNube(BASE, NUBE_URL, NUBE_CLAVE, sec_clientes)
    else:
        st.session_state.alm = Almacen(BASE, sec_clientes)
alm = st.session_state.alm
en_nube = isinstance(alm, AlmacenNube)
try:
    clientes = alm.listar()
except ErrorNube as ex:
    st.error(f"{ex}\n\nRevisá NUBE_URL y NUBE_CLAVE en los secrets, y que el Apps Script esté implementado.")
    st.stop()
cfg_mail = alm.config_mail(dict(_secretos("mail", {}) or {}))
mail_ok = bool(cfg_mail.get("servidor") and cfg_mail.get("password"))

# ------------------------------------------------------------------ barra lateral
with st.sidebar:
    st.markdown("## 🧾 Facturador ARCA\n**Estudio Marsan**")
    modo = st.radio("Modo", ["🧪 Simulación (no emite nada)", "✅ Emitir en ARCA"], index=0)
    simulacion = modo.startswith("🧪")
    if simulacion:
        st.info("Simulación: recorre todo el circuito y genera PDF marcados SIN VALIDEZ. No se conecta con ARCA.")
    else:
        st.warning("Se pide CAE real. En clientes en **producción** las facturas quedan emitidas "
                   "y solo se anulan con nota de crédito.")
    st.divider()
    if en_nube:
        st.caption("☁️ Datos en tu Google Drive (compartidos entre PC y web)")
        if st.button("🔄 Actualizar desde el Drive"):
            del st.session_state["alm"]
            st.rerun()
    else:
        st.caption("💻 Datos guardados solo en esta PC")
    st.caption(f"Clientes cargados: {len(clientes)}")
    for c in clientes.values():
        ent = "🟢 prod" if c.get("entorno") == "produccion" else "🟡 homo"
        st.caption(f"{ent} · {c.get('razon_social')} ({c['cuit']})")

def _enviar_mails(res: list[dict], simulacion: bool, forzar: bool = False):
    import datetime as _dt
    if simulacion:  # nunca mandar simulaciones a los receptores
        lote = [{**r, "emails": [cfg_mail.get("remitente_email")]} for r in res]
        mailer.enviar_lote({**cfg_mail, "copia_cliente": False, "copia_estudio": False}, lote, forzar)
        for r, l in zip(res, lote):
            if l["mail"]:
                r["mail"] = l["mail"] + " (simulación)"
        return
    mailer.enviar_lote(cfg_mail, res, forzar)
    for r in res:
        if r["mail"].startswith("Enviado") and r.get("pdf_nombre"):
            alm.marcar_mail(r["pdf_nombre"], _dt.datetime.now().strftime("%d/%m/%Y %H:%M"))


tab_fact, tab_cli, tab_hist, tab_mail = st.tabs(["📤 Facturar", "👥 Clientes", "📚 Historial", "✉️ Mail"])

# ================================================================== FACTURAR
with tab_fact:
    st.subheader("1 · Bajá la plantilla y completala")
    c_simple, c_compl = st.columns(2)
    with c_simple:
        st.download_button("⬇️ Plantilla SIMPLE (recomendada)", generar_plantilla_simple(list(clientes.values())),
                           "facturas_simple.xlsx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Una fila por factura: cliente, receptor, descripción e importe total. Lo demás se completa solo.")
    with c_compl:
        st.download_button("⬇️ Plantilla completa", generar_plantilla(list(clientes.values())),
                           "plantilla_facturacion.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Para varios ítems por factura, percepciones, notas de crédito, FCE o dólares.")

    st.subheader("2 · Subí el Excel")
    archivo = st.file_uploader("Excel de facturación", type=["xlsx", "xlsm"], label_visibility="collapsed")
    if archivo:
        try:
            cbtes = leer_excel(archivo.getvalue(), clientes)
        except Exception as ex:
            st.error(f"No pude leer el Excel: {ex}")
            st.stop()

        for c in cbtes:
            prev = alm.ya_emitido(archivo.name, c["id"], c["cuit"])
            if prev:
                c["errores"].append(f"Ya emitido antes desde este mismo archivo: {prev['tipo']} "
                                    f"{int(prev['pto_vta']):05d}-{int(prev['numero']):08d} (CAE {prev['cae']})")

        ok = [c for c in cbtes if not c["errores"]]
        mal = [c for c in cbtes if c["errores"]]
        st.subheader("3 · Revisá antes de emitir")
        m1, m2, m3 = st.columns(3)
        m1.metric("Comprobantes", len(cbtes))
        m2.metric("Listos", len(ok))
        m3.metric("Con errores", len(mal))

        tabla = pd.DataFrame([{
            "Estado": "✅" if not c["errores"] else "❌",
            "ID": c["id"], "Emisor": c["cliente"].get("razon_social", c["cuit"]),
            "Entorno": c["cliente"].get("entorno", ""), "Tipo": c["tipo"], "PV": c["pto_vta"],
            "Fecha": c["fecha"].strftime("%d/%m/%Y"), "Concepto": c["concepto_txt"],
            "Receptor": c["rec_nombre"], "Doc": c["doc_nro"], "Cond. IVA": c["rec_cond_iva"],
            "Neto": c["imp_neto"], "IVA": c["imp_iva"], "Exento": c["imp_opex"],
            "No grav.": c["imp_no_gravado"], "Tributos": c["imp_trib"], "Total": c["total"],
            "Moneda": c["moneda"], "Mail": "📧 " + ", ".join(c["emails"]) if c["enviar_mail"] else "",
        } for c in cbtes])
        st.dataframe(tabla, hide_index=True, width="stretch",
                     column_config={k: st.column_config.NumberColumn(format="%.2f")
                                    for k in ("Neto", "IVA", "Exento", "No grav.", "Tributos", "Total")})

        for c in cbtes:
            icono = "❌" if c["errores"] else ("⚠️" if c["avisos"] else "✅")
            with st.expander(f"{icono} ID {c['id']} · {c['tipo']} · {c['cliente'].get('razon_social', c['cuit'])} "
                             f"→ {c['rec_nombre']} · $ {c['total']:,.2f}"):
                for e in c["errores"]:
                    st.error(e)
                for a in c["avisos"]:
                    st.warning(a)
                st.dataframe(pd.DataFrame([{
                    "Código": l["codigo"], "Descripción": l["descripcion"], "Cant.": l["cantidad"],
                    "Unidad": l["unidad"], "Precio": l["precio"], "Bonif %": l["bonif"],
                    "Alícuota": l["alicuota"] if c["letra"] != "C" else "—", "Importe": l["importe"],
                } for l in c["lineas"]]), hide_index=True, width="stretch")
                if c["tributos"]:
                    st.caption("Tributos: " + ", ".join(f"{t['Desc']} $ {t['Importe']:,.2f}" for t in c["tributos"]))
                if c["leyendas"]:
                    st.caption("Leyendas: " + " / ".join(c["leyendas"]))

        st.subheader("4 · Emitir")
        hay_prod = any(c["cliente"].get("entorno") == "produccion" for c in ok)
        confirmar = True
        if not simulacion and hay_prod:
            confirmar = st.checkbox(f"Confirmo que revisé los {len(ok)} comprobantes y quiero emitirlos en PRODUCCIÓN")
        if mal:
            st.caption(f"Los {len(mal)} comprobantes con errores no se van a emitir. Corregilos y volvé a subir el Excel.")
        con_mail = [c for c in ok if c["enviar_mail"]]
        auto_mail = False
        if con_mail:
            if not mail_ok:
                st.warning(f"{len(con_mail)} comprobante(s) tienen 'Enviar por mail' = S, pero el mail no está "
                           "configurado (pestaña ✉️ Mail). Se emiten igual y los podés mandar después.")
            elif simulacion:
                auto_mail = st.checkbox(f"Mandarme la simulación a mí ({cfg_mail.get('remitente_email')}) para ver "
                                        "cómo llega el mail (no se envía a los receptores)")
            else:
                auto_mail = st.checkbox(f"Enviar por mail los {len(con_mail)} comprobante(s) marcados, apenas se autoricen",
                                        value=True)
        etiqueta = f"{'🧪 Simular' if simulacion else '🚀 Emitir'} {len(ok)} comprobante(s)"
        if st.button(etiqueta, type="primary", disabled=not ok or not confirmar):
            barra = st.progress(0.0, "Comenzando…")
            res = emitir_lote(ok, alm, simulacion, archivo.name,
                              lambda i, n, t: barra.progress(i / n, f"{i}/{n} · {t}"))
            barra.empty()
            if auto_mail:
                with st.spinner("Enviando mails…"):
                    _enviar_mails(res, simulacion)
            st.session_state.resultados = res
            st.session_state.res_simulacion = simulacion

    res = st.session_state.get("resultados")
    if res:
        st.subheader("Resultado")
        aut = [r for r in res if r["pdf"]]
        st.success(f"Autorizados: {len(aut)} de {len(res)}") if aut else None
        df = pd.DataFrame([{k: r[k] for k in ("estado", "id", "cliente", "tipo", "pto_vta", "numero", "total",
                                               "cae", "cae_vto", "mail", "mensaje")} for r in res])
        st.dataframe(df, hide_index=True, width="stretch")
        for r in res:
            if r["estado"] in ("Rechazado", "Error"):
                st.error(f"ID {r['id']} ({r['cliente']}): {r['mensaje']}")
        pend_mail = [r for r in aut if r["emails"] and not r["mail"].startswith("Enviado")]
        if pend_mail and mail_ok:
            sim_res = st.session_state.get("res_simulacion", False)
            m1, m2 = st.columns([1, 2])
            solo_marcados = m2.checkbox("Solo los que tienen 'Enviar por mail' = S", value=True)
            if m1.button(f"📧 Enviar por mail ({len(pend_mail)})" + (" — a mí, es simulación" if sim_res else "")):
                with st.spinner("Enviando…"):
                    _enviar_mails(res, sim_res, forzar=not solo_marcados)
                st.rerun()
        if aut:
            z = io.BytesIO()
            with zipfile.ZipFile(z, "w") as zf:
                for r in aut:
                    zf.writestr(r["pdf_nombre"], r["pdf"])
            c1, c2 = st.columns(2)
            c1.download_button("⬇️ Todos los PDF (.zip)", z.getvalue(), "comprobantes.zip", "application/zip")
            xl = io.BytesIO()
            df.to_excel(xl, index=False, sheet_name="Resultado")
            c2.download_button("⬇️ Resultado en Excel", xl.getvalue(), "resultado_emision.xlsx")
            with st.expander("Descargar PDF de a uno"):
                for r in aut:
                    st.download_button(r["pdf_nombre"], r["pdf"], r["pdf_nombre"], "application/pdf",
                                       key="pdf_" + r["pdf_nombre"])

# ================================================================== CLIENTES
with tab_cli:
    st.subheader("Clientes que facturan")
    if clientes:
        st.dataframe(pd.DataFrame([{
            "CUIT": c["cuit"], "Razón social": c.get("razon_social"), "Condición": c.get("cond_iva"),
            "PV": c.get("pto_vta"), "Concepto": c.get("concepto"), "Entorno": c.get("entorno"),
            "Certificado": "✅" if all(alm.credenciales(c["cuit"])) else "❌ falta",
            "Guardado en": c.get("_origen"),
        } for c in clientes.values()]), hide_index=True, width="stretch")

    if en_nube:
        local = Almacen(BASE)
        locales = [d for d in local.listar().values() if d.get("_origen") == "local" and d["cuit"] not in clientes]
        if locales:
            st.info(f"En esta PC hay {len(locales)} cliente(s) guardados solo acá. Subilos al Drive para "
                    "poder facturarlos también desde la web.")
            if st.button("☁️ Subir clientes de esta PC al Drive"):
                subidos = alm.subir_locales(local)
                st.success("Subidos: " + ", ".join(subidos))
                st.rerun()

    opciones = ["➕ Nuevo cliente"] + [f"{c.get('razon_social')} ({c['cuit']})" for c in clientes.values()]
    sel = st.session_state.pop("cliente_recien_guardado", None)
    idx = next((i for i, o in enumerate(opciones) if sel and o.endswith(f"({sel})")), 0)
    elegido = st.selectbox("Cliente", opciones, index=idx)
    actual = {} if elegido.startswith("➕") else clientes[elegido.split("(")[-1].rstrip(")")]

    st.markdown("#### 1 · Datos fiscales (salen en la factura)")
    with st.form("form_cliente"):
        a, b = st.columns(2)
        cuit = a.text_input("CUIT (sin guiones)", actual.get("cuit", ""), disabled=bool(actual))
        razon = b.text_input("Razón social", actual.get("razon_social", ""))
        fantasia = a.text_input("Nombre de fantasía (opcional)", actual.get("nombre_fantasia", ""))
        domicilio = b.text_input("Domicilio comercial", actual.get("domicilio", ""))
        cond = a.selectbox("Condición frente al IVA", T.COND_IVA_EMISOR,
                           index=T.COND_IVA_EMISOR.index(actual["cond_iva"]) if actual.get("cond_iva") in T.COND_IVA_EMISOR else 0)
        iibb = b.text_input("Ingresos Brutos (n° o 'Exento' / 'Convenio Multilateral ...')", actual.get("iibb", ""))
        inicio = a.text_input("Fecha de inicio de actividades (DD/MM/AAAA)", actual.get("inicio_act", ""))
        pv = b.number_input("Punto de venta Web Services (por defecto)", 0, 99999, int(actual.get("pto_vta") or 0))
        conceptos = list(T.CONCEPTOS)
        concepto = a.selectbox("Qué factura habitualmente", conceptos,
                               index=conceptos.index(actual["concepto"]) if actual.get("concepto") in conceptos else 1)
        entorno = b.selectbox("Entorno", ["homologacion", "produccion"],
                              index=1 if actual.get("entorno") == "produccion" else 0,
                              help="Homologación = pruebas de ARCA, sin validez fiscal. Pasar a producción cuando todo funcione.")
        email = a.text_input("Email del cliente (opcional — recibe copia y las respuestas de los mails)",
                             actual.get("email", ""))
        leyenda = st.text_area("Leyenda fija en todas sus facturas (opcional)", actual.get("leyenda_fija", ""),
                               placeholder="Ej: CBU para transferencias… / Exento IIBB según… / Factura de crédito electrónica…")
        st.markdown("**Factura de Crédito Electrónica MiPyME** (solo si este cliente la emite)")
        f1, f2, f3 = st.columns([2, 1.3, 1])
        cbu = f1.text_input("CBU del cliente (22 dígitos)", actual.get("cbu", ""))
        alias_cbu = f2.text_input("Alias del CBU", actual.get("alias_cbu", ""))
        transfs = list(T.FCE_TRANSFERENCIA)
        fce_tr = f3.selectbox("Transferencia", transfs, index=transfs.index(actual["fce_transferencia"])
                              if actual.get("fce_transferencia") in transfs else 0,
                              help="SCA = Sistema de Circulación Abierta · ADC = Agente de Depósito Colectivo")
        if st.form_submit_button("💾 Guardar datos", type="primary"):
            cuit_n = "".join(ch for ch in (cuit or actual.get("cuit", "")) if ch.isdigit())
            faltan = [n for n, v in (("Razón social", razon), ("Domicilio comercial", domicilio),
                                     ("Punto de venta (no puede ser 0)", pv)) if not v]
            if len(cuit_n) != 11:
                st.error(f"❌ NO SE GUARDÓ: el CUIT tiene que tener 11 números (tiene {len(cuit_n)}).")
            elif not cuit_valido(cuit_n):
                st.error(f"❌ NO SE GUARDÓ: el CUIT {cuit_n} no es válido (revisá el último dígito).")
            elif faltan:
                st.error("❌ NO SE GUARDÓ: falta completar " + ", ".join(faltan) + ".")
            elif cbu and len("".join(ch for ch in cbu if ch.isdigit())) != 22:
                st.error("❌ NO SE GUARDÓ: el CBU tiene que tener 22 dígitos (o dejalo vacío).")
            else:
                try:
                    alm.guardar({"cuit": cuit_n, "razon_social": razon, "nombre_fantasia": fantasia,
                                 "domicilio": domicilio, "cond_iva": cond, "iibb": iibb, "inicio_act": inicio,
                                 "pto_vta": int(pv), "concepto": concepto, "leyenda_fija": leyenda,
                                 "email": email, "entorno": entorno,
                                 "cbu": "".join(ch for ch in cbu if ch.isdigit()), "alias_cbu": alias_cbu,
                                 "fce_transferencia": fce_tr})
                except Exception as ex:
                    st.error(f"❌ NO SE GUARDÓ: error al guardar en el Drive: {ex}")
                    st.stop()
                st.session_state.cliente_recien_guardado = cuit_n
                st.toast("✅ Cliente guardado")
                st.rerun()

    if actual:
        cuit_a = actual["cuit"]
        cert, clave = alm.credenciales(cuit_a)
        st.markdown("#### 2 · Certificado digital del cliente")
        if cert and clave:
            inf = info_certificado(cert)
            (st.error if inf["vencido"] else st.success)(
                f"Certificado cargado · vence {inf['vence']} · {inf['sujeto']}")
        elif clave and not cert:
            st.info("Ya generaste la clave y el pedido (CSR). Falta subir el certificado que te dio ARCA.")
        else:
            st.warning("Este cliente todavía no tiene certificado.")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Opción A — generar el pedido acá (recomendado)**")
            alias = st.text_input("Alias del certificado", f"facturador{cuit_a[-4:]}", key="alias")
            if st.button("🔑 Generar clave + pedido de certificado (CSR)"):
                if cert and not st.session_state.get("conf_regen"):
                    st.session_state.conf_regen = True
                    st.warning("Ya tiene certificado. Si generás una clave nueva, el certificado actual deja de servir. "
                               "Tocá de nuevo para confirmar.")
                else:
                    k, csr = generar_clave_y_csr(cuit_a, actual.get("razon_social", ""), alias)
                    alm.guardar_clave_pendiente(cuit_a, k)
                    st.session_state.csr = (cuit_a, csr)
                    st.session_state.conf_regen = False
            if st.session_state.get("csr") and st.session_state.csr[0] == cuit_a:
                st.download_button("⬇️ Descargar pedido_" + cuit_a + ".csr", st.session_state.csr[1],
                                   f"pedido_{cuit_a}.csr")
                st.caption("Subí este archivo en ARCA → Administración de Certificados Digitales (o WSASS si es "
                           "homologación) con la clave fiscal del cliente, y descargá el certificado .crt.")
        with c2:
            st.markdown("**Subir el certificado (.crt) que dio ARCA**")
            up_cert = st.file_uploader("Certificado .crt / .pem", type=["crt", "pem", "cer"], key="up_cert")
            up_key = st.file_uploader("Clave privada .key (solo si la generaste fuera de esta app)",
                                      type=["key", "pem"], key="up_key")
            if st.button("📎 Guardar certificado"):
                nuevo_cert = up_cert.getvalue() if up_cert else None
                nueva_clave = up_key.getvalue() if up_key else clave
                if not nuevo_cert:
                    st.error("Subí el .crt")
                elif not nueva_clave:
                    st.error("Falta la clave privada: generá el CSR acá o subí el .key")
                else:
                    try:
                        if not clave_coincide(nuevo_cert, nueva_clave):
                            st.error("El certificado no corresponde a esta clave privada. "
                                     "¿Es el certificado del último pedido generado?")
                        else:
                            alm.guardar(actual, nuevo_cert, nueva_clave)
                            st.success("Certificado guardado ✅")
                            st.rerun()
                    except Exception as ex:
                        st.error(f"No pude leer el certificado: {ex}")

        st.markdown("#### 3 · Probar la conexión con ARCA")
        st.caption("Consulta el último comprobante emitido en el punto de venta. Si responde, el certificado, la "
                   "relación con el servicio y el punto de venta están bien.")
        if st.button("🔌 Probar conexión", disabled=not (cert and clave)):
            try:
                ws = WSFE(WSAA(cuit_a, cert, clave, actual.get("entorno", "homologacion"), alm))
                pvs = ws.puntos_venta()
                activos = [p["Nro"] for p in pvs if p["Bloqueado"] != "S" and p["Baja"] in ("", "NULL")]
                if pvs:
                    st.markdown("**Puntos de venta habilitados para web services en ARCA:**")
                    st.table(pd.DataFrame([{"Punto de venta": p["Nro"], "Tipo": p["Tipo"],
                                            "Bloqueado": p["Bloqueado"], "Baja": p["Baja"]} for p in pvs]))
                elif actual.get("entorno") == "produccion":
                    st.error("Este cliente NO tiene ningún punto de venta para web services en ARCA. Crealo en "
                             "'Administración de puntos de venta y domicilios' con el sistema 'RECE para aplicativo "
                             "y web services' (RI) o 'Factura Electrónica - Monotributo - Web Services'.")
                    st.stop()
                if pvs and actual["pto_vta"] not in activos:
                    st.error(f"El punto de venta cargado en la app ({actual['pto_vta']}) no está en esa lista. "
                             f"Cambialo por uno de estos: {', '.join(map(str, activos)) or 'ninguno activo'} "
                             "y guardá los datos del cliente.")
                    st.stop()
                letras = sorted(T.LETRAS_POR_EMISOR[actual["cond_iva"]] - {"M"})
                filas = []
                for letra in letras:
                    cod = T.TIPOS_CBTE["F" + letra][0]
                    filas.append({"Comprobante": f"Factura {letra}", "PV": actual["pto_vta"],
                                  "Último autorizado": ws.ultimo_autorizado(actual["pto_vta"], cod)})
                st.success(f"Conexión OK con ARCA ({actual.get('entorno')})")
                st.table(pd.DataFrame(filas))
            except ErrorARCA as ex:
                msg = str(ex)
                st.error(f"ARCA respondió: {msg}")
                if "cms.cert.untrusted" in msg or "Certificado no emitido" in msg:
                    st.info("El certificado no es del entorno elegido (homologación vs producción).")
                elif "notAuthorized" in msg or "no autorizado" in msg.lower():
                    st.info("Falta asociar el certificado al servicio 'Facturación Electrónica' (wsfe) en el "
                            "Administrador de Relaciones de Clave Fiscal del cliente.")
            except Exception as ex:
                st.error(f"No me pude conectar: {ex}")

        with st.expander("☁️ Pasar este cliente a la versión web (solo si NO usás el Drive)"):
            st.caption("La versión web no guarda archivos de forma permanente. Copiá este bloque en "
                       "Streamlit Cloud → Settings → Secrets (debajo de lo que ya tengas).")
            if cert and clave:
                campos = {k: actual.get(k, "") for k in ("cuit", "razon_social", "nombre_fantasia", "domicilio",
                                                         "cond_iva", "iibb", "inicio_act", "pto_vta", "concepto",
                                                         "leyenda_fija", "email", "entorno", "cbu",
                                                         "alias_cbu", "fce_transferencia")}
                lineas = [f"[clientes.{cuit_a}]"]
                for k, v in campos.items():
                    lineas.append(f"{k} = {v}" if isinstance(v, int) else f'{k} = """{v}"""')
                lineas.append(f'certificado = """{cert.decode().strip()}"""')
                lineas.append(f'clave = """{clave.decode().strip()}"""')
                st.code("\n".join(lineas), language="toml")
            else:
                st.info("Primero cargá el certificado.")

        if actual.get("_origen") in ("local", "drive"):
            with st.expander("🗑️ Eliminar cliente"):
                if st.button(f"Eliminar {actual.get('razon_social')} y su certificado"):
                    alm.eliminar(cuit_a)
                    st.rerun()

# ================================================================== HISTORIAL
with tab_hist:
    st.subheader("Comprobantes emitidos")
    h = alm.historial()
    if not h:
        st.info("Todavía no hay comprobantes emitidos (las simulaciones no se guardan).")
    else:
        df = pd.DataFrame(h)
        filtro = st.multiselect("Cliente", sorted(df["razon_social"].unique()))
        if filtro:
            df = df[df["razon_social"].isin(filtro)]
        st.dataframe(df.drop(columns=["archivo_excel"]), hide_index=True, width="stretch")
        xl = io.BytesIO()
        df.to_excel(xl, index=False)
        st.download_button("⬇️ Historial en Excel", xl.getvalue(), "historial_facturas.xlsx")
        registros = [f for f in df.to_dict("records") if f.get("pdf")]
        if registros:
            elegido_pdf = st.selectbox("Bajar o reenviar un PDF", [f["pdf"] for f in registros])
            fila_h = next(f for f in registros if f["pdf"] == elegido_pdf)
            if st.session_state.get("pdf_hist_nombre") != elegido_pdf:
                st.session_state.pdf_hist_nombre = elegido_pdf
                try:
                    st.session_state.pdf_hist = alm.leer_pdf(fila_h["cuit_emisor"], elegido_pdf)
                except Exception as ex:
                    st.session_state.pdf_hist = None
                    st.error(f"No pude traer el PDF: {ex}")
            pdf_bytes = st.session_state.get("pdf_hist")
            if pdf_bytes:
                st.download_button("⬇️ " + elegido_pdf, pdf_bytes, elegido_pdf, "application/pdf")
            else:
                st.caption("El PDF no está guardado en este lugar.")
            dest = st.text_input("Reenviar por mail a", fila_h.get("email", "") or "", key="reenvio")
            if st.button("📧 Reenviar", disabled=not (mail_ok and dest and pdf_bytes)):
                try:
                    destinos = [d.strip() for d in dest.replace(",", ";").split(";") if d.strip()]
                    cli_h = clientes.get(fila_h["cuit_emisor"], {})
                    with mailer.Mailer(cfg_mail) as m:
                        m.enviar(mailer.armar_mensaje(cfg_mail, mailer.datos_desde_historial(fila_h), cli_h,
                                                      pdf_bytes, elegido_pdf, destinos))
                    import datetime as _dt
                    alm.marcar_mail(elegido_pdf, _dt.datetime.now().strftime("%d/%m/%Y %H:%M"))
                    st.success("Enviado ✅")
                except Exception as ex:
                    st.error(f"No se pudo enviar: {ex}")

# ================================================================== MAIL
with tab_mail:
    st.subheader("Envío de comprobantes por mail")
    st.caption("Es opcional: solo se mandan los comprobantes con 'Enviar por mail' = S en el Excel, o los que "
               "reenvíes desde el Historial.")
    if mail_ok:
        st.success(f"Configurado: sale desde {cfg_mail.get('remitente_nombre')} <{cfg_mail.get('remitente_email')}>")
    else:
        st.info("Todavía no está configurado.")
    cfg = {**mailer.DEFECTO, **cfg_mail}
    provs = list(mailer.PROVEEDORES)
    prov = st.selectbox("¿Dónde está la casilla info@estudiomarsan.com.ar?", provs,
                        index=provs.index(cfg["proveedor"]) if cfg.get("proveedor") in provs else 0,
                        help="Si no sabés: si entrás al mail por Gmail con esa casilla, es Google Workspace; "
                             "si entrás por el webmail de Hostinger, es Hostinger.")
    with st.form("form_mail"):
        a, b, c3 = st.columns([2, 1, 1])
        pre = mailer.PROVEEDORES[prov]
        cambio = prov != cfg.get("proveedor")
        servidor = a.text_input("Servidor SMTP", key=f"srv_{prov}", value=pre["servidor"] if cambio and pre["servidor"] else cfg["servidor"])
        puerto = b.number_input("Puerto", 1, 65535, key=f"pto_{prov}", value= int(pre["puerto"] if cambio else cfg["puerto"]))
        segs = ["SSL", "STARTTLS", "Ninguna"]
        seg = c3.selectbox("Seguridad", segs, key=f"seg_{prov}", index=segs.index(pre["seguridad"] if cambio else cfg["seguridad"]))
        a, b = st.columns(2)
        usuario = a.text_input("Usuario", cfg["usuario"])
        password = b.text_input("Contraseña" + (" (Gmail: contraseña de aplicación)" if "Gmail" in prov else ""),
                                cfg.get("password", ""), type="password")
        a, b = st.columns(2)
        rem_email = a.text_input("Remitente (email)", cfg["remitente_email"])
        rem_nombre = b.text_input("Remitente (nombre que ve el cliente)", cfg["remitente_nombre"])
        a, b = st.columns(2)
        copia_cli = a.checkbox("Mandar copia al cliente del estudio (el que factura)", cfg["copia_cliente"])
        copia_est = b.checkbox("Mandar copia a la casilla del estudio", cfg["copia_estudio"])
        asunto = st.text_input("Asunto", cfg["asunto"])
        cuerpo = st.text_area("Texto del mail", cfg["cuerpo"], height=260)
        st.caption("Podés usar: {emisor} {receptor} {comprobante} {numero} {fecha} {total} {moneda} {cae}")
        if st.form_submit_button("💾 Guardar configuración", type="primary"):
            alm.guardar_config_mail({"proveedor": prov, "servidor": servidor, "puerto": int(puerto),
                                     "seguridad": seg, "usuario": usuario, "password": password,
                                     "remitente_email": rem_email, "remitente_nombre": rem_nombre,
                                     "copia_cliente": copia_cli, "copia_estudio": copia_est,
                                     "asunto": asunto, "cuerpo": cuerpo})
            st.success("Guardado")
            st.rerun()
    a, b = st.columns([2, 1])
    destino_prueba = a.text_input("Mandar un mail de prueba a", cfg_mail.get("remitente_email", ""))
    if b.button("✉️ Probar envío", disabled=not mail_ok):
        try:
            mailer.probar(cfg_mail, destino_prueba)
            st.success(f"Mail de prueba enviado a {destino_prueba}")
        except Exception as ex:
            st.error(f"No se pudo enviar: {ex}")
    with st.expander("☁️ Configuración para la versión web"):
        st.caption("Pegar en Streamlit Cloud → Settings → Secrets:")
        st.code("[mail]\n" + "\n".join(
            f"{k} = {v}" if isinstance(v, (int, bool)) and not isinstance(v, bool) else
            f"{k} = {'true' if v else 'false'}" if isinstance(v, bool) else f'{k} = """{v}"""'
            for k, v in cfg_mail.items()), language="toml")
