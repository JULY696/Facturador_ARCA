# Facturador ARCA — Estudio Marsan

Subís un Excel con las facturas de uno o varios clientes, la app las revisa, pide el CAE a ARCA con el certificado de **cada cliente** y devuelve los PDF con QR.

---

## A. Instalar en la PC (una sola vez)

1. Instalar **Python 3.11 o superior** desde python.org. En el instalador marcá **"Add Python to PATH"**.
2. Descomprimir la carpeta `facturador_arca` donde quieras (por ejemplo, `Documentos\Facturador`).
3. Doble clic en **`INICIAR Facturador.bat`**. La primera vez instala todo (tarda unos minutos) y después abre la app en el navegador.

Todo lo de cada cliente queda en la carpeta `datos\` (certificados, claves, PDF emitidos, historial). **Hacé backup de esa carpeta y no la compartas**: con esos archivos se puede facturar a nombre de tus clientes.

## B. Dar de alta un cliente (una vez por cliente)

En la app: pestaña **Clientes → Nuevo cliente**.

1. **Datos fiscales:** CUIT, razón social, domicilio, condición frente al IVA, IIBB, inicio de actividades, punto de venta, si factura productos o servicios y la leyenda fija (CBU, etc.). Dejalo en **homologación** al principio.
2. **Certificado:** tocá **"Generar clave + pedido de certificado (CSR)"** y bajá el `.csr`.
3. **En ARCA, con la clave fiscal del cliente (nivel 3):**

   **Para pruebas (homologación):**
   - En *Administrador de Relaciones de Clave Fiscal*, adherir el servicio **"WSASS – Autogestión Certificados Homologación"**.
   - Entrar a WSASS → *Nuevo Certificado* → poner un nombre y pegar el contenido del `.csr` → copiar el certificado que devuelve y guardarlo como `certificado.crt`.
   - En WSASS → *Crear autorización a servicio* → elegir el certificado y el servicio **wsfe**.
   - En homologación no hace falta crear el punto de venta (podés usar cualquier número, por ejemplo 1).

   **Para facturar de verdad (producción):**
   - **Administración de Certificados Digitales** → *Agregar alias* → subir el `.csr` → descargar el `.crt`.
   - **Administrador de Relaciones** → *Nueva Relación* → ARCA → Web Services → **Facturación Electrónica** → en *Representante* elegir el alias del certificado → confirmar.
   - **Administración de puntos de venta y domicilios** → agregar un punto de venta **nuevo** del tipo *"RECE para aplicativo y web services"* (RI) o *"Factura Electrónica – Monotributo – Web Services"* (monotributo). No sirve el de "Comprobantes en línea".
   - El certificado de producción es **otro** que el de homologación: al pasar a producción hay que generar un CSR nuevo.

4. Volver a la app → **Subir el certificado (.crt)** → **Guardar**. La app controla que el certificado corresponda a la clave.
5. **Probar conexión**: si muestra el último comprobante autorizado, está todo bien.

## C. Facturar

1. **Facturar → Descargar plantilla Excel** (ya trae los CUIT de tus clientes en los desplegables).
2. Completar:
   - **Comprobantes:** una fila por factura, nota de crédito o débito. En un mismo Excel puede haber facturas de distintos clientes.
   - **Items:** los renglones de cada factura (código, descripción, cantidad, unidad, precio, bonificación, alícuota). Se unen con el comprobante por el **ID**.
   - **Tributos:** percepciones de IIBB, de IVA, etc. (opcional).
3. Subir el Excel. La app muestra cada comprobante con ✅ o ❌ y el motivo, por ejemplo: CUIT mal, letra que no corresponde, falta de ítems o fecha fuera de rango.
4. Probar primero en **🧪 Simulación**: arma todo y genera PDF marcados "SIN VALIDEZ", sin conectarse con ARCA.
5. Pasar a **✅ Emitir en ARCA** → **Emitir**. Te devuelve el CAE de cada uno, el ZIP con los PDF y un Excel con el resultado.

**Reglas que aplica la app:**
- **Precios:** en A/B el precio va **sin IVA**, salvo que pongas `S` en "Precios con IVA incluido". En C es el precio final.
- **Servicios:** si no completás el período, toma el mes de la fecha. Si no completás el vencimiento de pago, usa la fecha de la factura.
- **Factura B:** agrega sola la leyenda del Régimen de Transparencia Fiscal al Consumidor (IVA contenido).
- **Doble emisión:** no deja emitir dos veces la misma fila del mismo archivo.
- **Corte de conexión:** antes de reintentar, consulta a ARCA si la factura llegó a autorizarse.

## D. Factura de Crédito Electrónica MiPyME (FCE)

1. En **Clientes**, cargar el **CBU** (22 dígitos), el alias y la opción de transferencia (**SCA** o **ADC**).
2. En el Excel, usar los tipos **FCEA / FCEB / FCEC**. Para las notas: **NCFCEA, NDFCEA, NCFCEB**, etc.
3. Reglas:
   - El receptor tiene que tener **CUIT**.
   - Si no ponés **Vto. pago**, la app lo calcula a 30 días.
   - El CBU, el alias y SCA/ADC se pueden cambiar en la fila. Si no los completás, se toman del cliente.
   - Las notas FCE llevan el comprobante FCE asociado **con su fecha** y la columna **FCE anulación** (S si es por rechazo o anulación de la factura).
4. Si cargás una Factura A común igual o mayor al monto mínimo de FCE, la app avisa para que verifiques si el receptor está en el listado de grandes empresas. El monto está en `facturador/tablas.py` (`MONTO_MINIMO_FCE`) y hay que actualizarlo cuando ARCA lo cambie.
5. En ARCA, el punto de venta de web services tiene que estar habilitado para FCE. Si al emitir da error de punto de venta, revisalo en *Administración de puntos de venta*.

## E. Facturas en dólares o euros

- En **Moneda** elegir *Dólares* o *Euros*. Los precios de los ítems van en esa moneda.
- **Cotización:** si la dejás vacía, la app usa la oficial que informa ARCA al momento de emitir. Si ponés una, usa la tuya.
- **Cancela en moneda extranjera:** S si te pagan en esa misma moneda.
- El PDF muestra el tipo de cambio y el equivalente en pesos.

## F. Envío por mail (opcional)

1. Pestaña **✉️ Mail**: elegir dónde está la casilla info@estudiomarsan.com.ar (Hostinger, Google Workspace u otro), poner la contraseña y guardar.
   - En **Google Workspace / Gmail** hay que usar una **contraseña de aplicación** (Cuenta de Google → Seguridad → Verificación en 2 pasos → Contraseñas de aplicaciones), no la contraseña normal.
2. Tocar **Probar envío**.
3. En el Excel: completar **Email receptor** (varios separados por `;`) y poner **S** en **Enviar por mail**. Los que no tengan S no se mandan.
4. Al emitir, tildar *"Enviar por mail…"*. También podés mandarlos después con el botón **📧 Enviar por mail**, o reenviar cualquiera desde **Historial**.
5. Cómo sale el mail:
   - Remitente: **ESTUDIO MARSAN <info@estudiomarsan.com.ar>**, con el PDF adjunto.
   - Copia al cliente que factura (si tiene email cargado). Si el receptor responde, la respuesta le llega a ese cliente.
   - El asunto y el texto se editan en la pestaña Mail.
6. **En modo simulación nunca se envía nada a los receptores ni a los clientes.** Solo podés mandarte la prueba a vos.

## G. Usarlo online, desde cualquier PC

Todo (clientes, certificados, historial, PDF, mail y tickets de ARCA) se guarda en **tu Google Drive**, en la carpeta **"Facturador ARCA"**. La versión de la PC y la web leen y escriben ahí, así que da lo mismo desde dónde factures.

### G1. Instalar el Apps Script (una sola vez, 5 minutos)
1. Entrá a **script.google.com** con tu cuenta de Google → **Nuevo proyecto** → ponele de nombre *Facturador ARCA – Almacén*.
2. Borrá lo que aparece y pegá todo el contenido de `apps_script/Codigo.gs`. Guardá (💾).
3. Tocá **⚙️ Configuración del proyecto → Propiedades del script → Agregar propiedad**:
   - Propiedad: `CLAVE`
   - Valor: una clave larga que solo vos conozcas (por ejemplo, 25 caracteres mezclados).
4. Tocá **Implementar → Nueva implementación** → ⚙️ **Aplicación web**:
   - *Ejecutar como*: **Yo**
   - *Quién tiene acceso*: **Cualquier usuario**. Igual nadie puede entrar sin la CLAVE.
5. Tocá **Implementar** y autorizá los permisos (Drive y Hojas de cálculo). Si aparece "Google no verificó esta app", tocá **Configuración avanzada → Ir a…**.
6. Copiá la **URL de la aplicación web** (termina en `/exec`). Si la abrís en el navegador, tiene que decir *"almacenamiento activo"*.

### G2. Conectar la PC al Drive
1. En la carpeta del programa, dentro de `.streamlit`, copiá `secrets_EJEMPLO.toml` y renombralo **`secrets.toml`**.
2. Completá `NUBE_URL` con la URL y `NUBE_CLAVE` con la CLAVE. `APP_PASSWORD` lo podés borrar en la PC.
3. Abrí el programa. En la barra lateral tiene que decir **"☁️ Datos en tu Google Drive"**.
4. Si ya tenías clientes cargados en la PC, en **Clientes** aparece el botón **"☁️ Subir clientes de esta PC al Drive"**.

### G3. Publicar la versión web (igual que el conciliador)
1. Creá un repositorio **privado** en GitHub y subí la carpeta del programa. El `.gitignore` ya deja afuera `datos/` y `secrets.toml`: **nunca subas esos dos**.
2. En **share.streamlit.io → Create app**, elegí el repositorio, el archivo principal `app.py`, y entrá a **Advanced settings → Secrets**.
3. Pegá ahí `APP_PASSWORD`, `NUBE_URL` y `NUBE_CLAVE` (ver `secrets_EJEMPLO.toml`).
4. Tocá **Deploy**. Te queda un link `https://….streamlit.app` que podés abrir desde cualquier PC o desde el celular, entrando con APP_PASSWORD. También lo podés enlazar en la sección Apps de tu web.

### G4. Cómo se trabaja después
- **Un cliente nuevo:** se da de alta desde la PC o desde la web. Todo el trámite del certificado se puede hacer online, porque la clave privada queda guardada en el Drive.
- **Si cargás algo en la web y no lo ves en la PC** (o al revés): tocá **"🔄 Actualizar desde el Drive"**.
- **El historial** está en la planilla *"Historial Facturador"* y **los PDF** en `Facturador ARCA/comprobantes/<CUIT>/` de tu Drive.
- **Si cambiás el Apps Script más adelante:** **Implementar → Gestionar implementaciones → ✏️ → Versión: nueva**. Así la URL no cambia.

### Seguridad
- La carpeta "Facturador ARCA" tiene las claves privadas de tus clientes: **no la compartas con nadie**.
- Tené la verificación en 2 pasos activada en tu cuenta de Google.
- Si sospechás que alguien conoce la CLAVE, cambiala en las propiedades del script y en los secrets.
- La alternativa sin Drive (pegar cada cliente en los Secrets de Streamlit) sigue funcionando, pero ahí la PC y la web no comparten historial ni tickets de ARCA. Por eso es mejor el Drive.
