# ☁️ Guía Rápida: Configuración de Cloudflare R2 para Render

Cloudflare R2 te da **10 GB de almacenamiento gratuito cada mes sin vencimiento y sin costo por transferencia**. Aquí tienes el paso a paso detallado para configurarlo en 5 minutos.

---

## Paso 1: Crear una cuenta en Cloudflare y entrar a R2

1. Ve a **[dash.cloudflare.com](https://dash.cloudflare.com/)** e inicia sesión (o regístrate gratis si no tienes cuenta).
2. En el menú lateral izquierdo, haz clic en **R2** (o **R2 Object Storage**).
3. Si es tu primera vez usando R2, Cloudflare te pedirá habilitarlo (puede pedirte vincular una tarjeta para verificar identidad en algunos casos, pero **no cobra nada**, el plan gratuito de 10 GB es 100% permanente).

---

## Paso 2: Crear el Bucket de Almacenamiento

1. Dentro de la sección **R2**, haz clic en el botón azul **"Create bucket"** (Crear bucket).
2. Asigna un nombre al bucket (por ejemplo: `asistencia-qr`).
3. Deja la ubicación en **Automatic** (o la que prefieras).
4. Haz clic en **"Create bucket"**.
5. ¡Listo! Ya tienes tu bucket creado.

---

## Paso 3: Subir los 3 Archivos Excel Iniciales

Dentro de tu bucket recién creado (`asistencia-qr`):
1. Haz clic en **"Upload"** -> **"Upload files"**.
2. Selecciona y sube los 3 archivos que tienes en la carpeta `archivos_excel/`:
   - `IIAT31-OPTIMIZACION-A1.xlsx`
   - `IIAT32-TEORIA DE LA DECISION-A2.xlsx`
   - `IIAT33-INVESTIGACION DE OPERACIONES-A3.xlsx`
3. Verifica que los nombres coincidan exactamente (con mayúsculas y extensión `.xlsx`).

---

## Paso 4: Obtener las Credenciales de la API (Tokens)

1. En el panel principal de **R2** (puedes volver dando clic a **R2** en el menú izquierdo):
2. En la parte derecha, busca el enlace **"Manage R2 API Tokens"** (Administrar tokens de API de R2) o **"Create API Token"**.
3. Haz clic en **"Create API token"**.
4. Configúralo así:
   - **Token name**: `render-asistencia-token`
   - **Permissions**: Selecciona **Object Read & Write** (Lectura y Escritura de objetos).
   - **Apply to specific buckets only** (opcional): Puedes elegir tu bucket `asistencia-qr` o dejarlo en todos los buckets.
   - **TTL**: Déjalo en Forever o según prefieras.
5. Haz clic en **"Create API Token"** al final de la página.
6. En la pantalla siguiente aparecerán tus credenciales:
   - **Access Key ID**: Cópialo.
   - **Secret Access Key**: Cópialo (solo se muestra una vez).

---

## Paso 5: Obtener tu Account ID (ID de Cuenta)

1. Regresa a la página principal de **R2**.
2. En la barra lateral derecha verás un cuadro titulado **"Account ID"** (ID de cuenta) con una cadena de letras y números (ejemplo: `a1b2c3d4e5f6...`).
3. Cópialo.

---

## Paso 6: Configurar las Variables en Render

En tu servicio web de **Render**:
Ve a tu proyecto -> Pestaña **Environment** -> Agrega estas 4 variables:

| Variable | Valor |
|---|---|
| `R2_ACCOUNT_ID` | Tu Account ID copiado de Cloudflare R2 |
| `R2_ACCESS_KEY_ID` | Tu Access Key ID copiado del token |
| `R2_SECRET_ACCESS_KEY` | Tu Secret Access Key copiado del token |
| `R2_BUCKET_NAME` | El nombre de tu bucket (ej. `asistencia-qr`) |

---

## 🎉 ¡Listo!

Cuando Render arranque tu servidor:
1. El script detectará las credenciales y se conectará automáticamente al bucket R2.
2. Cada vez que alguien pase lista o escanee un QR, el Excel en R2 se actualizará de inmediato.
3. Si el servidor gratuito de Render se apaga o duerme por inactividad, al despertar descargará la versión más reciente sin perder ninguna asistencia.
