# 🎓 Portal Académico QR — Cloudflare R2 + Render

Sistema de control de asistencia con código QR y gestión de noticias para profesores, desplegado en **Render** con almacenamiento persistente en **Cloudflare R2**.

---

## 🌟 Características de esta versión

- **Persistencia garantizada**: Resuelve el problema del disco efímero de Render.
- **Formato nativo Excel (`.xlsx`)**: Mantiene fórmulas, formato, celdas coloreadas en verde/rojo y fuentes usando `openpyxl`.
- **Cero costo**: 10 GB de almacenamiento gratuito mensual y 0 costo de transferencia con Cloudflare R2.
- **Frontend idéntico**: Funciona con el mismo diseño y escáner de cámara del `index.html` original.

---

## 📁 Estructura del Proyecto

```
QR-R2/
├── app.py                     # Servidor Flask con lógica openpyxl + sync R2
├── r2_storage.py              # Capa de sincronización S3 con Cloudflare R2
├── index.html                 # Frontend con escáner QR y panel de materias
├── requirements.txt           # Dependencias Python (Flask, openpyxl, gunicorn, boto3)
├── Procfile                   # Archivo de inicio para Render (gunicorn app:app)
├── render.yaml                # Especificación de variables para Render
├── .gitignore                 # Filtro para Git
├── cloudflare_r2_guide.md     # Guía detallada para configurar Cloudflare R2
├── archivos_excel/            # Plantillas iniciales de los 3 archivos Excel
└── README.md
```

---

## ⚙️ Variables de Entorno en Render

| Variable | Descripción |
|---|---|
| `R2_ACCOUNT_ID` | ID de cuenta de Cloudflare (aparece en el dashboard de R2) |
| `R2_ACCESS_KEY_ID` | Access Key generado en el token de R2 |
| `R2_SECRET_ACCESS_KEY` | Secret Key generado en el token de R2 |
| `R2_BUCKET_NAME` | Nombre de tu bucket en R2 (ej. `asistencia-qr`) |

> Consulta el archivo [`cloudflare_r2_guide.md`](file:///home/ebany/Escritorio/5/Servicio%20Social/QR-R2/cloudflare_r2_guide.md) para ver la guía paso a paso con capturas y explicaciones.

---

## 🚀 Despliegue en Render

1. Sube esta carpeta `QR-R2` a un repositorio en **GitHub**:
   ```bash
   cd "QR-R2"
   git init
   git add .
   git commit -m "Portal de Asistencia con Cloudflare R2"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
   git push -u origin main
   ```
2. En [Render.com](https://render.com), haz clic en **New +** -> **Web Service**.
3. Conecta tu repositorio de GitHub.
4. Render detectará automáticamente el archivo `render.yaml` (o selecciona Python, build command `pip install -r requirements.txt`, start command `gunicorn app:app`).
5. En la sección **Environment**, ingresa las 4 variables de entorno indicadas arriba.
6. Haz clic en **Deploy Web Service**.
