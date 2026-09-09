"""
r2_storage.py — Capa de persistencia con Cloudflare R2 (API compatible con AWS S3)
Descarga los archivos Excel a /tmp/ para que openpyxl los lea/modifique,
y los resube inmediatamente a Cloudflare R2 para asegurar la persistencia en Render.
"""

import os
import boto3
from botocore.config import Config
from botocore.exceptions import ClientError

# Directorio temporal local en Render/Linux
LOCAL_TMP_DIR = os.environ.get('TMPDIR', '/tmp/archivos_excel')
os.makedirs(LOCAL_TMP_DIR, exist_ok=True)

# Variables de entorno para Cloudflare R2
R2_ACCOUNT_ID = os.environ.get('R2_ACCOUNT_ID', '').strip()
R2_ACCESS_KEY_ID = os.environ.get('R2_ACCESS_KEY_ID', '').strip()
R2_SECRET_ACCESS_KEY = os.environ.get('R2_SECRET_ACCESS_KEY', '').strip()
R2_BUCKET_NAME = os.environ.get('R2_BUCKET_NAME', '').strip()

_s3_client = None


def is_r2_configured() -> bool:
    """Verifica si las credenciales de Cloudflare R2 están configuradas."""
    return bool(R2_ACCOUNT_ID and R2_ACCESS_KEY_ID and R2_SECRET_ACCESS_KEY and R2_BUCKET_NAME)


def get_s3_client():
    """Inicializa y retorna el cliente de boto3 conectado a Cloudflare R2."""
    global _s3_client
    if _s3_client is not None:
        return _s3_client

    if not is_r2_configured():
        print("⚠️ Variables de Cloudflare R2 incompletas. Operando en modo local (/tmp).")
        return None

    endpoint_url = f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com"
    _s3_client = boto3.client(
        's3',
        endpoint_url=endpoint_url,
        aws_access_key_id=R2_ACCESS_KEY_ID,
        aws_secret_access_key=R2_SECRET_ACCESS_KEY,
        config=Config(signature_version='s3v4'),
        region_name='auto'  # Cloudflare R2 requiere 'auto'
    )
    return _s3_client


def get_local_path(filename: str) -> str:
    """Retorna la ruta absoluta del archivo en la carpeta temporal local."""
    return os.path.join(LOCAL_TMP_DIR, filename)


def ensure_file_local(filename: str, force_download: bool = False) -> str:
    """
    Asegura que el archivo exista localmente en /tmp/.
    Si no existe o si force_download=True, lo descarga desde Cloudflare R2.
    """
    local_path = get_local_path(filename)

    # Si ya existe localmente y no forzamos descarga, reutilizar
    if os.path.exists(local_path) and not force_download:
        return local_path

    s3 = get_s3_client()
    if s3:
        try:
            print(f"📥 Descargando '{filename}' desde Cloudflare R2 a {local_path}...")
            s3.download_file(R2_BUCKET_NAME, filename, local_path)
            print(f"✅ '{filename}' descargado con éxito.")
            return local_path
        except ClientError as e:
            error_code = e.response.get('Error', {}).get('Code', 'Unknown')
            if error_code == '404' or error_code == 'NoSuchKey':
                print(f"⚠️ El archivo '{filename}' no existe en el bucket R2 '{R2_BUCKET_NAME}'.")
            else:
                print(f"❌ Error al descargar de R2 ({filename}): {e}")
    
    # Respaldo si no está en R2 ni configurado: buscar en carpeta de origen si existe
    fallback_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'archivos_excel')
    fallback_path = os.path.join(fallback_dir, filename)
    if os.path.exists(fallback_path):
        import shutil
        shutil.copy2(fallback_path, local_path)
        print(f"📋 Copiado archivo base desde {fallback_path} a {local_path}")

    return local_path


def upload_file_to_r2(filename: str) -> bool:
    """
    Sube el archivo modificado de /tmp/ a Cloudflare R2 para asegurar la persistencia.
    """
    s3 = get_s3_client()
    if not s3:
        print(f"⚠️ R2 no configurado. El archivo '{filename}' solo se guardó en el disco local efímero.")
        return True

    local_path = get_local_path(filename)
    if not os.path.exists(local_path):
        print(f"❌ Error: El archivo local no existe para subir a R2: {local_path}")
        return False

    try:
        print(f"📤 Subiendo '{filename}' actualizado a Cloudflare R2...")
        s3.upload_file(local_path, R2_BUCKET_NAME, filename)
        print(f"✅ '{filename}' sincronizado exitosamente en Cloudflare R2.")
        return True
    except Exception as e:
        print(f"❌ Error al subir a Cloudflare R2 ({filename}): {e}")
        return False
