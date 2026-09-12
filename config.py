"""Parámetros generales de la aplicación."""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


class Config:
    PUERTO = 5050
    LIMITE_ARCHIVO_MB = 10
    MAX_CONTENT_LENGTH = LIMITE_ARCHIVO_MB * 1024 * 1024  # Flask rechaza cargas mayores (413)
    CARPETA_PLANTILLAS = BASE_DIR / "plantillas"
    # Copia marcada de la plantilla oficial (se crea con plantillas/marcar_plantilla.py).
    PLANTILLA_CONTRATO = CARPETA_PLANTILLAS / "Contrato tiempo indeterminado (marcado).docx"
    CARPETA_PERFILES = BASE_DIR / "perfiles de puesto"  # «Perfil de Puesto <puesto>.docx»
    CARPETA_SALIDAS = BASE_DIR / "contratos generados"
