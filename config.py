"""Parámetros generales de la aplicación."""
import os
from datetime import timedelta
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


class Config:
    # 0.0.0.0: accesible desde la red del despacho (protegida con contraseña).
    HOST = os.environ.get("ORIENS_HOST", "0.0.0.0")
    PUERTO = int(os.environ.get("ORIENS_PUERTO", "5050"))
    LIMITE_ARCHIVO_MB = 10
    MAX_CONTENT_LENGTH = LIMITE_ARCHIVO_MB * 1024 * 1024  # Flask rechaza cargas mayores (413)

    CARPETA_PLANTILLAS = BASE_DIR / "plantillas"  # plantillas marcadas; se elige una en cada generación
    FORMATO_CAPTURA = CARPETA_PLANTILLAS / "Formato de captura de trabajadores.xlsx"
    CARPETA_PERFILES = BASE_DIR / "perfiles de puesto"  # «Perfil de Puesto <puesto>.docx»
    CARPETA_SALIDAS = BASE_DIR / "contratos generados"  # una carpeta por cliente y fecha

    # Contraseña (cifrada) y llave de sesión; no se versiona.
    ARCHIVO_ACCESO = Path(os.environ.get("ORIENS_ACCESO", BASE_DIR / "acceso.json"))
    PERMANENT_SESSION_LIFETIME = timedelta(hours=12)
    SESSION_COOKIE_NAME = os.environ.get("ORIENS_COOKIE", "oriens_sesion")
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
