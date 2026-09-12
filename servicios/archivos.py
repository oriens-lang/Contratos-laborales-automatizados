"""Revisión básica del archivo recibido (formato), sin leer su contenido."""
from __future__ import annotations

# Un .xlsx es un paquete ZIP: sus primeros bytes siempre son "PK\x03\x04".
FIRMA_ZIP = b"PK\x03\x04"


def revisar_archivo_xlsx(nombre: str, contenido: bytes) -> list[str]:
    """Devuelve la lista de errores de formato; vacía si el archivo es aceptable."""
    if not nombre.lower().endswith(".xlsx"):
        return ["El archivo debe tener extensión .xlsx."]
    if not contenido.startswith(FIRMA_ZIP):
        return ["El archivo está vacío o no es un libro de Excel válido (.xlsx)."]
    return []


def revisar_archivo_docx(nombre: str, contenido: bytes) -> list[str]:
    """Igual que el anterior, para documentos Word (.docx), que también son paquetes ZIP."""
    if not nombre.lower().endswith(".docx"):
        return ["El perfil debe ser un documento Word (.docx)."]
    if not contenido.startswith(FIRMA_ZIP):
        return ["El archivo está vacío o no es un documento Word válido (.docx)."]
    return []
