"""Plantillas de contrato disponibles: copias marcadas con {{ campo }} en «plantillas/»."""
from __future__ import annotations

import re
from pathlib import Path

from servicios.generador import ErrorGeneracion, revisar_plantilla


def listar(carpeta: Path) -> list[str]:
    """Plantillas utilizables (se excluyen los originales sin marcar y los temporales de Word)."""
    return sorted(ruta.name for ruta in carpeta.glob("*.docx")
                  if not ruta.name.startswith("~$") and "(original)" not in ruta.name)


def describir(carpeta: Path) -> list[dict]:
    """Plantillas con lo que muestra la pantalla: cuántos datos usa y si lleva el perfil como ANEXO UNO."""
    descripcion = []
    for nombre in listar(carpeta):
        try:
            marcadores, error = revisar_plantilla((carpeta / nombre).read_bytes()), None
        except ErrorGeneracion as problema:
            marcadores, error = set(), str(problema)
        descripcion.append({
            "nombre": nombre,
            "titulo": Path(nombre).stem,
            "campos": len(marcadores),
            "anexo": any(m.startswith("actividad_") for m in marcadores),
            "error": error,
        })
    return descripcion


def ruta(carpeta: Path, nombre: str) -> Path:
    """Ruta de una plantilla de la lista; impide usar archivos fuera de la carpeta."""
    if nombre not in listar(carpeta):
        raise ErrorGeneracion(f"No existe la plantilla «{nombre}».")
    return carpeta / nombre


def guardar(carpeta: Path, nombre_archivo: str, contenido: bytes) -> str:
    """Revisa los marcadores de una plantilla subida y la guarda (reemplaza la del mismo nombre)."""
    nombre = re.sub(r'[\\/:*?"<>|_\x00-\x1f]+', " ", Path(nombre_archivo).stem)
    nombre = re.sub(r"\s+", " ", nombre).strip(" .")
    if not nombre:
        raise ErrorGeneracion("La plantilla no tiene nombre.")
    if "(original)" in nombre:
        raise ErrorGeneracion("El nombre no puede contener «(original)»: así se identifica el formato sin marcar.")
    revisar_plantilla(contenido)
    destino = carpeta / f"{nombre}.docx"
    destino.write_bytes(contenido)
    return destino.name
