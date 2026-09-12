"""Perfiles de puesto: relación entre el PUESTO del Excel y su perfil en Word.

Cada perfil se guarda en la carpeta «perfiles de puesto» con el nombre
«Perfil de Puesto <nombre del puesto>.docx». El puesto se compara sin distinguir
mayúsculas, acentos ni espacios repetidos («Diseñadora» = «DISEÑADORA»).

De cada perfil se toman, sin modificarlas, las cinco actividades de la sección
«Cinco actividades principales». Solo se quita el número de lista («1. ») porque
la cláusula PRIMERA del contrato ya numera los cinco espacios.
"""
from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Union

from docx import Document

from servicios.lector_excel import normalizar

ACTIVIDADES_REQUERIDAS = 5
PATRON_NOMBRE = re.compile(r"^\s*perfil\s+de\s+puesto\s+(.+?)\s*$", re.IGNORECASE)
PATRON_ENUMERACION = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+)$", re.DOTALL)
SECCION_ACTIVIDADES = "CINCO ACTIVIDADES PRINCIPALES"


class ErrorPerfil(Exception):
    """El perfil existe pero no se puede usar (p. ej., no tiene las cinco actividades)."""


def cargar_catalogo(carpeta: Path) -> dict[str, dict]:
    """Puesto normalizado → {puesto, archivo, actividades, error}, uno por perfil de la carpeta."""
    catalogo = {}
    if not carpeta.is_dir():
        return catalogo
    for ruta in sorted(carpeta.glob("*.docx")):
        coincidencia = PATRON_NOMBRE.match(ruta.stem)
        if ruta.name.startswith("~$") or not coincidencia:  # temporales de Word u otros archivos
            continue
        puesto = coincidencia.group(1)
        try:
            actividades, error = extraer_actividades(ruta), None
        except ErrorPerfil as problema:
            actividades, error = [], str(problema)
        catalogo[normalizar(puesto)] = {
            "puesto": puesto, "archivo": ruta, "actividades": actividades, "error": error,
        }
    return catalogo


def guardar_perfil(carpeta: Path, puesto: str, contenido: bytes) -> dict:
    """Valida un perfil subido desde la pantalla y lo guarda como «Perfil de Puesto <puesto>.docx».

    Si el puesto ya tenía perfil (aunque con otro nombre de archivo), se reemplaza.
    """
    puesto = re.sub(r'[\\/:*?"<>|_\x00-\x1f]+', " ", puesto)
    puesto = re.sub(r"\s+", " ", puesto).strip(" .")
    if not puesto:
        raise ErrorPerfil("falta el nombre del puesto")
    actividades = extraer_actividades(BytesIO(contenido))  # si no hay cinco actividades, no se guarda

    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"Perfil de Puesto {puesto}.docx"
    for ruta in carpeta.glob("*.docx"):
        coincidencia = PATRON_NOMBRE.match(ruta.stem)
        if coincidencia and normalizar(coincidencia.group(1)) == normalizar(puesto) and ruta != destino:
            ruta.unlink()
    destino.write_bytes(contenido)
    return {"puesto": puesto, "archivo": destino.name, "actividades": actividades}


def extraer_actividades(origen: Union[Path, BinaryIO]) -> list[str]:
    """Las actividades numeradas que siguen al título «Cinco actividades principales»."""
    try:
        parrafos = Document(origen).paragraphs
    except Exception as error:  # archivo dañado o que no es Word
        raise ErrorPerfil("el archivo no es un documento Word (.docx) válido") from error

    inicio = next((i for i, p in enumerate(parrafos) if SECCION_ACTIVIDADES in normalizar(p.text)), None)
    if inicio is None:
        raise ErrorPerfil("no tiene la sección «Cinco actividades principales»")

    actividades = []
    for parrafo in parrafos[inicio + 1:]:
        texto = parrafo.text.strip()
        if not texto:
            if actividades:
                break
            continue
        enumerada = PATRON_ENUMERACION.match(texto)
        lista_de_word = parrafo._p.pPr is not None and parrafo._p.pPr.numPr is not None
        if enumerada and int(enumerada.group(1)) == len(actividades) + 1:
            actividades.append(enumerada.group(2).strip())
        elif lista_de_word and not enumerada:
            actividades.append(texto)
        else:  # siguiente título u otro contenido: termina la lista
            break

    if len(actividades) != ACTIVIDADES_REQUERIDAS:
        raise ErrorPerfil(f"la sección «Cinco actividades principales» tiene {len(actividades)} "
                          f"actividades numeradas; se requieren {ACTIVIDADES_REQUERIDAS}")
    return actividades


def relacionar(registros: list[dict], catalogo: dict[str, dict]) -> dict:
    """Agrupa a los trabajadores por puesto y asigna a cada puesto su perfil, si existe."""
    puestos: dict[str, dict] = {}
    sin_puesto = []
    for reg in registros:
        puesto = reg["valores"].get("puesto", "")
        trabajador = {"fila": reg["fila"], "nombre": reg["valores"].get("nombre") or "(sin nombre)"}
        if not puesto:
            sin_puesto.append(trabajador)
            continue
        grupo = puestos.setdefault(normalizar(puesto), {"puesto": puesto, "trabajadores": []})
        grupo["trabajadores"].append(trabajador)

    filas = []
    for clave, grupo in puestos.items():
        perfil = catalogo.get(clave)
        if perfil is None:
            estado, detalle = "pendiente", f"Perfil de puesto pendiente: {grupo['puesto']}"
        elif perfil["error"]:
            estado, detalle = "invalido", f"Perfil de {grupo['puesto']} no utilizable: {perfil['error']}"
        else:
            estado, detalle = "disponible", None
        filas.append({
            "clave": clave,
            "puesto": grupo["puesto"],
            "trabajadores": grupo["trabajadores"],
            "estado": estado,
            "detalle": detalle,
            "perfil": perfil["archivo"].name if perfil else None,
            "actividades": perfil["actividades"] if estado == "disponible" else [],
        })

    return {
        "puestos": filas,
        "sin_puesto": sin_puesto,
        "disponibles": [f["puesto"] for f in filas if f["estado"] == "disponible"],
        "pendientes": [f["puesto"] for f in filas if f["estado"] != "disponible"],
    }
