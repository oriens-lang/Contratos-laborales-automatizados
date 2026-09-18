"""Perfiles de puesto: relación entre el PUESTO del Excel y su perfil en Word.

Cada perfil se guarda en la carpeta «perfiles de puesto» con el nombre
«Perfil de Puesto <nombre del puesto>.docx». El puesto se compara sin distinguir
mayúsculas, acentos ni espacios repetidos («Diseñadora» = «DISEÑADORA»).

De cada perfil se toman, sin modificarlas, las actividades de la sección o tabla
«ACTIVIDADES PRINCIPALES» (también «Cinco actividades principales») o, si no existe,
de la tabla «FUNCIONES DEL PUESTO». Las «ACTIVIDADES COMPLEMENTARIAS» no se usan en el
contrato. A la cláusula PRIMERA van las cinco primeras, en el orden del perfil;
el perfil completo va como ANEXO UNO. Solo se quita el número de lista («1. »)
porque la cláusula PRIMERA del contrato ya numera los cinco espacios.
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
SECCION_ACTIVIDADES = "ACTIVIDADES PRINCIPALES"  # incluye «Cinco actividades principales»
# Título (primera fila) de la tabla de actividades, en orden de preferencia.
TITULOS_TABLA = ("ACTIVIDADES PRINCIPALES", "FUNCIONES PRINCIPALES", "FUNCIONES", "ACTIVIDADES")
EXCLUIDAS = "COMPLEMENTARIA"
SUBTITULOS_PRINCIPALES = {"ACTIVIDADES PRINCIPALES", "FUNCIONES PRINCIPALES", "CINCO ACTIVIDADES PRINCIPALES"}  # «ACTIVIDADES/FUNCIONES COMPLEMENTARIAS» no van a la cláusula PRIMERA
# Datos que no conviene anexar al contrato: se avisan al subir el perfil.
# (patrón sobre el texto normalizado, descripción). El rango de edad solo si fija años («Indistinto» no).
DATOS_SENSIBLES = [(r"\bSUELDO", "sueldos"), (r"\bSALARIO", "salarios"),
                   (r"RANGO DE EDAD\s+(DE\s+)?\d", "rango de edad")]


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
            todas, error = extraer_actividades(ruta), None
        except ErrorPerfil as problema:
            todas, error = [], str(problema)
        catalogo[normalizar(puesto)] = {
            "puesto": puesto, "archivo": ruta, "actividades": todas[:ACTIVIDADES_REQUERIDAS],
            "total_actividades": len(todas), "error": error,
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
    todas = extraer_actividades(BytesIO(contenido))  # si no hay al menos cinco actividades, no se guarda
    advertencias = datos_sensibles(BytesIO(contenido))

    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"Perfil de Puesto {puesto}.docx"
    for ruta in carpeta.glob("*.docx"):
        coincidencia = PATRON_NOMBRE.match(ruta.stem)
        if coincidencia and normalizar(coincidencia.group(1)) == normalizar(puesto) and ruta != destino:
            ruta.unlink()
    destino.write_bytes(contenido)
    return {"puesto": puesto, "archivo": destino.name, "actividades": todas[:ACTIVIDADES_REQUERIDAS],
            "total_actividades": len(todas), "advertencias": advertencias}


def datos_sensibles(origen: Union[Path, BinaryIO]) -> list[str]:
    """Datos del perfil que pasarían al contrato en el ANEXO UNO y conviene revisar (sueldo, edad)."""
    documento = Document(origen)
    textos = [p.text for p in documento.paragraphs]
    textos += [c.text for t in documento.tables for fila in t.rows for c in fila.cells]
    completo = normalizar(" ".join(textos))
    return [descripcion for patron, descripcion in DATOS_SENSIBLES if re.search(patron, completo)]


def extraer_actividades(origen: Union[Path, BinaryIO]) -> list[str]:
    """Todas las actividades del perfil, en su orden (se exigen al menos cinco)."""
    try:
        documento = Document(origen)
    except Exception as error:  # archivo dañado o que no es Word
        raise ErrorPerfil("el archivo no es un documento Word (.docx) válido") from error

    parrafos = documento.paragraphs
    # Título de sección: renglón corto (no una frase del texto) que menciona las actividades principales.
    inicio = next((i for i, p in enumerate(parrafos)
                   if SECCION_ACTIVIDADES in normalizar(p.text) and len(p.text.strip()) <= 60), None)
    if inicio is None:
        actividades = _actividades_de_tabla(documento)
        if actividades is None:
            raise ErrorPerfil("no se encontraron sus actividades: debe tener un apartado «ACTIVIDADES "
                              "PRINCIPALES» o una tabla «FUNCIONES DEL PUESTO»")
        return _exigir_minimo(actividades, "la tabla de actividades")

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

    return _exigir_minimo(actividades, "la sección «Cinco actividades principales»")


def _actividades_de_tabla(documento) -> list[str] | None:
    """Actividades de las tablas del perfil.

    1.º Las filas bajo un subtítulo «ACTIVIDADES PRINCIPALES» (p. ej., dentro de la tabla
        «VI. ACTIVIDADES DEL PUESTO»), hasta «ACTIVIDADES COMPLEMENTARIAS» u otro subtítulo.
    2.º Si no hay ese subtítulo, las filas de la tabla cuyo título es «FUNCIONES DEL PUESTO»,
        «ACTIVIDADES…», etc., sin las complementarias.
    """
    tablas = [_filas(t) for t in documento.tables]
    for filas in tablas:
        for n, fila in enumerate(filas):
            if _titulo(fila) in SUBTITULOS_PRINCIPALES:
                return _lista(filas[n + 1:])
    candidatas = []
    for filas in tablas:
        if not filas:
            continue
        titulo = _titulo(filas[0])
        prioridad = next((n for n, t in enumerate(TITULOS_TABLA) if titulo.startswith(t)), None)
        if prioridad is not None and EXCLUIDAS not in titulo:
            candidatas.append((prioridad, filas))
    if not candidatas:
        return None
    return _lista(min(candidatas, key=lambda c: c[0])[1][1:])


def _filas(tabla) -> list[str]:
    """Texto de cada fila (las celdas combinadas se repiten en python-docx: se toman una vez)."""
    filas = []
    for fila in tabla.rows:
        textos = dict.fromkeys(c.text.strip() for c in fila.cells if c.text.strip())
        filas.append(" ".join(textos))
    return filas


def _titulo(texto: str) -> str:
    """Texto normalizado sin el numeral romano del apartado: «VI.   ACTIVIDADES DEL PUESTO» → «ACTIVIDADES DEL PUESTO»."""
    return re.sub(r"^[IVXLC]+\s+", "", normalizar(texto))


def _es_subtitulo(texto: str) -> bool:
    """Renglón corto en mayúsculas y sin numerar, p. ej. «ACTIVIDADES COMPLEMENTARIAS»."""
    return texto.isupper() and len(texto) <= 60 and not PATRON_ENUMERACION.match(texto)


def _lista(filas: list[str]) -> list[str]:
    actividades = []
    for texto in filas:
        if not texto:
            continue
        if EXCLUIDAS in normalizar(texto) or _es_subtitulo(texto):
            break
        enumerada = PATRON_ENUMERACION.match(texto)
        actividades.append(enumerada.group(2).strip() if enumerada else texto)
    return actividades


def _exigir_minimo(actividades: list[str], donde: str) -> list[str]:
    if len(actividades) < ACTIVIDADES_REQUERIDAS:
        raise ErrorPerfil(f"{donde} tiene {len(actividades)} actividades; se requieren al menos "
                          f"{ACTIVIDADES_REQUERIDAS}")
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
