"""Motor de generación de contratos en Word (.docx).

Por cada trabajador:
1. llena la plantilla elegida (marcadores {{ campo }}, docxtpl) con sus datos,
   los del patrón y las cinco actividades de SU perfil;
2. agrega SU perfil de puesto completo como ANEXO UNO, en una sección nueva que
   conserva la página, los márgenes y el pie del perfil.

Si el puesto no tiene perfil, las actividades quedan como ____________ y no hay
anexo (no se inventan funciones). Cada contrato se arma desde cero (plantilla y
perfil nuevos), así que no se mezclan datos entre trabajadores. Cada generación
se guarda en una carpeta nueva «<razón social> <fecha y hora>».
"""
from __future__ import annotations

import re
import zipfile
from copy import deepcopy
from datetime import datetime
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docxcompose.composer import Composer
from docxtpl import DocxTemplate
from jinja2 import TemplateError

from servicios.lector_excel import normalizar
from servicios.mapeo_contrato import FIRMA_ALTA_IMSS, contexto_contrato, en_blanco, fecha_iso_con_letra
from servicios.perfiles import relacionar
from servicios.validador import FALTANTE

ATRIBUTOS_PAGINA = ("orientation", "page_width", "page_height", "left_margin", "right_margin",
                    "top_margin", "bottom_margin", "header_distance", "footer_distance", "gutter")


class ErrorGeneracion(Exception):
    """No se pueden generar los contratos (sin plantilla, plantilla inválida o sin trabajadores)."""


def generar_contratos(libro: dict, plantilla: Path, carpeta_salidas: Path, archivo_origen: str,
                      catalogo: dict[str, dict], fecha_firma: str = "") -> dict:
    """Genera un .docx por trabajador en una carpeta nueva dentro de carpeta_salidas."""
    if not plantilla.is_file():
        raise ErrorGeneracion(f"No se encontró la plantilla del contrato ({plantilla.name}).")
    hoja_t = libro["trabajadores"]
    if not hoja_t["registros"]:
        raise ErrorGeneracion("No hay trabajadores capturados; no se generó ningún contrato.")

    hoja_p = libro["patron"]
    valores_p = hoja_p["registros"][0]["valores"] if hoja_p and hoja_p["registros"] else {}

    relacion = relacionar(hoja_t["registros"], catalogo)
    con_perfil = {fila["clave"] for fila in relacion["puestos"] if fila["estado"] == "disponible"}

    base = plantilla.read_bytes()
    marcadores = revisar_plantilla(base)
    # El perfil se usa (actividades y ANEXO UNO) solo si la plantilla trae las actividades del puesto.
    con_anexo = any(m.startswith("actividad_") for m in marcadores)

    # Todo se genera en memoria y se escribe al final: si algo falla, no quedan archivos a medias.
    contratos, usados, detalle, sin_perfil = [], set(), [], []
    for reg in hoja_t["registros"]:
        clave = normalizar(reg["valores"].get("puesto", ""))
        perfil = catalogo[clave] if con_anexo and clave in con_perfil else None
        contexto, avisos = contexto_contrato(reg["valores"], valores_p,
                                             perfil["actividades"] if perfil else [], fecha_firma)

        doc = DocxTemplate(BytesIO(base))
        doc.render(contexto, autoescape=True)
        contrato = BytesIO()
        doc.save(contrato)
        contenido = anexar_perfil(contrato.getvalue(), perfil["archivo"]) if perfil else contrato.getvalue()

        nombre = _nombre_archivo(reg["valores"].get("nombre", ""), reg["fila"], usados)
        contratos.append((nombre, contenido))
        blancos = en_blanco(contexto, marcadores)
        if perfil:
            detalle.append(f"- {nombre} (fila {reg['fila']}) · perfil: {perfil['archivo'].name}")
        elif not con_anexo:
            detalle.append(f"- {nombre} (fila {reg['fila']})")
        else:
            puesto = reg["valores"].get("puesto") or "sin puesto capturado"
            detalle.append(f"- {nombre} (fila {reg['fila']}) · PERFIL DE PUESTO PENDIENTE: {puesto} "
                           "(sin actividades en la cláusula PRIMERA y sin ANEXO UNO)")
            sin_perfil.append({"fila": reg["fila"], "nombre": reg["valores"].get("nombre") or "(sin nombre)",
                               "puesto": reg["valores"].get("puesto") or ""})
        detalle.append("    Datos en blanco: " + (", ".join(blancos) + "." if blancos else "ninguno."))
        detalle += [f"    Aviso: {aviso}" for aviso in avisos]

    resumen = [
        "ORIENS | Generador Inteligente de Contratos",
        f"Fecha de generación: {datetime.now():%d/%m/%Y %H:%M}",
        f"Archivo de origen: {archivo_origen}",
        f"Plantilla: {plantilla.name}",
        "ANEXO UNO (perfil de puesto): " + ("sí" if con_anexo else "no; la plantilla no usa las actividades del perfil"),
        "Fecha de firma: " + ("la FECHA ALTA IMSS de cada trabajador" if fecha_firma == FIRMA_ALTA_IMSS
                              else fecha_iso_con_letra(fecha_firma) or "en blanco"),
        f"Contratos generados: {len(contratos)}",
        "",
        f"Los datos faltantes se escribieron como {FALTANTE} y deben completarse antes de firmar.",
        "",
        "CONTRATOS",
        *detalle,
    ]

    carpeta = _carpeta_lote(carpeta_salidas, valores_p.get("razon_social", ""))
    for nombre, contenido in contratos:
        (carpeta / nombre).write_bytes(contenido)
    (carpeta / "Resumen de generación.txt").write_text("\n".join(resumen) + "\n", encoding="utf-8")
    return {"carpeta": carpeta, "archivos": [nombre for nombre, _ in contratos], "sin_perfil": sin_perfil}


def revisar_plantilla(base: bytes) -> set[str]:
    """La plantilla debe abrir, tener marcadores y usar solo campos que el mapeo sabe llenar."""
    disponibles = set(contexto_contrato({}, {}, [])[0])
    try:
        usados = DocxTemplate(BytesIO(base)).get_undeclared_template_variables()
    except TemplateError as error:
        raise ErrorGeneracion(f"La plantilla tiene un marcador mal escrito: {error}.") from error
    except (zipfile.BadZipFile, KeyError) as error:
        raise ErrorGeneracion("La plantilla no es un documento Word (.docx) válido.") from error

    if not usados:
        if re.search(r"_{3,}", _texto_de_docx(base)):
            raise ErrorGeneracion(
                "La plantilla todavía tiene líneas en blanco (____) y ningún marcador {{ campo }}: la aplicación "
                "no sabe qué dato va en cada línea. Sustituye cada línea por su marcador (ver «Cómo preparar una "
                "plantilla nueva») y vuelve a subirla.")
        raise ErrorGeneracion("La plantilla no contiene marcadores {{ campo }}; no hay dónde insertar los datos.")
    desconocidos = sorted(usados - disponibles)
    if desconocidos:
        lista = ", ".join("{{ " + d + " }}" for d in desconocidos)
        raise ErrorGeneracion(f"La plantilla usa marcadores que la aplicación no conoce: {lista}.")
    return usados


def _texto_de_docx(base: bytes) -> str:
    doc = Document(BytesIO(base))
    return "\n".join(p.text for p in doc.paragraphs)


def anexar_perfil(contrato: bytes, perfil: Path) -> bytes:
    """Agrega el perfil completo como ANEXO UNO en una sección nueva (página, márgenes y pie del perfil)."""
    doc = Document(BytesIO(contrato))
    anexo = Document(perfil)
    origen = anexo.sections[-1]

    seccion = doc.add_section(WD_SECTION.NEW_PAGE)
    for atributo in ATRIBUTOS_PAGINA:
        valor = getattr(origen, atributo)
        if valor is not None:
            setattr(seccion, atributo, valor)
    for parte in ("header", "footer"):
        _copiar_encabezado_o_pie(getattr(origen, parte), getattr(seccion, parte))

    Composer(doc).append(anexo)
    salida = BytesIO()
    doc.save(salida)
    return salida.getvalue()


def _copiar_encabezado_o_pie(fuente, destino) -> None:
    """Copia el encabezado o pie del perfil a la sección del anexo (sin heredar el del contrato)."""
    destino.is_linked_to_previous = False
    elemento = destino._element
    for hijo in list(elemento):
        elemento.remove(hijo)
    contenido = [] if fuente.is_linked_to_previous else list(fuente._element)
    # Solo texto: si el pie trae imágenes u otras referencias, se deja en blanco para no romper el archivo.
    if any(e.xpath(".//@r:id | .//@r:embed") for e in contenido):
        contenido = []
    for hijo in contenido:
        elemento.append(deepcopy(hijo))
    if not len(elemento):
        elemento.append(elemento.makeelement(qn("w:p"), {}))


def _limpiar_nombre(texto: str, largo: int, final: bool = True) -> str:
    """Espacios normales, sin guiones bajos ni caracteres que no admite un nombre de archivo.

    Si el texto va al final del nombre (final=True) se quitan los puntos finales,
    que algunos sistemas no admiten; si después sigue la fecha, se conservan («S.A. de C.V.»).
    """
    limpio = re.sub(r'[\\/:*?"<>|_\x00-\x1f]+', " ", texto)
    limpio = re.sub(r"\s+", " ", limpio).strip()[:largo].strip()
    return limpio.strip(" .") if final else limpio


def _nombre_archivo(nombre: str, fila: int, usados: set[str]) -> str:
    """'Contrato Juan Pérez García.docx'."""
    limpio = _limpiar_nombre(nombre, 80)
    base = f"Contrato {limpio}" if limpio else f"Contrato sin nombre (fila {fila})"
    candidato, n = base, 2
    while candidato.lower() in usados:  # nombres repetidos: "Contrato X (2)"
        candidato, n = f"{base} ({n})", n + 1
    usados.add(candidato.lower())
    return f"{candidato}.docx"


def _carpeta_lote(carpeta_salidas: Path, razon_social: str) -> Path:
    """Carpeta nueva por generación: 'Servicios Administrativos Horizonte, S.A. de C.V 2026-09-12 18.30'."""
    razon = _limpiar_nombre(razon_social, 60, final=False) or "Sin patrón"
    base = carpeta_salidas / f"{razon} {datetime.now():%Y-%m-%d %H.%M}"
    carpeta, n = base, 2
    while carpeta.exists():
        carpeta, n = base.with_name(f"{base.name} ({n})"), n + 1
    carpeta.mkdir(parents=True)
    return carpeta
