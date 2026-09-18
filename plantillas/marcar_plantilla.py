"""Crea la copia marcada de un formato de contrato por tiempo indeterminado.

Sustituye solo los espacios variables (líneas de guiones bajos y campos de
combinación) por marcadores {{ campo }}; el texto jurídico queda intacto. El
original no se modifica. Al terminar verifica que no cambió nada más.

Se marcan todos los espacios variables; la línea «-» del bloque de firma de la
empresa se conserva tal cual.

Si la copia marcada ya existe, el script no la rehace (se perderían los cambios que
se le hayan hecho en Word) salvo que se pida expresamente:

Formatos conocidos (ver FORMATOS): «oficial» (predeterminado) y «regularizacion».

Uso:  .venv/bin/python plantillas/marcar_plantilla.py [formato] [--forzar]
"""
import difflib
import re
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

CARPETA = Path(__file__).resolve().parent

# (texto que identifica el párrafo, {n.º de línea de guiones bajos: texto que la sustituye}).
# Una clave (a, b) sustituye desde la línea a hasta la b, con el texto que haya entre ellas.
MARCAS_OFICIAL = [
    ("QUE CELEBRAN POR UNA PARTE", {1: "{{ patron_razon_social }}", 2: "{{ nombre }}"}),
    # La línea del domicilio va pegada a «Con el Registro…»: se agrega el punto y el espacio.
    ("Tiene su domicilio en", {1: "{{ patron_domicilio }}. ", 2: "{{ patron_rfc }}"}),
    ("actividad principal es", {1: "{{ patron_actividad }}"}),
    ("originario de", {1: "{{ lugar_nacimiento }}", 2: "{{ fecha_nacimiento_letra }}", 3: "{{ sexo }}",
                       4: "{{ estado_civil }}", 5: "{{ domicilio_calle }}", 6: "{{ domicilio_numero }}",
                       7: "{{ domicilio_colonia }}", 8: "{{ rfc }}", 9: "{{ curp }}",
                       10: "{{ seguro_social }}", 11: "{{ credencial_elector }} "}),
    ("adiestramiento, capacitación y experiencia", {1: "{{ experiencia }}"}),
    # Cláusula CUARTA: lunes a viernes (entrada, inicio de comida, fin de comida, salida) y sábado.
    ("de lunes a viernes de las", {1: "{{ lv_entrada }}", 2: "{{ comida_inicio }}", 3: "{{ comida_fin }}",
                                   4: "{{ lv_salida }}", 5: "{{ sabado_entrada }}", 6: "{{ sabado_salida }}"}),
    ("PRIMERA. - Por virtud", {1: "{{ puesto }}", 2: "{{ actividad_1 }}"}),
    ("2.  _", {1: "{{ actividad_2 }}"}),
    ("3. _", {1: "{{ actividad_3 }}"}),
    ("4. _", {1: "{{ actividad_4 }}"}),
    ("5. _", {1: "{{ actividad_5 }}"}),
    ("interrumpida durante 60 minutos", {1: "{{ comida_inicio }}", 2: "{{ comida_fin }}"}),
    ("salario diario de $", {1: "{{ salario_diario }}", 2: "{{ salario_letra }}"}),
    ("siendo este _", {1: "{{ correo }}"}),
    ("fecha de ingreso del trabajador fue el día", {1: "{{ fecha_ingreso_letra }}"}),
    ("testigos que firman al calce", {1: "{{ fecha_firma_letra }}"}),
    ("Representante legal de", {1: "{{ patron_razon_social }}"}),
]

# Variante de regularización (relaciones de trabajo ya vigentes): «Mexicano de nacimiento,
# ___ (edad)», domicilio completo, fecha de inicio de la relación y lugar de firma = domicilio del patrón.
MARCAS_REGULARIZACION = [
    ("QUE CELEBRAN POR UNA PARTE", {1: "{{ patron_razon_social }}", 2: "{{ nombre }}"}),
    ("Tiene su domicilio en", {1: "{{ patron_domicilio }}. ", 2: "{{ patron_rfc }}"}),
    ("actividad principal es", {1: "{{ patron_actividad }}"}),
    # «con domicilio en ___ número ___, Colonia ___»: el Excel trae el domicilio en una sola columna.
    ("Mexicano de nacimiento", {1: "{{ edad }}", 2: "{{ estado_civil }}", (3, 5): "{{ domicilio }}",
                                6: "{{ rfc }}", 7: "{{ curp }}", 8: "{{ seguro_social }}"}),
    ("adiestramiento, capacitación y experiencia", {1: "{{ experiencia }}"}),
    ("Declaran ambos contratantes con fecha", {1: "{{ fecha_ingreso_letra }}"}),
    ("PRIMERA. - Por virtud", {1: "{{ puesto }}", 2: "{{ actividad_1 }}"}),
    ("2.  _", {1: "{{ actividad_2 }}"}),
    ("3. _", {1: "{{ actividad_3 }}"}),
    ("4. _", {1: "{{ actividad_4 }}"}),
    ("5. _", {1: "{{ actividad_5 }}"}),
    ("salario diario de $", {1: "{{ salario_diario }}", 2: "{{ salario_letra }}"}),
    ("siendo este _", {1: "{{ correo }}"}),
    ("fecha de ingreso del trabajador fue el día", {1: "{{ fecha_ingreso_letra }}"}),
    # El contrato se firma en el domicilio del patrón.
    ("testigos que firman al calce", {1: "{{ fecha_firma_letra }}, en {{ lugar_firma }}"}),
    ("Representante legal de", {1: "{{ patron_razon_social }}"}),
]

# Cambios de redacción autorizados por el abogado responsable (patrón → texto nuevo), además de las marcas:
# jornada solo de lunes a viernes (el sábado se reparte conforme al art. 59 LFT, como ya dice el contrato)
# descanso para comida con la duración que indica el Excel, sin horas fijas, y horas semanales según el horario.
REEMPLAZOS_REGULARIZACION = [
    # Sin credencial de elector en los datos del cliente: la declaración termina con el número del IMSS.
    (r"\s*identificándose con credencial para votar con fotografía numero _+\s*expedida a su favor por el "
     r"Instituto Nacional Electoral INE\.", "."),
    # Horas semanales calculadas con el horario del Excel (lunes a viernes, sin el descanso para comida).
    (r"será de 48 horas a la semana", "será de {{ horas_semana }} horas a la semana"),
    (r"de lunes a viernes de las _+ horas\s+a las _+ horas y de las _+ a las _+ horas y los días sábados "
     r"de las _+ horas a las _+ horas\s*$",
     "de lunes a viernes de las {{ lv_entrada }} horas a las {{ lv_salida }} horas."),
    (r"interrumpida durante 60 minutos, todos\s+los días laborables,\s+comprendiendo de las _+\s+horas "
     r"a las _+\s+horas, durante",
     "interrumpida durante {{ comida_duracion }}, todos los días laborables, durante"),
]

FORMATOS = {
    "oficial": {
        "original": CARPETA / "FORMATO CONTRATO TIEMPO INDETERMINADO (original).docx",
        "destino": CARPETA / "Contrato tiempo indeterminado (marcado).docx",
        "marcas": MARCAS_OFICIAL,
        "reemplazos": [],
        "firma_con_nombre": True,
        "representante_en_firma": False,
    },
    "regularizacion": {
        "original": CARPETA / "FORMATO CONTRATO TIEMPO INDETERMINADO REGULARIZACION (original).docx",
        "destino": CARPETA / "Contrato tiempo indeterminado regularización (marcado).docx",
        "marcas": MARCAS_REGULARIZACION,
        "reemplazos": REEMPLAZOS_REGULARIZACION,
        "firma_con_nombre": False,  # el formato no trae «(NOMBRE COMPLETO DEL TRABAJADOR )»
        # Bajo la línea de firma de la empresa: nombre del representante legal y, debajo, la razón social.
        "representante_en_firma": True,
    },
}


def todos_los_parrafos(doc):
    yield from doc.paragraphs
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                yield from celda.paragraphs


def reemplazar_linea(parrafo, numero, texto_nuevo: str) -> None:
    """Sustituye la n-ésima línea de guiones bajos del párrafo, aunque esté repartida en varios runs.

    Con numero = (a, b) sustituye desde la línea a hasta la b, incluido el texto entre ellas.
    """
    primera, ultima = numero if isinstance(numero, tuple) else (numero, numero)
    lineas = list(re.finditer(r"_+", "".join(run.text for run in parrafo.runs)))
    reemplazar_tramo(parrafo, lineas[primera - 1].start(), lineas[ultima - 1].end(), texto_nuevo)


def reemplazar_texto(parrafo, patron: str, texto_nuevo: str) -> None:
    """Sustituye el texto que coincide con el patrón, aunque esté repartido en varios runs."""
    coincidencia = re.search(patron, "".join(run.text for run in parrafo.runs))
    if not coincidencia:
        sys.exit(f"ERROR: no se encontró el texto a sustituir: {patron!r}")
    reemplazar_tramo(parrafo, coincidencia.start(), coincidencia.end(), texto_nuevo)


def reemplazar_tramo(parrafo, inicio: int, fin: int, texto_nuevo: str) -> None:
    """Sustituye los caracteres [inicio, fin) del párrafo; el texto nuevo toma el formato del primer run."""
    runs = parrafo.runs
    textos = [run.text for run in runs]
    posicion, colocado = 0, False
    for run, texto in zip(runs, textos):
        a, b = posicion, posicion + len(texto)
        posicion = b
        if b <= inicio or a >= fin:
            continue
        izquierda, derecha = texto[:max(0, inicio - a)], texto[min(len(texto), fin - a):]
        run.text = izquierda + ("" if colocado else texto_nuevo) + derecha
        colocado = True


def reemplazar_campo_combinacion(parrafo, campo: str, texto_nuevo: str) -> bool:
    """Sustituye un MERGEFIELD por texto, conservando el formato del resultado visible."""
    instrucciones = [r for r in parrafo.runs if r._r.find(qn("w:instrText")) is not None]
    if not any(campo in r._r.find(qn("w:instrText")).text for r in instrucciones):
        return False
    for run in parrafo.runs:
        if run._r.find(qn("w:fldChar")) is not None or run._r.find(qn("w:instrText")) is not None:
            run._r.getparent().remove(run._r)
        elif run.text:
            run.text = texto_nuevo
    return True


def poner_representante(parrafos) -> None:
    """Escribe {{ patron_representante }} en el renglón vacío que precede a «Representante legal de …»."""
    indice = next((i for i, p in enumerate(parrafos) if p.text.startswith("Representante legal de")), None)
    if indice is None or parrafos[indice - 1].text.strip():
        sys.exit("ERROR: no hay un renglón vacío antes de «Representante legal de» para el nombre del representante.")
    modelo = parrafos[indice].runs[0]._r
    run = parrafos[indice - 1].add_run("{{ patron_representante }}")
    if modelo.rPr is not None:
        run._r.insert(0, deepcopy(modelo.rPr))


def marcar(formato: dict) -> None:
    doc = Document(formato["original"])
    parrafos = list(todos_los_parrafos(doc))
    for ancla, lineas in formato["marcas"]:
        encontrados = [p for p in parrafos if ancla in p.text]
        if len(encontrados) != 1:
            sys.exit(f"ERROR: «{ancla}» aparece {len(encontrados)} veces; se esperaba 1.")
        # De atrás hacia adelante para que la numeración de las líneas no cambie.
        for numero in sorted(lineas, key=lambda n: n if isinstance(n, tuple) else (n, n), reverse=True):
            reemplazar_linea(encontrados[0], numero, lineas[numero])

    for patron, texto_nuevo in formato["reemplazos"]:
        encontrados = [p for p in parrafos if re.search(patron, p.text)]
        if len(encontrados) != 1:
            sys.exit(f"ERROR: el texto {patron!r} aparece {len(encontrados)} veces; se esperaba 1.")
        reemplazar_texto(encontrados[0], patron, texto_nuevo)

    # Campo de combinación «BENEFICIARIOS» en la tabla de la cláusula VIGÉSIMA.
    if not any(reemplazar_campo_combinacion(p, "BENEFICIARIOS", "{{ beneficiarios }}") for p in parrafos):
        sys.exit("ERROR: no se encontró el campo de combinación BENEFICIARIOS.")

    if formato["representante_en_firma"]:
        poner_representante(parrafos)

    # «(NOMBRE COMPLETO DEL TRABAJADOR )» en el bloque de firma.
    if not formato["firma_con_nombre"]:
        doc.save(formato["destino"])
        print(f"Creada: {formato['destino'].name}")
        return
    firma = [p for p in parrafos if p.text.strip() == "(NOMBRE COMPLETO DEL TRABAJADOR )"]
    if len(firma) != 1:
        sys.exit("ERROR: no se encontró «(NOMBRE COMPLETO DEL TRABAJADOR )» en el bloque de firma.")
    firma[0].runs[0].text = "{{ nombre }}"
    for run in firma[0].runs[1:]:
        run.text = ""

    doc.save(formato["destino"])
    print(f"Creada: {formato['destino'].name}")


def verificar(formato: dict) -> None:
    """Compara original y copia: solo deben cambiar los espacios variables."""
    originales = [p.text for p in todos_los_parrafos(Document(formato["original"]))]
    marcados = [p.text for p in todos_los_parrafos(Document(formato["destino"]))]
    if len(originales) != len(marcados):
        sys.exit("ERROR: cambió el número de párrafos.")
    cambios = 0
    for antes, despues in zip(originales, marcados):
        if antes == despues:
            continue
        cambios += 1
        # Líneas y marcadores se igualan a «¤»; lo que quede distinto se muestra.
        a = re.sub(r"_+", "¤", antes)
        b = re.sub(r"_+", "¤", re.sub(r"\{\{ \w+ \}\}", "¤", despues))
        diferencias = [f"{op}: {a[i1:i2]!r} → {b[j1:j2]!r}"
                       for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b).get_opcodes() if op != "equal"]
        print(f"  {despues[:90]!r}")
        print(f"      {'; '.join(diferencias) if diferencias else 'solo espacios variables'}")
    print(f"Párrafos modificados: {cambios} de {len(originales)}; el resto quedó idéntico.")


if __name__ == "__main__":
    nombres = [a for a in sys.argv[1:] if not a.startswith("--")] or ["oficial"]
    if nombres[0] not in FORMATOS:
        sys.exit(f"Formato desconocido «{nombres[0]}». Disponibles: {', '.join(FORMATOS)}.")
    elegido = FORMATOS[nombres[0]]
    if elegido["destino"].exists() and "--forzar" not in sys.argv:
        sys.exit(f"Ya existe «{elegido['destino'].name}». Si la editaste en Word, rehacerla borraría esos cambios.\n"
                 f"Para rehacerla desde el original: .venv/bin/python plantillas/marcar_plantilla.py {nombres[0]} --forzar")
    marcar(elegido)
    verificar(elegido)
