"""Crea la copia marcada de la plantilla oficial del contrato por tiempo indeterminado.

Sustituye solo los espacios variables (líneas de guiones bajos y campos de
combinación) por marcadores {{ campo }}; el texto jurídico queda intacto. El
original no se modifica. Al terminar verifica que no cambió nada más.

Se marcan todos los espacios variables; la línea «-» del bloque de firma de la
empresa se conserva tal cual.

Si la copia marcada ya existe, el script no la rehace (se perderían los cambios que
se le hayan hecho en Word) salvo que se pida expresamente:

Uso:  .venv/bin/python plantillas/marcar_plantilla.py [--forzar]
"""
import difflib
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

CARPETA = Path(__file__).resolve().parent
ORIGINAL = CARPETA / "FORMATO CONTRATO TIEMPO INDETERMINADO (original).docx"
DESTINO = CARPETA / "Contrato tiempo indeterminado (marcado).docx"

# (texto que identifica el párrafo, {n.º de línea de guiones bajos: texto que la sustituye})
MARCAS = [
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


def todos_los_parrafos(doc):
    yield from doc.paragraphs
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                yield from celda.paragraphs


def reemplazar_linea(parrafo, numero: int, texto_nuevo: str) -> None:
    """Sustituye la n-ésima línea de guiones bajos del párrafo, aunque esté repartida en varios runs."""
    runs = parrafo.runs
    textos = [run.text for run in runs]
    lineas = list(re.finditer(r"_+", "".join(textos)))
    inicio, fin = lineas[numero - 1].span()
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


def marcar() -> None:
    doc = Document(ORIGINAL)
    parrafos = list(todos_los_parrafos(doc))
    for ancla, lineas in MARCAS:
        encontrados = [p for p in parrafos if ancla in p.text]
        if len(encontrados) != 1:
            sys.exit(f"ERROR: «{ancla}» aparece {len(encontrados)} veces; se esperaba 1.")
        # De atrás hacia adelante para que la numeración de las líneas no cambie.
        for numero in sorted(lineas, reverse=True):
            reemplazar_linea(encontrados[0], numero, lineas[numero])

    # Campo de combinación «BENEFICIARIOS» en la tabla de la cláusula VIGÉSIMA.
    if not any(reemplazar_campo_combinacion(p, "BENEFICIARIOS", "{{ beneficiarios }}") for p in parrafos):
        sys.exit("ERROR: no se encontró el campo de combinación BENEFICIARIOS.")

    # «(NOMBRE COMPLETO DEL TRABAJADOR )» en el bloque de firma.
    firma = [p for p in parrafos if p.text.strip() == "(NOMBRE COMPLETO DEL TRABAJADOR )"]
    if len(firma) != 1:
        sys.exit("ERROR: no se encontró «(NOMBRE COMPLETO DEL TRABAJADOR )» en el bloque de firma.")
    firma[0].runs[0].text = "{{ nombre }}"
    for run in firma[0].runs[1:]:
        run.text = ""

    doc.save(DESTINO)
    print(f"Creada: {DESTINO.name}")


def verificar() -> None:
    """Compara original y copia: solo deben cambiar los espacios variables."""
    originales = [p.text for p in todos_los_parrafos(Document(ORIGINAL))]
    marcados = [p.text for p in todos_los_parrafos(Document(DESTINO))]
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
    if DESTINO.exists() and "--forzar" not in sys.argv:
        sys.exit(f"Ya existe «{DESTINO.name}». Si la editaste en Word, rehacerla borraría esos cambios.\n"
                 "Para rehacerla desde el original: .venv/bin/python plantillas/marcar_plantilla.py --forzar")
    marcar()
    verificar()
