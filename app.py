"""ORIENS | Generador Inteligente de Contratos — servidor web local.

Ejecutar:  .venv/bin/python app.py   →   http://127.0.0.1:5050
"""
import subprocess
import sys

from flask import Flask, jsonify, render_template, request

from config import BASE_DIR, Config
from servicios.archivos import revisar_archivo_docx, revisar_archivo_xlsx
from servicios.generador import ErrorGeneracion, generar_contratos
from servicios.lector_excel import ErrorLectura, leer_libro
from servicios.perfiles import ErrorPerfil, cargar_catalogo, guardar_perfil
from servicios.validador import validar_libro

app = Flask(__name__)
app.config.from_object(Config)


class ArchivoInvalido(Exception):
    """No se recibió un .xlsx aceptable."""


def _recibir_xlsx() -> tuple[str, bytes]:
    """Nombre y contenido del Excel enviado. Se procesa en memoria: no se guarda en disco."""
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        raise ArchivoInvalido("No se recibió ningún archivo.")
    contenido = archivo.read()
    errores = revisar_archivo_xlsx(archivo.filename, contenido)
    if errores:
        raise ArchivoInvalido(" ".join(errores))
    return archivo.filename, contenido


@app.get("/")
def inicio():
    return render_template("index.html", limite_mb=Config.LIMITE_ARCHIVO_MB)


@app.post("/api/validar")
def validar():
    nombre, contenido = _recibir_xlsx()
    libro = leer_libro(contenido)
    return jsonify(
        ok=True,
        archivo={"nombre": nombre, "tamano_bytes": len(contenido)},
        plantilla=Config.PLANTILLA_CONTRATO.name,
        **validar_libro(libro, cargar_catalogo(Config.CARPETA_PERFILES)),
    )


@app.post("/api/generar")
def generar():
    nombre, contenido = _recibir_xlsx()
    libro = leer_libro(contenido)
    lote = generar_contratos(libro, Config.PLANTILLA_CONTRATO, Config.CARPETA_SALIDAS, nombre,
                             cargar_catalogo(Config.CARPETA_PERFILES))
    return jsonify(
        ok=True,
        total=len(lote["archivos"]),
        carpeta=str(lote["carpeta"].relative_to(BASE_DIR)),
        archivos=lote["archivos"],
        sin_perfil=lote["sin_perfil"],
    )


@app.post("/api/perfiles")
def subir_perfil():
    """Recibe el perfil (.docx) de un puesto y lo guarda en «perfiles de puesto»."""
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        raise ArchivoInvalido("No se recibió ningún archivo.")
    contenido = archivo.read()
    errores = revisar_archivo_docx(archivo.filename, contenido)
    if errores:
        raise ArchivoInvalido(" ".join(errores))
    try:
        perfil = guardar_perfil(Config.CARPETA_PERFILES, request.form.get("puesto", ""), contenido)
    except ErrorPerfil as error:
        return jsonify(ok=False, error=f"No se pudo usar el perfil: {error}."), 422
    return jsonify(ok=True, **perfil)


@app.post("/api/salidas/abrir")
def abrir_salida():
    """Abre en Finder la carpeta «contratos generados»."""
    if not Config.CARPETA_SALIDAS.is_dir():
        return jsonify(ok=False, error="Todavía no se ha generado ningún contrato."), 404
    if sys.platform != "darwin":
        return jsonify(ok=False, error="Abrir la carpeta solo está disponible en macOS."), 501
    subprocess.run(["open", str(Config.CARPETA_SALIDAS)], check=False)
    return jsonify(ok=True)


@app.errorhandler(ArchivoInvalido)
def archivo_invalido(error):
    return jsonify(ok=False, error=str(error)), 400


@app.errorhandler(ErrorLectura)
@app.errorhandler(ErrorGeneracion)
def error_de_proceso(error):
    return jsonify(ok=False, error=str(error)), 422


@app.errorhandler(413)
def archivo_demasiado_grande(_error):
    return jsonify(
        ok=False, error=f"El archivo excede el límite de {Config.LIMITE_ARCHIVO_MB} MB."
    ), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=Config.PUERTO, debug=True)
