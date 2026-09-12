"""ORIENS | Generador Inteligente de Contratos — servidor web del despacho.

Ejecutar:  .venv/bin/python app.py   →   http://127.0.0.1:5050
También queda accesible en la red del despacho, protegido con contraseña.
"""
from __future__ import annotations

import io
import os
import socket
import subprocess
import sys
import zipfile

from flask import Flask, jsonify, redirect, render_template, request, send_file, session, url_for

from config import BASE_DIR, Config
from servicios import acceso as control
from servicios import plantillas
from servicios.archivos import revisar_archivo_docx, revisar_archivo_xlsx
from servicios.generador import ErrorGeneracion, generar_contratos
from servicios.lector_excel import ErrorLectura, leer_libro
from servicios.mapeo_contrato import MARCADORES
from servicios.perfiles import ErrorPerfil, cargar_catalogo, guardar_perfil
from servicios.validador import validar_libro

app = Flask(__name__)
app.config.from_object(Config)
app.secret_key = control.llave_de_sesion(Config.ARCHIVO_ACCESO)

RUTAS_ABIERTAS = {"entrar", "static"}


class ArchivoInvalido(Exception):
    """No se recibió un archivo aceptable."""


def _es_local() -> bool:
    """La petición viene de la computadora donde corre la aplicación."""
    return request.remote_addr in ("127.0.0.1", "::1")


def _direccion_en_red() -> str | None:
    """Dirección para entrar desde otra computadora de la oficina (p. ej. http://192.168.1.20:5050/)."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as prueba:
            prueba.connect(("10.254.254.254", 1))  # no envía datos; solo elige la interfaz de red
            return f"http://{prueba.getsockname()[0]}:{Config.PUERTO}/"
    except OSError:
        return None


def _recibir(tipo: str) -> tuple[str, bytes]:
    """Nombre y contenido del archivo enviado (xlsx o docx). Se procesa en memoria."""
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        raise ArchivoInvalido("No se recibió ningún archivo.")
    contenido = archivo.read()
    revisar = revisar_archivo_xlsx if tipo == "xlsx" else revisar_archivo_docx
    errores = revisar(archivo.filename, contenido)
    if errores:
        raise ArchivoInvalido(" ".join(errores))
    return archivo.filename, contenido


def _carpeta_de_lote(lote: str):
    carpeta = (Config.CARPETA_SALIDAS / lote).resolve()
    if carpeta.parent != Config.CARPETA_SALIDAS.resolve() or not carpeta.is_dir():
        return None
    return carpeta


# ---------- Acceso ----------

@app.before_request
def exigir_acceso():
    if request.endpoint in RUTAS_ABIERTAS or session.get("autorizado"):
        return None
    if request.path.startswith("/api/"):
        return jsonify(ok=False, error="La sesión expiró; vuelve a entrar."), 401
    return redirect(url_for("entrar"))


@app.route("/entrar", methods=["GET", "POST"])
def entrar():
    primera_vez = not control.hay_contrasena(Config.ARCHIVO_ACCESO)
    error = None
    if request.method == "POST":
        contrasena = request.form.get("contrasena", "")
        if primera_vez:
            if not _es_local():
                error = "La contraseña se crea en la computadora donde corre la aplicación."
            elif len(contrasena) < control.LONGITUD_MINIMA:
                error = f"La contraseña debe tener al menos {control.LONGITUD_MINIMA} caracteres."
            elif contrasena != request.form.get("confirmacion", ""):
                error = "Las contraseñas no coinciden."
            else:
                control.establecer_contrasena(Config.ARCHIVO_ACCESO, contrasena)
        elif not control.verificar(Config.ARCHIVO_ACCESO, contrasena):
            error = "Contraseña incorrecta."
        if error is None:
            session.clear()
            session.permanent = True
            session["autorizado"] = True
            return redirect(url_for("inicio"))
    return render_template("entrar.html", primera_vez=primera_vez, error=error,
                           es_local=_es_local(), direccion_red=_direccion_en_red())


@app.get("/salir")
def salir():
    session.clear()
    return redirect(url_for("entrar"))


# ---------- Pantalla principal ----------

@app.get("/")
def inicio():
    guia = [{"marcador": "{{ %s }}" % clave, "descripcion": descripcion} for clave, descripcion in MARCADORES.items()]
    return render_template("index.html", limite_mb=Config.LIMITE_ARCHIVO_MB, marcadores=guia,
                           es_local=_es_local(), direccion_red=_direccion_en_red())


@app.get("/api/formato")
def descargar_formato():
    return send_file(Config.FORMATO_CAPTURA, as_attachment=True)


@app.post("/api/validar")
def validar():
    nombre, contenido = _recibir("xlsx")
    libro = leer_libro(contenido)
    return jsonify(
        ok=True,
        archivo={"nombre": nombre, "tamano_bytes": len(contenido)},
        **validar_libro(libro, cargar_catalogo(Config.CARPETA_PERFILES)),
    )


@app.post("/api/generar")
def generar():
    nombre, contenido = _recibir("xlsx")
    libro = leer_libro(contenido)
    if not request.form.get("plantilla"):
        raise ErrorGeneracion("Elige la plantilla del contrato (paso 1).")
    plantilla = plantillas.ruta(Config.CARPETA_PLANTILLAS, request.form["plantilla"])
    lote = generar_contratos(libro, plantilla, Config.CARPETA_SALIDAS, nombre,
                             cargar_catalogo(Config.CARPETA_PERFILES), request.form.get("fecha_firma", ""))
    return jsonify(
        ok=True,
        total=len(lote["archivos"]),
        lote=lote["carpeta"].name,
        carpeta=str(lote["carpeta"].relative_to(BASE_DIR)),
        archivos=lote["archivos"],
        sin_perfil=lote["sin_perfil"],
    )


# ---------- Plantillas y perfiles ----------

@app.get("/api/plantillas")
def listar_plantillas():
    return jsonify(ok=True, plantillas=plantillas.describir(Config.CARPETA_PLANTILLAS))


@app.post("/api/plantillas")
def subir_plantilla():
    nombre, contenido = _recibir("docx")
    return jsonify(ok=True, nombre=plantillas.guardar(Config.CARPETA_PLANTILLAS, nombre, contenido))


@app.get("/api/plantillas/<path:nombre>/descargar")
def descargar_plantilla(nombre):
    return send_file(plantillas.ruta(Config.CARPETA_PLANTILLAS, nombre), as_attachment=True)


@app.post("/api/perfiles")
def subir_perfil():
    """Recibe el perfil (.docx) de un puesto y lo guarda en «perfiles de puesto»."""
    _, contenido = _recibir("docx")
    try:
        perfil = guardar_perfil(Config.CARPETA_PERFILES, request.form.get("puesto", ""), contenido)
    except ErrorPerfil as error:
        return jsonify(ok=False, error=f"No se pudo usar el perfil: {error}."), 422
    return jsonify(ok=True, **perfil)


# ---------- Contratos generados ----------

@app.get("/api/salidas/<path:lote>/zip")
def descargar_lote(lote):
    carpeta = _carpeta_de_lote(lote)
    if carpeta is None:
        return jsonify(ok=False, error="No existe esa carpeta de contratos."), 404
    memoria = io.BytesIO()
    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as comprimido:
        for ruta in sorted(carpeta.iterdir()):
            if ruta.is_file() and not ruta.name.startswith(("~$", ".")):
                comprimido.write(ruta, ruta.name)
    memoria.seek(0)
    return send_file(memoria, mimetype="application/zip", as_attachment=True, download_name=f"{carpeta.name}.zip")


@app.post("/api/salidas/<path:lote>/abrir")
def abrir_salida(lote):
    """Abre en Finder la carpeta del lote (solo en la computadora donde corre la aplicación)."""
    carpeta = _carpeta_de_lote(lote)
    if carpeta is None:
        return jsonify(ok=False, error="No existe esa carpeta de contratos."), 404
    if not _es_local() or sys.platform != "darwin":
        return jsonify(ok=False, error="La carpeta solo se abre en la computadora donde corre la aplicación; "
                                       "usa «Descargar contratos»."), 403
    subprocess.run(["open", str(carpeta)], check=False)
    return jsonify(ok=True)


# ---------- Errores ----------

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
    # Modo de desarrollo (recarga automática) solo si se pide: ORIENS_DEBUG=1
    app.run(host=Config.HOST, port=Config.PUERTO, debug=os.environ.get("ORIENS_DEBUG") == "1")
