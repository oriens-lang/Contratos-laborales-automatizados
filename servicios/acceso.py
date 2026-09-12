"""Contraseña compartida del despacho y llave de sesión.

Se guardan en «acceso.json» (fuera de git): la contraseña solo como hash, nunca en
texto. Para restablecerla, se borra ese archivo en la computadora donde corre la
aplicación y se crea una nueva al entrar.
"""
from __future__ import annotations

import json
import os
import secrets
from pathlib import Path

from werkzeug.security import check_password_hash, generate_password_hash

LONGITUD_MINIMA = 8


def llave_de_sesion(ruta: Path) -> str:
    datos = _leer(ruta)
    if not datos.get("llave"):
        datos["llave"] = secrets.token_hex(32)
        _escribir(ruta, datos)
    return datos["llave"]


def hay_contrasena(ruta: Path) -> bool:
    return bool(_leer(ruta).get("contrasena"))


def establecer_contrasena(ruta: Path, contrasena: str) -> None:
    datos = _leer(ruta)
    datos["contrasena"] = generate_password_hash(contrasena, method="pbkdf2:sha256")
    _escribir(ruta, datos)


def verificar(ruta: Path, contrasena: str) -> bool:
    guardada = _leer(ruta).get("contrasena")
    return bool(guardada) and check_password_hash(guardada, contrasena)


def _leer(ruta: Path) -> dict:
    return json.loads(ruta.read_text(encoding="utf-8")) if ruta.exists() else {}


def _escribir(ruta: Path, datos: dict) -> None:
    ruta.write_text(json.dumps(datos, indent=2), encoding="utf-8")
    os.chmod(ruta, 0o600)  # solo el usuario de la computadora puede leerlo
