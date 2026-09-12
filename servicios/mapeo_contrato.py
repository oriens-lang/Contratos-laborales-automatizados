"""Mapeo Excel → plantilla oficial «Contrato por tiempo indeterminado».

Llena cada marcador {{ … }} de «plantillas/Contrato tiempo indeterminado (marcado).docx».
Un dato que no está en el Excel queda como ____________. Cuando el dato existe
pero no se puede interpretar (p. ej., un horario sin horas), también queda en
blanco y se reporta: nunca se completa por suposición.
"""
from __future__ import annotations

import re
from datetime import date
from decimal import Decimal, InvalidOperation

from servicios.validador import FALTANTE

MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre")

# Marcador → descripción (para los reportes de datos en blanco).
MARCADORES = {
    "patron_razon_social": "razón social del patrón",
    "patron_domicilio": "domicilio del patrón",
    "patron_rfc": "RFC del patrón",
    "patron_actividad": "actividad de la empresa",
    "nombre": "nombre del trabajador",
    "lugar_nacimiento": "lugar de nacimiento",
    "fecha_nacimiento_letra": "fecha de nacimiento",
    "sexo": "sexo",
    "estado_civil": "estado civil",
    "domicilio_calle": "calle del domicilio",
    "domicilio_numero": "número del domicilio",
    "domicilio_colonia": "colonia y municipio del domicilio",
    "rfc": "RFC del trabajador",
    "curp": "CURP",
    "seguro_social": "número de seguro social",
    "credencial_elector": "credencial de elector",
    "experiencia": "experiencia (declaración II b)",
    "puesto": "puesto",
    **{f"actividad_{n}": f"actividad {n} del perfil" for n in range(1, 6)},
    "lv_entrada": "hora de entrada de lunes a viernes",
    "comida_inicio": "inicio del horario de comida",
    "comida_fin": "fin del horario de comida",
    "lv_salida": "hora de salida de lunes a viernes",
    "sabado_entrada": "hora de entrada del sábado",
    "sabado_salida": "hora de salida del sábado",
    "salario_diario": "salario diario (número)",
    "salario_letra": "salario diario (letra)",
    "correo": "correo electrónico",
    "beneficiarios": "beneficiarios",
    "fecha_ingreso_letra": "fecha de ingreso",
    "fecha_firma_letra": "fecha de firma",
}

DIRECTOS = ("nombre", "lugar_nacimiento", "sexo", "estado_civil", "domicilio_calle", "domicilio_numero",
            "rfc", "curp", "seguro_social", "credencial_elector", "puesto", "correo", "beneficiarios",
            "sabado_entrada", "sabado_salida")


def contexto_contrato(trabajador: dict, patron: dict, actividades: list[str],
                      fecha_firma: str = "") -> tuple[dict, list[str]]:
    """Valores de los marcadores y avisos de datos que existen pero no se pudieron interpretar.

    fecha_firma llega de la pantalla en formato AAAA-MM-DD (puede venir vacía).
    """
    avisos: list[str] = []

    def dato(valores: dict, clave: str) -> str:
        return valores.get(clave) or FALTANTE

    contexto = {
        "patron_razon_social": dato(patron, "razon_social"),
        "patron_domicilio": _domicilio_patron(patron),
        "patron_rfc": dato(patron, "rfc"),
        "patron_actividad": dato(patron, "actividad"),
        **{clave: dato(trabajador, clave) for clave in DIRECTOS},
        "domicilio_colonia": _colonia_y_municipio(trabajador),
        # La declaración II b) se refiere a la experiencia en el puesto que se contrata.
        "experiencia": dato(trabajador, "puesto"),
        "fecha_nacimiento_letra": _interpretar(trabajador, "fecha_nacimiento", _fecha_con_letra,
                                               "FECHA DE NACIMIENTO", avisos),
        "fecha_ingreso_letra": _interpretar(trabajador, "fecha_ingreso", _fecha_con_letra, "FECHA DE INGRESO", avisos),
        "salario_diario": _interpretar(trabajador, "salario_numero", _importe, "SALARIO DIARIO IMSS (número)", avisos),
        "salario_letra": _interpretar(trabajador, "salario_letra", _letra_sin_fraccion,
                                      "SALARIO DIARIO IMSS (letra)", avisos),
        "fecha_firma_letra": _fecha_iso_con_letra(fecha_firma) or FALTANTE,
    }

    horas = _interpretar(trabajador, "horario_comida", _dos_horas, "HORARIO DE COMIDA", avisos)
    contexto["comida_inicio"], contexto["comida_fin"] = horas if horas != FALTANTE else (FALTANTE, FALTANTE)

    # Entrada y salida de lunes a viernes: columnas propias o, si no existen, las dos horas de JORNADA.
    entrada, salida = trabajador.get("lv_entrada"), trabajador.get("lv_salida")
    if not (entrada or salida) and trabajador.get("jornada"):
        horas = _dos_horas(trabajador["jornada"])
        if horas:
            entrada, salida = horas
        else:
            avisos.append(f"No se pudo interpretar JORNADA DE TRABAJO («{trabajador['jornada']}»); "
                          "el horario de la cláusula CUARTA quedó en blanco.")
    contexto["lv_entrada"], contexto["lv_salida"] = entrada or FALTANTE, salida or FALTANTE

    for n in range(1, 6):
        contexto[f"actividad_{n}"] = actividades[n - 1] if n <= len(actividades) else FALTANTE
    return contexto, avisos


def en_blanco(contexto: dict, usados: set[str] | None = None) -> list[str]:
    """Descripción de los marcadores (de los que usa la plantilla) que quedaron como ____________."""
    return [MARCADORES[clave] for clave in MARCADORES
            if contexto.get(clave) == FALTANTE and (usados is None or clave in usados)]


def fecha_iso_con_letra(texto: str) -> str | None:
    """'2026-09-12' → '12 de septiembre de 2026' (para mostrar la fecha de firma elegida)."""
    return _fecha_iso_con_letra(texto)


def _interpretar(valores: dict, clave: str, funcion, etiqueta: str, avisos: list[str]):
    texto = valores.get(clave)
    if not texto:
        return FALTANTE
    resultado = funcion(texto)
    if resultado is None:
        avisos.append(f"No se pudo interpretar {etiqueta} («{texto}»); se dejó en blanco.")
        return FALTANTE
    return resultado


def _domicilio_patron(patron: dict) -> str:
    partes = [patron.get(c) for c in ("domicilio_calle", "domicilio_numero", "domicilio_colonia",
                                      "domicilio_municipio", "domicilio_cp")]
    if not any(partes):
        return FALTANTE
    calle, numero, colonia, municipio, cp = (p or FALTANTE for p in partes)
    return f"{calle} número {numero}, colonia {colonia}, {municipio}, C.P. {cp}"


def _colonia_y_municipio(trabajador: dict) -> str:
    """La plantilla solo tiene espacio para la colonia: el municipio se agrega a continuación."""
    colonia, municipio = trabajador.get("domicilio_colonia"), trabajador.get("domicilio_municipio")
    if not (colonia or municipio):
        return FALTANTE
    return f"{colonia or FALTANTE}, {municipio}" if municipio else colonia


def _fecha_con_letra(texto: str) -> str | None:
    """'12/02/2024' → '12 de febrero de 2024'."""
    coincidencia = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", texto.strip())
    if not coincidencia:
        return None
    dia, mes, anio = (int(n) for n in coincidencia.groups())
    return f"{dia} de {MESES[mes - 1]} de {anio}" if 1 <= mes <= 12 and 1 <= dia <= 31 else None


def _fecha_iso_con_letra(texto: str) -> str | None:
    try:
        fecha = date.fromisoformat(texto.strip())
    except ValueError:
        return None
    return f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"


def _importe(texto: str) -> str | None:
    """'620' → '620.00'; '1234.5' → '1,234.50'."""
    try:
        valor = Decimal(re.sub(r"[$,\s]", "", texto))
    except InvalidOperation:
        return None
    return f"{valor:,.2f}" if valor > 0 else None


def _letra_sin_fraccion(texto: str) -> str | None:
    """La plantilla ya escribe «/100 M.N.)»: 'SEISCIENTOS VEINTE PESOS 00/100 M.N.' → '… PESOS 00'."""
    limpio = re.sub(r"\s*/\s*100\s*M\.?\s*N\.?\s*\)?\s*$", "", texto.strip(), flags=re.IGNORECASE)
    return limpio if limpio != texto.strip() else None


def _dos_horas(texto: str) -> tuple[str, str] | None:
    """'14:00 a 15:00 horas' → ('14:00', '15:00')."""
    horas = re.findall(r"\b\d{1,2}:\d{2}\b", texto)
    return (horas[0], horas[1]) if len(horas) == 2 else None
