"""Mapeo Excel → marcadores {{ … }} de las plantillas de contrato.

Llena cada marcador {{ … }} de las plantillas marcadas de «plantillas/».
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
    "patron_representante": "representante legal del patrón",
    "nombre": "nombre del trabajador",
    "lugar_nacimiento": "lugar de nacimiento",
    "fecha_nacimiento_letra": "fecha de nacimiento",
    "sexo": "sexo",
    "edad": "edad («52 años de edad»)",
    "estado_civil": "estado civil",
    "domicilio": "domicilio del trabajador (completo)",
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
    "comida_duracion": "duración de la comida («30 minutos»)",
    "horas_semana": "horas de trabajo a la semana (lunes a viernes, sin la comida)",
    "lv_salida": "hora de salida de lunes a viernes",
    "sabado_entrada": "hora de entrada del sábado",
    "sabado_salida": "hora de salida del sábado",
    "salario_diario": "salario diario (número)",
    "salario_letra": "salario diario (letra)",
    "correo": "correo electrónico",
    "beneficiarios": "beneficiarios",
    "fecha_ingreso_letra": "fecha de ingreso",
    "fecha_firma_letra": "fecha de firma",
    "lugar_firma": "lugar de firma (domicilio del patrón)",
}

# Valor de fecha_firma que indica usar la FECHA ALTA IMSS de cada trabajador.
FIRMA_ALTA_IMSS = "imss"

DIRECTOS = ("nombre", "lugar_nacimiento", "sexo", "estado_civil", "domicilio_calle", "domicilio_numero",
            "rfc", "curp", "seguro_social", "credencial_elector", "puesto", "correo", "beneficiarios",
            "sabado_entrada", "sabado_salida")


def contexto_contrato(trabajador: dict, patron: dict, actividades: list[str],
                      fecha_firma: str = "") -> tuple[dict, list[str]]:
    """Valores de los marcadores y avisos de datos que existen pero no se pudieron interpretar.

    fecha_firma llega de la pantalla: AAAA-MM-DD (una fecha para todos), «imss» (la FECHA ALTA
    IMSS de cada trabajador) o vacía.
    """
    avisos: list[str] = []

    def dato(valores: dict, clave: str) -> str:
        return valores.get(clave) or FALTANTE

    contexto = {
        "patron_razon_social": dato(patron, "razon_social"),
        "patron_domicilio": _domicilio_patron(patron),
        "patron_rfc": dato(patron, "rfc"),
        "patron_actividad": dato(patron, "actividad"),
        "patron_representante": dato(patron, "representante_legal"),
        **{clave: dato(trabajador, clave) for clave in DIRECTOS},
        "domicilio_colonia": _colonia_y_municipio(trabajador),
        "domicilio": _domicilio_trabajador(trabajador),
        "edad": _edad(trabajador.get("edad")),
        # La declaración II b) se refiere a la experiencia en el puesto que se contrata.
        "experiencia": dato(trabajador, "puesto"),
        "fecha_nacimiento_letra": _interpretar(trabajador, "fecha_nacimiento", _fecha_con_letra,
                                               "FECHA DE NACIMIENTO", avisos),
        "fecha_ingreso_letra": _fecha_de_ingreso(trabajador, avisos),
        "salario_diario": _interpretar(trabajador, "salario_numero", _importe, "SALARIO DIARIO IMSS (número)", avisos),
        "salario_letra": _interpretar(trabajador, "salario_letra", _letra_sin_fraccion,
                                      "SALARIO DIARIO IMSS (letra)", avisos),
        "fecha_firma_letra": _fecha_de_firma(trabajador, fecha_firma, avisos),
        # El contrato se firma en el domicilio del patrón.
        "lugar_firma": _domicilio_patron(patron),
    }

    comida = trabajador.get("horario_comida")
    horas = _dos_horas(comida) if comida else None
    contexto["comida_inicio"], contexto["comida_fin"] = horas or (FALTANTE, FALTANTE)
    contexto["comida_duracion"] = _duracion(comida) if comida else None
    if comida and not (horas or contexto["comida_duracion"]):
        avisos.append(f"No se pudo interpretar HORARIO DE COMIDA («{comida}»); se dejó en blanco.")
    contexto["comida_duracion"] = contexto["comida_duracion"] or FALTANTE

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
    contexto["horas_semana"] = _horas_semana(entrada, salida, contexto["comida_duracion"])

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


def _fecha_de_ingreso(trabajador: dict, avisos: list[str]) -> str:
    """FECHA DE INGRESO; si el Excel no la trae, la FECHA ALTA IMSS (criterio del despacho)."""
    if trabajador.get("fecha_ingreso") or not trabajador.get("fecha_alta_imss"):
        return _interpretar(trabajador, "fecha_ingreso", _fecha_con_letra, "FECHA DE INGRESO", avisos)
    fecha = _interpretar(trabajador, "fecha_alta_imss", _fecha_con_letra, "FECHA ALTA IMSS (fecha de ingreso)", avisos)
    if fecha != FALTANTE:
        avisos.append("Sin FECHA DE INGRESO en el Excel: se usó la FECHA ALTA IMSS.")
    return fecha


def _fecha_de_firma(trabajador: dict, fecha_firma: str, avisos: list[str]) -> str:
    if fecha_firma == FIRMA_ALTA_IMSS:
        return _interpretar(trabajador, "fecha_alta_imss", _fecha_con_letra, "FECHA ALTA IMSS (fecha de firma)", avisos)
    return _fecha_iso_con_letra(fecha_firma) or FALTANTE


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
    estado = f", {patron['domicilio_estado']}" if patron.get("domicilio_estado") else ""
    return f"{calle} número {numero}, colonia {colonia}, {municipio}{estado}, C.P. {cp}"


def _domicilio_trabajador(trabajador: dict) -> str:
    """Domicilio en una sola columna tal como se capturó, o armado con calle, número, colonia y municipio."""
    if trabajador.get("domicilio"):
        return trabajador["domicilio"]
    partes = [trabajador.get(c) for c in ("domicilio_calle", "domicilio_numero", "domicilio_colonia")]
    if not any(partes):
        return FALTANTE
    calle, numero, colonia = (p or FALTANTE for p in partes)
    municipio = f", {trabajador['domicilio_municipio']}" if trabajador.get("domicilio_municipio") else ""
    return f"{calle} número {numero}, Colonia {colonia}{municipio}"


def _edad(texto: str | None) -> str:
    """'52' → '52 años de edad'; si ya viene redactada («52 años») se respeta."""
    if not texto:
        return FALTANTE
    return f"{texto} años de edad" if re.fullmatch(r"\d{1,3}", texto.strip()) else texto


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


def _duracion(texto: str) -> str | None:
    """'30 min' → '30 minutos'; '1 hora' → '60 minutos'; '14:00 a 14:30' → '30 minutos'."""
    horas = _dos_horas(texto)
    if horas:
        (h1, m1), (h2, m2) = (map(int, h.split(":")) for h in horas)
        minutos = (h2 * 60 + m2) - (h1 * 60 + m1)
        return f"{minutos} minutos" if minutos > 0 else None
    if re.fullmatch(r"\s*media\s+hora\s*", texto, re.I):
        return "30 minutos"
    coincidencia = re.fullmatch(r"\s*(\d+(?:[.,]\d+)?)\s*(m|min|mins|minutos?|h|hr|hrs|horas?)\.?\s*", texto, re.I)
    if not coincidencia:
        return None
    cantidad = float(coincidencia.group(1).replace(",", "."))
    minutos = cantidad * 60 if coincidencia.group(2).lower().startswith("h") else cantidad
    return f"{minutos:g} minutos" if minutos > 0 else None


def _horas_semana(entrada: str | None, salida: str | None, comida: str) -> str:
    """(salida − entrada − comida) × 5 días: 8:30 a 18:00 con 30 minutos de comida → '45'.

    El descanso no se cuenta porque el trabajador puede salir del centro de trabajo (art. 64 LFT).
    """
    minutos = [_minutos_del_dia(h) for h in (entrada, salida)]
    duracion = re.fullmatch(r"(\d+(?:\.\d+)?) minutos", comida)
    if None in minutos or not duracion:
        return FALTANTE
    diarias = (minutos[1] - minutos[0] - float(duracion.group(1))) / 60
    return f"{diarias * 5:g}" if diarias > 0 else FALTANTE


def _minutos_del_dia(hora: str | None) -> int | None:
    coincidencia = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", hora or "")
    return int(coincidencia.group(1)) * 60 + int(coincidencia.group(2)) if coincidencia else None


def _dos_horas(texto: str) -> tuple[str, str] | None:
    """'14:00 a 15:00 horas' → ('14:00', '15:00')."""
    horas = re.findall(r"\b\d{1,2}:\d{2}\b", texto)
    return (horas[0], horas[1]) if len(horas) == 2 else None
