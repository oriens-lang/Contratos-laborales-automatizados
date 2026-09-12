"""Validación de los datos leídos del Excel y de los perfiles de puesto.

Los datos faltantes y las inconsistencias solo generan advertencias: no impiden
el procesamiento ni se completan con información que no esté en el archivo.
Solo se pueden generar los contratos de trabajadores cuyo puesto tenga perfil.
"""
from __future__ import annotations

import re

from servicios.perfiles import relacionar

FALTANTE = "____________"

# (patrón, mensaje) para revisar el formato de los datos capturados.
REGLAS_TRABAJADOR = {
    "curp": (r"[A-Z]{4}\d{6}[HMX][A-Z]{5}[A-Z0-9]\d",
             "La CURP no tiene la estructura oficial de 18 caracteres"),
    "rfc": (r"[A-ZÑ&]{4}\d{6}[A-Z0-9]{3}",
            "El RFC no tiene la estructura de persona física (13 caracteres)"),
    "seguro_social": (r"\d{11}", "El número de seguro social debe tener 11 dígitos"),
    "correo": (r"[^@\s]+@[^@\s]+\.[^@\s]+", "El correo electrónico no tiene un formato válido"),
}
REGLAS_PATRON = {
    "rfc": (r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}", "El RFC del patrón no tiene una estructura válida"),
}
CAMPOS_SIN_ESPACIOS = {"curp", "rfc", "seguro_social"}
CAMPOS_UNICOS = {"curp": "La CURP", "rfc": "El RFC", "seguro_social": "El número de seguro social"}


def validar_libro(libro: dict, catalogo: dict[str, dict]) -> dict:
    """Arma el resumen, las tablas y las observaciones que muestra la interfaz."""
    hoja_t = libro["trabajadores"]
    etiquetas = {col["clave"]: col["etiqueta"] for col in hoja_t["columnas"]}

    relacion = relacionar(hoja_t["registros"], catalogo)
    perfil_por_fila = {t["fila"]: fila for fila in relacion["puestos"] for t in fila["trabajadores"]}

    duplicados = _buscar_duplicados(hoja_t["registros"])
    registros, por_trabajador = [], []
    for reg in hoja_t["registros"]:
        valores = reg["valores"]
        faltantes = [col["clave"] for col in hoja_t["columnas"] if not valores.get(col["clave"])]
        advertencias = [
            f"«{etiquetas[clave]}» contiene el texto del formato («{texto}»), no un dato."
            for clave, texto in reg["auxiliares"].items()
        ]
        advertencias += _revisar_formatos(valores, REGLAS_TRABAJADOR)
        advertencias += duplicados.get(reg["fila"], [])

        grupo = perfil_por_fila.get(reg["fila"])
        perfil = grupo["estado"] if grupo else "sin_puesto"
        if grupo is None:
            advertencias.append("Sin PUESTO capturado: no se le puede asignar perfil ni generar contrato.")
        elif grupo["detalle"]:
            advertencias.append(grupo["detalle"])

        registros.append({"fila": reg["fila"], "valores": valores, "faltantes": faltantes, "perfil": perfil})
        if faltantes or advertencias:
            por_trabajador.append({
                "fila": reg["fila"],
                "nombre": valores.get("nombre") or "(sin nombre)",
                "faltantes": [etiquetas[clave] for clave in faltantes],
                "advertencias": advertencias,
            })

    patron, patron_faltantes, obs_patron = _armar_patron(libro["patron"])

    total = len(registros)
    completos = sum(1 for r in registros if not r["faltantes"])
    generables = sum(1 for r in registros if r["perfil"] == "disponible")
    observaciones = (_observaciones_generales(libro, total) + _observaciones_perfiles(relacion, catalogo)
                     + obs_patron)

    hay_pendientes = (total == 0 or por_trabajador or patron_faltantes or patron is None
                      or relacion["pendientes"] or relacion["sin_puesto"])
    return {
        "hojas": {
            "leidas": libro["hojas_leidas"],
            "omitidas": [
                f"{h['nombre']} (vacía)" if h["vacia"] else h["nombre"] for h in libro["hojas_omitidas"]
            ],
        },
        "resumen": {
            "patron": patron["razon_social"] if patron else None,
            "trabajadores": total,
            "completos": completos,
            "con_faltantes": total - completos,
            "puestos": len(relacion["puestos"]),
            "perfiles_disponibles": len(relacion["disponibles"]),
            "perfiles_pendientes": len(relacion["pendientes"]),
            "generables": generables,
        },
        "perfiles": {
            "puestos": [{k: fila[k] for k in ("puesto", "trabajadores", "estado", "detalle", "perfil", "actividades")}
                        for fila in relacion["puestos"]],
            "sin_puesto": relacion["sin_puesto"],
            "en_carpeta": sorted(perfil["puesto"] for perfil in catalogo.values()),
        },
        "trabajadores": {
            "hoja": hoja_t["hoja"],
            "columnas": [
                {k: col[k] for k in ("clave", "letra", "grupo", "encabezado")} for col in hoja_t["columnas"]
            ],
            "registros": registros,
        },
        "patron": patron,
        "validacion": {
            "estado": "pendiente" if hay_pendientes else "correcto",
            "observaciones": observaciones,
            "por_trabajador": por_trabajador,
        },
        # Solo se generan los contratos de trabajadores cuyo puesto tiene perfil.
        # Los datos faltantes no bloquean: el contrato sale con ____________ en esos datos.
        "puede_generar": total > 0,
    }


def _observaciones_perfiles(relacion: dict, catalogo: dict[str, dict]) -> list[str]:
    obs = [f"Puestos distintos: {len(relacion['puestos'])}."]
    if relacion["disponibles"]:
        obs.append(f"Perfiles disponibles ({len(relacion['disponibles'])}): {', '.join(relacion['disponibles'])}.")
    for fila in relacion["puestos"]:
        if fila["estado"] != "disponible":
            n = len(fila["trabajadores"])
            obs.append(f"{fila['detalle']} ({n} trabajador{'es' if n != 1 else ''}; su contrato saldrá "
                       "sin actividades en la cláusula PRIMERA y sin ANEXO UNO).")
    if relacion["sin_puesto"]:
        filas = ", ".join(str(t["fila"]) for t in relacion["sin_puesto"])
        obs.append(f"Trabajadores sin PUESTO capturado (filas {filas}): no se generará su contrato.")
    usados = {fila["clave"] for fila in relacion["puestos"]}
    sin_uso = sorted(p["puesto"] for clave, p in catalogo.items() if clave not in usados)
    if sin_uso:
        obs.append(f"Perfiles en la carpeta que no corresponden a ningún trabajador de este Excel: {', '.join(sin_uso)}.")
    return obs


def _armar_patron(hoja_p: dict | None):
    """Toma el primer registro de la hoja del patrón y marca sus campos vacíos."""
    if hoja_p is None:
        return None, [], []

    registro = hoja_p["registros"][0] if hoja_p["registros"] else None
    valores = registro["valores"] if registro else {}
    campos = [
        {"grupo": col["grupo"], "encabezado": col["encabezado"], "valor": valores.get(col["clave"], "")}
        for col in hoja_p["columnas"]
    ]
    faltantes = [col["etiqueta"] for col in hoja_p["columnas"] if not valores.get(col["clave"])]

    observaciones = []
    if registro is None:
        observaciones.append(f"La hoja «{hoja_p['hoja']}» no tiene datos capturados; solo encabezados.")
    else:
        if len(hoja_p["registros"]) > 1:
            observaciones.append(
                f"La hoja «{hoja_p['hoja']}» tiene {len(hoja_p['registros'])} registros; "
                f"se toma el de la fila {registro['fila']}."
            )
        if faltantes:
            observaciones.append(f"Datos del patrón pendientes ({len(faltantes)}): {', '.join(faltantes)}.")
        observaciones += [f"Patrón: {a}" for a in _revisar_formatos(valores, REGLAS_PATRON)]
    observaciones += _observaciones_de_hoja(hoja_p)

    patron = {
        "hoja": hoja_p["hoja"],
        "fila": registro["fila"] if registro else None,
        "encontrado": registro is not None,
        "razon_social": valores.get("razon_social"),
        "campos": campos,
    }
    return patron, faltantes, observaciones


def _observaciones_generales(libro: dict, total: int) -> list[str]:
    hoja_t = libro["trabajadores"]
    obs = [f"Hojas leídas: {', '.join(f'«{h}»' for h in libro['hojas_leidas'])}."]
    for h in libro["hojas_omitidas"]:
        obs.append(f"Se omitió la hoja «{h['nombre']}»" + (" (vacía)." if h["vacia"] else "."))

    filas = hoja_t["filas_encabezado"]
    en_filas = f"en las filas {filas[0]} y {filas[1]}" if len(filas) > 1 else f"en la fila {filas[0]}"
    obs.append(f"«{hoja_t['hoja']}»: {len(hoja_t['columnas'])} columnas detectadas (encabezados {en_filas}).")
    obs.append(f"Trabajadores detectados: {total}." if total
               else f"No hay trabajadores capturados en la hoja «{hoja_t['hoja']}».")
    obs += _observaciones_de_hoja(hoja_t)
    obs += libro["avisos"]
    return obs


def _observaciones_de_hoja(hoja: dict) -> list[str]:
    """Filas ignoradas y detalles de estructura de una hoja."""
    obs = []
    por_motivo: dict[str, list[int]] = {}
    for fila in hoja["ignoradas"]:
        por_motivo.setdefault(fila["motivo"], []).append(fila["fila"])
    for motivo, filas in por_motivo.items():
        ignoradas = "ignorada" if len(filas) == 1 else "ignoradas"
        obs.append(f"«{hoja['hoja']}», {_rangos(filas)} {ignoradas}: {motivo}.")
    if hoja["columnas_sin_encabezado"]:
        obs.append(f"«{hoja['hoja']}»: las columnas {', '.join(hoja['columnas_sin_encabezado'])} "
                   "tienen datos pero no encabezado; no se leyeron.")
    if hoja["columnas_duplicadas"]:
        obs.append(f"«{hoja['hoja']}»: encabezados repetidos en {', '.join(hoja['columnas_duplicadas'])}.")
    if hoja["no_encontrados"]:
        obs.append(f"«{hoja['hoja']}»: no se encontraron las columnas {', '.join(hoja['no_encontrados'])}. "
                   "Si el contrato las necesita, usa el formato de captura actualizado.")
    return obs


def _revisar_formatos(valores: dict, reglas: dict) -> list[str]:
    advertencias = []
    for clave, (patron, mensaje) in reglas.items():
        valor = valores.get(clave)
        if not valor:
            continue
        limpio = re.sub(r"[\s-]", "", valor).upper() if clave in CAMPOS_SIN_ESPACIOS else valor.strip()
        if re.fullmatch(patron, limpio):
            continue
        if clave == "seguro_social" and limpio.isdigit() and len(limpio) < 11:
            mensaje += f"; tiene {len(limpio)} (revisa si Excel eliminó ceros iniciales)"
        advertencias.append(f"{mensaje}: «{valor}».")
    return advertencias


def _buscar_duplicados(registros: list[dict]) -> dict[int, list[str]]:
    """CURP, RFC o NSS repetidos entre trabajadores → advertencia en cada fila afectada."""
    avisos: dict[int, list[str]] = {}
    for clave, nombre in CAMPOS_UNICOS.items():
        filas_por_valor: dict[str, list[int]] = {}
        for reg in registros:
            valor = re.sub(r"[\s-]", "", reg["valores"].get(clave, "")).upper()
            if valor:
                filas_por_valor.setdefault(valor, []).append(reg["fila"])
        for filas in filas_por_valor.values():
            if len(filas) > 1:
                for fila in filas:
                    otras = ", ".join(str(f) for f in filas if f != fila)
                    avisos.setdefault(fila, []).append(f"{nombre} se repite en la fila {otras}.")
    return avisos


def _rangos(filas: list[int]) -> str:
    """[4, 5, 6, 9] → 'filas 4 a 6 y 9'."""
    tramos = []
    inicio = fin = filas[0]
    for fila in filas[1:] + [None]:
        if fila is not None and fila == fin + 1:
            fin = fila
            continue
        tramos.append(str(inicio) if inicio == fin else f"{inicio} a {fin}")
        inicio = fin = fila
    texto = tramos[0] if len(tramos) == 1 else ", ".join(tramos[:-1]) + " y " + tramos[-1]
    return ("fila " if len(filas) == 1 else "filas ") + texto
