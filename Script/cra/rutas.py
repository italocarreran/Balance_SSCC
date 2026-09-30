# -*- coding: utf-8 -*-
"""
Rutas de un caso CRA: todo se deriva de la carpeta base y del periodo.
Nunca rutas absolutas ni relativas al .py (mismo criterio que el BESS).
"""

from pathlib import Path

from ..nucleo.utiles import ErrorEntrada
from . import parametros as p
from .lectura import buscar_por_prefijo


def validar_aamm(aamm):
    """'2608' -> (2026, 8). ErrorEntrada si no son 4 digitos validos."""

    aamm = str(aamm or "").strip()

    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        raise ErrorEntrada(
            f"Periodo invalido: '{aamm}'. Tienen que ser 4 digitos AAMM "
            f"(por ejemplo 2608 para agosto de 2026)."
        )

    return 2000 + int(aamm[:2]), int(aamm[2:])


def resolver_rutas(carpeta_base):
    """Las carpetas del caso y el archivo de salida (existan o no)."""

    base = Path(carpeta_base)

    return {
        "base": base,
        "energia": base / p.CARPETA_ENERGIA,
        "fp": base / p.CARPETA_FP,
        "co": base / p.CARPETA_CO,
        "sc_co": base / p.CARPETA_SC_CO,
        "salida": base / p.ARCHIVO_SALIDA,
    }


# Cada entrada: (id, carpeta de resolver_rutas, prefijo, descripcion).
# El prefijo puede llevar {aamm}.
ENTRADAS = [
    ("energia", "energia", p.PREFIJO_ENERGIA,
     "energia (Formato_Solicitud_SSAA_SSCC_Hidro_*)"),
    ("fp", "fp", p.PREFIJO_FP, "factores de penalizacion (fp_*)"),
    ("co", "co", p.PREFIJO_CO, "costos de operacion (cvar_cra_AAMM_*)"),
    ("reporte_cra", "sc_co", p.PREFIJO_REPORTE_CRA,
     "CO de SC y CO (Reporte_CRA*)"),
    ("sobrecostos", "sc_co", p.PREFIJO_SOBRECOSTOS,
     "SC de SC y CO (Cálculo_SobrecostosSSCC_*)"),
]


def buscar_entrada(rutas, id_entrada, aamm):
    """La ruta del archivo de esa entrada, o None si no esta."""

    for id_, carpeta, prefijo, descripcion in ENTRADAS:
        if id_ == id_entrada:
            return buscar_por_prefijo(
                rutas[carpeta], prefijo.format(aamm=aamm), descripcion
            )

    raise KeyError(id_entrada)


def exigir_entrada(rutas, id_entrada, aamm):
    """Como buscar_entrada(), pero ErrorEntrada si no esta."""

    ruta = buscar_entrada(rutas, id_entrada, aamm)

    if ruta is None:
        for id_, carpeta, prefijo, descripcion in ENTRADAS:
            if id_ == id_entrada:
                raise ErrorEntrada(
                    f"Falta el archivo de {descripcion} en "
                    f"{rutas[carpeta]} (nombre que empiece con "
                    f"'{prefijo.format(aamm=aamm)}')."
                )

    return ruta


def crear_carpetas_caso(carpeta_base):
    """Crea las subcarpetas que falten. No toca lo que ya hay."""

    base = Path(carpeta_base)
    creadas = []

    for nombre in p.SUBCARPETAS_CASO:
        ruta = base / nombre
        if not ruta.exists():
            ruta.mkdir(parents=True)
            creadas.append(nombre)

    return creadas
