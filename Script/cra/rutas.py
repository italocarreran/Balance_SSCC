# -*- coding: utf-8 -*-
"""
Rutas de un caso CRA: todo se deriva de la carpeta base y del periodo.
Nunca rutas absolutas ni relativas al .py (mismo criterio que el BESS).
"""

from pathlib import Path

from ..nucleo.prorrata_retiros import buscar_archivo_prorrata
from ..nucleo.rutas import buscar_archivo_sscc_desempeno
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
        "fd": base / p.CARPETA_FD,
        "prorrata": base / p.CARPETA_PRORRATA,
        "salida": base / p.ARCHIVO_SALIDA,
    }


# ============================================================
# LAS ENTRADAS
#
# Cada una: id, carpeta (clave de resolver_rutas), como se muestra en el
# arbol, que se busca (para el mensaje de error) y la funcion que la
# encuentra: buscar(carpeta, aamm) -> ruta o None.
# ============================================================

def prefijos_energia(aamm):
    """'2608' -> ('formato_solicitud_ssaa_sscc_hidro_agosto2026',)"""

    anio, mes = validar_aamm(aamm)
    meses = (p.MESES[mes - 1],) + p.MESES_ALTERNATIVOS.get(mes, ())

    return tuple(
        p.PREFIJO_ENERGIA.format(mes=nombre, anio=anio) for nombre in meses
    )


def _buscar_energia(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, prefijos_energia(aamm), "energia (Formato_Solicitud_*)"
    )


def _buscar_fp(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_FP.format(aamm=aamm), "factores de penalizacion"
    )


def _buscar_co(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_CO.format(aamm=aamm), "costos de operacion"
    )


def _buscar_reporte_cra(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_REPORTE_CRA, "Reporte_CRA",
        extensiones=p.EXTENSIONES_REPORTE_CRA,
    )


def _buscar_sobrecostos(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_SOBRECOSTOS, "Cálculo_SobrecostosSSCC"
    )


def _buscar_sscc_desempeno(carpeta, aamm):
    # El mismo buscador del BESS: si hay varios, el mas reciente (asi
    # lo hace la macro original).
    return buscar_archivo_sscc_desempeno(carpeta)


def _buscar_prorrata(carpeta, aamm):
    return buscar_archivo_prorrata(carpeta, aamm)


ENTRADAS = [
    ("energia", "energia",
     "Formato_Solicitud_SSAA_SSCC_Hidro_<Mes><AAAA>.xlsx", _buscar_energia),
    ("fp", "fp", "fp_<AAMM>*.xlsx", _buscar_fp),
    ("co", "co", "cvar_cra_<AAMM>_*.xlsx", _buscar_co),
    ("reporte_cra", "sc_co", "Reporte_CRA*.csv", _buscar_reporte_cra),
    ("sobrecostos", "sc_co", "Cálculo_SobrecostosSSCC_*.xlsm",
     _buscar_sobrecostos),
    ("sscc_desempeno", "fd", "SSCC_Desempeño_*.xlsx", _buscar_sscc_desempeno),
    ("prorrata", "prorrata", "Prorrata_Retiros_<AAMM>_pre/_def.xlsx",
     _buscar_prorrata),
]


def _entrada(id_entrada):
    for entrada in ENTRADAS:
        if entrada[0] == id_entrada:
            return entrada

    raise KeyError(id_entrada)


def buscar_entrada(rutas, id_entrada, aamm):
    """La ruta del archivo de esa entrada, o None si no esta."""

    _, carpeta, _, buscar = _entrada(id_entrada)

    return buscar(rutas[carpeta], aamm)


def exigir_entrada(rutas, id_entrada, aamm):
    """Como buscar_entrada(), pero ErrorEntrada si no esta."""

    ruta = buscar_entrada(rutas, id_entrada, aamm)

    if ruta is None:
        _, carpeta, patron, _ = _entrada(id_entrada)
        raise ErrorEntrada(
            f"Falta el archivo {patron.replace('<AAMM>', aamm)} en "
            f"{rutas[carpeta]}."
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
