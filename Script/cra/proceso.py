# -*- coding: utf-8 -*-
"""
El proceso del CRA: arma las hojas pedidas y las escribe en
Balance_CRA.xlsx conservando las demas.

Por ahora solo existen las hojas de entrada con carga definida
(SECCIONES). El calculo (CÁLCULO_CRA -> PRORRATA_RETIROS -> RESUMEN)
se agrega cuando esten cerradas sus entradas pendientes.
"""

from . import parametros as p
from .escritura import escribir_hojas
from .hojas_entrada import (
    construir_co,
    construir_energia,
    construir_fp,
    construir_sc_co,
)
from .rutas import exigir_entrada, resolver_rutas, validar_aamm


def _nada(*_args, **_kwargs):
    pass


def _hoja_energia(rutas, aamm, registrar):
    anio, mes = validar_aamm(aamm)
    ruta = exigir_entrada(rutas, "energia", aamm)
    registrar(f"  Leyendo {ruta.name}")
    return construir_energia(ruta, anio, mes, registrar)


def _hoja_fp(rutas, aamm, registrar):
    ruta = exigir_entrada(rutas, "fp", aamm)
    registrar(f"  Leyendo {ruta.name}")
    return construir_fp(ruta, registrar)


def _hoja_co(rutas, aamm, registrar):
    ruta = exigir_entrada(rutas, "co", aamm)
    registrar(f"  Leyendo {ruta.name}")
    return construir_co(ruta, registrar)


def _hoja_sc_co(rutas, aamm, registrar):
    reporte = exigir_entrada(rutas, "reporte_cra", aamm)
    sobrecostos = exigir_entrada(rutas, "sobrecostos", aamm)
    registrar(f"  Leyendo {reporte.name} y {sobrecostos.name}")
    return construir_sc_co(reporte, sobrecostos, registrar)


# id de seccion -> (hoja de salida, funcion que la arma). El orden es el
# de ejecucion (de la entrada al calculo).
SECCIONES = {
    "energia": (p.HOJA_ENERGIA, _hoja_energia),
    "fp": (p.HOJA_FP, _hoja_fp),
    "co": (p.HOJA_CO, _hoja_co),
    "sc_co": (p.HOJA_SC_CO, _hoja_sc_co),
}


def generar_balance_cra(
    carpeta_base, aamm, secciones=None, registrar=_nada, progreso=_nada
):
    """
    Arma las `secciones` pedidas (todas si es None) y las escribe en
    Balance_CRA.xlsx. Una seccion que falla corta todo antes de
    escribir: el libro nunca queda con la mitad de una corrida.
    """

    validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)
    secciones = list(SECCIONES) if secciones is None else list(secciones)

    hojas = {}
    for i, seccion in enumerate(secciones):
        hoja, construir = SECCIONES[seccion]
        registrar(f"Hoja '{hoja}'...")
        df = construir(rutas, aamm, registrar)
        registrar(f"  {len(df)} fila(s).")
        hojas[hoja] = df
        progreso(int(90 * (i + 1) / len(secciones)))

    registrar(f"Escribiendo {rutas['salida'].name}...")
    escribir_hojas(rutas["salida"], hojas)
    progreso(100)
    registrar("Listo.")

    return rutas["salida"]
