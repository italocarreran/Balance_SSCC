# -*- coding: utf-8 -*-
"""
El proceso del CRA: arma las hojas pedidas y las escribe en
Balance_CRA.xlsx conservando las demas.

Por ahora solo existen las hojas de entrada con carga definida
(SECCIONES). El calculo (CÁLCULO_CRA -> PRORRATA_RETIROS -> RESUMEN)
se agrega cuando esten cerradas sus entradas pendientes.
"""

from ..nucleo.externos import indicadores_dco
from ..nucleo.utiles import ErrorEntrada
from . import parametros as p
from .escritura import escribir_hojas
from .fuentes_bess import construir_fd, construir_matriz_prorrata
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
    return construir_sc_co(
        reporte, sobrecostos, registrar, periodo=validar_aamm(aamm)
    )


def _hoja_fd(seccion):
    def construir(rutas, aamm, registrar):
        ruta = exigir_entrada(rutas, "sscc_desempeno", aamm)
        registrar(f"  Leyendo {ruta.name}")
        return construir_fd(ruta, seccion, registrar)
    return construir


def _hoja_prorrata(rutas, aamm, registrar):
    anio, mes = validar_aamm(aamm)
    ruta = exigir_entrada(rutas, "prorrata", aamm)
    registrar(f"  Leyendo {ruta.name}")
    return construir_matriz_prorrata(ruta, anio, mes, registrar)


# id de seccion -> (hoja de salida, funcion que la arma). El orden es el
# de ejecucion (de la entrada al calculo).
SECCIONES = {
    "energia": (p.HOJA_ENERGIA, _hoja_energia),
    "fp": (p.HOJA_FP, _hoja_fp),
    "co": (p.HOJA_CO, _hoja_co),
    "sc_co": (p.HOJA_SC_CO, _hoja_sc_co),
    **{
        seccion: (hoja, _hoja_fd(seccion))
        for seccion, (hoja, _, _) in p.HOJAS_FD.items()
    },
    "prorrata": (p.HOJA_PRORRATA, _hoja_prorrata),
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


def traer_fd(carpeta_base, aamm, registrar=_nada, progreso=_nada):
    """
    Boton "Traer" de FD/: copia el SSCC_Desempeño_* del periodo desde el
    arbol de indicadores del DCO (y lo descomprime si viene en zip). Es
    la misma funcion que usa el BESS, con destino la carpeta FD/ del
    caso CRA.
    """

    validar_aamm(aamm)
    destino = resolver_rutas(carpeta_base)["fd"]

    if not destino.is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta {destino}")

    progreso(5)
    try:
        copiados, extraidos, version = indicadores_dco.traer_fd(
            destino, aamm, registrar=registrar
        )
    except indicadores_dco.ErrorFd as error:
        raise ErrorEntrada(str(error)) from error
    except OSError as error:
        raise ErrorEntrada(f"No se pudo traer el FD: {error}") from error

    progreso(100)
    registrar(
        f"Listo: {len(copiados)} archivo(s) desde {version.name}"
        + (f" y {len(extraidos)} descomprimido(s)" if extraidos else "")
        + f" en {destino}"
    )

    return destino
