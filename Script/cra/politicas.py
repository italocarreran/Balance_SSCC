# -*- coding: utf-8 -*-
"""
Botones "Generar" de FP/ y CO/: arman fp_<AAMM>_1_<N>.xlsx y
cvar_cra_<AAMM>_1_<N>.xlsx desde las politicas de operacion de la red
(Script/Politicas/Costos_Variables.py) y los dejan en su carpeta del
caso, donde despues los toman las hojas FP y CO.
"""

from pathlib import Path

from ..Politicas import Costos_Variables as costos_variables
from ..nucleo.utiles import ErrorEntrada
from .maestros import leer_configuraciones
from .rutas import buscar_entrada, exigir_entrada, resolver_rutas, validar_aamm

# que -> (id de la entrada, carpeta de resolver_rutas, prefijo del archivo)
SALIDAS = {
    "fp": ("fp", "fp", "fp"),
    "cvar": ("co", "co", "cvar_cra"),
}


def _nada(*_args, **_kwargs):
    pass


def generar_politicas(carpeta_base, aamm, que=("fp", "cvar"),
                      registrar=_nada, progreso=_nada):
    """
    Lee el mes de politicas (y PID) y escribe lo pedido en `que`
    ("fp", "cvar"). Devuelve las rutas escritas.

    Si en la carpeta ya hay un archivo del periodo con OTRO nombre (por
    ejemplo uno dejado a mano), se para: no deja dos, porque la hoja
    no sabria cual leer.
    """

    validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)
    que = list(que)

    destinos = {}
    for cual in que:
        id_entrada, carpeta, prefijo = SALIDAS[cual]
        destino = rutas[carpeta] / costos_variables.nombre_salida(prefijo, aamm)

        if not rutas[carpeta].is_dir():
            raise ErrorEntrada(f"No se encontro la carpeta {rutas[carpeta]}")

        existente = buscar_entrada(rutas, id_entrada, aamm)
        if existente is not None and existente.name != destino.name:
            raise ErrorEntrada(
                f"Ya hay un archivo del periodo en {rutas[carpeta]}: "
                f"{existente.name}. Borralo o movelo si queres generar "
                f"{destino.name} (no pueden quedar los dos)."
            )
        destinos[cual] = destino

    configuraciones = None
    if "cvar" in que:
        maestro = exigir_entrada(rutas, "centrales_cra", aamm)
        configuraciones = leer_configuraciones(maestro)
        registrar(
            f"{len(configuraciones)} configuracion(es) del CRA en "
            f"{maestro.name}"
        )

    registrar(f"Leyendo las politicas del periodo {aamm}...")
    try:
        tablas = costos_variables.construir(
            aamm, que=que, configuraciones=configuraciones,
            registrar=registrar, progreso=progreso,
        )
    except costos_variables.ErrorPoliticas as error:
        raise ErrorEntrada(str(error)) from error

    for cual, destino in destinos.items():
        registrar(f"Escribiendo {destino.name} ({len(tablas[cual]):,} filas)")
        try:
            tablas[cual].to_excel(destino, index=False)
        except PermissionError as error:
            raise ErrorEntrada(
                f"No se pudo guardar {destino.name}: esta abierto en Excel."
            ) from error

    progreso(100)
    registrar("Listo.")

    return [Path(d) for d in destinos.values()]
