# -*- coding: utf-8 -*-
"""
Calculo del CRA (remuneracion CRA): replica de
5_REMUNERACIÓN_CRA_<AAMM>_Definitivo.xlsx, hoja por hoja.

Es un paquete hermano de `nucleo` (el del Balance BESS): usa sus piezas
genericas (normalizar, ErrorEntrada, formato de hojas) pero `nucleo`
nunca importa nada de aca. La ventana es Balance_CRA.py.

    parametros.py     Carpetas, patrones de nombre, hojas y mapeos de columnas.
    lectura.py        Busqueda de archivos y lectura por letra/fila de Excel.
    rutas.py          Rutas del caso y validacion del AAMM.
    hojas_entrada.py  ENERGIA, FP, CO y SC y CO.
    fuentes_bess.py   FD_CPF/CSF/CTF y la matriz de PRORRATA_RETIROS (fuentes del BESS).
    escritura.py      Balance_CRA.xlsx conservando las hojas que no se tocan.
    proceso.py        Las secciones y generar_balance_cra().
    estructura.py     El arbol que dibuja la ventana.

Referencia de dominio: docs/Trazabilidad_CRA_Periodo_Generico_v3.md.
"""

from ..nucleo.utiles import ErrorEntrada
from . import parametros
from .estructura import HOJAS_PENDIENTES, revisar_estructura
from .hojas_entrada import (
    construir_co,
    construir_energia,
    construir_fp,
    construir_sc_co,
    construir_sc_co_desde_reporte,
    construir_sc_co_desde_sobrecostos,
    participacion_por_servicio,
)
from .fuentes_bess import construir_fd, construir_matriz_prorrata
from .proceso import SECCIONES, generar_balance_cra, traer_fd
from .rutas import (
    buscar_entrada,
    crear_carpetas_caso,
    resolver_rutas,
    validar_aamm,
)

__all__ = [
    "ErrorEntrada", "parametros", "HOJAS_PENDIENTES", "revisar_estructura",
    "construir_co", "construir_energia", "construir_fp", "construir_sc_co",
    "construir_sc_co_desde_reporte", "construir_sc_co_desde_sobrecostos",
    "participacion_por_servicio", "SECCIONES", "generar_balance_cra",
    "construir_fd", "construir_matriz_prorrata", "traer_fd",
    "buscar_entrada", "crear_carpetas_caso", "resolver_rutas", "validar_aamm",
]
