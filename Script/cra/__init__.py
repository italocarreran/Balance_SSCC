# -*- coding: utf-8 -*-
"""
Calculo del CRA (remuneracion CRA): replica de
5_REMUNERACIÓN_CRA_<AAMM>_Definitivo.xlsx, hoja por hoja.

Es un paquete hermano de `nucleo` (el del Balance BESS): usa sus piezas
genericas (normalizar, ErrorEntrada, formato de hojas) pero `nucleo`
nunca importa nada de aca. La ventana es Balance_CRA.py.

    parametros.py     Carpetas, patrones de nombre, hojas y mapeos de columnas.
    lectura.py        Busqueda de archivos y lectura por letra/fila de Excel.
    rutas.py          Rutas del caso (<mes>/CRA/ + las comunes con el BESS) y AAMM.
    hojas_entrada.py  ENERGIA, FP, CO y SC y CO.
    fuentes_bess.py   FD_CPF/CSF/CTF y la matriz de PRORRATA_RETIROS (fuentes del BESS).
    politicas.py      Botones Generar de fp_ y cvar_cra_ (Script/Politicas).
    medidas_cra.py    Medidas_CRA.xlsx desde PRMTE (boton Generar del Formato_Solicitud).
    maestros.py       Auxiliares/centrales_cra.xlsx (configuraciones, empresas, diccionario FD).
    escritura.py      5_REMUNERACION_CRA_AAMM.xlsx (cuadros y escritura rapida del BESS).
    proceso.py        Las secciones y generar_balance_cra().
    estructura.py     El arbol que dibuja la ventana.

Referencia de dominio: docs/Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_Formulas.md.
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
from .maestros import (
    leer_configuraciones,
    leer_empresas,
    leer_unidades_fd,
)
from .medidas_cra import construir_medidas_cra, generar_medidas_cra
from .politicas import generar_politicas
from .proceso import SECCIONES, generar_balance_cra, traer_fd
from .rutas import (
    buscar_entrada,
    crear_carpetas_caso,
    nombre_salida,
    raiz_cra,
    resolver_rutas,
    subcarpetas_del_caso,
    validar_aamm,
)

__all__ = [
    "ErrorEntrada", "parametros", "HOJAS_PENDIENTES", "revisar_estructura",
    "construir_co", "construir_energia", "construir_fp", "construir_sc_co",
    "construir_sc_co_desde_reporte", "construir_sc_co_desde_sobrecostos",
    "participacion_por_servicio", "SECCIONES", "generar_balance_cra",
    "construir_fd", "construir_matriz_prorrata", "traer_fd",
    "generar_politicas", "generar_medidas_cra", "construir_medidas_cra",
    "leer_configuraciones", "leer_empresas",
    "leer_unidades_fd",
    "buscar_entrada", "crear_carpetas_caso", "resolver_rutas", "validar_aamm",
    "nombre_salida", "raiz_cra", "subcarpetas_del_caso",
]
