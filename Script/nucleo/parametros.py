# -*- coding: utf-8 -*-
"""
Nombres de archivo, carpeta y hoja del caso; mapeo A:I.
"""

import re

from .externos import ofertas_adj


# ============================================================
# PARAMETROS FIJOS
# ============================================================

# Equivale a Medidores!S1 en la planilla 11.
INICIO_VENTANA = 10

# Umbral de la columna O. La planilla usa 6%.
UMBRAL_SOC = 0.06


# ============================================================
# NOMBRES DE ARCHIVO Y CARPETA
# ============================================================

CARPETA_MEDIDAS = "Medidas"
CARPETA_AUXILIARES = "Auxiliares"
CARPETA_OFERTAS = "Ofertas"
CARPETA_CMG = "Cmg"
# El usuario le cambio el nombre a esta carpeta: antes "SSCC_Desempeño"
# y ahora "FD y FMA", porque adentro van las dos cosas -el archivo
# SSCC_Desempeño_* del que sale la hoja FD, y las salidas de los FMA
# (fma_cpf_*, fma_csf_*, fma_cft_*) de las que sale Subastas!FMA-.
# Se sigue aceptando el nombre viejo si la carpeta nueva no existe, para
# que los casos ya armados no se rompan (ver resolver_rutas).
CARPETA_FD_FMA = "FD y FMA"
CARPETA_FD_FMA_ANTIGUA = "SSCC_Desempeño"
CARPETA_SUBASTAS = "Subastas"
CARPETA_PRORRATA_RETIROS = "Prorrata retiros"

ARCHIVO_MEDIDAS_SAE = "Medidas_SAE.xlsx"
HOJA_MEDIDAS_SAE = "Medidas"

# Medidas_SAE.xlsx tampoco se arma a mano (igual que cmg.xlsx): lo
# genera el programa desde las dos APIs del Coordinador, y los pasos
# intermedios (lotes descargados, marca de reanudacion) van a esta
# carpeta, que la ventana NO muestra a pedido del usuario -- no son
# entradas ni salidas del caso, son andamios.
CARPETA_TRABAJO_MEDIDAS = "_trabajo"

ARCHIVO_CENTRALES = "Centrales.xlsx"
HOJA_RESUMEN_BESS = "Resumen BESS"
HOJA_DICCIONARIO = "Diccionario"

# Nombre literal y fijo (a diferencia de SoC/Ofertas/SSCC_Desempeño/
# Subastas): asi lo exige Cargar_CMg_Desde_Archivo.
ARCHIVO_CMG = "cmg.xlsx"
HOJA_CMG_ORIGEN = "CMg"

# El CSV 15-minutal del que sale cmg.xlsx vive en la carpeta Cmg/ del
# caso, al lado de cmg.xlsx, y se baja ahi desde la unidad de red con
# el boton "Traer cmg_15min" (ver traer_csv_cmg). El nombre del
# archivo, la ruta de red y el formato del CSV los conoce
# Script/Cmg/Extrae_CMG_barras.py, no este modulo.

HOJA_CPF_HORARIO = "CPF Horario"
HOJA_CSF_HORARIO = "CSF Horario"

# Las subastas ya no salen de la planilla 3: salen de los Access
# OfertasSSCCAdj*.accdb, que son su origen real (la planilla 3 tambien
# se arma pegando lo que sale de ellos). Se copian de la unidad de red
# a <CARPETA_BASE>/Subastas/DB subastas/ con el boton "Traer
# subastas". El nombre de la carpeta, la ruta de red, los nombres de
# archivo y la consulta SQL los conoce Script/Subastas/
# Ofertas_Adjudicadas.py, no este modulo.
CARPETA_DB_SUBASTAS = ofertas_adj.CARPETA_DB_SUBASTAS

# UNA sola planilla de salida (pedido del usuario: "combinar el
# consolidado entradas con pagos bess"). Antes eran dos
# (Consolidado_entradas.xlsx + Pagos_BESS.xlsx); ahora las entradas y
# el calculo viven en el mismo libro, ordenado de FIN A INICIO: el
# Resumen primero y las entradas al final (ver ORDEN_HOJAS_SALIDA).
ARCHIVO_SALIDA = "Balance_BESS.xlsx"

# Las hojas de control de la corrida (Alertas, Ejecucion, Log) ya no
# ensucian la planilla de trabajo: viven en su propio archivo, al lado
# (pedido del usuario).
ARCHIVO_CONTROL = "Control_corrida.xlsx"

# --- hojas de la planilla de salida --------------------------------
# Las entradas (el "inicio").
HOJA_MEDIDORES = "Medidores"
HOJA_OFERTAS_SSCC = "Ofertas SSCC"
HOJA_CMG = "CMg"
HOJA_FD = "FD"
HOJA_SUBASTAS = "Subastas"

# El calculo y el cierre (el "fin").
HOJA_CALCULO_ECOSTOS = "Calculo E Costos"
HOJA_CALCULO_RE545 = "Calculo RE545"
HOJA_COMPENSACION_CENTRAL = "COMPENSACION_CENTRAL"
HOJA_PRORRATA_RETIROS = "PRORRATA_RETIROS"
HOJA_RESUMEN = "Resumen"

# El orden en que quedan las hojas del libro, de FIN A INICIO: se abre
# en el Resumen (quien paga y quien recibe, que es lo que se mira
# primero) y se termina en las entradas de las que sale todo. Es el
# orden que usan la escritura (escritura.py) y el diagrama de la
# ventana (estructura.py): si se agrega una hoja, va aca.
ORDEN_HOJAS_SALIDA = (
    HOJA_RESUMEN,
    HOJA_PRORRATA_RETIROS,
    HOJA_COMPENSACION_CENTRAL,
    HOJA_CALCULO_RE545,
    HOJA_CALCULO_ECOSTOS,
    HOJA_SUBASTAS,
    HOJA_FD,
    HOJA_CMG,
    HOJA_OFERTAS_SSCC,
    HOJA_MEDIDORES,
)

# --- hojas del archivo de control ----------------------------------
HOJA_EJECUCION = "Ejecucion"
HOJA_ALERTAS = "Alertas"
HOJA_LOG = "Log"

ORDEN_HOJAS_CONTROL = (HOJA_EJECUCION, HOJA_ALERTAS, HOJA_LOG)

# El periodo AAMM (ej. "2607") ya no se infiere del nombre del archivo:
# lo ingresa el usuario en la ventana. El archivo de SoC solo debe
# contener "SOC" y el AAMM en su nombre (plan, seccion 19.1) - no existe
# un nombre de archivo literal fijo.
PATRON_AAMM = re.compile(r"^\d{4}$")

# Extensiones de Excel aceptadas para los archivos que se buscan por
# patron de nombre (OfertasSSCC, SSCC_Desempeño_*) - no para SoC
# (siempre .xlsx) ni para cmg.xlsx (nombre literal fijo).
EXTENSIONES_EXCEL = {".xlsx", ".xlsm", ".xlsb", ".xls"}

# Se derivan con .lower() en vez de transcribir el literal a mano: con
# letras dobles/triples seguidas ("Ofertas"+"SSCC") es facil perder una
# al tipear (ya paso una vez, ver METODOLOGIA.md #7).
PATRON_NOMBRE_OFERTAS = "OfertasSSCC".lower()
PATRON_NOMBRE_SSCC_DESEMPENO = "SSCC_Desempeño_".lower()


# ============================================================
# MAPEO DE COLUMNAS A:I
# ============================================================

# Orden de Medidas_SAE.xlsx, que alimenta Medidores!A:I.
COLUMNAS_AI = [
    "Mes",
    "Dia",
    "Hora",
    "Minutos",
    "Hora Mes",
    "Cuarto de Hora",
    "clave",
    "intervalo",
    "Gen_Unidad",
]

# Nombre logico de cada letra de Excel, para poder comparar contra la
# hoja original columna por columna. El orden de insercion de este
# dict ES el orden final de columnas de Medidores.
#
# V, W, X, Y, AB, AC, AD, AE NO estan aca: en la planilla original no
# son una columna por fila de Medidores, son tablas auxiliares de otro
# largo (una fila por central x dia, o por central x ventana) que solo
# viven en esas letras de columna porque ahi habia espacio libre. En
# Python se escriben como hoja propia de Consolidado_entradas.xlsx en
# vez de forzarlas a columnas del mismo largo que A:U (ver ejecutar()).
LETRA_A_CAMPO = {
    "A": "Mes",
    "B": "Dia",
    "C": "Hora",
    "D": "Minutos",
    "E": "Hora Mes",
    "F": "Cuarto de Hora",
    "G": "clave",
    "H": "intervalo",
    "I": "Gen_Unidad",
    "J": "SoC",
    "K": "Copia_Ventana",
    "L": "Ventana",
    "M": "M_VACIA",
    "N": "Clave_Dia_HoraMes",
    "O": "Indicador_SoC",
    "P": "P_VACIA",
    "Q": "Q_VACIA",
    "U": "U_VACIA",
}

# R, S y T de la planilla original: las tres SALEN DE OFERTAS SSCC y
# por eso ya no viven en la hoja Medidores.
#
# Pedido explicito del usuario: "para construir Medidas se leen las
# ofertas, me gustaria sacar lo de ofertas de esa hoja y dejarlas en la
# hoja de ofertas para independizar Medidas de ofertas". Ahora el boton
# "Actualizar" de Medidores no abre el archivo *OfertasSSCC* para nada;
# la hoja "Ofertas SSCC" del consolidado guarda las dos tablas de las
# que se derivan las tres columnas:
#
#   R (Oferta_Completa_Dia)      <- "Ofertas SSCC por dia" (central+dia)
#   S (Indicador_Ventana_Oferta) <- R + la Ventana de Medidores
#   T (Ventana_No_Completa)      <- "Resumen ventana oferta" (central+ventana)
#
# Quien las necesita (el calculo de Pagos_BESS.xlsx: el reparto entre
# "Calculo E Costos" y "Calculo RE545" se hace con T) las reconstruye
# con completar_ofertas_en_medidores() a partir de esas dos tablas, sin
# volver a leer el archivo de ofertas.
COLUMNAS_OFERTAS_EN_MEDIDORES = [
    "Oferta_Completa_Dia",
    "Indicador_Ventana_Oferta",
    "Ventana_No_Completa",
]

# Columnas que el plan define como deliberadamente vacias (plan
# seccion 16.3): no son trabajo pendiente, es el diseño confirmado.
COLUMNAS_VACIAS = [
    "M_VACIA",
    "P_VACIA",
    "Q_VACIA",
    "U_VACIA",
]

# Auxiliares de la hoja Medidores que NO se escriben en
# Consolidado_entradas.xlsx (pedido del usuario: "quita los auxiliares
# innecesarios"). Se siguen calculando igual que antes -- esto es solo
# que dejen de ocupar una columna de la hoja:
#
#   - las cuatro COLUMNAS_VACIAS: existen en la planilla original
#     porque ahi la letra de Excel manda (M, P, Q y U estan en blanco);
#     en la salida de Python no hay nada que alinear, y una columna
#     vacia con nombre "M_VACIA" solo invita a preguntar que falta.
#   - Clave_Dia_HoraMes (N): clave auxiliar Dia|Hora Mes. No la lee
#     nadie: ni las hojas de calculo ni los cruces.
#   - Copia_Ventana (K): copia fila a fila de Ventana (L). Las dos
#     hojas de calculo si la usan, pero como "Ciclo de Carga del mes",
#     y la reponen al leer Medidores con reponer_auxiliares_medidores()
#     -- es una copia, no un dato propio.
COLUMNAS_AUXILIARES_MEDIDORES = COLUMNAS_VACIAS + [
    "Clave_Dia_HoraMes",
    "Copia_Ventana",
]

# Orden final de columnas de la hoja Medidores: el de LETRA_A_CAMPO
# (que es el de la planilla original) sin los auxiliares de arriba.
COLUMNAS_MEDIDORES_SALIDA = [
    campo for campo in LETRA_A_CAMPO.values()
    if campo not in COLUMNAS_AUXILIARES_MEDIDORES
]
