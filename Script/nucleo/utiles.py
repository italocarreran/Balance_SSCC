# -*- coding: utf-8 -*-
"""
Normalizacion y el error de entrada: lo que usa todo.
"""

import pandas as pd
import re
import unicodedata


# ============================================================
# UTILIDADES
# ============================================================

def normalizar(texto):
    """Minusculas, sin tildes, espacios colapsados."""

    if texto is None:
        return ""

    # Una celda vacia llega como NaN y str(NaN) es 'nan', que
    # se confundiria con un nombre real.
    try:
        if pd.isna(texto):
            return ""
    except (TypeError, ValueError):
        pass

    nfkd = unicodedata.normalize("NFKD", str(texto))

    sin_tildes = "".join(
        c for c in nfkd
        if not unicodedata.combining(c)
    )

    return re.sub(r"\s+", " ", sin_tildes).strip().lower()


class ErrorEntrada(Exception):
    """Problema en las entradas que impide continuar."""


# ============================================================
# OFERTAS SSCC
#
# Replica, en este orden, las macros y formulas de la planilla 11:
#   1. Generar_Resumen_Ofertas_SSCC        -> construir_resumen_ofertas_sscc
#   2. OSSCC_CargarResumenEnMedidores      -> cargar_resumen_en_medidores
#   3. Formula de Medidores!V y R          -> calcular_r
#   4. Formula de Medidores!S              -> calcular_s
#   5. Resumir_Medidores_Central_Ventana_
#      Oferta_Completa                     -> construir_resumen_ventana_oferta
#   6. Formula de Medidores!T              -> calcular_t
# ============================================================

def _texto_seguro(valor):
    """Replica OSSCC_TextoSeguro / RESOF: '' para vacio/NaN, si no str().strip()."""

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    return str(valor).strip()


def _tiene_valor(valor):
    """Replica OSSCC_TieneValor / RESOF_TieneValor."""

    return len(_texto_seguro(valor)) > 0


def _es_numero(valor):
    if not _tiene_valor(valor):
        return False

    try:
        float(valor)
        return True
    except (TypeError, ValueError):
        return False


def _valor_clave(valor):
    """
    Replica OSSCC_ValorClave / RESOF_ValorClave: formatea un numero
    sin ceros/decimales sobrantes, o el texto tal cual si no es numero.
    """

    if not _es_numero(valor):
        return _texto_seguro(valor)

    numero = float(valor)

    if numero == int(numero):
        return str(int(numero))

    return repr(numero)


def _entero_a_texto(valor):
    """
    1 -> '1' (no '1.0'): la Clave horaria es un pegado de textos y un
    decimal de mas la dejaria distinta de la de la planilla.
    """

    if pd.isna(valor):
        return ""

    try:
        return str(int(valor))
    except (TypeError, ValueError):
        return _texto_seguro(valor)


# ============================================================
# CALCULO E COSTOS (etapa 2): L, M, N, O, R, S, T, U, W, X, Y, AB,
# AC, AD, AE, AF
#
# Replica esa parte de Actualizar_Calculos_Columnas (modulo
# J_Calculo_Ecostos, ver plan seccion 25.6/25.7). El usuario confirmo
# con un archivo real que la hoja "Resumen" del libro original (que
# M usa para el umbral, y AE/AF para el factor por central) es la
# MISMA tabla que Centrales.xlsx!Resumen BESS -- no hacia falta una
# hoja nueva. Ver construir_dic_resumen_factor().
#
# Quedan FUERA de esta etapa (bloqueados): AG:AZ -- dependen de la
# tabla dinamica "Prorrata SSCC" (todavia no se construye en Python,
# aunque el usuario ya confirmo su estructura: Filas: Configuración,
# Hora_mes / Columnas: Control / Valores: Cuenta de Sub_Baj) y de un
# umbral de subida/bajada por central+ventana (en el .xlsm original
# vive en Subastas!R:V o U:W segun la fuente -- la posicion exacta
# todavia no esta clara ni siquiera con el archivo de encabezados
# real, ver BITACORA) y de una categoria "CTF" en FD que no existe en
# nuestra hoja FD (que solo tiene CSF/CPF).
#
# ADVERTENCIA sobre L (parcialmente resuelta, ver BITACORA): el VBA
# original arma la clave de match contra Subastas usando columnas por
# posicion que, en el archivo de trazabilidad, no coincidian con los
# encabezados reales. El usuario confirmo que el "tipo" (BAJADA/
# SUBIDA) esta en Subastas!Sub_Baj. La central equivalente se uso
# como Subastas!Configuración, y un archivo real posterior confirmo
# que "Calculo E Costos"!G se llama literalmente "Configuracion" --
# el mismo campo en ambas hojas, lo que da bastante mas confianza en
# esta homologacion (aunque no es una confirmacion letra por letra
# del match, solo de que el NOMBRE del campo coincide en las dos
# hojas). Si al correr esto la cantidad de filas con L=1 sale
# sospechosamente baja o en cero, sigue siendo la primera sospechosa
# a revisar.
# ============================================================

def _normaliza_valor_vba(valor):
    """
    Replica NormalizarValor: texto en mayusculas y recortado. Los
    numeros se renderizan sin decimales de mas (igual que CStr en
    VBA: 7 -> "7", no "7.0"), para poder armar claves compuestas
    comparables entre hojas.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, bool):
        texto = str(valor)
    elif isinstance(valor, (int, float)):
        numero = float(valor)
        texto = str(int(numero)) if numero.is_integer() else str(numero)
    else:
        texto = str(valor)

    return texto.strip().upper()


def _columna_clave_vba(serie):
    """
    serie.map(_normaliza_valor_vba), pero forzando el resultado a
    dtype string SIEMPRE, incluso cuando serie esta vacia (0 filas).

    Trampa real (encontrada con datos reales, no en los sinteticos):
    pandas.Series.map() sobre una Series vacia es un no-op que NO
    llama a la funcion -- devuelve una Series vacia con el MISMO
    dtype que tenia antes de mapear. Si esa columna original era
    numerica (ej. "Mes"/"Hora Mes" leida como int64 desde Excel) y
    el resultado se concatena con "+" contra una Series de texto (u
    otro separador), la suma falla: numpy no sabe sumar int64 con
    texto, aunque las dos esten vacias.

    Pasa cuando el filtro previo (ej. Subastas!Sub_Baj en {BAJADA,
    SUBIDA}) no encuentra ninguna fila -- un caso real y valido (no
    hay ninguna subasta en el periodo), no un error de datos. Forzar
    .astype(str) despues del .map() corrige el dtype en los dos
    casos (vacio o no), sin cambiar ningun valor.
    """

    return serie.map(_normaliza_valor_vba).astype(str)
