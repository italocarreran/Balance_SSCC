# -*- coding: utf-8 -*-
"""
L, M y N/O: iguales en las dos hojas de calculo.
"""

import numpy as np
import pandas as pd

from .utiles import _columna_clave_vba, _normaliza_valor_vba


def _construir_set_subastas_tipo(df_subastas):
    """
    Conjunto de claves "central¦mes¦dia¦hora" que SI participaron en
    una subasta de subida o bajada (Subastas!Sub_Baj en {BAJADA,
    SUBIDA}), para la columna L.
    """

    tipo = df_subastas["Sub_Baj"].map(_normaliza_valor_vba)
    filtro = tipo.isin(["BAJADA", "SUBIDA"])

    sub = df_subastas.loc[filtro]

    claves = (
        _columna_clave_vba(sub["Configuración"])
        + "¦" + _columna_clave_vba(sub["Mes"])
        + "¦" + _columna_clave_vba(sub["Dia"])
        + "¦" + _columna_clave_vba(sub["Hora_dia"])
    )

    return set(claves)


def calcular_l(df_ecostos, df_subastas):
    """
    Replica la columna L: 1 si la central+mes+dia+hora de la fila
    existe en Subastas como registro BAJADA o SUBIDA, si no 0.
    """

    claves_subasta = _construir_set_subastas_tipo(df_subastas)

    clave_fila = (
        _columna_clave_vba(df_ecostos["clave"])
        + "¦" + _columna_clave_vba(df_ecostos["Mes"])
        + "¦" + _columna_clave_vba(df_ecostos["Dia"])
        + "¦" + _columna_clave_vba(df_ecostos["Hora"])
    )

    return clave_fila.isin(claves_subasta).astype("int64")


def calcular_n_o(df_ecostos):
    """
    Replica N y O: por grupo (central=clave, ventana=Copia_Ventana),
    ordenando por 'Cuarto de Hora' descendente, suma acumulada de I
    (N) y de -J (O), solo contando filas con L=1, y repartida a TODAS
    las filas que comparten el mismo 'Cuarto de Hora' dentro del
    grupo (no solo a las que tienen L=1).
    """

    # Vectorizado (antes, un groupby.apply por central + ventana: varios
    # segundos por corrida). Mismas sumas en el mismo orden: por grupo y
    # Cuarto de Hora se suma en el orden de las filas, y se acumula de
    # mayor a menor Cuarto de Hora. Las filas sin clave, ventana o cuarto
    # quedan en NaN, como antes (el groupby las dejaba fuera).
    valido = df_ecostos["L"] == 1
    tabla = pd.DataFrame({
        "g1": df_ecostos["clave"],
        "g2": df_ecostos["Copia_Ventana"],
        "q": df_ecostos["Cuarto de Hora"],
        "i": df_ecostos["Energia_Positiva"].where(valido, 0.0),
        "j": df_ecostos["Energia_Negativa"].where(valido, 0.0),
    })
    sumas = tabla.groupby(["g1", "g2", "q"], sort=False)[["i", "j"]].sum()
    sumas = sumas.sort_index(level="q", ascending=False, sort_remaining=False)
    sumas = sumas.sort_index(level=["g1", "g2"], sort_remaining=False, kind="mergesort")
    # cumsum de numpy por grupo (como el Series.cumsum de antes): el
    # cumsum de un groupby usa otra suma (compensada) y difiere en el
    # ultimo decimal.
    acumulado = sumas.copy()
    grupos = sumas.index.droplevel("q")
    cortes = np.flatnonzero(grupos[1:] != grupos[:-1]) + 1
    for columna in ("i", "j"):
        valores = sumas[columna].to_numpy(dtype=float)
        acumulado[columna] = np.concatenate([
            np.cumsum(tramo) for tramo in np.split(valores, cortes)
        ]) if len(valores) else valores

    claves = pd.MultiIndex.from_arrays([tabla["g1"], tabla["g2"], tabla["q"]])
    n = pd.Series(acumulado["i"].reindex(claves).to_numpy(), index=df_ecostos.index)
    o = -pd.Series(acumulado["j"].reindex(claves).to_numpy(), index=df_ecostos.index)
    # Como el groupby de antes: las filas sin clave o sin ventana no
    # vienen en el resultado (en los datos reales no las hay).
    con_grupo = tabla["g1"].notna() & tabla["g2"].notna()
    return n[con_grupo], o[con_grupo]
