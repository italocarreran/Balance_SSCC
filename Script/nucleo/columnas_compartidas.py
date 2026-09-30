# -*- coding: utf-8 -*-
"""
L, M y N/O: iguales en las dos hojas de calculo.
"""

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

    def _por_grupo(grupo):

        valido = grupo["L"] == 1

        i_valido = grupo["Energia_Positiva"].where(valido, 0.0)
        j_valido = grupo["Energia_Negativa"].where(valido, 0.0)

        suma_i_por_f = i_valido.groupby(grupo["Cuarto de Hora"]).sum()
        suma_j_por_f = j_valido.groupby(grupo["Cuarto de Hora"]).sum()

        acumulado_i = suma_i_por_f.sort_index(ascending=False).cumsum()
        acumulado_j = suma_j_por_f.sort_index(ascending=False).cumsum()

        n = grupo["Cuarto de Hora"].map(acumulado_i)
        o = -grupo["Cuarto de Hora"].map(acumulado_j)

        return pd.DataFrame({"N": n, "O": o}, index=grupo.index)

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["N"], resultado["O"]


def calcular_m(df_ecostos, umbral_soc_minimo):
    """
    Replica M ("SoC sobre el minimo"): 1 si SoC > umbral_soc_minimo
    (ver construir_dic_resumen_factor), si no 0.
    """

    soc = pd.to_numeric(df_ecostos["SoC"], errors="coerce").fillna(0.0)

    return (soc > umbral_soc_minimo).astype("int64")
