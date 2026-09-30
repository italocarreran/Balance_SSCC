# -*- coding: utf-8 -*-
"""
E Costos, etapa 2: R, S, T, U, W, X, Y, AB:AF.
"""

import math
import pandas as pd

from .alertas import ALTA, Alerta, anotar_muchas
from .utiles import _texto_seguro, _tiene_valor, normalizar


def calcular_r_ecostos(df_ecostos):
    """
    Replica "Calculo E Costos"!R: ranking por grupo (central+
    ventana), ordenando por CMg descendente y 'Cuarto de Hora'
    descendente; las filas empatadas en ambos comparten el mismo
    ranking (la posicion donde empieza el empate), igual que un
    RANK() de Excel con "competition ranking" (no denso).

    Nombre con sufijo _ecostos a proposito: Medidores ya tiene su
    propia calcular_r() (Oferta_Completa_Dia, logica no relacionada)
    -- no renombrarla ni fusionarlas, son dos columnas "R" de hojas
    distintas con formulas distintas.
    """

    def _por_grupo(grupo):

        orden = grupo.sort_values(
            ["CMg", "Cuarto de Hora"],
            ascending=[False, False],
            kind="mergesort",
        )

        posiciones = pd.Series(range(1, len(orden) + 1), index=orden.index)

        posicion_min = posiciones.groupby(
            [orden["CMg"], orden["Cuarto de Hora"]]
        ).transform("min")

        return pd.DataFrame(
            {"R": posicion_min.reindex(grupo.index)}, index=grupo.index
        )

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["R"]


def calcular_s_t_u(df_ecostos):
    """
    Replica S (=I*Q), T (=J*Q) y U (=L*(S+T)). No dependen de
    agrupar por central/ventana.
    """

    cmg_num = pd.to_numeric(df_ecostos["CMg"], errors="coerce").fillna(0.0)

    s = df_ecostos["Energia_Positiva"].fillna(0.0) * cmg_num
    t = df_ecostos["Energia_Negativa"].fillna(0.0) * cmg_num
    u = df_ecostos["L"].astype(float) * (s + t)

    return s, t, u


def calcular_w_x(df_ecostos):
    """
    Replica X (=P, copia de Copia_Ventana) y W (contador que se
    reinicia a 1 cada vez que cambia X respecto de la fila anterior,
    GLOBAL -- no por grupo). La primera fila es una excepcion fiel al
    original: W toma el valor de 'Hora' en vez de 1, y esa diferencia
    se arrastra en el resto de su bloque (W sigue siendo "contador
    que suma 1", solo que ese primer bloque no arranca en 1).
    """

    x = df_ecostos["Copia_Ventana"].reset_index(drop=True)

    cambia = x.ne(x.shift())
    bloque = cambia.cumsum()

    w = (
        x.groupby(bloque)
        .cumcount()
        .add(1)
        .astype(float)
    )

    if len(w):
        primera_hora = pd.to_numeric(
            df_ecostos["Hora"].iloc[:1], errors="coerce"
        ).fillna(0.0).iloc[0]
        offset = primera_hora - 1.0
        primer_bloque = bloque.iloc[0]
        w = w.mask(bloque == primer_bloque, w + offset)

    w.index = df_ecostos.index
    x.index = df_ecostos.index

    return w, x


def calcular_y_ab_ac_ad(df_ecostos):
    """
    Replica Y, AB (a partir de las filas con L=1 e I!=0, ordenadas
    por CMg descendente) y AC, AD (analogo con J!=0, ordenadas por
    CMg ASCENDENTE), por grupo (central+ventana). Cada fila del grupo
    (en su orden original) recibe los valores de la fila en esa
    posicion dentro del orden calificado; si el grupo tiene menos
    filas calificadas que filas totales, las posiciones sobrantes
    toman los valores de las filas NO calificadas en su orden
    original (sin ordenar).
    """

    def _por_grupo(grupo):

        l_uno = grupo["L"] == 1
        calif_i = l_uno & (grupo["Energia_Positiva"] != 0)
        calif_j = l_uno & (grupo["Energia_Negativa"] != 0)

        orden_i = list(
            grupo.loc[calif_i]
            .sort_values("CMg", ascending=False, kind="mergesort")
            .index
        )
        fuente_i = orden_i + list(grupo.index[~calif_i])
        cantidad_calif_i = len(orden_i)

        orden_j = list(
            grupo.loc[calif_j]
            .sort_values("CMg", ascending=True, kind="mergesort")
            .index
        )
        fuente_j = orden_j + list(grupo.index[~calif_j])
        cantidad_calif_j = len(orden_j)

        destino = list(grupo.index)

        y, ab, ac, ad = [], [], [], []

        for posicion, idx_origen in enumerate(fuente_i):
            y.append(grupo.at[idx_origen, "Cuarto de Hora"])
            ab.append(
                grupo.at[idx_origen, "CMg"]
                if posicion < cantidad_calif_i
                else pd.NA
            )

        for posicion, idx_origen in enumerate(fuente_j):
            ac.append(grupo.at[idx_origen, "Cuarto de Hora"])
            ad.append(
                grupo.at[idx_origen, "CMg"]
                if posicion < cantidad_calif_j
                else pd.NA
            )

        return pd.DataFrame(
            {"Y": y, "AB": ab, "AC": ac, "AD": ad}, index=destino
        )

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["Y"], resultado["AB"], resultado["AC"], resultado["AD"]


def _calcular_asignacion_energia(bloque, energia_maxima, factor):
    """
    Replica CalcularAsignacionEnergia (AE/AF): distribuye
    energia_maxima en bloques de 15 minutos segun 'factor' (Pmax de
    la central) -- el bloque asigna 1 (completo) si cae dentro de la
    cantidad de bloques llenos, una fraccion al siguiente bloque si
    sobra un resto, y 0 al resto. Int() de VBA redondea hacia abajo
    incluso con numeros negativos, igual que math.floor.
    """

    cantidad_bloques = 4.0 * energia_maxima / factor / 1000.0
    parte_entera = math.floor(cantidad_bloques)
    fraccion = cantidad_bloques - parte_entera

    if bloque <= cantidad_bloques:
        proporcion = 1.0
    elif fraccion != 0 and bloque == parte_entera + 1.0:
        proporcion = fraccion
    else:
        proporcion = 0.0

    return proporcion * factor / 4.0 * 1000.0


def calcular_ae_af(df_ecostos, dic_factor, registrar=print):
    """
    Replica AE ("Energía descargada") y AF ("Energía cargada"):
    asigna, dentro de cada grupo (central+ventana), la energia
    maxima acumulada (N para AE, O para AF) en bloques segun el
    orden W de cada fila y un 'factor' por central (Resumen BESS!
    Pmax (MW), ver construir_dic_resumen_factor). Requiere que N, O
    y W ya esten calculados en df_ecostos.

    Si no hay factor para la central (no encontrada) o el factor es
    0 o no numerico, AE/AF quedan en blanco (pd.NA) -- equivalente a
    los #N/A / #VALOR! / #DIV/0! del original, sin fabricar un tipo
    de error de Excel en Python.
    """

    maximo_n = (
        df_ecostos.groupby(["clave", "Copia_Ventana"])["N"].transform("max")
    )
    maximo_o = (
        df_ecostos.groupby(["clave", "Copia_Ventana"])["O"].transform("max")
    )

    factor = df_ecostos["clave"].map(
        lambda valor: dic_factor.get(normalizar(valor), pd.NA)
    )

    # _avisar_claves_sin_mapeo() cubre el Pmax ausente o en blanco; el
    # Pmax que SI esta pero vale 0 (o no es numero) tambien deja AE/AF
    # vacias y no lo veria nadie, asi que se avisa aparte.
    sin_factor_util = sorted({
        _texto_seguro(central)
        for central, f in zip(df_ecostos["clave"], factor)
        if _tiene_valor(central)
        and _tiene_valor(f)
        and (not isinstance(f, (int, float)) or f == 0)
    })

    if sin_factor_util:
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "MAE-003", ALTA, "Calculo E Costos",
                    "Pmax (MW) en 0 o no numerico.",
                    central=central, valor_encontrado="0 o no numerico",
                    valor_esperado="un Pmax (MW) numerico y distinto de 0",
                    accion="AE y AF quedan vacias para esa central",
                    origen_control="CATALOGO AUX-007",
                )
                for central in sin_factor_util
            ],
            f"  [{ALTA}] MAE-003: Calculo E Costos: Pmax (MW) en 0 o no "
            f"numerico para {len(sin_factor_util):,} central(es) "
            f"({', '.join(repr(v) for v in sin_factor_util[:15])}); "
            f"AE y AF quedan vacias para esas centrales.",
        )

    ae, af = [], []

    for w, mn, mo, f in zip(df_ecostos["W"], maximo_n, maximo_o, factor):

        if pd.isna(f) or not isinstance(f, (int, float)) or f == 0:
            ae.append(pd.NA)
            af.append(pd.NA)
            continue

        ae.append(_calcular_asignacion_energia(w, mn, f))
        af.append(-_calcular_asignacion_energia(w, mo, f))

    return (
        pd.Series(ae, index=df_ecostos.index),
        pd.Series(af, index=df_ecostos.index),
    )
