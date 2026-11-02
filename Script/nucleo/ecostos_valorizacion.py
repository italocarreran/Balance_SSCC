# -*- coding: utf-8 -*-
"""
E Costos, etapa 4: acumulados, asignacion y valorizacion de la energia
ajustada (AU:BA) y el monto a compensar (BC).
"""

import pandas as pd

from .alertas import MEDIA, Alerta, anotar
from .columnas_compartidas import calcular_n_o
from .ecostos_columnas import _calcular_asignacion_energia
from .utiles import normalizar


# ============================================================
# CALCULO E COSTOS (etapa 4): AU:BA y BC -- metodo nuevo
#
# Macro SSCC_Calcular_Integrado. Con la energia ya ajustada por
# prorrata y FD (AS, AT, ver ecostos_prorratas), se repite el esquema
# de siempre, pero sobre esa energia:
#
#   AU / AV  acumulados de AS / AT por central + ciclo, igual que N / O
#            (bloque horario descendente, solo filas con L = 1; AV con
#            el signo cambiado)
#   AW / AX  el maximo de AU / AV del grupo repartido en bloques de
#            Pmax/4 kWh segun el "Bloque ordenado" (W), igual que
#            AE / AF
#   AY / AZ  AW / AX por el promedio del CMg de la curva monotona (AB /
#            AD) en las filas con asignacion <> 0. La curva es la de
#            siempre: no se vuelve a ordenar con la energia ajustada.
#            Se mantiene la condicion global por ciclo: descarga total
#            > 10 y carga total < -10, sumando todas las centrales.
#   BA       AY + AZ. Ya NO hay "Descuento FD" (el FD esta dentro de la
#            energia) y por eso tampoco la tabla de umbrales de
#            subida/bajada (Subastas!S:W) ni el Ciclo de Subastas.
#   BC       Monto a compensar = MAX(0, (SUMA(BA) - SUMA(U)) / filas)
#            por grupo, el mismo valor en todas sus filas.
#
# El metodo anterior (descuento FD sobre AE / AF) queda en la rama
# Definitivo-Agosto.
# ============================================================

# Por debajo de esto un maximo de AU / AV se toma como "sin energia".
_CERO = 1e-12


def _grupo(df_ecostos):
    return [df_ecostos["clave"], df_ecostos["Copia_Ventana"]]


def calcular_acumulados_ajustados(df_ecostos):
    """
    Replica AU y AV: N y O (calcular_n_o) calculados sobre la energia
    ajustada AS / AT en vez de la medida.
    """

    n, o = calcular_n_o(
        df_ecostos.assign(
            Energia_Positiva=df_ecostos["AS"],
            Energia_Negativa=df_ecostos["AT"],
        )
    )

    return n.reindex(df_ecostos.index), o.reindex(df_ecostos.index)


def calcular_asignacion_ajustada(df_ecostos, dic_factor):
    """
    Replica AW ("Asignacion Descarga kWh") y AX ("Asignacion Carga
    kWh"): el mismo reparto de AE / AF (_calcular_asignacion_energia)
    con el maximo de AU / AV del grupo.

    Un grupo sin energia ajustada queda en 0 sin pedir Pmax. Si la
    central no tiene Pmax valida (> 0) o la fila no tiene W, tambien 0:
    la macro nueva no deja errores (la falta de Pmax ya la avisan
    MAE-002 / MAE-003).
    """

    grupo = _grupo(df_ecostos)
    au = pd.to_numeric(df_ecostos["AU"], errors="coerce").fillna(0.0)
    av = pd.to_numeric(df_ecostos["AV"], errors="coerce").fillna(0.0)

    maximo_au = au.groupby(grupo).transform("max")
    maximo_av = av.groupby(grupo).transform("max")

    factor = pd.to_numeric(
        df_ecostos["clave"].map(
            lambda valor: dic_factor.get(normalizar(valor), pd.NA)
        ),
        errors="coerce",
    )
    w = pd.to_numeric(df_ecostos["W"], errors="coerce")

    aw, ax = [], []

    for bloque, mn, mo, f in zip(w, maximo_au, maximo_av, factor):

        hay_energia = abs(mn) > _CERO or abs(mo) > _CERO

        if not hay_energia or pd.isna(f) or f <= 0 or pd.isna(bloque) \
                or bloque < 1:
            aw.append(0.0)
            ax.append(0.0)
            continue

        aw.append(_calcular_asignacion_energia(bloque, mn, f))
        ax.append(-_calcular_asignacion_energia(bloque, mo, f))

    return (
        pd.Series(aw, index=df_ecostos.index),
        pd.Series(ax, index=df_ecostos.index),
    )


def calcular_valorizacion(df_ecostos, registrar=print):
    """
    Replica AY ("Valorizacion Descarga [$]"), AZ ("Valorizacion Carga
    [$]") y BA ("Total [CLP]"):

      AY = promedio de AB (curva CMg descarga) entre las filas del
           grupo con AW <> 0, por AW; solo si la descarga total del
           ciclo (todas las centrales) es > 10.
      AZ = idem con AD (curva CMg carga), AX y carga total < -10.
      BA = AY + AZ, sin descuento FD.

    Una fila con asignacion pero sin precio en la curva (mas bloques
    asignados que horas con subasta) no entra al promedio y se avisa.
    """

    grupo = _grupo(df_ecostos)

    aw = pd.to_numeric(df_ecostos["AW"], errors="coerce").fillna(0.0)
    ax = pd.to_numeric(df_ecostos["AX"], errors="coerce").fillna(0.0)
    ab = pd.to_numeric(df_ecostos["AB"], errors="coerce")
    ad = pd.to_numeric(df_ecostos["AD"], errors="coerce")

    ab_contable = ab.where(aw.ne(0))
    ad_contable = ad.where(ax.ne(0))

    def _promedio(serie):
        cantidad = serie.groupby(grupo).transform("count")
        suma = serie.groupby(grupo).transform("sum")
        return (suma / cantidad.replace(0, pd.NA)).astype("Float64")

    promedio_ab = _promedio(ab_contable)
    promedio_ad = _promedio(ad_contable)

    # Suma GLOBAL por ciclo (P), con todas las centrales y la energia
    # medida de la hoja, igual que antes.
    ciclo = df_ecostos["Copia_Ventana"]
    suma_i = df_ecostos["Energia_Positiva"].groupby(ciclo).transform("sum")
    suma_j = df_ecostos["Energia_Negativa"].groupby(ciclo).transform("sum")

    ay = (promedio_ab * aw).where(promedio_ab.notna() & (suma_i > 10), 0.0)
    az = (promedio_ad * ax).where(promedio_ad.notna() & (suma_j < -10), 0.0)
    ay = ay.astype(float)
    az = az.astype(float)

    sin_precio = int((aw.ne(0) & ab.isna()).sum() + (ax.ne(0) & ad.isna()).sum())
    if sin_precio:
        anotar(registrar, Alerta(
            "VAL-001", MEDIA, "Calculo E Costos",
            f"{sin_precio:,} fila(s) con energia asignada (AW / AX) y sin "
            f"precio en la curva monotona (AB / AD).",
            valor_encontrado=f"{sin_precio:,} filas",
            accion="esas filas no entran al promedio de CMg de su grupo",
            origen_control="CONTROL NUEVO",
        ))

    return ay, az, ay + az


def calcular_monto_a_compensar(df_ecostos):
    """
    Replica BC ("Monto a compensar"): por grupo (central + ciclo),
    (suma de BA - suma de U) / cantidad de filas del grupo, nunca
    negativo, y el MISMO valor en todas las filas del grupo. U no
    numerica cuenta como 0.
    """

    grupo = _grupo(df_ecostos)

    ba = pd.to_numeric(df_ecostos["BA"], errors="coerce").fillna(0.0)
    u = pd.to_numeric(df_ecostos["U"], errors="coerce").fillna(0.0)

    suma_ba = ba.groupby(grupo).transform("sum")
    suma_u = u.groupby(grupo).transform("sum")
    cantidad = ba.groupby(grupo).transform("size")

    return ((suma_ba - suma_u) / cantidad).clip(lower=0.0)
