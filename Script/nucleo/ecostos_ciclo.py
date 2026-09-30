# -*- coding: utf-8 -*-
"""
E Costos, etapa 4: Ciclo, umbrales, AW, AX, AZ.
"""

import pandas as pd

from .utiles import (
    _columna_clave_vba, _normaliza_valor_vba, _tiene_valor,
    _valor_clave,
)


# ============================================================
# CALCULO E COSTOS (etapa 4): Subastas!N (Ciclo), umbrales, AW, AX, AZ
#
# El usuario entrego el documento de trazabilidad completo
# (docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md), que trae las
# tres piezas que faltaban y que estaban anotadas como bloqueantes:
#
# 1. Subastas!N ("Ciclo" en el archivo real -- ver correccion de
#    nombres en NOMBRES_SUBASTAS; se penso "Energía SSCC" hasta que
#    Libro1.xlsx trajo el encabezado real). Formula del libro
#    original (seccion 5.3 del documento):
#      =IFERROR(XLOOKUP(1,
#          ('Calculo E Costos'!$D$2:$D$50000=J3)*
#          ('Calculo E Costos'!$G$2:$G$50000=K3),
#          'Calculo E Costos'!$P$2:$P$50000,""),"")
#    O sea: es el "Ciclo de Carga del mes" (Calculo E Costos!P =
#    Copia_Ventana) de la primera fila de Calculo E Costos que
#    coincide en "Hora mes" (D) y "Configuracion" (G) -- coincide
#    exactamente con que el encabezado real de esta columna sea
#    "Ciclo", no una energia.
#
# 2. La tabla de umbrales de subida/bajada, que era EL bloqueante.
#    Vive en Subastas!S:W del libro original (confirmado con
#    Libro1.xlsx: S=Configuración, T=Ciclo, U=Clave, V=SUBIDA,
#    W=BAJADA):
#      U3 = S3&"&"&T3                              (clave)
#      V3 = COUNTIFS($N:$N,$T3,$D:$D,V$2,K:K,S3)   (V$2 = "SUBIDA")
#      W3 = COUNTIFS($N:$N,$T3,$D:$D,W$2,K:K,S3)   (W$2 = "BAJADA")
#    No es un archivo externo ni una hoja aparte: se deriva de
#    Subastas + Subastas!N (Ciclo), igual que la Prorrata SSCC. Y la
#    "dependencia circular" que se habia anotado NO existe: N
#    depende de Calculo E Costos!P (Copia_Ventana), que viene de
#    Medidores y ya esta disponible desde la etapa base, antes de
#    cualquier columna calculada.
#
# 3. El bloque "AU, AV, AW, AX Y AZ" de Actualizar_Calculos_Columnas
#    (modulo J_Calculo_Ecostos) + CrearDiccionarioUmbralesSubastas.
#
# LETRAS: se homologa por NOMBRE de columna real de Subastas
# (Sub_Baj, Hora_mes, Configuración -- confirmados y corregidos con
# Libro1.xlsx, ver NOMBRES_SUBASTAS), no por posicion, igual que en
# todo el resto del proyecto.
#
# Sigue fuera de alcance: toda la hoja "Calculo RE545". La columna
# AY del archivo real tampoco se calcula aca: no la escribe la macro
# J (no hay region de formulas para AY4:AY26787 en el documento, ver
# seccion 5.4) -- la macro salta de AX a AZ, y este codigo tambien.
# ============================================================

def calcular_subastas_ciclo(df_subastas, df_ecostos):
    """
    Resuelve Subastas!N ("Ciclo" en el archivo real -- antes
    documentado como "Energía SSCC" por error, corregido con
    Libro1.xlsx). Replica el XLOOKUP de arriba: para cada fila de
    Subastas, el Copia_Ventana ("Ciclo de Carga del mes") de la
    PRIMERA fila de Calculo E Costos que coincide en Hora Mes +
    central. Sin coincidencia -> "" (vacio), como el IFERROR original.

    df_ecostos: solo necesita las columnas de la etapa base ("Hora
    Mes", "clave", "Copia_Ventana") -- no depende de ninguna columna
    calculada, por eso no hay circularidad.

    La comparacion se hace con _normaliza_valor_vba en los dos lados
    (mismo criterio que calcular_l): el "=" de Excel compara valores,
    no texto crudo, asi que 7 y "7" tienen que cruzar igual.
    """

    indice = {}

    for hora_mes, central, ventana in zip(
        df_ecostos["Hora Mes"],
        df_ecostos["clave"],
        df_ecostos["Copia_Ventana"],
    ):
        clave = (
            _normaliza_valor_vba(hora_mes)
            + "¦" + _normaliza_valor_vba(central)
        )

        # XLOOKUP sin modo de busqueda: gana la primera coincidencia.
        if clave not in indice:
            indice[clave] = ventana

    claves = (
        _columna_clave_vba(df_subastas["Hora_mes"])
        + "¦" + _columna_clave_vba(df_subastas["Configuración"])
    )

    return claves.map(lambda clave: indice.get(clave, ""))


def _clave_central_ciclo(central, ciclo):
    """
    Clave "central&ciclo" de la tabla de umbrales. Replica las dos
    puntas de la homologacion, que en el original se arman distinto
    pero tienen que dar lo mismo:
      - Subastas!U = S & "&" & T (concatenacion de Excel)
      - Calculo E Costos = NormalizarValor(
            TextoSeguro(G) & "&" & TextoSeguro(P))
    _valor_clave() es lo que hace que un 3 y un 3.0 den "3" en las
    dos, igual que CStr/la concatenacion de Excel.
    """

    return _normaliza_valor_vba(
        _valor_clave(central) + "&" + _valor_clave(ciclo)
    )


def construir_dic_umbrales_subastas(df_subastas, ciclo_subastas):
    """
    Replica la tabla auxiliar Subastas!S:W (Configuración, Ciclo,
    Clave, SUBIDA, BAJADA -- confirmado con Libro1.xlsx) y
    CrearDiccionarioUmbralesSubastas: devuelve
    {"CENTRAL&CICLO": (umbral_subida, umbral_bajada)}.

    Cada umbral es el COUNTIFS del original: cuantas filas de
    Subastas tienen ese ciclo (columna "Ciclo" -- antes documentada
    por error como "Energía SSCC", ver calcular_subastas_ciclo), esa
    central (Configuración) y ese tipo (Sub_Baj = SUBIDA / BAJADA).

    La tabla del libro original es una lista fija de central x ciclo
    escrita a mano; aca las combinaciones salen de los datos. Es
    equivalente: una combinacion que la lista tiene pero los datos no
    daria 0/0, y con umbral 0 ninguna fila pasa el filtro
    "W <= umbral*4" (W arranca en 1), o sea AW = 0 igual.

    ciclo_subastas: la Serie que devuelve calcular_subastas_ciclo().
    """

    dic = {}

    tipos = df_subastas["Sub_Baj"].map(_normaliza_valor_vba)

    for central, ciclo, tipo in zip(
        df_subastas["Configuración"], ciclo_subastas, tipos
    ):
        if not _tiene_valor(ciclo):
            continue

        clave = _clave_central_ciclo(central, ciclo)

        if not clave:
            continue

        conteo = dic.setdefault(clave, [0, 0])

        if tipo == "SUBIDA":
            conteo[0] += 1
        elif tipo == "BAJADA":
            conteo[1] += 1

    return {clave: (valores[0], valores[1]) for clave, valores in dic.items()}


def calcular_aw_ax(df_ecostos, dic_umbrales):
    """
    Replica AW ("Descuento FD") y AX ("Total") del bloque
    "AU, AV, AW, AX Y AZ" de Actualizar_Calculos_Columnas.

    Por grupo (central=clave + ventana=Copia_Ventana):
      - clave de umbrales = central & "&" & ventana. Si no esta en el
        diccionario, el grupo no tiene umbrales validos y AW = 0.
      - promedioABW = promedio de AB entre las filas del grupo con
        W (Bloque ordenado) <= umbralBajada * 4.
      - promedioADW = promedio de AD entre las filas del grupo con
        W <= umbralSubida * 4.
        (Si, van cruzados: AB con el umbral de BAJADA y AD con el de
        SUBIDA. Asi esta en el VBA original, no es un tipeo.)
      - AW = (AE - AS) * promedioABW - (AF - AT) * promedioADW, y solo
        si hay umbrales validos, las dos cantidades promediadas son > 0
        y AE, AF, AS y AT son numeros validos. Si no, AW = 0.

    AX = AU + AV - AW, siempre (sin condiciones).
    """

    df = df_ecostos

    grupo = [df["clave"], df["Copia_Ventana"]]

    clave_umbral = pd.Series(
        [
            _clave_central_ciclo(central, ventana)
            for central, ventana in zip(df["clave"], df["Copia_Ventana"])
        ],
        index=df.index,
    )

    umbral_subida = clave_umbral.map(
        lambda clave: dic_umbrales.get(clave, (None, None))[0]
    )
    umbral_bajada = clave_umbral.map(
        lambda clave: dic_umbrales.get(clave, (None, None))[1]
    )

    umbrales_validos = umbral_subida.notna() & umbral_bajada.notna()

    w = pd.to_numeric(df["W"], errors="coerce")
    ab = pd.to_numeric(df["AB"], errors="coerce")
    ad = pd.to_numeric(df["AD"], errors="coerce")

    ab_contable = ab.where(
        umbrales_validos
        & (w <= pd.to_numeric(umbral_bajada, errors="coerce") * 4.0)
        & ab.notna()
    )
    ad_contable = ad.where(
        umbrales_validos
        & (w <= pd.to_numeric(umbral_subida, errors="coerce") * 4.0)
        & ad.notna()
    )

    cantidad_ab = ab_contable.groupby(grupo).transform("count")
    cantidad_ad = ad_contable.groupby(grupo).transform("count")

    promedio_ab = (
        ab_contable.groupby(grupo).transform("sum")
        / cantidad_ab.replace(0, pd.NA)
    ).fillna(0.0)
    promedio_ad = (
        ad_contable.groupby(grupo).transform("sum")
        / cantidad_ad.replace(0, pd.NA)
    ).fillna(0.0)

    ae = pd.to_numeric(df["AE"], errors="coerce")
    af = pd.to_numeric(df["AF"], errors="coerce")
    as_ = pd.to_numeric(df["AS"], errors="coerce")
    at = pd.to_numeric(df["AT"], errors="coerce")

    condicion_aw = (
        umbrales_validos
        & (cantidad_ab > 0)
        & (cantidad_ad > 0)
        & ae.notna()
        & af.notna()
        & as_.notna()
        & at.notna()
    )

    aw = (
        (ae - as_) * promedio_ab - (af - at) * promedio_ad
    ).where(condicion_aw, 0.0).fillna(0.0)

    ax = (
        pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)
        + pd.to_numeric(df["AV"], errors="coerce").fillna(0.0)
        - aw
    )

    return aw, ax


def calcular_az(df_ecostos):
    """
    Replica AZ ("Monto a compensar"): por grupo (central + ventana),
    (suma de AX - suma de U) / cantidad de filas del grupo, nunca
    negativo (max(0, ...)), y el MISMO valor en todas las filas del
    grupo. U no numerica cuenta como 0 (NumeroSeguro del original).
    """

    grupo = [df_ecostos["clave"], df_ecostos["Copia_Ventana"]]

    ax = pd.to_numeric(df_ecostos["AX"], errors="coerce").fillna(0.0)
    u = pd.to_numeric(df_ecostos["U"], errors="coerce").fillna(0.0)

    suma_ax = ax.groupby(grupo).transform("sum")
    suma_u = u.groupby(grupo).transform("sum")
    cantidad = ax.groupby(grupo).transform("size")

    return ((suma_ax - suma_u) / cantidad).clip(lower=0.0)
