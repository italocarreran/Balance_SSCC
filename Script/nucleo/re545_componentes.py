# -*- coding: utf-8 -*-
"""
RE545: Componente 1 y Componente 2 (BI:CE).
"""

import pandas as pd

from .avisos import _avisar_claves_sin_mapeo
from .parametros import (
    ARCHIVO_CENTRALES, HOJA_CALCULO_RE545, HOJA_RESUMEN_BESS,
)
from .re545_resumen import (
    _clave_grupo_re545, _mapa_resumen_por_grupo, calcular_bv_re545,
)
from .utiles import _normaliza_valor_vba, normalizar


def calcular_bi_bj_re545(df_re545):
    """
    Replica BI ("Orden") y BJ ("Periodo"), las dos con el mismo salto
    de 4 filas del original (4 bloques de 15 minutos por hora):
    BI arranca en 1 y suma 1 cada vez que la fila de 4 mas arriba
    tiene la misma ventana; BJ repite el periodo de esa misma fila.

    Las 4 primeras filas son constantes en el .xlsm (BI = 1 y BJ = 0,
    15, 30, 45). Aca se toman de la propia columna "Minutos" para no
    hardcodear una grilla de 15 minutos que el resto del codigo no
    asume en ningun lado.
    """

    df = df_re545.reset_index(drop=True)
    n = len(df)

    ventana = list(df["T"])
    minutos = list(df["Minutos"])

    bi = [1.0] * n
    bj = [None] * n

    for i in range(n):

        if i < 4:
            bi[i] = 1.0
            bj[i] = minutos[i]
            continue

        bj[i] = bj[i - 4]

        if _normaliza_valor_vba(ventana[i]) == _normaliza_valor_vba(
            ventana[i - 4]
        ):
            bi[i] = bi[i - 4] + 1.0
        else:
            bi[i] = 1.0

    return (
        pd.Series(bi, index=df_re545.index),
        pd.Series(bj, index=df_re545.index),
    )


def calcular_bk_bl_bm_bs_re545(df_re545):
    """
    Replica BK, BL, BM y BS.

    TRAMPA REAL (encontrada comparando fila a fila contra la hoja
    "RE545 P11" que el usuario pego -- planilla 11 real -- y
    confirmada contra las formulas guardadas del archivo real,
    docs/Calculo_RE545_reducido_para_IA.xlsx, hoja Mapa_Formulas):

        BK4 = SUMIFS(R:R, G:G,G4, S:S,BI4, T:T,T4, E:E,BJ4)
        BL4 = SUMIFS(Q:Q, S:S,BI4, E:E,BJ4, G:G,G4, T:T,T4)
        BS4 = INDEX(I, MATCH(1, (BR=BR4)*(S=BI4)*(G=G4)*(E=BJ4), 0))

    Las tres NO agrupan comparando BI contra BI: el criterio de cada
    SUMIFS/MATCH es "S:S,BI4" -- la columna S (ranking cmg) de las
    OTRAS filas contra el BI (Orden) de LA FILA ACTUAL. O sea hay dos
    claves distintas: la clave de ACUMULACION (para saber que sumar)
    usa el S de cada fila; la clave de BUSQUEDA (para saber que leer)
    usa el BI de la fila. Antes se usaba BI de los dos lados por
    error (mismo "trampa de letras" que ya afecto a otras columnas de
    este proyecto -- R/S/T/U/V significan cosas distintas segun la
    hoja, y aca ademas se cruzan entre si DENTRO de la misma hoja).

      BK = suma de R (CMg Promedio) de las filas cuyo S == BI(fila)
      BL = suma de Q (CMg) de las filas cuyo S == BI(fila)
      BM = el k-esimo valor mas grande de BL entre TODAS las filas
           (de toda la hoja, no del grupo) cuyo BK es igual al de la
           fila, con k = Periodo/15 + 1
      BS = el I (Descarga kWh) de la PRIMERA fila cuyo S == BI(fila)
    """

    df = df_re545.reset_index(drop=True)

    r = pd.to_numeric(df["R"], errors="coerce").fillna(0.0)
    q = pd.to_numeric(df["CMg"], errors="coerce").fillna(0.0)
    i_energia = pd.to_numeric(df["Energia_Positiva"], errors="coerce")

    # Clave de ACUMULACION: se arma con S (ranking cmg) de cada fila,
    # que es lo que el SUMIFS/MATCH real compara contra "BI4".
    claves_acumulacion = [
        (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ranking),
            _normaliza_valor_vba(ventana),
            _normaliza_valor_vba(periodo),
        )
        for central, ranking, ventana, periodo in zip(
            df["clave"], df["S"], df["T"], df["BJ"]
        )
    ]

    # Clave de BUSQUEDA (una por fila): se arma con el BI (Orden) de
    # esa misma fila, el criterio fijo del SUMIFS/MATCH real.
    claves_busqueda = [
        (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(orden),
            _normaliza_valor_vba(ventana),
            _normaliza_valor_vba(periodo),
        )
        for central, orden, ventana, periodo in zip(
            df["clave"], df["BI"], df["T"], df["BJ"]
        )
    ]

    suma_r = {}
    suma_q = {}
    primer_i = {}

    for posicion, clave in enumerate(claves_acumulacion):
        suma_r[clave] = suma_r.get(clave, 0.0) + float(r.iloc[posicion])
        suma_q[clave] = suma_q.get(clave, 0.0) + float(q.iloc[posicion])
        if clave not in primer_i:
            primer_i[clave] = i_energia.iloc[posicion]

    # SUMIFS sin match da 0; el INDEX/MATCH de BS, en cambio, cae en
    # el IFERROR real y da "" (blanco), no 0.
    bk = [suma_r.get(clave, 0.0) for clave in claves_busqueda]
    bl = [suma_q.get(clave, 0.0) for clave in claves_busqueda]
    bs = [primer_i.get(clave, pd.NA) for clave in claves_busqueda]

    # BM: LARGE(IF(BK = BK(i), BL), Periodo/15 + 1). El IF recorre
    # TODA la columna, no el grupo: se indexa por valor de BK.
    por_bk = {}

    for valor_bk, valor_bl in zip(bk, bl):
        por_bk.setdefault(round(float(valor_bk), 9), []).append(float(valor_bl))

    for lista in por_bk.values():
        lista.sort(reverse=True)

    bm = []

    for posicion, valor_bk in enumerate(bk):

        periodo = pd.to_numeric(
            pd.Series([df["BJ"].iloc[posicion]]), errors="coerce"
        ).iloc[0]

        if pd.isna(periodo):
            bm.append(pd.NA)
            continue

        k = int(periodo / 15) + 1
        lista = por_bk.get(round(float(valor_bk), 9), [])

        # LARGE con k fuera de rango da #NUM! -> el IFERROR lo deja "".
        bm.append(lista[k - 1] if 1 <= k <= len(lista) else pd.NA)

    indice = df_re545.index

    return (
        pd.Series(bk, index=indice),
        pd.Series(bl, index=indice),
        pd.Series(bm, index=indice),
        pd.Series(bs, index=indice),
    )


def calcular_componentes_re545(
    df_re545, df_resumen_re545, dic_factor, registrar=print
):
    """
    Replica BI:CE (menos BV, que ya calcula calcular_bv_re545) y
    devuelve un diccionario columna interna -> Serie. Ver el
    comentario de seccion para la formula de cada una.

    dic_factor: central -> "Pmax (MW)" (indice 2 del VLOOKUP de BN).
    df_resumen_re545: la tabla AW:BG ya construida (de ahi salen BC,
    BF y BG por central+ventana).
    """

    df = df_re545.reset_index(drop=True).copy()
    n = len(df)

    # Sin Pmax, BN queda vacio; BQ = BS + BN lo convierte en 0 con
    # fillna(0.0), asi que el faltante desaparece sin dejar rastro.
    # El aviso va antes de calcular nada para que salga igual si algo
    # mas adelante falla.
    _avisar_claves_sin_mapeo(
        df["clave"], dic_factor, "Central sin Pmax (MW)",
        f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}", registrar,
        id_alerta="MAE-002", etapa=HOJA_CALCULO_RE545,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="BN queda vacia y BQ = BS + BN la toma como 0",
        origen_control="CATALOGO AUX-008",
    )

    bi, bj = calcular_bi_bj_re545(df)
    df["BI"] = bi
    df["BJ"] = bj

    bk, bl, bm, bs = calcular_bk_bl_bm_bs_re545(df)
    df["BK"], df["BL"], df["BM"], df["BS"] = bk, bl, bm, bs

    br = df["T"]
    claves = [
        _clave_grupo_re545(central, ventana)
        for central, ventana in zip(df["clave"], br)
    ]

    dic_bc = _mapa_resumen_por_grupo(df_resumen_re545, "Edisp_T")
    dic_bf = _mapa_resumen_por_grupo(df_resumen_re545, "Margen ultima hora")
    dic_bg = _mapa_resumen_por_grupo(df_resumen_re545, "flag ultima hora")

    bv = calcular_bv_re545(df)

    suma_bv = {}
    for clave, valor in zip(claves, bv):
        suma_bv[clave] = suma_bv.get(clave, 0.0) + float(valor)

    pmax = [
        dic_factor.get(normalizar(central), pd.NA) for central in df["clave"]
    ]

    bs_num = pd.to_numeric(df["BS"], errors="coerce")

    # --- BN: recursion hacia ABAJO (suma de los BN anteriores) ---
    bn = [0.0] * n
    acumulado_bn = {}

    for posicion in range(n):

        clave = claves[posicion]
        capacidad_pmax = pmax[posicion]
        valor_bs = bs_num.iloc[posicion]

        if pd.isna(capacidad_pmax) or pd.isna(valor_bs):
            bn[posicion] = pd.NA
            continue

        disponible = max(
            0.0, float(capacidad_pmax) * 1000.0 / 4.0 - float(valor_bs)
        )
        techo = dic_bc.get(clave, 0.0) - acumulado_bn.get(clave, 0.0)

        valor = max(0.0, min(disponible, techo))

        bn[posicion] = valor
        acumulado_bn[clave] = acumulado_bn.get(clave, 0.0) + valor

    bn = pd.Series(bn, index=df.index)

    bm_num = pd.to_numeric(df["BM"], errors="coerce")
    bn_num = pd.to_numeric(bn, errors="coerce")

    bo = bn_num * bm_num
    bq = bs_num.fillna(0.0) + bn_num.fillna(0.0)

    # --- BU: recursion hacia ARRIBA (BT mira las filas siguientes) ---
    bt = [0.0] * n
    bu = [0.0] * n
    acumulado_bu = {}

    for posicion in range(n - 1, -1, -1):

        clave = claves[posicion]

        # BT = suma de los BU de las filas POSTERIORES del grupo.
        bt[posicion] = acumulado_bu.get(clave, 0.0)

        valor_bs = bs_num.iloc[posicion]

        if pd.isna(valor_bs) or float(valor_bs) == 0.0:
            bu[posicion] = 0.0
        else:
            techo = (
                dic_bc.get(clave, 0.0)
                - bt[posicion]
                - suma_bv.get(clave, 0.0)
            )
            bu[posicion] = max(0.0, min(float(bq.iloc[posicion]), techo))

        acumulado_bu[clave] = acumulado_bu.get(clave, 0.0) + bu[posicion]

    bt = pd.Series(bt, index=df.index)
    bu = pd.Series(bu, index=df.index)

    bw = (
        pd.to_numeric(df["Energia_Positiva"], errors="coerce").fillna(0.0)
        + pd.to_numeric(df["Energia_Negativa"], errors="coerce").fillna(0.0)
    )

    bx = pd.Series([dic_bf.get(clave, 0.0) for clave in claves], index=df.index)

    bz = pd.Series(
        [
            (0.0 if bw.iloc[posicion] < 0 else 8.0)
            * dic_bg.get(claves[posicion], 0.0)
            for posicion in range(n)
        ],
        index=df.index,
    )

    # --- BY: suma de los BZ ANTERIORES del grupo (BZ ya esta) ---
    by = [0.0] * n
    acumulado_bz = {}

    for posicion in range(n):

        clave = claves[posicion]

        if bw.iloc[posicion] < 0:
            by[posicion] = float(bx.iloc[posicion])
        else:
            by[posicion] = acumulado_bz.get(clave, 0.0)

        acumulado_bz[clave] = acumulado_bz.get(clave, 0.0) + float(
            bz.iloc[posicion]
        )

    by = pd.Series(by, index=df.index)

    ca = bu + bz
    cc = ca * bm_num

    # --- CE: por grupo ---
    au = pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)

    suma_bo = {}
    suma_cc = {}
    suma_au = {}

    for posicion, clave in enumerate(claves):
        valor_bo = bo.iloc[posicion]
        valor_cc = cc.iloc[posicion]
        suma_bo[clave] = suma_bo.get(clave, 0.0) + (
            0.0 if pd.isna(valor_bo) else float(valor_bo)
        )
        suma_cc[clave] = suma_cc.get(clave, 0.0) + (
            0.0 if pd.isna(valor_cc) else float(valor_cc)
        )
        suma_au[clave] = suma_au.get(clave, 0.0) + float(au.iloc[posicion])

    ce = []

    for posicion, clave in enumerate(claves):

        total_au = suma_au.get(clave, 0.0)

        if total_au == 0.0:
            # division por cero -> el IFERROR original devuelve 0
            ce.append(0.0)
            continue

        ce.append(
            max(suma_bo.get(clave, 0.0) - suma_cc.get(clave, 0.0), 0.0)
            * float(au.iloc[posicion])
            / total_au
        )

    indice = df_re545.index

    def _serie(valores):
        return pd.Series(list(valores), index=indice)

    return {
        "BI": _serie(df["BI"]),
        "BJ": _serie(df["BJ"]),
        "BK": _serie(df["BK"]),
        "BL": _serie(df["BL"]),
        "BM": _serie(df["BM"]),
        "BN": _serie(bn),
        "BO": _serie(bo),
        "BQ": _serie(bq),
        "BR": _serie(br),
        "BS": _serie(df["BS"]),
        "BT": _serie(bt),
        "BU": _serie(bu),
        "BV": _serie(bv),
        "BW": _serie(bw),
        "BX": _serie(bx),
        "BY": _serie(by),
        "BZ": _serie(bz),
        "CA": _serie(ca),
        "CC": _serie(cc),
        "CE": _serie(ce),
    }
