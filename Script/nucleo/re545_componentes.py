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
      BM = el k-esimo valor mas grande de BL entre las filas de la
           misma central + ventana + BI, con k = Periodo/15 + 1
      BS = el I (Descarga kWh) del MISMO bloque del que sale BM: el
           k-esimo de mayor CMg de la hora (central + ventana + S == BI)

    BM (observacion 4 de la reliquidacion de enero 2026): el .xlsm
    original hacia LARGE(IF(BK = BK(i), BL), k) sobre TODA la columna,
    agrupando por el valor de BK (suma de CMg promedio). Dos BESS en la
    misma barra tienen el mismo BK y sus bloques se mezclaban, corriendo
    los precios de la curva. El usuario acogio la observacion y corrigio
    la macro con la clave G + T + BI (central + ventana + orden); aca se
    replica esa clave.

    BS (pedido del usuario, 2026-11-02): el .xlsm tomaba el I del bloque
    en la posicion k EN EL TIEMPO (minuto BJ), no del bloque que tiene el
    k-esimo CMg. Si la inyeccion cambia dentro de la hora, el espacio
    libre de la caja (Pmax/4 - BS) se calculaba con la inyeccion de otro
    bloque. Ahora BS sale del mismo bloque que BM; con empate de CMg gana
    el bloque mas temprano.
    """

    df = df_re545.reset_index(drop=True)

    # Rendimiento (2026-11-03): mismo calculo que antes, pero con listas
    # de Python en vez de .iloc y pd.Series por fila (25 s -> < 1 s en un
    # mes real). _normaliza_valor_vba es una funcion pura y los valores
    # se repiten mucho: se memoriza por (tipo, valor).
    memoria = {}

    def norm(valor):
        clave = (type(valor), valor)
        try:
            return memoria[clave]
        except KeyError:
            resultado = memoria[clave] = _normaliza_valor_vba(valor)
            return resultado
        except TypeError:  # valor no hasheable
            return _normaliza_valor_vba(valor)

    r = pd.to_numeric(df["R"], errors="coerce").fillna(0.0).tolist()
    q = pd.to_numeric(df["CMg"], errors="coerce").fillna(0.0).tolist()
    # numpy (np.float64), igual que lo que devolvia .iloc.
    i_energia = pd.to_numeric(df["Energia_Positiva"], errors="coerce").to_numpy()
    # El periodo de cada fila, convertido una vez para toda la columna
    # (antes, un pd.to_numeric de una Series de 1 elemento por fila).
    periodos = [
        pd.to_numeric(pd.Series([valor]), errors="coerce").iloc[0]
        if not isinstance(valor, (int, float)) else valor
        for valor in df["BJ"]
    ]

    centrales = [norm(v) for v in df["clave"]]
    rankings = [norm(v) for v in df["S"]]
    ventanas = [norm(v) for v in df["T"]]
    ordenes = [norm(v) for v in df["BI"]]
    periodos_txt = [norm(v) for v in df["BJ"]]

    # Clave de ACUMULACION: se arma con S (ranking cmg) de cada fila,
    # que es lo que el SUMIFS/MATCH real compara contra "BI4".
    claves_acumulacion = list(zip(centrales, rankings, ventanas, periodos_txt))

    # Clave de BUSQUEDA (una por fila): se arma con el BI (Orden) de
    # esa misma fila, el criterio fijo del SUMIFS/MATCH real.
    claves_busqueda = list(zip(centrales, ordenes, ventanas, periodos_txt))

    suma_r = {}
    suma_q = {}

    for posicion, clave in enumerate(claves_acumulacion):
        suma_r[clave] = suma_r.get(clave, 0.0) + float(r[posicion])
        suma_q[clave] = suma_q.get(clave, 0.0) + float(q[posicion])

    # SUMIFS sin match da 0.
    bk = [suma_r.get(clave, 0.0) for clave in claves_busqueda]
    bl = [suma_q.get(clave, 0.0) for clave in claves_busqueda]

    # BS: bloques de cada hora (central + ventana + S) de mayor a menor
    # CMg; el puesto k de la hora toma la inyeccion del k-esimo. Sin
    # bloque queda en blanco, como el IFERROR del original.
    bloques_hora = {}
    for posicion, clave in enumerate(zip(centrales, ventanas, rankings)):
        bloques_hora.setdefault(clave, []).append(
            (float(q[posicion]), i_energia[posicion])
        )
    for lista in bloques_hora.values():
        lista.sort(key=lambda bloque: bloque[0], reverse=True)

    def puesto(periodo):
        return None if pd.isna(periodo) else int(periodo / 15) + 1

    bs = []
    for clave, periodo in zip(zip(centrales, ventanas, ordenes), periodos):
        lista = bloques_hora.get(clave, [])
        k = puesto(periodo)
        bs.append(lista[k - 1][1] if k and 1 <= k <= len(lista) else pd.NA)

    # BM: k-esimo mayor BL dentro de central + ventana + BI (obs. 4).
    claves_bm = list(zip(centrales, ventanas, ordenes))

    por_grupo = {}

    for clave, valor_bl in zip(claves_bm, bl):
        por_grupo.setdefault(clave, []).append(float(valor_bl))

    for lista in por_grupo.values():
        lista.sort(reverse=True)

    bm = []

    for clave, periodo in zip(claves_bm, periodos):

        if pd.isna(periodo):
            bm.append(pd.NA)
            continue

        k = int(periodo / 15) + 1
        lista = por_grupo.get(clave, [])

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
