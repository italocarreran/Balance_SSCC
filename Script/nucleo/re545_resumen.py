# -*- coding: utf-8 -*-
"""
RE545: la tabla por central+ventana (AW:BG) y BV.
"""

import pandas as pd

from .parametros import INICIO_VENTANA
from .utiles import _normaliza_valor_vba, normalizar


# ------------------------------------------------------------
# CALCULO RE545 (etapa 3): AW:BG -- resumen por central + ventana
#
# OTRA TABLA, no mas columnas de la misma: en el original tiene 288
# filas (9 centrales x 32 ventanas) contra las 26.787 del bloque
# principal, compartiendo hoja de la fila 4 para abajo. Es el mismo
# patron de "dos tablas de distinto largo en una hoja" que ya
# aparecio en FD (bloques CSF/CPF) y en Ofertas SSCC.
#
# AW (central), AX (Ventana) y AY (Oferta Completa) NO son formulas
# ni las escribe ninguna macro: en el .xlsm son constantes. AW/AX son
# el cruce central x ventana, y AY es la marca de "oferta completa"
# de esa ventana -- que en esta migracion NO hay que inventar: es la
# columna "Completa" de construir_resumen_ventana_oferta(), la misma
# que ya alimenta Medidores!T (T = 1 - Completa). Coincide el nombre,
# la clave (central+ventana), el dominio (0/1) y el sentido: RE545 se
# queda con las filas de ventana NO completa y la marca en AY.
#
# Formulas replicadas:
#   AZ = primer U del grupo (G=AW, T=AX)   [INDEX/AGGREGATE(15,6,...,1)]
#   BA = primer V del mismo grupo
#   BB = SUMIFS(AU:AU, T:T,AX, G:G,AW)
#   BC = MIN(MAX(MIN(AZ+BA, Capacidad*1000), BA), BB) * AY * (AX<>31)
#   BF = SUMIFS(BV:BV, BR:BR,AX, G:G,AW)   (BV, ver calcular_bv_re545)
#   BG = AND(AZ+BA - BF(ventana anterior) > Capacidad*1000, BF_previa<>0)*1
#
# BD ("check 1") y BE ("check 2") quedan para la etapa siguiente:
# dependen de BN y BU, que son columnas del bloque principal todavia
# sin implementar. Son columnas de CONTROL, no entran en ningun
# calculo posterior.
# ------------------------------------------------------------

NOMBRES_RESUMEN_RE545 = {
    "AW": "",
    "AX": "Ventana",
    "AY": "Oferta Completa",
    "AZ": "EiniT",
    "BA": "EalmT",
    "BB": "Total Reservas* FD *FMA",
    "BC": "Edisp_T",
    "BD": "check 1",
    "BE": "check 2",
    "BF": "Margen ultima hora",
    "BG": "flag ultima hora",
}


def calcular_bv_re545(df_re545, inicio_ventana=INICIO_VENTANA):
    """
    Replica BV ("SSCC ultima hora"):

        =IF(AND(C4 = Medidores!$S$1 - 1, AU4 <> 0), AU4, 0)

    Medidores!S1 es la hora en que arranca la ventana (el mismo dato
    que la constante INICIO_VENTANA: la formula de Medidores!L
    incrementa la ventana justo cuando la hora es igual a S1), asi
    que la condicion es "la hora anterior al inicio de la ventana",
    o sea la ultima hora de la ventana que termina.

    Se devuelve como Serie suelta, sin guardarla todavia en el
    DataFrame: BV vive en el bloque BQ:CE, que es de una etapa
    posterior, y meterla antes desordenaria las columnas de salida.
    """

    au = pd.to_numeric(df_re545["AU"], errors="coerce").fillna(0.0)
    hora = pd.to_numeric(df_re545["Hora"], errors="coerce")

    return au.where((hora == inicio_ventana - 1) & (au != 0), 0.0)


def construir_resumen_ventanas_re545(
    df_re545, resumen_ventana_oferta, dic_capacidad, registrar=print
):
    """
    Arma la tabla AW:BG (una fila por central + ventana de
    valorizacion). Ver el comentario de seccion para el detalle de
    cada columna.

    resumen_ventana_oferta: el DataFrame de
    construir_resumen_ventana_oferta() (columnas Central, Ventana T,
    Oferta, Completa) -- de ahi salen AW, AX y AY.
    dic_capacidad: central -> "Capacidad (MWh)"
    (construir_dic_resumen_capacidad) -- es el VLOOKUP con indice 4
    de BC, que NO es la Pmax.
    """

    base = (
        resumen_ventana_oferta[["Central", "Ventana T", "Completa"]]
        .copy()
        .sort_values(["Central", "Ventana T"], kind="mergesort")
        .reset_index(drop=True)
    )

    df = df_re545.reset_index(drop=True)

    u = pd.to_numeric(df["U"], errors="coerce")
    v = pd.to_numeric(df["V"], errors="coerce")
    au = pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)
    bv = calcular_bv_re545(df)

    primer_u = {}
    primer_v = {}
    suma_au = {}
    suma_bv = {}

    for posicion, (central, ventana) in enumerate(zip(df["clave"], df["T"])):

        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )

        # INDEX + AGGREGATE(15,6,...,1): gana la PRIMERA fila del grupo.
        if clave not in primer_u:
            primer_u[clave] = u.iloc[posicion]
            primer_v[clave] = v.iloc[posicion]

        suma_au[clave] = suma_au.get(clave, 0.0) + float(au.iloc[posicion])
        suma_bv[clave] = suma_bv.get(clave, 0.0) + float(bv.iloc[posicion])

    filas = []

    for _, fila in base.iterrows():

        central = fila["Central"]
        ventana = fila["Ventana T"]

        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )

        az = primer_u.get(clave, pd.NA)
        ba = primer_v.get(clave, pd.NA)
        bb = suma_au.get(clave, 0.0)
        bf = suma_bv.get(clave, 0.0)

        completa = fila["Completa"]
        completa = 0.0 if pd.isna(completa) else float(completa)

        factor = dic_capacidad.get(normalizar(central), pd.NA)

        filas.append(
            {
                "AW": central,
                "AX": ventana,
                "AY": completa,
                "AZ": az,
                "BA": ba,
                "BB": bb,
                "BF": bf,
                "_factor": factor,
            }
        )

    tabla = pd.DataFrame(
        filas,
        columns=["AW", "AX", "AY", "AZ", "BA", "BB", "BF", "_factor"],
    )

    if tabla.empty:
        tabla = tabla.assign(BC=[], BD=[], BE=[], BG=[])
        return tabla[list(NOMBRES_RESUMEN_RE545)].rename(
            columns=NOMBRES_RESUMEN_RE545
        )

    az_num = pd.to_numeric(tabla["AZ"], errors="coerce")
    ba_num = pd.to_numeric(tabla["BA"], errors="coerce")
    factor_num = pd.to_numeric(tabla["_factor"], errors="coerce")
    ventana_num = pd.to_numeric(tabla["AX"], errors="coerce")

    capacidad = factor_num * 1000.0

    # BC = MIN(MAX(MIN(AZ+BA, Pmax*1000), BA), BB) * AY * (AX<>31)
    tabla["BC"] = (
        pd.concat(
            [
                pd.concat(
                    [
                        pd.concat([az_num + ba_num, capacidad], axis=1).min(axis=1),
                        ba_num,
                    ],
                    axis=1,
                ).max(axis=1),
                pd.to_numeric(tabla["BB"], errors="coerce"),
            ],
            axis=1,
        ).min(axis=1)
        * tabla["AY"]
        * ventana_num.ne(31).astype(float)
    )

    # BG: la resta usa el BF de la MISMA central en la ventana anterior,
    # y ademas exige que el BF de la FILA ANTERIOR de esta tabla sea
    # distinto de 0. En el original la primera fila apunta a la fila de
    # encabezados (texto), que en Excel tambien cumple "<> 0".
    bf_por_clave = {
        (_normaliza_valor_vba(c), _normaliza_valor_vba(x)): float(b)
        for c, x, b in zip(tabla["AW"], tabla["AX"], tabla["BF"])
    }

    bg = []

    for posicion in range(len(tabla)):

        central = tabla["AW"].iloc[posicion]
        ventana = ventana_num.iloc[posicion]

        bf_anterior_ventana = bf_por_clave.get(
            (
                _normaliza_valor_vba(central),
                _normaliza_valor_vba(
                    ventana - 1 if pd.notna(ventana) else ventana
                ),
            ),
            0.0,
        )

        if posicion == 0:
            bf_fila_previa_no_cero = True
        else:
            bf_fila_previa_no_cero = float(tabla["BF"].iloc[posicion - 1]) != 0.0

        suma = az_num.iloc[posicion] + ba_num.iloc[posicion]
        limite = capacidad.iloc[posicion]

        if pd.isna(suma) or pd.isna(limite):
            bg.append(0)
            continue

        bg.append(
            int(
                (suma - bf_anterior_ventana > limite)
                and bf_fila_previa_no_cero
            )
        )

    tabla["BG"] = bg

    # Dependen de BN/BU, del bloque principal (etapa siguiente).
    tabla["BD"] = pd.NA
    tabla["BE"] = pd.NA

    registrar(
        f"  Calculo RE545: {len(tabla):,} fila(s) en el resumen por "
        f"central+ventana (AW:BG); BD/BE quedan pendientes."
    )

    tabla = tabla[list(NOMBRES_RESUMEN_RE545)]

    return tabla.rename(columns=NOMBRES_RESUMEN_RE545)


# ------------------------------------------------------------
# CALCULO RE545 (etapa 4): BI:CE -- Componente 1 y Componente 2
#
# Ultimo bloque del bloque principal. Formulas (fila 4 del original):
#
#   BI "Orden"    = 1 en las 4 primeras filas; despues
#                   IF(T(i)=T(i-4), BI(i-4)+1, 1)  <- salto de 4 filas
#   BJ "Periodo"  = 0,15,30,45 en las 4 primeras; despues BJ(i-4)
#   BK            = SUMIFS(R, G=G, S=BI, T=T, E=BJ)
#   BL            = SUMIFS(Q, S=BI, E=BJ, G=G, T=T)
#   BM "Curva Cmg Decendente"
#                 = LARGE(IF(BK_todas = BK(i), BL_todas), BJ(i)/15 + 1)
#   BN "Edisp_Asig"
#                 = MAX(0, MIN(MAX(0, Pmax*1000/4 - BS),
#                              BC(central,ventana) - suma de los BN
#                              ANTERIORES del mismo (T,G)))
#   BO "Total C1_545" = BN * BM
#   BQ "Energia Total" = BS + BN
#   BR "Ventana de Valorizacion" = T
#   BS "inyeccion en el periodo del Cmg Descendente"
#                 = I de la primera fila con BR=BR(i), S=BI(i),
#                   G=G(i), E=BJ(i)
#   BT "Energía ya Asignada" = suma de los BU POSTERIORES del mismo
#                   (BR, G)   <- mira hacia adelante
#   BU "Asignacion Edisponible"
#                 = IF(BS=0, 0, MAX(0, MIN(BQ, BC(central,ventana)
#                       - BT - suma de BV del grupo)))
#   BV "SSCC ultima hora" (ver calcular_bv_re545, ya usada por BF)
#   BW "inyeccion orden cronologico" = I + J
#   BX "Energia Ultima hora" = BF del resumen (central, ventana)
#   BY "Energía ya Asignada ultima hora"
#                 = IF(BW<0, BX, suma de los BZ ANTERIORES del grupo)
#   BZ "Energia Asignada Ultima hora"
#                 = IF(BW<0, 0, 8) * BG del resumen (central, ventana)
#   CA            = BU + BZ
#   CC "Total C2_545" = CA * BM
#   CE "Monto a compensar"
#                 = MAX(suma(BO del grupo) - suma(CC del grupo), 0)
#                   * AU / suma(AU del grupo)
#
# Dos recursiones que hay que resolver en orden, no vectorizables de
# una: BN necesita los BN anteriores de su grupo (se recorre de
# arriba hacia abajo) y BU necesita los BU POSTERIORES del suyo (se
# recorre de abajo hacia arriba). BY necesita los BZ anteriores, pero
# BZ no depende de BY, asi que ahi alcanza con calcular BZ primero.
#
# OJO con el VLOOKUP de BN: usa el indice 2 del rango B:J, o sea
# "Pmax (MW)" (dic_factor) -- distinto del indice 4 ("Capacidad
# (MWh)", dic_capacidad) que usan U y BC.
# ------------------------------------------------------------

def _clave_grupo_re545(central, ventana):
    return (_normaliza_valor_vba(central), _normaliza_valor_vba(ventana))


def _mapa_resumen_por_grupo(df_resumen_re545, columna):
    """
    (central, ventana) -> valor de una columna de la tabla resumen
    AW:BG, que ya viene con los nombres reales. Se toma por posicion
    porque AW no tiene encabezado en el archivo real.
    """

    centrales = df_resumen_re545.iloc[:, 0]
    ventanas = df_resumen_re545.iloc[:, 1]
    valores = pd.to_numeric(df_resumen_re545[columna], errors="coerce")

    mapa = {}

    for central, ventana, valor in zip(centrales, ventanas, valores):
        mapa[_clave_grupo_re545(central, ventana)] = (
            0.0 if pd.isna(valor) else float(valor)
        )

    return mapa
