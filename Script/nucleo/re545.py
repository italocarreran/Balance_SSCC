# -*- coding: utf-8 -*-
"""
Calculo RE545: traspaso base, S, U, V y orquestacion.
"""

import pandas as pd

from .alertas import ALTA, Alerta, anotar
from .avisos import _avisar_claves_sin_mapeo
from .columnas_compartidas import calcular_l, calcular_m, calcular_n_o
from .diccionarios import _buscar_cmg
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, HOJA_CALCULO_RE545,
    HOJA_RESUMEN_BESS,
)
from .re545_reservas import (
    calcular_reservas_re545, construir_dic_reservas_subastas,
)
from .re545_resumen import _clave_grupo_re545
from .utiles import _normaliza_valor_vba, normalizar


# ============================================================
# CALCULO RE545 (etapa base): A:V
#
# La otra hoja de calculo del libro, hermana de "Calculo E Costos".
# Las dos se alimentan de la MISMA macro de traspaso
# (Traspasar_Medidores_A_Calculos_Rapido), que reparte cada fila de
# Medidores a una o a la otra segun Medidores!L (Ventana_No_Completa):
#   Ventana_No_Completa = 1  -> la energia va a "Calculo E Costos"
#   cualquier otro valor, o vacio -> va a "Calculo RE545"
# Las columnas A:G, K y P se escriben IGUALES en las dos hojas (no se
# reparten): lo unico que cambia es I/J, que quedan en 0 en la hoja
# que no corresponde.
#
# Diferencias propias de RE545 respecto de E Costos:
#   - T ("Ventana de valorizacion") = Medidores!L. En E Costos no
#     existe: la macro solo escribe T en RE545.
#   - R ("CMg Promedio") = CMg!I, via el mismo diccionario de
#     Asignar_CMg_a_Calculos_Turbo, que para esta hoja se llama con
#     escribirR:=True.
#   - Las columnas calculadas son OTRAS y, cuando comparten letra con
#     E Costos, casi nunca significan lo mismo (R, S, T, U, V son el
#     ejemplo claro). Los nombres reales estan en NOMBRES_CALCULO_RE545,
#     confirmados contra el archivo "Calculo_RE545_reducido_para_IA.xlsx"
#     que entrego el usuario (fila 3 del original = nombres, filas 1-2
#     = titulos de grupo).
#
# Igual que en E Costos, las columnas vacias del original (W:AB, AV,
# BH, BP, BS...) no se escriben: la hoja de salida no reproduce la
# letra de Excel, solo el orden y el contenido.
#
# Ver plan seccion 26.
# ============================================================

# Nombres reales de columna de "Calculo RE545" (fila 3 del archivo
# real). Mismo criterio que NOMBRES_CALCULO_E_COSTOS: se calcula todo
# con las letras/nombres internos y se renombra recien al final.
NOMBRES_CALCULO_RE545 = {
    "Mes": "Mes",
    "Dia": "Dia",
    "Hora": "Hora",
    "Hora Mes": "Hora mes",
    "Minutos": "Minuto",
    "Cuarto de Hora": "Bloque horario",
    "clave": "Configuracion",
    "Barra": "Barra",
    "Energia_Positiva": "Descarga kWh",
    "Energia_Negativa": "Carga kWh",
    "SoC": "SoC %",
    "L": "Adj SSCC",
    "M": "SoC sobre el minimo",
    "N": "Energía SSCC (-) por remunerar",
    "O": "Energía SSCC (+) por remunerar",
    "Copia_Ventana": "Ciclo de Carga del mes",
    "CMg": "CMg",
    "R": "CMg Promedio",
    "S": "ranking cmg",
    "T": "Ventana de valorizacion",
    "U": "EiniT",
    "V": "EalmT",
    "AC": "CPF(-)",
    "AD": "CSF(-)",
    "AE": "CTF(-)",
    "AF": "CPF(+)",
    "AG": "CSF(+)",
    "AH": "CTF(+)",
    "AI": "CPF(-)",
    "AJ": "CSF(-)",
    "AK": "CTF(-)",
    "AL": "CPF(+)",
    "AM": "CSF(+)",
    "AN": "CTF(+)",
    "AO": "CPF(-)",
    "AP": "CSF(-)",
    "AQ": "CTF(-)",
    "AR": "CPF(+)",
    "AS": "CSF(+)",
    "AT": "CTF(+)",
    "AU": "SUMA Reservas*FMA*FD",
    "BI": "Orden",
    "BJ": "Periodo",
    "BK": "Curva Cmg Decendente promedio horario",
    "BL": "",
    "BM": "Curva Cmg Decendente",
    "BN": "Edisp_Asig",
    "BO": "Total  C1_545",
    "BQ": "Energia Total",
    "BR": "Ventana de Valorizacion",
    "BS": "inyeccion en el periodo del Cmg Descendente",
    "BT": "Energía ya Asignada",
    "BU": "Asignacion Edisponible",
    "BV": "SSCC ultima hora",
    "BW": "inyeccion orden cronologico",
    "BX": "Energia Ultima hora",
    "BY": "Energía ya Asignada ultima hora",
    "BZ": "Energia Asignada Ultima hora",
    "CA": "Asignacion Edisponible+SSCC ultima hora",
    "CC": "Total C2_545",
    "CE": "Monto a compensar",
}

# Encabezados de grupo (celdas combinadas arriba de los nombres de
# columna) de "Calculo RE545", confirmados contra
# docs/Calculo_RE545_reducido_para_IA.xlsx (fila 3 del archivo real,
# aunque esa fila viene sin el merge en si -- se perdio en la
# reduccion -- el texto queda solo en la celda de mas a la izquierda
# de cada grupo, igual que en un merge real). Cada tupla es
# (etiqueta, [claves internas del grupo, en orden]); las claves usan
# el mismo diccionario NOMBRES_CALCULO_RE545 de arriba, asi que la
# posicion real en la hoja de salida se calcula en el momento de
# escribir (no depende de la letra del archivo real, que es distinta
# de la nuestra -- ver el comentario de escribir_pagos_bess()).
#
# "Componente 1" y "Componente 2" no arrancan en BI/BQ (las primeras
# columnas de cada seccion): el archivo real las deja sueltas, sin
# grupo, y el merge arranca recien en BK/BS -- igual que en
# GRUPOS_CALCULO_E_COSTOS, donde "Componente 1"/"Componente 2" tampoco
# incluyen la columna final "Total"... salvo que aca SI la incluyen
# (BO y CC quedan dentro del grupo); la unica columna que queda
# siempre afuera de cualquier grupo es la ultima de toda la hoja
# ("Monto a compensar").
# Auxiliares de calculo que NO se escriben (pedido del usuario:
# "quita los auxiliares innecesarios"). Se siguen calculando igual
# que antes -- salen de la hoja, no del calculo:
#
#   BL: en el archivo real es una columna SIN NOMBRE (la suma de CMg
#     del bloque) que solo existe para que BM saque de ahi su
#     k-esimo mayor. En la salida quedaba como una columna de numeros
#     con el encabezado en blanco.
#   BR "Ventana de Valorizacion": es T ("Ventana de valorizacion")
#     copiada tal cual, con el mismo nombre escrito distinto.
#
# El resto de los intermedios (BK, BM, BN, BS:CA, U/V, S) se quedan:
# son el paso a paso del Componente 1 y 2 que un coordinado necesita
# para seguir el calculo.
COLUMNAS_AUXILIARES_RE545 = ("BL", "BR")

# Orden final de columnas internas de la hoja (el de
# NOMBRES_CALCULO_RE545 sin los auxiliares). Lo usan
# renombrar_calculo_re545() y, para ubicar los encabezados de grupo,
# escribir_pagos_bess().
COLUMNAS_SALIDA_RE545 = [
    clave for clave in NOMBRES_CALCULO_RE545
    if clave not in COLUMNAS_AUXILIARES_RE545
]


GRUPOS_CALCULO_RE545 = (
    ("Dia", ["Mes", "Dia", "Hora"]),
    ("Nombre", ["clave", "Barra"]),
    ("BESS", ["Energia_Positiva", "Energia_Negativa", "SoC"]),
    ("Subastas", ["AC", "AD", "AE", "AF", "AG", "AH"]),
    ("FD", ["AI", "AJ", "AK", "AL", "AM", "AN"]),
    ("FMA", ["AO", "AP", "AQ", "AR", "AS", "AT"]),
    ("Componente 1", ["BK", "BM", "BN", "BO"]),
    (
        "Componente 2",
        ["BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CC"],
    ),
)


def construir_calculo_re545(
    df_medidores, mapa_barra, dic_cmg, registrar=print
):
    """
    Etapa base de "Calculo RE545": traspaso desde Medidores (A:G, I/J,
    K, P, T) + H (Barra) + Q/R (CMg y CMg Promedio).

    Es la hoja espejo de construir_calculo_e_costos(): misma macro de
    traspaso, misma A:G, misma K/P, y la energia repartida al reves
    (aca entra la de las filas con Ventana_No_Completa <> 1, incluido
    el caso "vacio o no numerico", que el VBA manda explicitamente a
    RE545).
    """

    n = len(df_medidores)
    df_medidores = df_medidores.reset_index(drop=True)

    df = pd.DataFrame(index=range(n))

    df["Mes"] = df_medidores["Mes"]
    df["Dia"] = df_medidores["Dia"]
    df["Hora"] = df_medidores["Hora"]
    df["Hora Mes"] = df_medidores["Hora Mes"]
    df["Minutos"] = df_medidores["Minutos"]
    df["Cuarto de Hora"] = df_medidores["Cuarto de Hora"]
    df["clave"] = df_medidores["clave"]

    df["Barra"] = df["clave"].map(
        lambda valor: mapa_barra.get(normalizar(valor), "")
    )

    _avisar_claves_sin_mapeo(
        df["clave"], mapa_barra, "Central sin barra de inyeccion",
        f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}", registrar,
        id_alerta="MAE-001", etapa=HOJA_CALCULO_RE545,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="la Barra queda vacia y con ella el CMg de esas filas",
        origen_control="CATALOGO AUX-006",
    )

    energia = pd.to_numeric(
        df_medidores["Gen_Unidad"], errors="coerce"
    ).fillna(0.0)

    # El VBA manda a RE545 todo lo que NO tiene T = 1 numerico:
    # distinto de 1, vacio, no numerico o error.
    va_a_ecostos = pd.to_numeric(
        df_medidores["Ventana_No_Completa"], errors="coerce"
    ).eq(1)
    va_a_re545 = ~va_a_ecostos

    df["Energia_Positiva"] = energia.where(energia > 0, 0.0).where(
        va_a_re545, 0.0
    )
    df["Energia_Negativa"] = energia.where(energia < 0, 0.0).where(
        va_a_re545, 0.0
    )

    df["SoC"] = df_medidores["SoC"]
    df["Copia_Ventana"] = df_medidores["Copia_Ventana"]

    # Medidores L -> RE545 T (solo esta hoja lo recibe).
    df["T"] = df_medidores["Ventana"]

    pares = [
        _buscar_cmg(dic_cmg, barra, cuarto_hora)
        for barra, cuarto_hora in zip(df["Barra"], df["Cuarto de Hora"])
    ]
    df["CMg"] = [par[0] for par in pares]
    df["R"] = [par[1] for par in pares]

    sin_cmg = int(df["CMg"].isna().sum())
    sin_cmg_promedio = int(df["R"].isna().sum())
    if sin_cmg or sin_cmg_promedio:
        anotar(registrar, Alerta(
            "CMG-004", ALTA, HOJA_CALCULO_RE545,
            f"{sin_cmg:,} fila(s) sin CMg y {sin_cmg_promedio:,} sin "
            f"CMg Promedio (sin match Barra+Cuarto de Hora en "
            f"{ARCHIVO_CMG}).",
            valor_encontrado=f"{sin_cmg:,} / {sin_cmg_promedio:,} filas",
            accion="los calculos posteriores los toman como 0",
            archivo=ARCHIVO_CMG,
            origen_control="CATALOGO CMG-004",
        ))

    filas_con_energia = int(va_a_re545.sum())

    registrar(
        f"  Calculo RE545: {n:,} fila(s) traspasadas desde Medidores "
        f"({filas_con_energia:,} con energia; el resto la tiene "
        f"'Calculo E Costos')."
    )

    return df


def calcular_s_re545(df_re545):
    """
    Replica S ("ranking cmg") de RE545:

        =(COUNTIFS($T:$T,T4,$R:$R,">"&R4,G:G,G4)
        + COUNTIFS($T:$T,T4,$R:$R,R4,$C:$C,">"&C4,G:G,G4))/4 + 1

    O sea, dentro del grupo central (G) + ventana de valorizacion (T):
    cuantas filas tienen CMg Promedio (R) mayor, mas cuantas lo tienen
    igual pero con Hora (C) mayor; todo eso dividido por 4 y +1.

    Ojo: NO es la misma columna que el "ranking cmg" de Calculo E
    Costos (esa es la R de esa hoja, agrupa por P y ordena por CMg, no
    por CMg Promedio).
    """

    df = df_re545.reset_index(drop=True)

    grupos = {}

    for posicion, (central, ventana) in enumerate(
        zip(df["clave"], df["T"])
    ):
        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )
        grupos.setdefault(clave, []).append(posicion)

    r = pd.to_numeric(df["R"], errors="coerce")
    c = pd.to_numeric(df["Hora"], errors="coerce")

    resultado = [pd.NA] * len(df)

    for posiciones in grupos.values():

        for posicion in posiciones:

            r_fila = r.iloc[posicion]
            c_fila = c.iloc[posicion]

            if pd.isna(r_fila):
                # COUNTIFS contra un blanco no cuenta nada: queda el +1.
                resultado[posicion] = 1.0
                continue

            mayores = 0
            empates = 0

            for otra in posiciones:

                r_otra = r.iloc[otra]

                if pd.isna(r_otra):
                    continue

                if r_otra > r_fila:
                    mayores += 1
                elif r_otra == r_fila:
                    c_otra = c.iloc[otra]
                    if pd.notna(c_otra) and pd.notna(c_fila) and c_otra > c_fila:
                        empates += 1

            resultado[posicion] = (mayores + empates) / 4.0 + 1.0

    return pd.Series(resultado, index=df_re545.index)


def calcular_u_v_re545(df_re545, dic_capacidad, dic_eficiencia):
    """
    Replica U ("EiniT") y V ("EalmT") de RE545:

        U = K * VLOOKUP(G, Resumen!B:J, 4, 0) * 1000
            (SoC % x "Capacidad (MWh)" x 1000 -- la 4ta columna del
             rango B:J es la Capacidad, no la Pmax; ver
             construir_dic_resumen_capacidad)

        V = -SUMIFS(J:J, T:T, T4, G:G, G4)
            * VLOOKUP(G, Resumen!B:J, 9, 0)
            (la carga total del grupo central+ventana, cambiada de
             signo, por la "Eficiencia" de esa central)

    Una central que no esta en "Resumen BESS" deja las dos en blanco
    (equivale al #N/A del VLOOKUP original).
    """

    df = df_re545

    capacidad = df["clave"].map(
        lambda valor: dic_capacidad.get(normalizar(valor), pd.NA)
    )
    eficiencia = df["clave"].map(
        lambda valor: dic_eficiencia.get(normalizar(valor), pd.NA)
    )

    soc = pd.to_numeric(df["SoC"], errors="coerce")
    capacidad_num = pd.to_numeric(capacidad, errors="coerce")
    eficiencia_num = pd.to_numeric(eficiencia, errors="coerce")

    u = soc * capacidad_num * 1000.0

    carga = pd.to_numeric(df["Energia_Negativa"], errors="coerce").fillna(0.0)
    suma_carga = carga.groupby([df["clave"], df["T"]]).transform("sum")

    v = -suma_carga * eficiencia_num

    return u, v


def completar_checks_resumen_re545(df_resumen_re545, df_re545):
    """
    Completa BD ("check 1") y BE ("check 2") de la tabla resumen, que
    dependen de BN y BU del bloque principal:

        BD = SUMIFS(BN, BR=AX, G=AW) - BC
        BE = SUMIFS(BU, BR=AX, G=AW) - BC + BF

    Son columnas de control: no alimentan ningun calculo posterior.
    """

    tabla = df_resumen_re545.copy()

    suma_bn = {}
    suma_bu = {}

    bn = pd.to_numeric(df_re545["BN"], errors="coerce").fillna(0.0)
    bu = pd.to_numeric(df_re545["BU"], errors="coerce").fillna(0.0)

    for posicion, (central, ventana) in enumerate(
        zip(df_re545["clave"], df_re545["BR"])
    ):
        clave = _clave_grupo_re545(central, ventana)
        suma_bn[clave] = suma_bn.get(clave, 0.0) + float(bn.iloc[posicion])
        suma_bu[clave] = suma_bu.get(clave, 0.0) + float(bu.iloc[posicion])

    claves = [
        _clave_grupo_re545(central, ventana)
        for central, ventana in zip(tabla.iloc[:, 0], tabla.iloc[:, 1])
    ]

    bc = pd.to_numeric(tabla["Edisp_T"], errors="coerce").fillna(0.0)
    bf = pd.to_numeric(tabla["Margen ultima hora"], errors="coerce").fillna(0.0)

    tabla["check 1"] = [
        suma_bn.get(clave, 0.0) - float(bc.iloc[posicion])
        for posicion, clave in enumerate(claves)
    ]
    tabla["check 2"] = [
        suma_bu.get(clave, 0.0) - float(bc.iloc[posicion])
        + float(bf.iloc[posicion])
        for posicion, clave in enumerate(claves)
    ]

    return tabla


def completar_calculo_re545(
    df_re545, df_subastas, umbral_soc_minimo, dic_capacidad,
    dic_eficiencia, registrar=print,
):
    """
    Agrega a la etapa base de RE545 las columnas calculadas L, M, N,
    O, S, U, V y AC:AU, en el orden final de la hoja pero todavia
    con los nombres internos (el renombre a los nombres reales lo
    hace renombrar_calculo_re545, despues de la tabla resumen).

    L y M son literalmente las mismas formulas que en Calculo E Costos
    (mismo COUNTIFS contra Subastas, mismo 1*(SoC > umbral)), asi que
    se reusan calcular_l() y calcular_m(). N y O tambien: la formula
    de RE545 (SUMIFS del grupo menos SUMIFS de los bloques anteriores)
    es la version en formula de lo mismo que calcula calcular_n_o()
    para E Costos -- suma, dentro del grupo central+ciclo y contando
    solo filas con L=1, la energia de los bloques horarios >= al de la
    fila. El resto (S, U, V) es propio de esta hoja.

    Agrega tambien AC:AU (los tres bloques de reservas por subasta y
    el SUMPRODUCT que los combina, ver la seccion de mas arriba).

    Todavia FUERA de esta etapa (ver plan seccion 26.3): AW:BG (el
    resumen por central+ventana, que es una tabla de otro largo) y
    BI:CE (Componentes 1 y 2).
    """

    df = df_re545.reset_index(drop=True).copy()

    df["L"] = calcular_l(df, df_subastas)
    df["M"] = calcular_m(df, umbral_soc_minimo)

    n, o = calcular_n_o(df)
    df["N"] = n.reset_index(drop=True)
    df["O"] = o.reset_index(drop=True)

    df["S"] = calcular_s_re545(df)

    u, v = calcular_u_v_re545(df, dic_capacidad, dic_eficiencia)
    df["U"] = u
    df["V"] = v

    _avisar_claves_sin_mapeo(
        df["clave"], dic_capacidad, "Central sin Capacidad (MWh)",
        f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}", registrar,
        id_alerta="MAE-004", etapa=HOJA_CALCULO_RE545,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="U, BC y BN quedan vacias para esa central",
        origen_control="CATALOGO AUX-006",
    )
    _avisar_claves_sin_mapeo(
        df["clave"], dic_eficiencia, "Central sin Eficiencia",
        f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}", registrar,
        id_alerta="MAE-005", etapa=HOJA_CALCULO_RE545,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="V queda vacia para esa central",
        origen_control="CATALOGO AUX-006",
    )

    dics_reservas = construir_dic_reservas_subastas(df_subastas)

    for interno, serie in calcular_reservas_re545(
        df, dics_reservas, registrar=registrar
    ).items():
        df[interno] = serie

    participa = int(df["L"].sum())
    registrar(
        f"  Calculo RE545: {participa:,} de {len(df):,} fila(s) "
        f"marcadas como 'participa en subasta' (L=1)."
    )

    return df


def renombrar_calculo_re545(df_re545):
    """
    Deja las columnas en el orden final de la hoja y las pasa de los
    nombres internos (letras) a los nombres reales. Se hace al final
    de todo y aparte, porque la tabla resumen AW:BG y las columnas
    BI:CE necesitan el DataFrame con los nombres internos.
    """

    return (
        df_re545[COLUMNAS_SALIDA_RE545]
        .rename(columns=NOMBRES_CALCULO_RE545)
    )
