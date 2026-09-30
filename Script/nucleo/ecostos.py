# -*- coding: utf-8 -*-
"""
Calculo E Costos: traspaso base y orquestacion.
"""

import pandas as pd

from .alertas import ALTA, Alerta, anotar
from .avisos import _avisar_claves_sin_mapeo
from .columnas_compartidas import calcular_l, calcular_m, calcular_n_o
from .diccionarios import _buscar_cmg
from .ecostos_ciclo import (
    calcular_aw_ax, calcular_az, calcular_subastas_ciclo,
    construir_dic_umbrales_subastas,
)
from .ecostos_columnas import (
    calcular_ae_af, calcular_r_ecostos, calcular_s_t_u, calcular_w_x,
    calcular_y_ab_ac_ad,
)
from .ecostos_prorratas import (
    calcular_as_at, calcular_au_av, calcular_fd_prorrateado,
    calcular_prorratas, construir_dic_fd_bloque,
    construir_dic_mapeo_diccionario, construir_dic_prorrata,
    construir_prorrata_sscc, unidades_bloque_fd,
)
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, HOJA_CALCULO_ECOSTOS,
    HOJA_RESUMEN_BESS,
)
from .utiles import _tiene_valor, normalizar


def construir_calculo_e_costos(
    df_medidores, mapa_barra, dic_cmg, registrar=print
):
    """
    Etapa base de "Calculo E Costos": traspaso desde Medidores (A:G
    con D<->E invertidas, I/J, K, P) + H (Barra, homologada por
    nombre) + Q (CMg, homologado por Barra+Cuarto de Hora). Ver el
    comentario de seccion mas arriba para el detalle de cada macro
    replicada.

    Solo cubre "Calculo E Costos" (Ventana_No_Completa == 1);
    "Calculo RE545" (Ventana_No_Completa <> 1) queda fuera de esta
    etapa.
    """

    n = len(df_medidores)
    df_medidores = df_medidores.reset_index(drop=True)

    df = pd.DataFrame(index=range(n))

    df["Mes"] = df_medidores["Mes"]
    df["Dia"] = df_medidores["Dia"]
    df["Hora"] = df_medidores["Hora"]

    # D <-> E invertidas: Destino D = Medidores E, Destino E = Medidores D.
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
        id_alerta="MAE-001", etapa=HOJA_CALCULO_ECOSTOS,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="la Barra queda vacia y con ella el CMg de esas filas",
        origen_control="CATALOGO AUX-006",
    )

    energia = pd.to_numeric(
        df_medidores["Gen_Unidad"], errors="coerce"
    ).fillna(0.0)

    va_a_ecostos = pd.to_numeric(
        df_medidores["Ventana_No_Completa"], errors="coerce"
    ).eq(1)

    energia_positiva = energia.where(energia > 0, 0.0).where(va_a_ecostos, 0.0)
    energia_negativa = energia.where(energia < 0, 0.0).where(va_a_ecostos, 0.0)

    df["Energia_Positiva"] = energia_positiva
    df["Energia_Negativa"] = energia_negativa

    # Medidores J (SoC) -> Destino K; Medidores K (Copia_Ventana) -> Destino P.
    df["SoC"] = df_medidores["SoC"]
    df["Copia_Ventana"] = df_medidores["Copia_Ventana"]

    df["CMg"] = [
        _buscar_cmg(dic_cmg, barra, cuarto_hora)[0]
        for barra, cuarto_hora in zip(df["Barra"], df["Cuarto de Hora"])
    ]

    sin_barra = int((df["Barra"] == "").sum())
    sin_cmg = int(df["CMg"].isna().sum())

    if sin_barra:
        anotar(registrar, Alerta(
            "MAE-001", ALTA, HOJA_CALCULO_ECOSTOS,
            f"{sin_barra:,} fila(s) sin barra de inyeccion (central no "
            f"encontrada en '{HOJA_RESUMEN_BESS}').",
            valor_encontrado=f"{sin_barra:,} filas",
            accion="la Barra queda vacia y con ella el CMg de esas filas",
            archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
            origen_control="CATALOGO AUX-006",
        ))

    if sin_cmg:
        anotar(registrar, Alerta(
            "CMG-004", ALTA, HOJA_CALCULO_ECOSTOS,
            f"{sin_cmg:,} fila(s) sin CMg (sin match Barra+Cuarto de "
            f"Hora en {ARCHIVO_CMG}).",
            valor_encontrado=f"{sin_cmg:,} filas",
            accion="el CMg queda vacio; S y T lo toman como 0",
            archivo=ARCHIVO_CMG,
            origen_control="CATALOGO CMG-004",
        ))

    registrar(
        f"  Calculo E Costos: {n:,} fila(s) traspasadas desde Medidores."
    )

    return df


# Nombres reales de columna de "Calculo E Costos" (confirmados por el
# usuario contra un archivo real, hoja "E COSTOS"). Se calcula todo
# con los nombres/letras internos usados hasta aca y se renombra
# recien al final, mismo criterio que NOMBRES_FD_CSF/NOMBRES_SUBASTAS.
#
# OJO: AG:AL ("Prorratas") y AM:AR ("FD") comparten los mismos 6
# nombres cortos (CPF(-), CSF(-), CTF(-), CPF(+), CSF(+), CTF(+)) --
# asi esta en el archivo real (se distinguen por un encabezado de
# grupo en las filas 1-2 que no se replica en este esquema de una
# sola fila de encabezado, igual que Subastas!Q sin nombre o el
# "Hora Mes" duplicado de FD). No es un error de tipeo. Lo mismo pasa
# con "Total": es el nombre real de U y tambien de AX (el total del
# grupo "Componente 1"). Para llegar sin ambiguedad a una de esas
# columnas hay que ir por posicion, no por nombre.
NOMBRES_CALCULO_E_COSTOS = {
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
    "Copia_Ventana": "Ciclo de Carga del mes",
    "CMg": "CMg",
    "L": "Adj SSCC",
    "M": "SoC sobre el minimo",
    "N": "Energía SSCC (-) por remunerar",
    "O": "Energía SSCC (+) por remunerar",
    "R": "ranking cmg",
    "S": "Valorizacion Descarga",
    "T": "Valorizacion Carga",
    "U": "Total",
    "W": "Bloque ordenado",
    "X": "Ciclo",
    "Y": "Bloque Mes Descarga",
    "AB": "Curva monotona CMg Descarga",
    "AC": "Bloque Mes  Carga",
    "AD": "Curva monotona CMg Carga",
    "AE": "Energía descargada",
    "AF": "Energía cargada",
    "AG": "CPF(-)",
    "AH": "CSF(-)",
    "AI": "CTF(-)",
    "AJ": "CPF(+)",
    "AK": "CSF(+)",
    "AL": "CTF(+)",
    "AM": "CPF(-)",
    "AN": "CSF(-)",
    "AO": "CTF(-)",
    "AP": "CPF(+)",
    "AQ": "CSF(+)",
    "AR": "CTF(+)",
    "AS": "Energía descarga con FD",
    "AT": "Energía carga con FD",
    "AU": "Ingreso descarga",
    "AV": "Costo carga",
    "AW": "Descuento FD",
    "AX": "Total",
    "AZ": "Monto a compensar",
}

# Auxiliares de calculo que NO se escriben (pedido del usuario:
# "quita los auxiliares innecesarios"). Se siguen calculando igual
# que antes -- salen de la hoja, no del calculo:
#
#   X "Ciclo": es Copia_Ventana copiada tal cual, y Copia_Ventana ya
#     esta en la hoja como "Ciclo de Carga del mes". Dos columnas con
#     el mismo numero, una al lado de la otra.
#
# NO se sacan las otras columnas intermedias (W "Bloque ordenado", Y,
# AB, AC, AD de las curvas monotonas, R "ranking cmg", AE/AF): son el
# paso a paso con el que un coordinado sigue de donde sale cada peso.
# Tampoco las CTF en 0 (AI, AL, AO, AR): que el CTF sea 0 es un dato,
# no una columna de relleno.
COLUMNAS_AUXILIARES_E_COSTOS = ("X",)

# Orden final de columnas internas de la hoja (el de
# NOMBRES_CALCULO_E_COSTOS sin los auxiliares). Lo usan
# renombrar_calculo_e_costos() y, para ubicar los encabezados de
# grupo, escribir_pagos_bess().
COLUMNAS_SALIDA_E_COSTOS = [
    clave for clave in NOMBRES_CALCULO_E_COSTOS
    if clave not in COLUMNAS_AUXILIARES_E_COSTOS
]


# Encabezados de grupo de "Calculo E Costos", confirmados contra
# docs/Libro1_Subastas_real.xlsx (hoja "E COSTOS", fila 2 real -- la
# unica hoja de las que tenemos como referencia real que conserva los
# merges de Excel tal cual, via ws.merged_cells). Mismo criterio que
# GRUPOS_CALCULO_RE545: (etiqueta, [claves internas]), resuelto contra
# la posicion real en la salida al momento de escribir.
#
# "Prorratas (-)"/"Prorratas (+)" y el bloque "FD" (FD homologado)
# comparten los mismos 6 nombres de columna (CPF/CSF/CTF por
# direccion) que las Prorratas -- por eso los grupos se arman con las
# claves internas (unicas), nunca buscando por nombre de columna
# (duplicado a proposito, como el resto de la hoja).
GRUPOS_CALCULO_E_COSTOS = (
    ("Dia", ["Mes", "Dia", "Hora"]),
    ("Nombre", ["clave", "Barra"]),
    ("BESS", ["Energia_Positiva", "Energia_Negativa", "SoC"]),
    ("Componente 2", ["S", "T", "U"]),
    ("Prorratas (-)", ["AG", "AH", "AI"]),
    ("Prorratas (+)", ["AJ", "AK", "AL"]),
    ("FD", ["AM", "AN", "AO", "AP", "AQ", "AR"]),
    ("Componente 1", ["AU", "AV", "AW", "AX"]),
)


def completar_calculo_e_costos_grupos(
    df_ecostos,
    df_subastas,
    dic_factor,
    umbral_soc_minimo,
    diccionario,
    df_fd_csf,
    df_fd_cpf,
    registrar=print,
):
    """
    Etapas 2, 3 y 4 de "Calculo E Costos": agrega L, M, N, O, R, S,
    T, U, W, X, Y, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM,
    AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX y AZ a df_ecostos
    (que ya viene con la etapa base
    de construir_calculo_e_costos). Devuelve la hoja con los nombres
    INTERNOS de columna: el renombre a los nombres reales es un paso
    aparte (renombrar_calculo_e_costos), igual que en RE545, porque
    todo lo que consume esta hoja despues -la conciliacion de energia,
    las prorratas- busca por nombre interno.
    Ver los comentarios de seccion mas arriba para el detalle y las
    advertencias de cada columna.

    dic_factor, umbral_soc_minimo: de construir_dic_resumen_factor().
    diccionario: hoja Diccionario de Centrales.xlsx (header=None).
    df_fd_csf, df_fd_cpf: de construir_fd() (nombres reales ya
    aplicados).

    Fuera de esta funcion: la columna AY del archivo real (que la
    macro original tampoco escribe) y toda la hoja Calculo RE545.
    """

    df = df_ecostos.reset_index(drop=True).copy()

    df["L"] = calcular_l(df, df_subastas)
    df["M"] = calcular_m(df, umbral_soc_minimo)

    n, o = calcular_n_o(df)
    df["N"] = n.reset_index(drop=True)
    df["O"] = o.reset_index(drop=True)

    df["R"] = calcular_r_ecostos(df).reset_index(drop=True)

    s, t, u = calcular_s_t_u(df)
    df["S"] = s
    df["T"] = t
    df["U"] = u

    w, x = calcular_w_x(df)
    df["W"] = w
    df["X"] = x

    y, ab, ac, ad = calcular_y_ab_ac_ad(df)
    df["Y"] = y.reset_index(drop=True)
    df["AB"] = ab.reset_index(drop=True)
    df["AC"] = ac.reset_index(drop=True)
    df["AD"] = ad.reset_index(drop=True)

    ae, af = calcular_ae_af(df, dic_factor, registrar=registrar)
    df["AE"] = ae
    df["AF"] = af

    _avisar_claves_sin_mapeo(
        df["clave"], dic_factor, "Central sin Pmax (MW)",
        f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}", registrar,
        id_alerta="MAE-002", etapa=HOJA_CALCULO_ECOSTOS,
        archivo=ARCHIVO_CENTRALES, hoja=HOJA_RESUMEN_BESS,
        accion="AE y AF quedan vacias (el #N/D del original)",
        origen_control="CATALOGO AUX-008",
    )

    tabla_prorrata = construir_prorrata_sscc(df_subastas)
    dic_prorrata = construir_dic_prorrata(tabla_prorrata, registrar=registrar)

    ag, ah = calcular_prorratas(df, dic_prorrata, registrar=registrar)
    df["AG"] = ag
    df["AH"] = ah
    df["AI"] = 0.0
    df["AJ"] = df["AG"]
    df["AK"] = df["AH"]
    df["AL"] = 0.0

    dic_mapeo = construir_dic_mapeo_diccionario(diccionario)
    dic_fd_csf = construir_dic_fd_bloque(df_fd_csf, "id", "CSF(+)", "CSF(-)")
    dic_fd_cpf = construir_dic_fd_bloque(df_fd_cpf, "id", "CPF(+)", "CPF(-)")

    am, an, ap, aq = calcular_fd_prorrateado(
        df, dic_mapeo, dic_fd_csf, dic_fd_cpf, registrar=registrar,
        unidades_fd_csf=unidades_bloque_fd(df_fd_csf),
        unidades_fd_cpf=unidades_bloque_fd(df_fd_cpf),
    )
    df["AM"] = am
    df["AN"] = an
    df["AO"] = 0.0
    df["AP"] = ap
    df["AQ"] = aq
    df["AR"] = 0.0

    as_, at = calcular_as_at(df)
    df["AS"] = as_
    df["AT"] = at

    au, av = calcular_au_av(df)
    df["AU"] = au
    df["AV"] = av

    ciclo_subastas = calcular_subastas_ciclo(df_subastas, df)
    dic_umbrales = construir_dic_umbrales_subastas(df_subastas, ciclo_subastas)

    con_ciclo = int(ciclo_subastas.map(_tiene_valor).sum())
    registrar(
        f"  Subastas: {con_ciclo:,} de {len(df_subastas):,} fila(s) "
        f"cruzaron con un ciclo de Calculo E Costos ('Ciclo'); "
        f"{len(dic_umbrales):,} par(es) central+ciclo con umbral "
        f"SUBIDA/BAJADA."
    )

    aw, ax = calcular_aw_ax(df, dic_umbrales)
    df["AW"] = aw
    df["AX"] = ax

    df["AZ"] = calcular_az(df)

    participa = int(df["L"].sum())
    registrar(
        f"  Calculo E Costos: {participa:,} de {len(df):,} fila(s) "
        f"marcadas como 'participa en subasta' (L=1)."
    )

    return df


def renombrar_calculo_e_costos(df_ecostos):
    """
    Deja las columnas en el orden final de la hoja y las pasa de los
    nombres internos a los nombres reales. Se hace aparte y al final
    de todo -mismo criterio que renombrar_calculo_re545()- porque
    NOMBRES_CALCULO_E_COSTOS tiene nombres REPETIDOS a proposito
    (AG:AL y AM:AR comparten los seis nombres CPF/CSF/CTF, y "Total"
    es U y AX): una vez renombrado, el DataFrame ya no se puede
    indexar por nombre sin ambiguedad.

    Devolver la hoja ya renombrada desde
    completar_calculo_e_costos_grupos() era justamente el origen del
    KeyError 'Energia_Positiva' de la conciliacion: esa columna pasa a
    llamarse "Descarga kWh" y, con nombres duplicados en el indice,
    pandas ni siquiera avisa "columna renombrada", tira KeyError.
    """

    return (
        df_ecostos[COLUMNAS_SALIDA_E_COSTOS]
        .rename(columns=NOMBRES_CALCULO_E_COSTOS)
    )
