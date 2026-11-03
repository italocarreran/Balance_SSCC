# -*- coding: utf-8 -*-
"""
Calculo E Costos: traspaso base y orquestacion.
"""

import pandas as pd

from .alertas import ALTA, Alerta, anotar
from .avisos import _avisar_claves_sin_mapeo
from .columnas_compartidas import calcular_l, calcular_m, calcular_n_o
from .diccionarios import _buscar_cmg
from .ecostos_valorizacion import (
    calcular_acumulados_ajustados, calcular_asignacion_ajustada,
    calcular_monto_a_compensar, calcular_valorizacion,
)
from .ecostos_columnas import (
    calcular_ae_af, calcular_r_ecostos, calcular_s_t_u, calcular_w_x,
    calcular_y_ab_ac_ad,
)
from .ecostos_prorratas import (
    calcular_energia_ajustada, calcular_fd_prorrateado,
    calcular_prorratas, construir_dic_fd_bloque,
    construir_dic_mapeo_diccionario, construir_dic_prorrata,
    construir_prorrata_sscc, dic_mapeo_desde_fd, unidades_bloque_fd,
)
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, HOJA_CALCULO_ECOSTOS,
    HOJA_RESUMEN_BESS,
)
from .lectura import ROL_FD
from .utiles import normalizar


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
# "Hora Mes" duplicado de FD). No es un error de tipeo.
#
# Desde el metodo nuevo del FD (macro SSCC_Calcular_Integrado) AG:AR
# son la prorrata y el FD de la HORA DE LA FILA, con (-) y (+) por
# separado, y AS:BC la energia ajustada, su asignacion y su
# valorizacion (ver ecostos_prorratas y ecostos_valorizacion). La
# columna BB de la macro es un separador vacio y no se escribe. El
# monto final conserva el nombre "Monto a compensar": lo buscan asi la
# compensacion por central y la prorrata de retiros.
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
    # Prorratas y FD llevan el prefijo en el nombre (antes se
    # distinguian por un encabezado de grupo).
    "AG": "Prorrata CPF(-)",
    "AH": "Prorrata CSF(-)",
    "AI": "CTF(-)",
    "AJ": "Prorrata CPF(+)",
    "AK": "Prorrata CSF(+)",
    "AL": "CTF(+)",
    "AM": "FD CPF(-)",
    "AN": "FD CSF(-)",
    "AO": "CTF(-)",
    "AP": "FD CPF(+)",
    "AQ": "FD CSF(+)",
    "AR": "CTF(+)",
    "AS": "Descarga ajustada FD-prorrata kWh",
    "AT": "Carga ajustada FD-prorrata kWh",
    "AU": "Energía descargada por remunerar",
    "AV": "Energía cargada por remunerar",
    "AW": "Asignación Descarga kWh",
    "AX": "Asignación Carga kWh",
    "AY": "Valorización Descarga [$]",
    "AZ": "Valorización Carga [$]",
    "BA": "Total [CLP]",
    "BC": "Monto a compensar",
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
#
# Pedido del usuario (2026-11-03, "ordenar la hoja de Ecostos") salen
# tambien, sin dejar de calcularse:
#   - las CTF (AI, AL, AO, AR): siempre 0, no se usan;
#   - las del metodo anterior que no llegan al monto: M "SoC sobre el
#     minimo", N / O "Energía SSCC (-)/(+) por remunerar" (sin FD) y
#     AE / AF "Energía descargada / cargada" (su reparto en Pmax).
COLUMNAS_AUXILIARES_E_COSTOS = (
    "X", "AI", "AL", "AO", "AR", "M", "N", "O", "AE", "AF",
)

# La hoja en tres cuadros, en el orden del calculo: lo que se trae de
# otras hojas y la energia ajustada por FD-prorrata; el Componente 1
# (curva de CMg, energia ajustada acumulada, asignacion y su
# valorizacion); el Componente 2 (energia medida a su CMg) y el monto
# = C1 - C2. (titulo, [claves internas]).
CUADROS_CALCULO_E_COSTOS = (
    ("Información y energía con FD-prorrata", [
        "Mes", "Dia", "Hora", "Hora Mes", "Minutos", "Cuarto de Hora",
        "clave", "Barra", "Energia_Positiva", "Energia_Negativa", "SoC",
        "Copia_Ventana", "CMg", "L",
        "AG", "AH", "AJ", "AK", "AM", "AN", "AP", "AQ", "AS", "AT",
    ]),
    ("Componente 1", [
        "R", "W", "Y", "AB", "AC", "AD", "AU", "AV", "AW", "AX",
        "AY", "AZ", "BA",
    ]),
    ("Componente 2", ["S", "T", "U"]),
    # Aparte y destacado (pedido del usuario): es el pago.
    ("Monto a compensar", ["BC"]),
)

# Columnas que se calculan con datos de la MISMA fila: van a la hoja
# como formula de Excel (pedido del usuario: "prefiero la formula"), con
# los nombres internos entre llaves; la escritura los cambia por la
# letra de la columna y el numero de la fila. Dan exactamente lo mismo
# que el calculo en Python (ver calcular_energia_ajustada,
# calcular_s_t_u y calcular_valorizacion), que sigue siendo el que
# alimenta el monto: una celda vacia vale 0 en los dos.
FORMULAS_E_COSTOS = {
    "AS": "={Energia_Positiva}*({AG}*{AM}+{AH}*{AN}+{AJ}*{AP}+{AK}*{AQ})",
    "AT": "={Energia_Negativa}*({AG}*{AM}+{AH}*{AN}+{AJ}*{AP}+{AK}*{AQ})",
    "S": "={Energia_Positiva}*{CMg}",
    "T": "={Energia_Negativa}*{CMg}",
    "U": "={L}*({S}+{T})",
    "BA": "={AY}+{AZ}",
}

# Columnas con su total en la fila 1, sobre el nombre de la columna.
TOTALES_E_COSTOS = ("BA", "U", "BC")

# Curvas monotonas de CMg: el puesto que ya no tiene una fila que cumpla
# (Adj SSCC = 1 y energia distinta de 0) se ve con un guion, no vacio
# (pedido del usuario). Solo en la hoja: el calculo usa el valor vacio y
# ninguna formula mira estas columnas.
SIN_PUESTO_E_COSTOS = ("AB", "AD")
GUION_SIN_PUESTO = "–"

# Orden final de columnas internas de la hoja: los tres cuadros
# seguidos. Lo usan renombrar_calculo_e_costos() y escribir_pagos_bess()
# (que la parte en sus cuadros).
COLUMNAS_SALIDA_E_COSTOS = [
    clave for _, claves in CUADROS_CALCULO_E_COSTOS for clave in claves
]

# Comentario del encabezado de las columnas que vienen de otra hoja (o
# que usan un dato de otra hoja que no esta como columna): de donde sale
# y con que condicion.
_MEDIDORES = "Hoja origen: Medidores, Cuadro 1"
_GRUPO = "central + ciclo (Configuracion + Ciclo de Carga del mes)"
ORIGEN_COLUMNAS_E_COSTOS = {
    # B:H (Mes ... Configuracion) sin comentario: pedido del usuario.
    "Barra": (
        "Hoja origen: Centrales, Cuadro 1 Resumen BESS (Barra inyección). "
        "Se busca por Configuracion."
    ),
    "Energia_Positiva": (
        f"{_MEDIDORES} (Gen_Unidad). Solo la parte positiva y solo en las "
        "filas que van a E Costos (Ventana_No_Completa = 1); en las demás, 0."
    ),
    "Energia_Negativa": (
        f"{_MEDIDORES} (Gen_Unidad). Solo la parte negativa y solo en las "
        "filas que van a E Costos (Ventana_No_Completa = 1); en las demás, 0."
    ),
    "SoC": f"{_MEDIDORES} (SoC). Copia directa.",
    "Copia_Ventana": f"{_MEDIDORES} (Copia_Ventana). Copia directa.",
    "CMg": (
        "Hoja origen: CMg, Cuadro 1 (F_CMg). Se busca por Barra + Bloque "
        "horario."
    ),
    "L": (
        "Hoja origen: Subastas, Cuadro 1. 1 si la central tiene subasta "
        "adjudicada (BAJADA o SUBIDA) en ese mes, día y hora; si no, 0."
    ),
    **{
        clave: (
            "Se calcula desde Subastas, Cuadro 1: por central y Hora mes, "
            f"filas de {servicio} / total de filas de esa central y hora "
            "(la fila suma 1)."
        )
        for clave, servicio in (
            ("AG", "CPF(-)"), ("AH", "CSF(-)"),
            ("AJ", "CPF(+)"), ("AK", "CSF(+)"),
        )
    },
    **{
        clave: (
            f"Hoja origen: FD, {cuadro} (columna {servicio}). Se busca por "
            "Hora mes + unidad de la central (columna Central de la hoja FD "
            "o Diccionario). Solo si Adj SSCC = 1 y la prorrata del servicio "
            "es distinta de 0; si no, 0."
        )
        for clave, cuadro, servicio in (
            ("AM", "Cuadro 2 FD CPF", "CPF(-)"),
            ("AN", "Cuadro 1 FD CSF", "CSF(-)"),
            ("AP", "Cuadro 2 FD CPF", "CPF(+)"),
            ("AQ", "Cuadro 1 FD CSF", "CSF(+)"),
        )
    },
    **{
        clave: (
            "Usa Pmax (MW) de Centrales, Cuadro 1 Resumen BESS (se busca por "
            "Configuracion): el máximo del acumulado de la energía ajustada "
            "del ciclo se reparte en bloques de Pmax/4 × 1000 kWh según el "
            "Bloque ordenado. Sin Pmax queda vacía."
        )
        for clave in ("AW", "AX")
    },
    # Cuadro 2: lo que se calcula mirando otras filas (no va formula).
    "R": (
        f"Puesto del CMg dentro de su {_GRUPO}: de mayor a menor CMg (en "
        "empate, Bloque horario de mayor a menor); los empates comparten "
        "el puesto."
    ),
    "W": (
        "Contador de bloques: sube de 1 en 1 y vuelve a 1 cada vez que "
        "cambia el Ciclo de Carga del mes respecto de la fila de arriba. Es "
        "la posición del bloque en la curva monótona (Asignación)."
    ),
    "Y": (
        f"Curva monótona de descarga de su {_GRUPO}: las filas con Adj "
        "SSCC = 1 y Descarga distinta de 0, ordenadas por CMg de mayor a "
        "menor; acá va el Bloque horario de la fila que cae en este puesto."
    ),
    "AB": (
        "El CMg de esa fila de la curva monótona de descarga (vacío si el "
        "puesto ya no tiene una fila con descarga)."
    ),
    "AC": (
        f"Curva monótona de carga de su {_GRUPO}: las filas con Adj SSCC = "
        "1 y Carga distinta de 0, ordenadas por CMg de menor a mayor; acá va "
        "el Bloque horario de la fila que cae en este puesto."
    ),
    "AD": (
        "El CMg de esa fila de la curva monótona de carga (vacío si el "
        "puesto ya no tiene una fila con carga)."
    ),
    "AU": (
        "Energía con FD-prorrata por remunerar: acumulado de la Descarga "
        f"ajustada FD-prorrata de su {_GRUPO}, desde el último Bloque "
        "horario hacia el primero, contando solo filas con Adj SSCC = 1. "
        "Las filas del mismo Bloque horario llevan el mismo valor."
    ),
    "AV": (
        "Energía con FD-prorrata por remunerar: acumulado de la Carga "
        f"ajustada FD-prorrata (con el signo cambiado) de su {_GRUPO}, "
        "desde el último Bloque horario hacia el primero, contando solo "
        "filas con Adj SSCC = 1. Las filas del mismo Bloque horario llevan "
        "el mismo valor."
    ),
    "AY": (
        "Asignación Descarga × promedio del CMg de la curva monótona de "
        f"descarga en las filas de su {_GRUPO} con asignación distinta de "
        "0. Solo si la Descarga medida total del ciclo (todas las "
        "centrales) supera 10 kWh; si no, 0."
    ),
    "AZ": (
        "Asignación Carga × promedio del CMg de la curva monótona de carga "
        f"en las filas de su {_GRUPO} con asignación distinta de 0. Solo si "
        "la Carga medida total del ciclo (todas las centrales) es menor que "
        "−10 kWh; si no, 0."
    ),
    "BC": (
        f"Por {_GRUPO}: (Σ Total C1 − Σ Total C2) / número de filas del "
        "grupo, nunca negativo; el mismo valor en todas las filas del "
        "grupo. Arriba, el total del mes."
    ),
}


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
# Desde que la hoja va en tres cuadros (CUADROS_CALCULO_E_COSTOS) no
# lleva encabezados de grupo: el titulo de cada cuadro hace ese papel y
# prorratas y FD llevan el prefijo en el nombre de la columna.
GRUPOS_CALCULO_E_COSTOS = ()


def _completar_mapeo_con_otros_nombres(
    df, dic_mapeo, diccionario, df_fd_csf, df_fd_cpf, registrar
):
    """
    Central -> unidad del FD: si la unidad del Diccionario (o la de la
    columna 'Central' de la hoja FD) no esta en la hoja FD, se prueba
    con los demas nombres de la central en el Diccionario (Subastas,
    Ofertas, ...). Si alguno esta, se usa y se avisa (DIC-002).
    """

    from .equivalencias import Equivalencias

    equivalencias = Equivalencias(diccionario)
    if not equivalencias:
        return
    disponibles = set(unidades_bloque_fd(df_fd_csf)) | set(
        unidades_bloque_fd(df_fd_cpf)
    )
    if not disponibles:
        return

    for central in pd.unique(df["clave"]):
        clave = normalizar(central)
        actual = dic_mapeo.get(clave)
        if actual is not None and normalizar(actual) in disponibles:
            continue
        unidad = equivalencias.buscar(
            central, ROL_FD, disponibles, "las unidades de la hoja FD"
        )
        if unidad is not None:
            dic_mapeo[clave] = unidad

    equivalencias.informar(registrar)


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
    AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ, BA y BC a
    df_ecostos (que ya viene con la etapa base
    de construir_calculo_e_costos). Devuelve la hoja con los nombres
    INTERNOS de columna: el renombre a los nombres reales es un paso
    aparte (renombrar_calculo_e_costos), igual que en RE545, porque
    todo lo que consume esta hoja despues -la conciliacion de energia,
    las prorratas- busca por nombre interno.
    Ver los comentarios de seccion mas arriba para el detalle y las
    advertencias de cada columna.

    dic_factor, umbral_soc_minimo: de construir_dic_resumen_factor().
    diccionario: hoja Diccionario (header=None), de la hoja
    "Centrales" de la planilla. Solo hace falta para las centrales que
    la columna 'Central' de la hoja FD no cubre; puede ser None.
    df_fd_csf, df_fd_cpf: los dos bloques de la hoja FD (nombres reales
    ya aplicados, y la columna 'Central' si la hoja la trae).

    Hasta AF es el calculo historico; de AG en adelante, el metodo
    nuevo del FD (prorrata y FD de la hora de la fila aplicados a la
    energia medida, sin descuento FD).
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

    # Metodo nuevo: prorrata y FD de la hora de la fila (solo L = 1),
    # aplicados a la energia medida antes de acumular y repartir.
    tabla_prorrata = construir_prorrata_sscc(df_subastas)
    dic_prorrata = construir_dic_prorrata(tabla_prorrata, registrar=registrar)

    ag, ah, aj, ak = calcular_prorratas(df, dic_prorrata, registrar=registrar)
    df["AG"] = ag
    df["AH"] = ah
    df["AI"] = 0.0
    df["AJ"] = aj
    df["AK"] = ak
    df["AL"] = 0.0

    # La hoja FD ya trae, al lado de cada fila, el nombre Balance_BESS
    # de su unidad (columna 'Central', puesta al importarla): ese cruce
    # manda. El Diccionario queda solo para lo que la hoja no cubra (un
    # libro de antes de la columna, o una central sin ninguna fila en
    # FD, que es lo que distingue FD-007 de DIC-001).
    dic_mapeo = (
        construir_dic_mapeo_diccionario(diccionario)
        if diccionario is not None else {}
    )
    dic_mapeo.update(dic_mapeo_desde_fd(df_fd_csf, df_fd_cpf))
    _completar_mapeo_con_otros_nombres(
        df, dic_mapeo, diccionario, df_fd_csf, df_fd_cpf, registrar
    )
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

    df["AS"], df["AT"] = calcular_energia_ajustada(df)
    df["AU"], df["AV"] = calcular_acumulados_ajustados(df)
    df["AW"], df["AX"] = calcular_asignacion_ajustada(df, dic_factor)
    df["AY"], df["AZ"], df["BA"] = calcular_valorizacion(
        df, registrar=registrar
    )
    df["BC"] = calcular_monto_a_compensar(df)

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
