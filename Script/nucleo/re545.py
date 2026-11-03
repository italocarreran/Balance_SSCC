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
    HOJA_RESUMEN_BESS, INICIO_VENTANA,
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
    # Reservas, FD y FMA llevan el prefijo en el nombre (antes se
    # distinguian por un encabezado de grupo).
    "AC": "Reserva CPF(-)",
    "AD": "Reserva CSF(-)",
    "AE": "CTF(-)",
    "AF": "Reserva CPF(+)",
    "AG": "Reserva CSF(+)",
    "AH": "CTF(+)",
    "AI": "FD CPF(-)",
    "AJ": "FD CSF(-)",
    "AK": "CTF(-)",
    "AL": "FD CPF(+)",
    "AM": "FD CSF(+)",
    "AN": "CTF(+)",
    "AO": "FMA CPF(-)",
    "AP": "FMA CSF(-)",
    "AQ": "CTF(-)",
    "AR": "FMA CPF(+)",
    "AS": "FMA CSF(+)",
    "AT": "CTF(+)",
    "AU": "Reserva con FD y FMA",
    "BI": "Orden",
    "BJ": "Periodo",
    "BK": "Curva Cmg Decendente promedio horario",
    "BL": "",
    "BM": "Curva Cmg Decendente",
    "BN": "Energía asignada",
    "BO": "Total C1",
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
    "CC": "Total C2",
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
#
# Pedido del usuario (2026-11-03, "ordenar RE545", igual que E Costos)
# salen tambien, sin dejar de calcularse:
#   - las CTF de Reservas, FD y FMA (AE, AH, AK, AN, AQ, AT);
#   - M "SoC sobre el minimo", N / O "Energía SSCC (-)/(+) por
#     remunerar": no las usa ninguna otra columna;
#   - BK "Curva Cmg Decendente promedio horario": desde la observacion
#     4 la curva ya no se agrupa por ella, solo se mostraba;
#   - BY "Energía ya Asignada ultima hora": no la usa ninguna columna.
COLUMNAS_AUXILIARES_RE545 = (
    "BL", "BR", "AE", "AH", "AK", "AN", "AQ", "AT", "M", "N", "O", "BK",
    "BY",
)
COLUMNAS_CTF_RE545 = ("AE", "AH", "AK", "AN", "AQ", "AT")

# La hoja en cuadros, en el orden del calculo (titulo, [claves]). El
# quinto (resumen por central + ventana, otro largo) lo agrega la
# escritura despues de estos.
CUADROS_CALCULO_RE545 = (
    ("Información y reservas", [
        "Mes", "Dia", "Hora", "Hora Mes", "Minutos", "Cuarto de Hora",
        "clave", "Barra", "Energia_Positiva", "Energia_Negativa", "SoC",
        "Copia_Ventana", "T", "CMg", "R", "L",
        "AC", "AD", "AF", "AG", "AI", "AJ", "AL", "AM",
        "AO", "AP", "AR", "AS", "AU", "U", "V",
    ]),
    ("Componente 1", ["S", "BI", "BJ", "BM", "BN", "BO"]),
    ("Componente 2", [
        "BS", "BT", "BQ", "BU", "BV", "BW", "BX", "BZ", "CA", "CC",
    ]),
    ("Monto a compensar", ["CE"]),
)
TITULO_RESUMEN_RE545 = "Resumen por central y ventana"

# Orden final de columnas internas de la hoja: los cuadros seguidos.
COLUMNAS_SALIDA_RE545 = [
    clave for _, claves in CUADROS_CALCULO_RE545 for clave in claves
]

# Sin encabezados de grupo desde que la hoja va en cuadros.
GRUPOS_CALCULO_RE545 = ()

# Lo que se calcula con datos de la MISMA fila va como formula de Excel
# (mismo criterio que FORMULAS_E_COSTOS). Dan lo mismo que el calculo en
# Python, que sigue siendo el que alimenta el monto. AU solo va como
# formula si las CTF (que no se escriben) son 0 en todas las filas: si
# no, la formula no las veria (ver renombrar_calculo_re545).
FORMULAS_RE545 = {
    "AU": "=({AC}*{AI}*{AO}+{AD}*{AJ}*{AP}+{AF}*{AL}*{AR}+{AG}*{AM}*{AS})"
          "/4*1000",
    "BO": '=IF(COUNT({BN},{BM})<2,"",{BN}*{BM})',
    "BQ": "={BS}+{BN}",
    "BV": f"=IF(AND({{Hora}}={INICIO_VENTANA - 1},{{AU}}<>0),{{AU}},0)",
    "BW": "={Energia_Positiva}+{Energia_Negativa}",
    "CA": "={BU}+{BZ}",
    "CC": '=IF(COUNT({CA},{BM})<2,"",{CA}*{BM})',
}
TOTALES_RE545 = ("BO", "CC", "CE")

_MEDIDORES = "Hoja origen: Medidores, Cuadro 1"
_GRUPO = "central + ventana (Configuracion + Ventana de valorizacion)"
_SUBASTAS = (
    "Se suma desde Subastas, Cuadro 1: filas con la misma Configuración, "
    "Hora_mes y Concepto"
)
ORIGEN_COLUMNAS_RE545 = {
    # B:H (Mes ... Configuracion) sin comentario, como en E Costos.
    "Barra": (
        "Hoja origen: Centrales, Cuadro 1 Resumen BESS (Barra inyección). "
        "Se busca por Configuracion."
    ),
    "Energia_Positiva": (
        f"{_MEDIDORES} (Gen_Unidad). Solo la parte positiva y solo en las "
        "filas que van a RE545 (Ventana_No_Completa distinta de 1); en las "
        "demás, 0."
    ),
    "Energia_Negativa": (
        f"{_MEDIDORES} (Gen_Unidad). Solo la parte negativa y solo en las "
        "filas que van a RE545 (Ventana_No_Completa distinta de 1); en las "
        "demás, 0."
    ),
    "SoC": f"{_MEDIDORES} (SoC). Copia directa.",
    "Copia_Ventana": f"{_MEDIDORES} (Copia_Ventana). Copia directa.",
    "T": f"{_MEDIDORES} (Ventana). Copia directa.",
    "CMg": (
        "Hoja origen: CMg, Cuadro 1 (F_CMg). Se busca por Barra + Bloque "
        "horario."
    ),
    "R": (
        "Hoja origen: CMg, Cuadro 1 (I_CMgProm, promedio de la hora). Se "
        "busca por Barra + Bloque horario."
    ),
    "L": (
        "Hoja origen: Subastas, Cuadro 1. 1 si la central tiene subasta "
        "adjudicada (BAJADA o SUBIDA) en ese mes, día y hora; si no, 0."
    ),
    **{
        clave: f"{_SUBASTAS} {servicio}: suma de Energía SSCC [MW]."
        for clave, servicio in (
            ("AC", "CPF(-)"), ("AD", "CSF(-)"),
            ("AF", "CPF(+)"), ("AG", "CSF(+)"),
        )
    },
    **{
        clave: (
            f"{_SUBASTAS} {servicio}: suma del FD de esas filas (el FD de "
            "cada fila sale de la hoja FD por Hora mes + unidad)."
        )
        for clave, servicio in (
            ("AI", "CPF(-)"), ("AJ", "CSF(-)"),
            ("AL", "CPF(+)"), ("AM", "CSF(+)"),
        )
    },
    **{
        clave: (
            f"{_SUBASTAS} {servicio}: suma del FMA de esas filas (el FMA de "
            "cada fila sale de la hoja FMA)."
        )
        for clave, servicio in (("AO", "CPF(-)"), ("AP", "CSF(-)"))
    },
    "AR": "Constante 1, como en la planilla original (FMA de subida).",
    "AS": "Constante 1, como en la planilla original (FMA de subida).",
    "U": (
        "SoC × Capacidad (MWh) × 1000. La Capacidad sale de Centrales, "
        "Cuadro 1 Resumen BESS (se busca por Configuracion); sin ella, vacía."
    ),
    "V": (
        f"−(suma de la Carga de su {_GRUPO}) × Eficiencia. La Eficiencia "
        "sale de Centrales, Cuadro 1 Resumen BESS (se busca por "
        "Configuracion); sin ella, vacía."
    ),
    "S": (
        f"Puesto de la hora en la curva de su {_GRUPO}: (filas con mayor CMg "
        "Promedio + filas con igual CMg Promedio y Hora mayor) / 4 + 1."
    ),
    "BI": (
        "Número de la hora dentro de la ventana: empieza en 1 y sube en 1 "
        "cada 4 filas (4 cuartos de hora) mientras la ventana sea la misma."
    ),
    "BJ": (
        "Minuto del cuarto de hora de esta fila dentro de su hora (0, 15, "
        "30, 45), tomado de la fila 4 más arriba."
    ),
    "BM": (
        f"Curva de CMg descendente de su {_GRUPO}: la hora cuyo ranking cmg "
        "es igual al Orden de esta fila; de sus cuartos de hora, el k-ésimo "
        "de mayor CMg, con k = Periodo / 15 + 1."
    ),
    "BN": (
        "Energía asignada a este puesto: lo que cabe en el cuarto de hora "
        "(Pmax/4 × 1000 − inyección del periodo), sin pasar el Edisp_T de su "
        f"{_GRUPO} (cuadro 5) menos lo asignado en las filas anteriores. "
        "Pmax de Centrales, Cuadro 1 Resumen BESS; sin Pmax, vacía."
    ),
    "BS": (
        "Descarga del mismo cuarto de hora del que sale la Curva Cmg "
        "Decendente (el k-ésimo de mayor CMg de la hora del puesto)."
    ),
    "BT": (
        f"Suma de la Asignacion Edisponible de las filas siguientes de su "
        f"{_GRUPO}."
    ),
    "BU": (
        "Si la inyección del periodo es 0, 0. Si no, la Energia Total sin "
        "pasar el tope: Edisp_T (cuadro 5) − Energía ya Asignada − suma de "
        f"SSCC ultima hora de su {_GRUPO}."
    ),
    "BX": f"Margen ultima hora de su {_GRUPO} (cuadro 5).",
    "BZ": (
        "8 × flag ultima hora de su central + ventana (cuadro 5) si la "
        "inyeccion orden cronologico es mayor o igual a 0; si no, 0."
    ),
    "CE": (
        f"Por {_GRUPO}: máx(0, Σ Total C1 − Σ Total C2), repartido entre sus "
        "filas en proporción a la Reserva con FD y FMA (la de la fila / la "
        "suma del grupo). Arriba, el total del mes."
    ),
}

# Formula de Excel EQUIVALENTE para el comentario (mismo criterio que
# FORMULAS_PRUEBA_E_COSTOS). Las columnas de la tabla resumen (cuadro 5)
# van con prefijo R_ (R_Central, R_Ventana, R_Edisp_T, R_Margen, R_flag);
# en las plantillas del cuadro 5 la fila de ejemplo es la de la tabla
# resumen. Las que necesitan otra hoja (Pmax, Capacidad, Eficiencia) no
# van.
_GRUPO_RE = "{clave:r},{clave},{T:r},{T}"
_VENTANA_RE = "{R_Central:r},{clave},{R_Ventana:r},{T}"
_DESDE_RESUMEN = "{T:r},{R_Ventana},{clave:r},{R_Central}"
FORMULAS_PRUEBA_RE545 = {
    "S": (
        "=(COUNTIFS({T:r},{T},{R:r},\">\"&{R},{clave:r},{clave})"
        "+COUNTIFS({T:r},{T},{R:r},{R},{Hora:r},\">\"&{Hora},"
        "{clave:r},{clave}))/4+1"
    ),
    "BI": "=IF({T}={T@-4},{BI@-4}+1,1)",
    "BJ": "={BJ@-4}",
    "BT": "=SUMIFS({BU:b},{clave:b},{clave},{T:b},{T})",
    "BU": (
        "=IF(OR({BS}=\"\",{BS}=0),0,MAX(0,MIN({BQ},"
        f"SUMIFS({{R_Edisp_T:r}},{_VENTANA_RE})-{{BT}}"
        f"-SUMIFS({{BV:r}},{_GRUPO_RE}))))"
    ),
    "BX": f"=SUMIFS({{R_Margen:r}},{_VENTANA_RE})",
    "BZ": f"=IF({{BW}}<0,0,8)*SUMIFS({{R_flag:r}},{_VENTANA_RE})",
    "CE": (
        f"=IF(SUMIFS({{AU:r}},{_GRUPO_RE})=0,0,MAX(SUMIFS({{BO:r}},"
        f"{_GRUPO_RE})-SUMIFS({{CC:r}},{_GRUPO_RE}),0)*{{AU}}"
        f"/SUMIFS({{AU:r}},{_GRUPO_RE}))"
    ),
}
FORMULAS_PRUEBA_RESUMEN_RE545 = {
    "EiniT": (
        "=INDEX({U:r},MATCH(1,({clave:r}={R_Central})*({T:r}={R_Ventana}),0))"
    ),
    "EalmT": (
        "=INDEX({V:r},MATCH(1,({clave:r}={R_Central})*({T:r}={R_Ventana}),0))"
    ),
    "Total Reservas* FD *FMA": f"=SUMIFS({{AU:r}},{_DESDE_RESUMEN})",
    "Margen ultima hora": f"=SUMIFS({{BV:r}},{_DESDE_RESUMEN})",
    "check 1": f"=SUMIFS({{BN:r}},{_DESDE_RESUMEN})-{{R_Edisp_T}}",
    "check 2": (
        f"=SUMIFS({{BU:r}},{_DESDE_RESUMEN})-{{R_Edisp_T}}+{{R_Margen}}"
    ),
}
MATRICIALES_RE545 = ("EiniT", "EalmT")
# Nombre de columna de la tabla resumen -> clave R_ de las plantillas.
CLAVES_RESUMEN_RE545 = {
    "Central": "R_Central",
    "Ventana": "R_Ventana",
    "Edisp_T": "R_Edisp_T",
    "Margen ultima hora": "R_Margen",
    "flag ultima hora": "R_flag",
}

# Comentarios de la tabla resumen (cuadro 5), por nombre de columna.
ORIGEN_RESUMEN_RE545 = {
    "Oferta Completa": (
        "Hoja origen: Ofertas SSCC, Cuadro 2 Resumen ventana oferta "
        "(Completa), por central + ventana."
    ),
    "EiniT": "El EiniT de la primera fila de la central + ventana (cuadro 1).",
    "EalmT": "El EalmT de la primera fila de la central + ventana (cuadro 1).",
    "Total Reservas* FD *FMA": (
        "Suma de la Reserva con FD y FMA (cuadro 1) de la central + ventana."
    ),
    "Edisp_T": (
        "MIN(MAX(MIN(EiniT + EalmT, Capacidad × 1000), EalmT), Total "
        "Reservas) × Oferta Completa. Capacidad de Centrales, Cuadro 1 "
        "Resumen BESS. Es el tope de energía de la ventana."
    ),
    "check 1": "Control: suma de Energía asignada de la ventana − Edisp_T.",
    "check 2": (
        "Control: suma de Asignacion Edisponible de la ventana − Edisp_T + "
        "Margen ultima hora."
    ),
    "Margen ultima hora": (
        "Suma de SSCC ultima hora (cuadro 3) de la central + ventana."
    ),
    "flag ultima hora": (
        "1 si EiniT + EalmT − Margen ultima hora de la ventana anterior "
        "supera la Capacidad × 1000 y ese margen es distinto de 0."
    ),
}


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

    salida = (
        df_re545[COLUMNAS_SALIDA_RE545]
        .rename(columns=NOMBRES_CALCULO_RE545)
    )
    # La formula de la reserva (AU) no ve las CTF, que no se escriben:
    # solo se usa si su aporte (reserva x FD x FMA de CTF(-) y CTF(+))
    # es 0 en todas las filas. El FMA CTF(+) (AT) es la constante 1.
    def _num(columna):
        if columna not in df_re545.columns:
            return pd.Series(0.0, index=df_re545.index)
        return pd.to_numeric(df_re545[columna], errors="coerce").fillna(0.0)

    aporte_ctf = (
        _num("AE") * _num("AK") * _num("AQ")
        + _num("AH") * _num("AN") * _num("AT")
    )
    salida.attrs["ctf_en_cero"] = bool(aporte_ctf.eq(0.0).all())
    return salida
