# -*- coding: utf-8 -*-
"""
Hojas CMg, FD y Subastas del consolidado.
"""

import pandas as pd
from pathlib import Path

from .parametros import (
    HOJA_CMG_ORIGEN, HOJA_CPF_HORARIO, HOJA_CSF_HORARIO,
)
from .utiles import ErrorEntrada, _texto_seguro


# ============================================================
# CMg, FD, SUBASTAS
#
# Replican las macros de carga de esas tres hojas (no las macros que
# las consumen despues, como Asignar_CMg_a_Calculos_Turbo o
# Actualizar_Calculos_Columnas, que pertenecen a una etapa posterior
# todavia no implementada):
#   - Cargar_CMg_Desde_Archivo               -> leer_cmg
#   - Cargar_SSCC_Desempeno_En_FD            -> construir_fd
#
# La tercera, Cargar_Remuneracion_Subastas_Rapido, leia la hoja DB de
# la planilla 3 (3_REMUNERACIÓN_SUBASTAS_E_ID_*). Esa planilla no es el
# origen de las subastas -se arma pegando lo que sale de los Access
# OfertasSSCCAdj*.accdb- y el usuario confirmo que ya no se usa, asi
# que se saco: la hoja Subastas la arma construir_subastas_desde_accdb
# (Script/nucleo/subastas_accdb.py) leyendo esos Access.
#
# Ninguna de las tres tiene una hoja de referencia de dominio tan
# detallada como la de Medidores (Plan_Traspaso...): no se conocen
# nombres de negocio para casi ninguna columna mas alla de lo que las
# formulas de Excel revelan. Por eso las columnas que solo se copian
# (no se calculan) se nombran con su letra de Excel tal cual, en vez
# de inventarles un nombre que no esta documentado en ningun lado.
# ============================================================

def _contiene_bess_o_sae_sin_bat(texto):
    """
    Replica EsBESSoSAE (macros F_Leer_FD y G_Lee_Subastas): contiene
    "BESS" o "SAE". A diferencia de OSSCC_ContieneBESSoSAE (Ofertas
    SSCC), esta NO incluye "BAT" - son dos filtros distintos aunque se
    parezcan, no simplificar a una sola funcion.
    """

    texto = texto.upper()
    return "BESS" in texto or "SAE" in texto


def _dia_hora_mes_fd(columna_fecha, columna_hora):
    """
    Replica la formula compartida de FD (B, M, R y AE):
        =(DAY(fecha)-1)*24 + hora + 1 + IF(DAY(fecha)>100,1,0)

    El termino IF(DAY(fecha)>100,...) nunca es verdadero para un dia
    de calendario real (DAY() da 1..31): se conserva tal cual, tal
    como esta en la formula original, en vez de "limpiarla".
    """

    fecha = pd.to_datetime(columna_fecha, errors="coerce")
    dia = fecha.dt.day
    hora = pd.to_numeric(columna_hora, errors="coerce")
    ajuste_dia_mayor_100 = (dia > 100).astype("Int64")

    return (dia - 1) * 24 + hora + 1 + ajuste_dia_mayor_100


# --------------------------------------------------------------
# CMg
# --------------------------------------------------------------

def leer_cmg(ruta_cmg, registrar=print):
    """
    Replica Cargar_CMg_Desde_Archivo: lee cmg.xlsx (hoja "CMg" si
    existe, si no la primera hoja), columnas A:I desde la fila 2, y
    las ordena por columna D ascendente y luego H ascendente - el
    mismo orden que la macro aplica sobre el origen antes de pegarlo.

    No se renombran las columnas: se preserva el encabezado real del
    archivo (fila 1), igual que hace la macro al no tocarlo.
    """

    excel = pd.ExcelFile(ruta_cmg)

    if not excel.sheet_names:
        raise ErrorEntrada(
            f"{Path(ruta_cmg).name} no contiene hojas."
        )

    nombre_hoja = (
        HOJA_CMG_ORIGEN
        if HOJA_CMG_ORIGEN in excel.sheet_names
        else excel.sheet_names[0]
    )

    df = pd.read_excel(ruta_cmg, sheet_name=nombre_hoja)

    if df.shape[1] < 9:
        raise ErrorEntrada(
            f"{Path(ruta_cmg).name} debe tener al menos 9 columnas "
            f"(A:I) en la hoja '{nombre_hoja}'; tiene {df.shape[1]}."
        )

    df = df.iloc[:, :9].copy()

    columna_d = df.columns[3]
    columna_h = df.columns[7]

    df = (
        df
        .sort_values(
            by=[columna_d, columna_h],
            kind="mergesort",
            na_position="last",
        )
        .reset_index(drop=True)
    )

    registrar(
        f"  CMg: {len(df):,} filas leidas de {Path(ruta_cmg).name} "
        f"(hoja '{nombre_hoja}')"
    )

    return df


# --------------------------------------------------------------
# FD
# --------------------------------------------------------------

# Encabezados reales de FD (confirmados por el usuario contra un caso
# real, no inventados). Se aplican al final de cada bloque, despues de
# calcular todo con los nombres de letra (asi se evitan columnas
# duplicadas en el DataFrame mientras se opera con el, ya que "Hora
# Mes" se repite dos veces en cada bloque real - B y M en el CSF, R y
# AE en el CPF - igual que en el archivo original).
NOMBRES_FD_CSF = {
    "A": "id",
    "B": "Hora Mes",
    "C": "Dia",
    "D": "Fecha",
    "E": "Hora",
    "F": "Unidad",
    "G": "Respuesta CSF\n (Fact_CSF)",
    "H": "Disponibilidad\n(Fdis_CSF)",
    "I": "Desempeño\n(DCSF)",
    "J": "Factor de Desempeño\n (Fd_CSF)",
    "K": "CSF(+)",
    "L": "CSF(-)",
    "M": "Hora Mes",
}

NOMBRES_FD_CPF = {
    "Q": "id",
    "R": "Hora Mes",
    "S": "Dia",
    "T": "Fecha",
    "U": "Hora",
    "V": "Unidad",
    "W": "Respuesta CPF+\n(Fact_CPF+)",
    "X": "Respuesta CPF-\n(Fact_CPF-)",
    "Y": "Disponibilidad\n(Fdis_CPF)",
    "Z": "Desempeño\n(DCPF)",
    "AA": "Factor de Desempeño\n(Fd_CPF)",
    "AB": "Cuenta con equipo\nregistrador validado",
    "AC": "CPF(+)",
    "AD": "CPF(-)",
    "AE": "Hora Mes",
}

def _filtrar_bess_sae_posicional(df_bloque, indice_columna_filtro):
    """
    Replica FiltrarFilasBESSoSAE: conserva las filas donde la columna
    dada (posicion 0-indexada dentro de df_bloque) contiene "BESS" o
    "SAE".
    """

    textos = df_bloque.iloc[:, indice_columna_filtro].map(_texto_seguro)
    mascara = textos.map(_contiene_bess_o_sae_sin_bat)

    return df_bloque[mascara].reset_index(drop=True)


def _construir_bloque_fd_csf(df_filtrado):
    """
    A partir del bloque ya filtrado (7 columnas, origen B:H de "CSF
    Horario" en ese orden), arma las columnas A:M de FD tal como las
    escribe Cargar_SSCC_Desempeno_En_FD + sus formulas (plan de
    migracion, seccion de formulas de FD).
    """

    df = df_filtrado.iloc[:, :7].copy()
    df.columns = ["D", "E", "F", "G", "H", "I", "J"]
    df = df.reset_index(drop=True)

    df["B"] = _dia_hora_mes_fd(df["D"], df["E"])
    df["C"] = pd.to_datetime(df["D"], errors="coerce").dt.day
    df["A"] = (
        df["B"].astype("Int64").astype(str)
        + df["F"].map(_texto_seguro)
    )
    df["K"] = df["J"]
    df["L"] = df["K"]
    df["M"] = df["B"]

    df = df[list("ABCDEFGHIJKLM")]

    # Los nombres reales duplican "Hora Mes" (B y M): se renombra al
    # final, ya con las columnas en su posicion definitiva.
    return df.set_axis(
        [NOMBRES_FD_CSF[letra] for letra in df.columns], axis=1
    )


def _construir_bloque_fd_cpf(df_filtrado):
    """
    A partir del bloque ya filtrado (9 columnas, origen B:J de "CPF
    Horario" en ese orden), arma las columnas Q:AE de FD tal como las
    escribe Cargar_SSCC_Desempeno_En_FD + sus formulas.
    """

    df = df_filtrado.iloc[:, :9].copy()
    df.columns = ["T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"]
    df = df.reset_index(drop=True)

    df["R"] = _dia_hora_mes_fd(df["T"], df["U"])
    df["S"] = pd.to_datetime(df["T"], errors="coerce").dt.day
    df["Q"] = (
        df["R"].astype("Int64").astype(str)
        + df["V"].map(_texto_seguro)
    )
    df["AC"] = df["AA"]
    df["AD"] = df["AC"]
    df["AE"] = df["R"]

    df = df[list("QRSTUVWXYZ") + ["AA", "AB", "AC", "AD", "AE"]]

    # Los nombres reales duplican "Hora Mes" (R y AE): se renombra al
    # final, ya con las columnas en su posicion definitiva.
    return df.set_axis(
        [NOMBRES_FD_CPF[letra] for letra in df.columns], axis=1
    )


def construir_fd(ruta_sscc, registrar=print):
    """
    Replica Cargar_SSCC_Desempeno_En_FD.

    Lee, del archivo SSCC_Desempeño_*, las hojas "CPF Horario" y "CSF
    Horario" desde la fila 12, filtra por BESS/SAE en la columna D de
    cada una, y arma dos bloques independientes (distinto largo cada
    uno, igual que en la planilla): A:M (desde CSF) y Q:AE (desde
    CPF), con sus nombres de columna reales (NOMBRES_FD_CSF/
    NOMBRES_FD_CPF, confirmados por el usuario). N:P quedan fuera de
    alcance (la macro no las toca).

    Devuelve (df_csf, df_cpf).
    """

    ruta_sscc = Path(ruta_sscc)

    excel = pd.ExcelFile(ruta_sscc)

    for hoja in (HOJA_CPF_HORARIO, HOJA_CSF_HORARIO):
        if hoja not in excel.sheet_names:
            raise ErrorEntrada(
                f"No existe la hoja '{hoja}' en {ruta_sscc.name}."
            )

    df_cpf_crudo = pd.read_excel(
        ruta_sscc, sheet_name=HOJA_CPF_HORARIO, header=None
    )
    df_csf_crudo = pd.read_excel(
        ruta_sscc, sheet_name=HOJA_CSF_HORARIO, header=None
    )

    # Fila 12 de Excel (1-indexada) = indice 11 (0-indexado).
    # CPF: columnas B:J (9); CSF: columnas B:H (7).
    bloque_cpf = df_cpf_crudo.iloc[11:, 1:10]
    bloque_csf = df_csf_crudo.iloc[11:, 1:8]

    # D es la 3ra columna de cada bloque (B, C, D -> indice 2).
    filtrado_cpf = _filtrar_bess_sae_posicional(bloque_cpf, 2)
    filtrado_csf = _filtrar_bess_sae_posicional(bloque_csf, 2)

    df_csf = _construir_bloque_fd_csf(filtrado_csf)
    df_cpf = _construir_bloque_fd_cpf(filtrado_cpf)

    registrar(
        f"  FD: {len(df_csf):,} fila(s) CSF Horario, "
        f"{len(df_cpf):,} fila(s) CPF Horario (filtro BESS/SAE)"
    )

    return df_csf, df_cpf


# --------------------------------------------------------------
# SUBASTAS
# --------------------------------------------------------------

# Encabezados reales de Subastas!B:Q -- CORREGIDOS con el archivo
# Libro1.xlsx que trae la macro Cargar_Remuneracion_Subastas_Rapido y
# formulas reales cruzadas contra encabezados reales (sesion de
# correccion, ver BITACORA). La version anterior de este diccionario
# tenia TODO corrido una posicion: le faltaba la columna "Concepto"
# (B), la primera de las 11 que copia DB!B:L, que hasta esta sesion
# se asumia (mal) que era "A" y que la macro no tocaba.
#
# La prueba definitiva: Subastas!B1 tiene la formula real
#   =IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",IF(AND(C1="CSF",D1="BAJADA"),
#     "CSF(-)",IF(AND(C1="CTF",D1="SUBIDA"),"CTF(+)",
#     IF(AND(C1="CTF",D1="BAJADA"),"CTF(-)","REVISAR"))))
# que arma "Concepto" (B) a partir de DOS insumos: "Control" (C, el
# tipo SIN direccion: CSF/CTF/CPF) y "Sub_Baj" (D, la direccion:
# SUBIDA/BAJADA). La version anterior solo tenia UNA columna ahi
# (llamada "Control", en la posicion de lo que en realidad es
# "Concepto") -- faltaba la columna "Control" real.
#
# "A" SIGUE sin usarse (confirmado: ninguna celda con datos en esa
# columna en el archivo real) -- no era un error de ubicacion, era
# que faltaba contar una columna mas dentro del bloque B:L.
NOMBRES_SUBASTAS = {
    "B": "Concepto",
    "C": "Control",
    "D": "Sub_Baj",
    "E": "Fecha",
    "F": "Año",
    "G": "Mes",
    "H": "Dia",
    "I": "Hora_dia",
    "J": "Hora_mes",
    "K": "Configuración",
    "L": "Propietario",
    "M": "Clave horaria",
    "N": "Ciclo",
    "O": "Energía SSCC",
    "P": "FD",
    "Q": "FMA",
}


def _ordenar_subastas_por_hora_mes(df):
    """
    Ya NO ordena: deja la hoja Subastas en el orden en que viene del
    origen (los Access de 'DB subastas/', o la planilla 3).

    En una sesion anterior el usuario pidio ordenarla por Hora_mes y
    en esta pidio deshacerlo explicitamente ("al final que no este
    ordenado por hora_mes, mala mia yo lo pedi pero no"). Se conserva
    la funcion -en vez de borrar los llamados- porque el orden de esta
    hoja es puramente de presentacion: ninguna columna de mas abajo
    depende de el (calcular_l, construir_prorrata_sscc y
    construir_dic_umbrales_subastas cruzan por clave), asi que el
    unico lugar donde se decide es aca.
    """

    return df.reset_index(drop=True)


# La hoja 'CMg' del consolidado: la copia ya ordenada de cmg.xlsx
# (leer_cmg escribe sus 9 columnas A:I tal cual, con el encabezado
# real del archivo).
HOJA_CMG_CONSOLIDADO = "CMg"


def leer_cmg_consolidado(ruta_consolidado, registrar=print, libro=None):
    """
    La hoja 'CMg' de la planilla de salida, tal como la dejo
    leer_cmg(): las 9 columnas A:I, ya ordenadas.

    Mismo criterio que leer_fd_consolidado() y que 'Medidores' y
    'Subastas': el consolidado es la UNICA foto de las entradas con la
    que se calculan los pagos. Hasta ahora "Calculo E Costos" y
    "Calculo RE545" volvian a abrir cmg.xlsx aunque el mismo dato ya
    estuviera en el consolidado -- una planilla mas abierta en cada
    corrida y, peor, la posibilidad de que el CMg de los pagos no
    fuera el de la hoja 'CMg' (basta con dejar un cmg.xlsx mas nuevo
    en la carpeta despues de generar el consolidado).

    Las columnas NO se renombran (igual que leer_cmg): todo lo que las
    consume -construir_dic_cmg()- las busca por POSICION.

    libro: un pd.ExcelFile ya abierto del consolidado (ver
    leer_ofertas_sscc_consolidado).
    """

    ruta_consolidado = Path(ruta_consolidado)
    fuente = ruta_consolidado if libro is None else libro

    try:
        df = pd.read_excel(fuente, sheet_name=HOJA_CMG_CONSOLIDADO)
    except ValueError as error:
        raise ErrorEntrada(
            f"{ruta_consolidado.name} no tiene la hoja "
            f"'{HOJA_CMG_CONSOLIDADO}' todavia. Genera "
            f"esa hoja primero (tildando 'CMg')."
        ) from error

    if df.empty:
        raise ErrorEntrada(
            f"La hoja '{HOJA_CMG_CONSOLIDADO}' de "
            f"{ruta_consolidado.name} esta vacia. Genera "
            f"esa hoja primero (tildando 'CMg')."
        )

    if df.shape[1] < 9:
        raise ErrorEntrada(
            f"La hoja '{HOJA_CMG_CONSOLIDADO}' de "
            f"{ruta_consolidado.name} tiene {df.shape[1]} columnas y se "
            f"esperaban al menos 9 (A:I). Regenerala con su boton "
            f"'Actualizar'."
        )

    df = df.iloc[:, :9].copy()

    registrar(f"  filas de CMg: {len(df):,}")

    return df


# La hoja 'FD' del consolidado: los dos bloques quedan separados por
# columnas vacias (escritura.py escribe el CSF en A y el CPF en Q).
HOJA_FD_CONSOLIDADO = "FD"


def _bloques_de_columnas(df):
    """
    Grupos de columnas contiguas NO vacias de un DataFrame leido con
    header=None. Mismo criterio que _bloques_columnas_diccionario():
    una columna separa dos bloques cuando esta vacia en todas sus
    filas.
    """

    bloques = []
    actual = []

    for col in range(df.shape[1]):

        vacia = df.iloc[:, col].map(lambda v: not _texto_seguro(v)).all()

        if vacia:
            if actual:
                bloques.append(actual)
                actual = []
        else:
            actual.append(col)

    if actual:
        bloques.append(actual)

    return bloques


def leer_fd_consolidado(ruta_consolidado, registrar=print, libro=None):
    """
    Los dos bloques de la hoja 'FD' de la planilla de salida, tal
    como los dejo construir_fd(): (df_csf, df_cpf).

    Por que existe esta funcion (pedido del usuario: "Ecostos: se leen
    los sscc_desempeño, no deberia apuntar al consolidado de
    entradas?"): "Calculo E Costos" necesita el FD para AM:AR, y hasta
    ahora se lo armaba releyendo el archivo SSCC_Desempeño_* con
    construir_fd(). Eso significaba que la hoja FD del consolidado y el
    FD que usaba E Costos podian NO ser el mismo dato -- bastaba con
    que alguien dejara un SSCC_Desempeño_* mas nuevo en la carpeta
    despues de generar el consolidado. Ahora E Costos consume la hoja
    ya generada, igual que ya hacia con 'Medidores' y 'Subastas': el
    consolidado es la unica foto de las entradas.

    Los encabezados vienen en la primera fila de cada bloque y pueden
    estar REPETIDOS a proposito ("Hora Mes" sale dos veces en cada
    bloque), asi que se leen con header=None y se aplican con
    set_axis(), nunca con el header= de read_excel.

    libro: un pd.ExcelFile ya abierto del consolidado (ver
    leer_ofertas_sscc_consolidado).
    """

    ruta_consolidado = Path(ruta_consolidado)
    fuente = ruta_consolidado if libro is None else libro

    try:
        crudo = pd.read_excel(
            fuente, sheet_name=HOJA_FD_CONSOLIDADO, header=None
        )
    except ValueError as error:
        raise ErrorEntrada(
            f"{ruta_consolidado.name} no tiene la hoja "
            f"'{HOJA_FD_CONSOLIDADO}' todavia. Genera "
            f"esa hoja primero (tildando 'FD')."
        ) from error

    bloques = [
        columnas for columnas in _bloques_de_columnas(crudo)
        if len(columnas) > 1
    ]

    esperados = {
        len(NOMBRES_FD_CSF): ("csf", NOMBRES_FD_CSF),
        len(NOMBRES_FD_CPF): ("cpf", NOMBRES_FD_CPF),
    }

    encontrados = {}

    for columnas in bloques:

        esperado = esperados.get(len(columnas))

        if esperado is None:
            continue

        tipo, nombres = esperado

        if tipo in encontrados:
            continue

        bloque = crudo.iloc[1:, columnas].reset_index(drop=True)

        # Los dos bloques tienen distinto largo y conviven en la misma
        # hoja, asi que el mas corto viene con filas de relleno vacias
        # al final (las que ocupa el otro): se cortan.
        con_datos = bloque.map(lambda v: bool(_texto_seguro(v))).any(axis=1)
        ultima = con_datos[con_datos].index.max()
        bloque = (
            bloque.iloc[: int(ultima) + 1]
            if pd.notna(ultima) else bloque.iloc[:0]
        )

        encontrados[tipo] = bloque.set_axis(list(nombres.values()), axis=1)

    faltan = [t for t in ("csf", "cpf") if t not in encontrados]

    if faltan:
        raise ErrorEntrada(
            f"La hoja '{HOJA_FD_CONSOLIDADO}' de "
            f"{ruta_consolidado.name} no trae el/los bloque(s) "
            f"{[t.upper() for t in faltan]} (se esperaban dos bloques de "
            f"{len(NOMBRES_FD_CSF)} y {len(NOMBRES_FD_CPF)} columnas, "
            f"separados por columnas vacias; se encontraron bloques de "
            f"{[len(c) for c in bloques]} columnas). Volve a generar la "
            f"hoja 'FD' con su boton 'Actualizar'."
        )

    registrar(
        f"  FD desde el consolidado: {len(encontrados['csf']):,} fila(s) "
        f"CSF, {len(encontrados['cpf']):,} fila(s) CPF."
    )

    return encontrados["csf"], encontrados["cpf"]
