# -*- coding: utf-8 -*-
"""
Hojas CMg, FD y Subastas del consolidado.
"""

import pandas as pd
from pathlib import Path

from .cambio_hora import hora_mes_correlativa
from .parametros import (
    COLUMNA_CENTRAL, HOJA_CMG_ORIGEN, HOJA_CPF_HORARIO, HOJA_CSF_HORARIO,
)
from .utiles import ErrorEntrada, _entero_a_texto, _texto_seguro


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


def _hora_mes_fd_correlativa(columna_fecha, columna_hora):
    """
    Hora Mes de FD contada igual que la de Medidores (ver
    cambio_hora.hora_mes_correlativa). La formula del Excel
    ((dia-1)*24 + hora + 1) supone dias de 24 horas y, desde el dia del
    cambio de hora, E Costos cruzaba el FD de la hora equivocada.

    Las horas del FD van de 0 a 23 (periodo = hora + 1); en abril el
    sabado del cambio llega a la 24 (periodo 25). La hora que no existe
    (la 0 del dia que se adelanta) queda sin Hora Mes.
    """

    return hora_mes_correlativa(
        columna_fecha, pd.to_numeric(columna_hora, errors="coerce") + 1
    )


def _poner_en_cero_hora_inexistente(df, columna_hora_mes, columna_fecha,
                                    columnas_valor):
    """La hora que no existe (sin Hora Mes) queda con sus valores en 0.

    El FD la trae ("No participó", y alguna unidad con valores): se
    descarta lo que traiga y se deja en 0, pedido del usuario. Devuelve
    (df, filas, filas que traian algun valor distinto de 0).
    """

    inexistente = (
        df[columna_hora_mes].isna()
        & pd.to_datetime(df[columna_fecha], errors="coerce").notna()
    )
    if not inexistente.any():
        return df, 0, 0

    numeros = df.loc[inexistente, columnas_valor].apply(
        pd.to_numeric, errors="coerce"
    )
    con_valor = int(numeros.fillna(0).ne(0).any(axis=1).sum())
    df[columnas_valor] = df[columnas_valor].astype(object)
    df.loc[inexistente, columnas_valor] = 0
    return df, int(inexistente.sum()), con_valor


def _id_fd(hora_mes, unidad):
    """Hora Mes & Unidad; vacio si la fila no tiene Hora Mes."""

    texto = hora_mes.astype("Int64").astype(str) + unidad.map(_texto_seguro)
    return texto.where(hora_mes.notna(), "")


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

    with pd.ExcelFile(ruta_cmg) as excel:
        hojas = list(excel.sheet_names)

    if not hojas:
        raise ErrorEntrada(
            f"{Path(ruta_cmg).name} no contiene hojas."
        )

    nombre_hoja = (
        HOJA_CMG_ORIGEN
        if HOJA_CMG_ORIGEN in hojas
        else hojas[0]
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

    df["B"] = _hora_mes_fd_correlativa(df["D"], df["E"])
    df["C"] = pd.to_datetime(df["D"], errors="coerce").dt.day
    df["A"] = _id_fd(df["B"], df["F"])
    df["K"] = df["J"]
    df["L"] = df["K"]
    df["M"] = df["B"]
    df, filas, con_valor = _poner_en_cero_hora_inexistente(
        df, "B", "D", list("GHIJKL")
    )
    df.attrs["hora_inexistente"] = (filas, con_valor)

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

    df["R"] = _hora_mes_fd_correlativa(df["T"], df["U"])
    df["S"] = pd.to_datetime(df["T"], errors="coerce").dt.day
    df["Q"] = _id_fd(df["R"], df["V"])
    df["AC"] = df["AA"]
    df["AD"] = df["AC"]
    df["AE"] = df["R"]
    df, filas, con_valor = _poner_en_cero_hora_inexistente(
        df, "R", "T", ["W", "X", "Y", "Z", "AA", "AC", "AD"]
    )
    df.attrs["hora_inexistente"] = (filas, con_valor)

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

    with pd.ExcelFile(ruta_sscc) as excel:
        hojas = list(excel.sheet_names)

    for hoja in (HOJA_CPF_HORARIO, HOJA_CSF_HORARIO):
        if hoja not in hojas:
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

    for nombre, df in (("CSF", df_csf), ("CPF", df_cpf)):
        filas, con_valor = df.attrs.get("hora_inexistente", (0, 0))
        if filas:
            registrar(
                f"  FD {nombre}: {filas:,} fila(s) en la hora que no existe "
                f"por el cambio de hora; quedan en 0 y sin Hora Mes"
                + (
                    f" ({con_valor:,} traian valores y se descartaron)."
                    if con_valor else "."
                )
            )

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
# correccion, ver BITACORA_BESS). La version anterior de este diccionario
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


# La hoja 'Subastas' del consolidado: sin las columnas que se arman de
# las otras (Control y Sub_Baj salen de Concepto, la Clave horaria de
# Configuración + Dia + Hora_dia) ni la de Ciclo, que va vacia (el ciclo
# se calcula en Calculo E Costos). Pedido del usuario: antes B, C, L y M
# de la hoja. Se reponen al leerla.
COLUMNAS_SUBASTAS_AUXILIARES = ("Control", "Sub_Baj", "Clave horaria", "Ciclo")


def subastas_para_hoja(df_subastas):
    """La hoja Subastas como se escribe: sin las auxiliares."""

    if df_subastas is None:
        return None
    return df_subastas.drop(
        columns=[c for c in COLUMNAS_SUBASTAS_AUXILIARES if c in df_subastas],
    )


def reponer_auxiliares_subastas(df_subastas):
    """
    Vuelve a armar Control, Sub_Baj, Clave horaria y Ciclo (vacia) de
    una hoja Subastas que no las trae, y deja las columnas en el orden
    de siempre (NOMBRES_SUBASTAS). Una hoja de antes, que si las trae,
    queda igual.
    """

    df = df_subastas.copy()
    concepto = df["Concepto"].map(_texto_seguro) if "Concepto" in df else None

    if "Control" not in df and concepto is not None:
        df["Control"] = concepto.str[:3]
    if "Sub_Baj" not in df and concepto is not None:
        # =IF(LEFT(RIGHT(C,2),1)="+","SUBIDA","BAJADA")
        df["Sub_Baj"] = concepto.map(
            lambda t: "SUBIDA" if t[-2:-1] == "+" else "BAJADA"
        )
    if "Clave horaria" not in df and {"Configuración", "Dia", "Hora_dia"} <= set(df):
        df["Clave horaria"] = (
            df["Configuración"].map(_texto_seguro)
            + df["Dia"].map(_entero_a_texto)
            + df["Hora_dia"].map(_entero_a_texto)
        )
    if "Ciclo" not in df:
        df["Ciclo"] = pd.NA

    orden = [n for n in NOMBRES_SUBASTAS.values() if n in df]
    return df[orden + [c for c in df.columns if c not in orden]]


# La hoja 'FD' del consolidado: los dos bloques quedan separados por
# una columna vacia. En la hoja NO van las columnas auxiliares (el "id"
# = Hora Mes & Unidad, y la copia de "Hora Mes" del final de cada
# bloque, antes A, M, Q y AE): se reponen al leer, asi que el calculo
# sigue viendo los mismos bloques. 'Central' queda al final.
HOJA_FD_CONSOLIDADO = "FD"


def fd_para_hoja(df_fd):
    """El bloque FD como va en la hoja: sin 'id' ni la copia de Hora Mes."""

    if df_fd is None:
        return None
    if len(df_fd.columns) == 0 or df_fd.columns[0] != "id":
        return df_fd
    nombres = list(df_fd.columns)
    # La copia de "Hora Mes" es la segunda aparicion del nombre.
    copia = [i for i, n in enumerate(nombres) if n == "Hora Mes"][1:]
    quedan = [i for i in range(1, len(nombres)) if i not in copia]
    return df_fd.iloc[:, quedan]


def _reponer_auxiliares_fd(bloque, nombres_completos, con_central):
    """De un bloque compacto (sin id ni copia) al bloque completo."""

    hora_mes = pd.to_numeric(bloque["Hora Mes"], errors="coerce")
    completo = pd.DataFrame(index=bloque.index)
    completo["id"] = _id_fd(hora_mes, bloque["Unidad"])
    for nombre in bloque.columns:
        if nombre == COLUMNA_CENTRAL:
            continue
        completo[nombre] = bloque[nombre]
    columnas = list(completo.columns) + ["Hora Mes (copia)"]
    completo = pd.concat(
        [completo, bloque["Hora Mes"].rename("Hora Mes (copia)")], axis=1
    ).set_axis(columnas, axis=1)
    if con_central:
        completo[COLUMNA_CENTRAL] = bloque[COLUMNA_CENTRAL]
    return completo.set_axis(
        nombres_completos + ([COLUMNA_CENTRAL] if con_central else []),
        axis=1,
    )


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

    # Dos formas de bloque: la de antes (con 'id' y la copia de Hora
    # Mes, empieza con "id") y la actual (sin las dos, empieza con
    # "Hora Mes"). Un CSF de antes y un CPF de ahora tienen el mismo
    # largo: se distinguen por la primera columna.
    esperados = {
        ("id", len(NOMBRES_FD_CSF)): ("csf", NOMBRES_FD_CSF, False),
        ("id", len(NOMBRES_FD_CPF)): ("cpf", NOMBRES_FD_CPF, False),
        ("Hora Mes", len(NOMBRES_FD_CSF) - 2): ("csf", NOMBRES_FD_CSF, True),
        ("Hora Mes", len(NOMBRES_FD_CPF) - 2): ("cpf", NOMBRES_FD_CPF, True),
    }

    encontrados = {}

    for columnas in bloques:

        # Desde que se importa con la columna 'Central' (el nombre
        # Balance_BESS de la unidad, ver agregar_central_fd), cada
        # bloque trae una columna mas al final. Un libro de antes no la
        # trae: los dos se leen igual.
        con_central = (
            _texto_seguro(crudo.iloc[0, columnas[-1]]) == COLUMNA_CENTRAL
        )
        primera = _texto_seguro(crudo.iloc[0, columnas[0]])
        esperado = esperados.get((primera, len(columnas) - int(con_central)))

        if esperado is None:
            continue

        tipo, nombres, compacto = esperado

        if tipo in encontrados:
            continue

        nombres_completos = list(nombres.values())
        nombres = nombres_completos + (
            [COLUMNA_CENTRAL] if con_central else []
        )
        if compacto:
            nombres = [
                _texto_seguro(crudo.iloc[0, c]) for c in columnas
            ]

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

        bloque = bloque.set_axis(nombres, axis=1)
        if compacto:
            bloque = _reponer_auxiliares_fd(
                bloque, nombres_completos, con_central
            )
        encontrados[tipo] = bloque

    faltan = [t for t in ("csf", "cpf") if t not in encontrados]

    if faltan:
        raise ErrorEntrada(
            f"La hoja '{HOJA_FD_CONSOLIDADO}' de "
            f"{ruta_consolidado.name} no trae el/los bloque(s) "
            f"{[t.upper() for t in faltan]} (se esperaban dos bloques de "
            f"{len(NOMBRES_FD_CSF) - 2} y {len(NOMBRES_FD_CPF) - 2} columnas "
            f"mas 'Central', "
            f"separados por columnas vacias; se encontraron bloques de "
            f"{[len(c) for c in bloques]} columnas). Volve a generar la "
            f"hoja 'FD' con su boton 'Actualizar'."
        )

    registrar(
        f"  FD desde el consolidado: {len(encontrados['csf']):,} fila(s) "
        f"CSF, {len(encontrados['cpf']):,} fila(s) CPF."
    )

    return encontrados["csf"], encontrados["cpf"]
