# -*- coding: utf-8 -*-
"""
Claves_Balance — de las mediciones crudas a la tabla por clave del
balance, con el calendario de cuartos de hora del mes.

Viene de "2_generacion_claves_Balance.py". Cambios respecto del
original:

  - no glob-ea el directorio actual ni escribe parquets/Excel de
    diagnostico: recibe el DataFrame ya descargado y devuelve
    DataFrames (quien llama decide que guardar y que mostrar);
  - el umbral de "punto de medida completo" ya no es la constante
    2976 (= 31 dias x 96) sino la cantidad de cuartos de hora que el
    propio mes descargado trae. 2976 estaba mal para cualquier mes de
    30 dias o menos, y para los meses con cambio de hora.

Lo que NO cambia, porque es la logica que ya funciona: el orden real
de los cuartos de hora sale de `intervaloUtc` (la hora local se repite
en el cambio de hora y desordenaria la numeracion), el signo sale de
`Flujo`, y la agrupacion final es por clave.
"""

import pandas as pd

from .comun import ErrorMedidas


COL_INTERVALO_LOCAL = "intervalo"
COL_INTERVALO_UTC = "intervaloUtc"

# Las 9 columnas de Medidas_SAE.xlsx, en su orden exacto (es el mismo
# de nucleo.COLUMNAS_AI: esta hoja alimenta Medidores!A:I).
COLUMNAS_SAE = [
    "Mes",
    "Dia",
    "Hora",
    "Minutos",
    "Hora Mes",
    "Cuarto de Hora",
    "clave",
    COL_INTERVALO_LOCAL,
    "Gen_Unidad",
]


def parse_fecha_mixta(serie):
    """Fechas que pueden venir en formatos distintos en el mismo lote."""

    try:
        return pd.to_datetime(
            serie, dayfirst=True, errors="coerce", format="mixed"
        )
    except (TypeError, ValueError):
        return pd.to_datetime(serie, dayfirst=True, errors="coerce")


def _exigir_columnas(df, columnas, origen):
    faltantes = [c for c in columnas if c not in df.columns]
    if faltantes:
        raise ErrorMedidas(
            f"{origen} no trae la(s) columna(s) {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}"
        )


def normalizar_fechas(df):
    """Deja `intervalo` e `intervaloUtc` como datetime."""

    _exigir_columnas(
        df, [COL_INTERVALO_LOCAL, COL_INTERVALO_UTC], "La descarga de medidas"
    )

    df = df.copy()
    df[COL_INTERVALO_LOCAL] = parse_fecha_mixta(df[COL_INTERVALO_LOCAL])
    df[COL_INTERVALO_UTC] = parse_fecha_mixta(df[COL_INTERVALO_UTC])

    return df


def cuartos_de_hora_del_mes(df):
    """
    Cuantos cuartos de hora trae realmente el mes descargado: 2880 en
    un mes de 30 dias, 2976 en uno de 31, +/- 4 si hay cambio de hora.
    Reemplaza a la constante hardcodeada del script original.
    """

    return int(df[COL_INTERVALO_UTC].dropna().nunique())


def detectar_incompletos(df, esperados):
    """
    Puntos de medida a los que les faltan cuartos de hora. Mismo
    criterio que el original (cuenta de canalVal1/canalVal3 no nulos
    por punto de medida), con el umbral calculado en vez de fijo.
    """

    _exigir_columnas(
        df, ["idPuntoMedida", "canalVal1", "canalVal3", "principal"],
        "La descarga de medidas",
    )

    principal = df[df["principal"] == True].copy()  # noqa: E712

    conteos = (
        principal
        .groupby("idPuntoMedida")
        .agg(
            canalVal1=("canalVal1", "count"),
            canalVal3=("canalVal3", "count"),
        )
        .reset_index()
    )

    incompletos = conteos[
        (conteos["canalVal1"] != esperados)
        | (conteos["canalVal3"] != esperados)
    ].copy()

    return conteos, incompletos


def construir_calendario(df):
    """
    Calendario de cuartos de hora del mes, numerado por `intervaloUtc`
    (el orden real) pero descrito con la hora LOCAL, que es la que ve
    el balance.

    Devuelve un DataFrame con intervalo, intervaloUtc, Cuarto de Hora,
    Mes, Dia, Hora, Minutos y Hora Mes. `Hora` se cuenta dentro del
    dia local, asi que en un dia de cambio de hora llega a 23 o a 25 en
    vez de a 24 -- igual que en la planilla.
    """

    cal = (
        df[[COL_INTERVALO_LOCAL, COL_INTERVALO_UTC]]
        .dropna()
        .drop_duplicates()
        .copy()
    )

    if cal.empty:
        raise ErrorMedidas(
            "No se pudo armar el calendario de cuartos de hora: la "
            "descarga no trae intervalos validos."
        )

    cal["Periodo"] = cal[COL_INTERVALO_LOCAL].dt.to_period("M")
    cal["Fecha_Local"] = cal[COL_INTERVALO_LOCAL].dt.date

    cal = (
        cal
        .sort_values(["Periodo", COL_INTERVALO_UTC])
        .reset_index(drop=True)
    )

    cal["Cuarto de Hora"] = cal.groupby("Periodo").cumcount() + 1
    cal["Hora Mes"] = ((cal["Cuarto de Hora"] - 1) // 4) + 1

    cal["Mes"] = cal[COL_INTERVALO_LOCAL].dt.month
    cal["Dia"] = cal[COL_INTERVALO_LOCAL].dt.day
    cal["Minutos"] = cal[COL_INTERVALO_LOCAL].dt.minute

    cal["QH_Dia"] = cal.groupby("Fecha_Local").cumcount() + 1
    cal["Hora"] = ((cal["QH_Dia"] - 1) // 4) + 1

    return cal[
        [
            COL_INTERVALO_LOCAL,
            COL_INTERVALO_UTC,
            "Cuarto de Hora",
            "Mes",
            "Dia",
            "Hora",
            "Minutos",
            "Hora Mes",
        ]
    ]


def construir_por_clave(df, df_homol, registrar=print):
    """
    Del crudo descargado a (df_por_clave, calendario, diagnostico).

    df_por_clave trae ya las 9 columnas de Medidas_SAE.xlsx.
    diagnostico trae las tablas que el script original exportaba a
    'reporte_medidas_consolidadas.xlsx' (conteos por punto de medida,
    incompletos, generacion total por clave): no se escriben como
    archivo -- quien llama las resume en el log.
    """

    df = normalizar_fechas(df)

    esperados = cuartos_de_hora_del_mes(df)
    registrar(f"  cuartos de hora del mes: {esperados:,}")

    conteos, incompletos = detectar_incompletos(df, esperados)
    registrar(f"  puntos de medida incompletos: {len(incompletos):,}")

    df = df[~df["idPuntoMedida"].isin(incompletos["idPuntoMedida"])].copy()

    df = df[df["principal"] == True].copy()  # noqa: E712

    if df.empty:
        raise ErrorMedidas(
            "Despues de descartar los puntos de medida incompletos no "
            "quedo ningun registro principal. Revisa el periodo y la "
            "cobertura de la descarga."
        )

    df["canalVal"] = df["canalVal1"].fillna(0) + df["canalVal3"].fillna(0)

    _exigir_columnas(df, ["slugCanal"], "La descarga de medidas")

    df_cruzado = df.merge(
        df_homol,
        left_on=["idPuntoMedida", "slugCanal"],
        right_on=["Punto de Medida", "Canal"],
        how="inner",
    )

    if df_cruzado.empty:
        raise ErrorMedidas(
            "Ningun punto de medida descargado cruza con el archivo de "
            "homologacion (idPuntoMedida + slugCanal contra 'Punto de "
            "Medida' + 'Canal'). Revisa que el archivo de Auxiliares/ "
            "sea el del periodo."
        )

    registrar(
        f"  registros homologados: {len(df_cruzado):,} "
        f"({df_cruzado['clave'].nunique()} clave(s))"
    )

    # El signo del flujo: la API entrega magnitudes, el balance
    # necesita inyeccion (+) y retiro (-).
    df_cruzado["canalVal"] = (
        df_cruzado["canalVal"].abs() * df_cruzado["Flujo"]
    )

    calendario = construir_calendario(df_cruzado)

    df_cruzado = df_cruzado.merge(
        calendario,
        on=[COL_INTERVALO_LOCAL, COL_INTERVALO_UTC],
        how="left",
    )

    df_por_clave = (
        df_cruzado
        .groupby(
            [
                "Mes", "Dia", "Hora", "Minutos", "Hora Mes",
                "Cuarto de Hora", "clave", COL_INTERVALO_LOCAL,
            ],
            as_index=False,
        )
        .agg(Gen_Unidad=("canalVal", "sum"))
    )

    diagnostico = {
        "conteos": conteos,
        "incompletos": incompletos,
        "cuartos_esperados": esperados,
        "generacion_total": (
            df_por_clave
            .groupby("clave", as_index=False)
            .agg(Gen_Unidad=("Gen_Unidad", "sum"))
        ),
    }

    return df_por_clave[COLUMNAS_SAE], calendario, diagnostico
