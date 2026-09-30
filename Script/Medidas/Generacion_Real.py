# -*- coding: utf-8 -*-
"""
Generacion_Real — las centrales cuya medida NO sale de la
homologacion por punto de medida, sino de la API de operacion real
(opreal), por nombre de topologia.

Viene de "3_Generacion_Real.py". Tres cambios de fondo, pedidos por el
usuario:

  - la lista de centrales ya no esta escrita en el codigo
    (FILTROS_TOPOLOGY): sale de la hoja "Gen real" del Excel de
    homologacion, que ademas dice con que `clave` tiene que aparecer
    cada una en Medidas_SAE.xlsx (ver Homologacion.HOJA_GEN_REAL);
  - ya no es un REEMPLAZO. Estas centrales se sacan de la hoja
    'homol', asi que no llegan por el otro camino: lo que hace este
    modulo es AGREGAR las filas de la lista a la tabla del balance;
  - no escribe Excel: devuelve DataFrames.

Tambien se dejo de generar el 'log_inconsistencias_medidas.xlsx' del
original, que comparaba el criterio de desempate viejo contra el de
mayor idMeasure: esa comparacion era una investigacion ya cerrada (el
propio script la titula "CRITERIO DEFINITIVO: mayor idMeasure"). Lo
que si se informa en el log es cuantos grupos venian duplicados.

La API entrega el dato HORARIO; el balance trabaja en cuartos de hora,
asi que cada hora se reparte en 4 partes iguales (energia/4), igual
que el script original.

Y lo entrega en **MWh**, mientras que el balance entero trabaja en
**kWh**: la conversion la hace expandir_a_cuartos() con la unidad que
dice la columna "Canal" de la hoja "Gen real" (ver UNIDADES_GEN_REAL en
Homologacion.py). Sin esa conversion estas centrales entraban mil veces
mas chicas que las que vienen por punto de medida -- se detecto con un
caso real, comparando Andes Solar III (Pmax 170,78 MW, maximo 44 por
cuarto de hora) contra Tocopilla (116 MW, maximo 29.493).
"""

import time

import pandas as pd

from .comun import ErrorMedidas, leer_clave_api, CLAVE_GENERACION_REAL
from .Homologacion import (
    ETIQUETA_UNIDAD_GEN_REAL, UNIDADES_GEN_REAL,
    UNIDAD_GEN_REAL_POR_DEFECTO,
)


URL_OPREAL = "https://operacion.api.coordinador.cl/opreal-medidas/v1/bydate"

TIPO = 1
TIMEOUT = 60
REINTENTOS = 3
ESPERA_REINTENTO = 1.0

COLUMNA_TOPOLOGY = "topologyName"
COLUMNA_HORA = "topologyTimeHour"
COLUMNA_VALOR = "measure"
COLUMNA_ID = "idMeasure"

COLUMNAS_GRUPO = ["fecha_consulta", COLUMNA_TOPOLOGY, COLUMNA_HORA]


def _json_a_dataframe(datos):
    """La respuesta puede venir como lista o envuelta en un dict."""

    if isinstance(datos, list):
        return pd.DataFrame(datos)

    if isinstance(datos, dict):
        for campo in ("data", "results", "content"):
            if campo in datos and isinstance(datos[campo], list):
                return pd.DataFrame(datos[campo])
        return pd.json_normalize(datos)

    return pd.DataFrame()


def descargar_mes(
    anio, mes, ultimo_dia, user_key=None, registrar=print,
    progreso=None, desde=0, hasta=100,
):
    """
    Baja el mes completo, dia por dia, sin filtrar: el filtro por
    nombre de topologia se aplica despues (asi el log puede decir que
    nombres trajo la API cuando alguno de la lista no aparece).
    """

    try:
        import requests
    except ImportError as error:
        raise ErrorMedidas(
            "Falta la libreria 'requests' (pip install -r "
            "requirements.txt)."
        ) from error

    user_key = user_key or leer_clave_api(CLAVE_GENERACION_REAL)

    sesion = requests.Session()
    partes = []

    for dia in range(1, ultimo_dia + 1):

        fecha = f"{anio}-{mes:02d}-{dia:02d}"
        params = {"created": fecha, "type": TIPO, "user_key": user_key}

        for intento in range(REINTENTOS):

            try:
                respuesta = sesion.get(URL_OPREAL, params=params, timeout=TIMEOUT)

                if respuesta.ok:
                    df_dia = _json_a_dataframe(respuesta.json())

                    if not df_dia.empty:
                        df_dia.insert(0, "fecha_consulta", fecha)
                        partes.append(df_dia)

                    break

                if intento == REINTENTOS - 1:
                    registrar(
                        f"  AVISO: {fecha} -> HTTP "
                        f"{respuesta.status_code}"
                    )
                else:
                    time.sleep(ESPERA_REINTENTO)

            except Exception as error:
                if intento == REINTENTOS - 1:
                    registrar(f"  AVISO: {fecha} -> {error}")
                else:
                    time.sleep(ESPERA_REINTENTO)

        if progreso:
            progreso(desde + (hasta - desde) * dia / ultimo_dia)

    if not partes:
        raise ErrorMedidas(
            f"La API de operacion real no devolvio datos para "
            f"{anio}-{mes:02d}. Revisa la clave (user_key) y que el "
            f"periodo ya este publicado."
        )

    df = pd.concat(partes, ignore_index=True)

    faltantes = [
        c for c in
        ["fecha_consulta", COLUMNA_TOPOLOGY, COLUMNA_HORA, COLUMNA_VALOR,
         COLUMNA_ID]
        if c not in df.columns
    ]

    if faltantes:
        raise ErrorMedidas(
            f"La respuesta de la API de operacion real no trae "
            f"{faltantes}. Columnas recibidas: {list(df.columns)}"
        )

    registrar(f"  registros descargados: {len(df):,}")

    return df


def filtrar_y_desempatar(df, centrales, registrar=print):
    """
    Deja una sola fila por dia + topologia + hora, quedandose con el
    idMeasure mayor (criterio definitivo del script original; ante el
    mismo idMeasure, prefiere una medida distinta de cero).

    centrales: filas de la hoja "Gen real" -- dicts con
    'topologyName', 'clave' y 'factor' (ver
    Homologacion.leer_gen_real).
    """

    if not centrales:
        return df.iloc[0:0].copy()

    nombres = [str(c["topologyName"]).strip() for c in centrales]

    df = df.copy()
    df[COLUMNA_TOPOLOGY] = df[COLUMNA_TOPOLOGY].astype(str).str.strip()

    buscados = {n.casefold() for n in nombres}
    df = df[df[COLUMNA_TOPOLOGY].str.casefold().isin(buscados)].copy()

    encontrados = {n.casefold() for n in df[COLUMNA_TOPOLOGY].unique()}
    sin_datos = [n for n in nombres if n.casefold() not in encontrados]

    if sin_datos:
        registrar(
            f"  AVISO: {len(sin_datos)} nombre(s) de la hoja 'Medidas "
            f"API' no aparecen en la API: {', '.join(sin_datos)}"
        )

    if df.empty:
        raise ErrorMedidas(
            "Ninguna de las centrales de la hoja 'Gen real' aparece "
            "en la API de operacion real. La columna 'Punto de Medida' "
            "de esa hoja tiene que traer el 'topologyName' exacto (por "
            "ejemplo 'SAE PFV Andes Solar III (Inyección)')."
        )

    df[COLUMNA_VALOR] = pd.to_numeric(df[COLUMNA_VALOR], errors="coerce")
    df[COLUMNA_ID] = pd.to_numeric(df[COLUMNA_ID], errors="coerce")
    df[COLUMNA_HORA] = pd.to_numeric(df[COLUMNA_HORA], errors="coerce")

    df["_no_cero"] = df[COLUMNA_VALOR].fillna(0).ne(0).astype(int)

    duplicados = (
        df.groupby(COLUMNAS_GRUPO, dropna=False).size().gt(1).sum()
    )

    if duplicados:
        registrar(
            f"  grupos dia/central/hora con mas de un registro: "
            f"{duplicados:,} (se toma el idMeasure mayor)"
        )

    df = (
        df
        .sort_values(
            by=COLUMNAS_GRUPO + [COLUMNA_ID, "_no_cero"],
            ascending=[True, True, True, False, False],
            na_position="last",
        )
        .drop_duplicates(subset=COLUMNAS_GRUPO, keep="first")
        .drop(columns=["_no_cero"])
        .reset_index(drop=True)
    )

    registrar(f"  registros horarios definitivos: {len(df):,}")

    return df


def expandir_a_cuartos(df_horario, centrales, registrar=print):
    """
    Cada fila horaria se abre en 4 cuartos de hora con la energia
    repartida en partes iguales, y se le pega la `clave` (y el factor
    de signo) que le corresponde segun la hoja "Gen real".

    Devuelve un DataFrame con `intervalo` (hora local, inicio del
    cuarto), `clave` y `Gen_Unidad`.
    """

    if df_horario.empty:
        return pd.DataFrame(columns=["intervalo", "clave", "Gen_Unidad"])

    # topologyName -> (clave, factor de signo, factor de unidad).
    #
    # El factor de unidad es la correccion que faltaba: la API entrega
    # MWh y el balance entero trabaja en kWh (ver UNIDADES_GEN_REAL).
    # Sin el, estas centrales entraban mil veces mas chicas que las que
    # vienen por punto de medida.
    mapa = {
        str(c["topologyName"]).strip().casefold(): (
            str(c["clave"]).strip(),
            float(c.get("factor", 1) or 1),
            UNIDADES_GEN_REAL[
                c.get("unidad") or UNIDAD_GEN_REAL_POR_DEFECTO
            ],
        )
        for c in centrales
    }

    df = df_horario.loc[df_horario.index.repeat(4)].copy()
    df["cuarto_hora"] = df.groupby(level=0).cumcount() + 1

    # topologyTimeHour: 1 -> 00:00-00:45, 24 -> 23:00-23:45.
    df["intervalo"] = (
        pd.to_datetime(df["fecha_consulta"])
        + pd.to_timedelta(df[COLUMNA_HORA] - 1, unit="h")
        + pd.to_timedelta((df["cuarto_hora"] - 1) * 15, unit="m")
    )

    claves_factores = (
        df[COLUMNA_TOPOLOGY].str.casefold().map(mapa)
    )

    df["clave"] = [par[0] if isinstance(par, tuple) else None
                   for par in claves_factores]
    factores = pd.Series(
        [par[1] if isinstance(par, tuple) else 1.0 for par in claves_factores],
        index=df.index,
    )
    unidades = pd.Series(
        [par[2] if isinstance(par, tuple) else 1.0 for par in claves_factores],
        index=df.index,
    )

    df["Gen_Unidad"] = (df[COLUMNA_VALOR] / 4) * factores * unidades

    df = df.dropna(subset=["clave"])

    salida = (
        df
        .groupby(["intervalo", "clave"], as_index=False)
        .agg(Gen_Unidad=("Gen_Unidad", "sum"))
    )

    registrar(
        f"  filas cuarto-horarias generadas: {len(salida):,} "
        f"({salida['clave'].nunique()} clave(s))"
    )

    # Que unidad se le aplico a cada central queda dicho en el log: es
    # un factor de 1000 y una equivocacion ahi no se ve a simple vista
    # en la hoja Medidores, solo al comparar contra otra central.
    for central in centrales:
        unidad = central.get("unidad") or UNIDAD_GEN_REAL_POR_DEFECTO
        registrar(
            f"    {central['clave']}: la medida se lee en "
            f"{ETIQUETA_UNIDAD_GEN_REAL[unidad]} y se pasa a kWh "
            f"(x{UNIDADES_GEN_REAL[unidad]:g})"
        )

    return salida


def pegar_calendario(df_cuartos, calendario, registrar=print):
    """
    Le pega a las filas de opreal el MISMO calendario de cuartos de
    hora que se armo con el otro camino (ver Claves_Balance).

    Es a proposito que no se numere aparte: "Cuarto de Hora" es un
    indice global del mes que despues cruza contra CMg, asi que las
    dos fuentes tienen que compartirlo o un dia de cambio de hora las
    desalinearia.

    La API de operacion real entrega fecha + hora local (no UTC), asi
    que el cruce es por hora local. En un dia de cambio de hora hacia
    atras la hora local se repite y el calendario tiene dos filas para
    ese instante: se toma la primera y se avisa.
    """

    if df_cuartos.empty:
        return df_cuartos.assign(**{
            "Mes": [], "Dia": [], "Hora": [], "Minutos": [],
            "Hora Mes": [], "Cuarto de Hora": [],
        })

    cal = calendario.drop_duplicates(subset=["intervalo"], keep="first")

    repetidos = len(calendario) - len(cal)
    if repetidos:
        registrar(
            f"  AVISO: {repetidos} cuarto(s) de hora con hora local "
            f"repetida (cambio de hora). Para las centrales de la hoja "
            f"'Gen real' se toma la primera ocurrencia: la API de "
            f"operacion real no entrega hora UTC para distinguirlas."
        )

    df = df_cuartos.merge(
        cal[["intervalo", "Cuarto de Hora", "Mes", "Dia", "Hora",
             "Minutos", "Hora Mes"]],
        on="intervalo",
        how="left",
    )

    sin_calendario = int(df["Cuarto de Hora"].isna().sum())

    if sin_calendario == len(df):
        raise ErrorMedidas(
            "Ninguna de las filas de la API de operacion real cruza "
            "contra el calendario de cuartos de hora armado con las "
            "medidas por punto de medida. Las dos fuentes tienen que "
            "referirse al mismo periodo y al mismo huso: revisa el "
            "AAMM antes de seguir."
        )

    if sin_calendario:
        registrar(
            f"  AVISO: {sin_calendario:,} fila(s) de la API de "
            f"operacion real quedaron fuera del calendario del mes y "
            f"no se agregan (probablemente de otro periodo)."
        )
        df = df.dropna(subset=["Cuarto de Hora"])

    return df
