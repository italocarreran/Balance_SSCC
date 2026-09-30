# -*- coding: utf-8 -*-
"""
Homologacion — el Excel que dice que punto de medida + canal es cada
clave del balance.

Viene de "0_diccionario_prmte_a_claves_balance.py", que leia
'Homologacion ClavesTF y PRMTE.xlsx' del directorio actual y lo
guardaba como homol.parquet para que el paso siguiente lo levantara.
Ese parquet intermedio ya no existe: el archivo se lee una vez, desde
Auxiliares/ (al lado de Centrales.xlsx), y queda en memoria.

El mismo archivo trae ademas la hoja "Gen real" (ver HOJA_GEN_REAL):
las centrales que se miden por la API de operacion real. Estaba
pensada como una hoja de Centrales.xlsx, pero el usuario pidio que
viva aca -- es homologacion, igual que 'homol', y se mantiene con el
mismo archivo.
"""

from pathlib import Path

import pandas as pd

from .comun import (
    ErrorMedidas, EXTENSIONES_EXCEL, columna_que_contenga, normalizar
)


# El nombre real es 'Homologacion ClavesTF y PRMTE.xlsx', pero se busca
# por patron (como *OfertasSSCC*) para no depender de tildes, guiones o
# de que alguien le agregue el periodo al nombre.
PATRON_NOMBRE = "homologacion"
HOJA_HOMOL = "homol"

# Segunda hoja del mismo archivo: las centrales cuya medida NO sale de
# la API por punto de medida sino de la API de operacion real. Vive
# aca, y no en Centrales.xlsx, a pedido del usuario: es homologacion,
# igual que 'homol', y se mantiene con el mismo archivo.
#
# Mismas cuatro columnas que 'homol' (clave / Punto de Medida / Canal /
# Flujo), con una lectura propia de cada una:
#
#   clave           la clave del balance, igual que en 'homol'
#   Punto de Medida el 'topologyName' EXACTO de la API de operacion
#                   real (ej. "SAE PFV Andes Solar III (Inyección)"):
#                   es lo que identifica a la central en esa API, que
#                   no tiene el concepto de punto de medida
#   Canal           la UNIDAD en la que viene la medida ("MWh" o
#                   "kWh"). Esa API no expone canales, asi que la
#                   columna quedaba sin uso; el usuario le escribio
#                   "MWh" y preguntó si afectaba. Ahora si: es lo que
#                   decide el factor de conversion (ver
#                   UNIDADES_GEN_REAL). Vacia o con cualquier otro
#                   texto = MWh, que es lo que devuelve la API
#   Flujo           +1 / -1, igual que en 'homol' (retiros en -1)
#
# La hoja es opcional: un caso sin centrales de este tipo es valido.
HOJA_GEN_REAL = "Gen real"

COLUMNA_PUNTO = "Punto de Medida"
COLUMNA_CANAL = "Canal"
COLUMNA_CLAVE = "clave"
COLUMNA_FLUJO = "Flujo"

# Unidades que se reconocen en la columna "Canal" de la hoja
# "Gen real", y el factor que lleva cada una a kWh, que es la unidad en
# la que trabaja TODO el balance (Medidores!Gen_Unidad y, mas
# adelante, "Descarga kWh"/"Carga kWh").
#
# La API de operacion real devuelve MWh: sin esta conversion, las
# centrales que se miden por ahi entran al balance mil veces mas
# chicas que las demas. Se vio con un caso real: Andes Solar III
# (Pmax 170,78 MW) llegaba con un maximo de 44 por cuarto de hora
# cuando Tocopilla (116 MW) llegaba con 29.493 -- los 44 son MWh y los
# 29.493 son kWh.
UNIDADES_GEN_REAL = {
    "mwh": 1000.0,
    "kwh": 1.0,
}

# Como se escribe cada una en el log.
ETIQUETA_UNIDAD_GEN_REAL = {"mwh": "MWh", "kwh": "kWh"}

# Lo que se usa cuando la columna "Canal" viene vacia o con un texto
# que no empieza con ninguna de las dos unidades: la API devuelve MWh.
UNIDAD_GEN_REAL_POR_DEFECTO = "mwh"


def unidad_desde_canal(valor):
    """
    La unidad que dice la columna "Canal" de la hoja "Gen real".

    Se mira SOLO EL PRINCIPIO del texto: en el archivo real el canal
    viene escrito "MWhD" / "MWhR" (y podria venir "kWhD" / "kWhR"), o
    sea la unidad con el tipo de medida pegado atras. Lo unico que
    decide aca es la unidad -si hay que multiplicar por mil o no-; la
    D y la R del final no significan nada para esta cuenta.

    Antes se comparaba el texto ENTERO contra "mwh"/"kwh", asi que
    cualquier canal con sufijo caia en el valor por defecto: cambiar la
    columna de MWhD a kWhD no cambiaba nada y la central entraba igual
    multiplicada por mil.

    Vacio, o un texto que no empieza con ninguna de las dos, vale MWh,
    que es lo que devuelve la API de operacion real.
    """

    texto = normalizar(valor)

    for unidad in UNIDADES_GEN_REAL:
        if texto.startswith(unidad):
            return unidad

    return UNIDAD_GEN_REAL_POR_DEFECTO


def buscar_archivo_homologacion(carpeta_auxiliares):
    """
    Archivo de homologacion dentro de Auxiliares/: cualquier Excel
    cuyo nombre contenga "homologacion" (sin tildes). Si hay varios,
    el mas reciente por fecha de modificacion. None si no hay ninguno.
    """

    carpeta = Path(carpeta_auxiliares)

    if not carpeta.is_dir():
        return None

    candidatos = [
        archivo for archivo in carpeta.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() in EXTENSIONES_EXCEL
        and PATRON_NOMBRE in normalizar(archivo.stem)
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda a: a.stat().st_mtime)


def leer_homologacion(ruta):
    """
    Lee la hoja 'homol' y devuelve el DataFrame con las cuatro
    columnas que usa el cruce posterior:

        Punto de Medida | Canal | clave | Flujo

    'clave' se fuerza a texto (asi lo hacia el script original: hay
    claves que son solo numeros y pandas las leeria como int, y
    despues no cruzarian contra el resto del balance).
    """

    ruta = Path(ruta)

    try:
        excel = pd.ExcelFile(ruta)
    except Exception as error:
        raise ErrorMedidas(
            f"No se pudo abrir {ruta.name}: {error}"
        ) from error

    hoja = _buscar_hoja(excel, HOJA_HOMOL)

    if hoja is None:
        raise ErrorMedidas(
            f"{ruta.name} no tiene la hoja '{HOJA_HOMOL}'. "
            f"Hojas encontradas: {excel.sheet_names}"
        )

    df = pd.read_excel(ruta, sheet_name=hoja)

    columnas = {}
    for interno, fragmentos in (
        (COLUMNA_PUNTO, ("punto", "medida")),
        (COLUMNA_CANAL, ("canal",)),
        (COLUMNA_CLAVE, ("clave",)),
        (COLUMNA_FLUJO, ("flujo",)),
    ):
        real = columna_que_contenga(df, *fragmentos)
        if real is None:
            raise ErrorMedidas(
                f"La hoja '{hoja}' de {ruta.name} no tiene una columna "
                f"'{interno}'. Columnas encontradas: {list(df.columns)}"
            )
        columnas[interno] = real

    df = df.rename(columns={real: interno for interno, real in columnas.items()})
    df[COLUMNA_CLAVE] = df[COLUMNA_CLAVE].astype(str)

    df = df.dropna(subset=[COLUMNA_PUNTO, COLUMNA_CANAL])

    if df.empty:
        raise ErrorMedidas(
            f"La hoja '{hoja}' de {ruta.name} no tiene ninguna fila con "
            f"'Punto de Medida' y 'Canal' cargados."
        )

    return df[[COLUMNA_PUNTO, COLUMNA_CANAL, COLUMNA_CLAVE, COLUMNA_FLUJO]]


def _buscar_hoja(excel, nombre):
    for hoja in excel.sheet_names:
        if normalizar(hoja) == normalizar(nombre):
            return hoja
    return None


def leer_gen_real(ruta):
    """
    Lee la hoja "Gen real" y devuelve una lista de dicts con
    'topologyName', 'clave', 'factor' y 'unidad' -- la forma que
    espera Generacion_Real.

    'unidad' sale del PRINCIPIO de la columna "Canal" ("MWhD",
    "MWhR", "kWh", ...): lo que decide es la unidad, no el sufijo (ver
    unidad_desde_canal). Vacia o con un texto que no empieza con
    ninguna de las dos vale MWh, que es lo que devuelve la API de
    operacion real (ver UNIDADES_GEN_REAL).

    Si la hoja no existe, devuelve lista vacia y el proceso sigue: un
    caso sin centrales de operacion real es valido.
    """

    ruta = Path(ruta)
    excel = pd.ExcelFile(ruta)

    hoja = _buscar_hoja(excel, HOJA_GEN_REAL)

    if hoja is None:
        return []

    df = pd.read_excel(ruta, sheet_name=hoja)

    columna_clave = columna_que_contenga(df, "clave")
    columna_punto = columna_que_contenga(df, "punto", "medida")
    columna_flujo = columna_que_contenga(df, "flujo")
    columna_canal = columna_que_contenga(df, "canal")

    if columna_clave is None or columna_punto is None:
        raise ErrorMedidas(
            f"La hoja '{hoja}' de {ruta.name} tiene que tener las mismas "
            f"columnas que '{HOJA_HOMOL}' ('clave', 'Punto de Medida', "
            f"'Canal', 'Flujo'). Columnas encontradas: "
            f"{list(df.columns)}"
        )

    centrales = []

    for _, fila in df.iterrows():

        clave = fila[columna_clave]
        topology = fila[columna_punto]

        if pd.isna(clave) or pd.isna(topology):
            continue

        factor = 1.0

        if columna_flujo is not None and not pd.isna(fila[columna_flujo]):
            try:
                factor = float(fila[columna_flujo])
            except (TypeError, ValueError):
                raise ErrorMedidas(
                    f"En la hoja '{hoja}' de {ruta.name}, la central "
                    f"'{topology}' tiene un Flujo no numerico "
                    f"({fila[columna_flujo]!r}). Usa 1 o -1."
                )

        unidad = (
            unidad_desde_canal(fila[columna_canal])
            if columna_canal is not None
            else UNIDAD_GEN_REAL_POR_DEFECTO
        )

        centrales.append(
            {
                "topologyName": str(topology).strip(),
                "clave": str(clave).strip(),
                "factor": factor,
                "unidad": unidad,
            }
        )

    return centrales


def puntos_de_medida(df_homol):
    """Los puntos de medida distintos a consultar en la API."""

    return list(pd.unique(df_homol[COLUMNA_PUNTO].dropna()))
