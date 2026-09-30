# -*- coding: utf-8 -*-
"""
Medidas_SAE.xlsx y el maestro Centrales.xlsx.
"""

import pandas as pd

from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_MEDIDAS_SAE, COLUMNAS_AI,
    HOJA_DICCIONARIO, HOJA_MEDIDAS_SAE, HOJA_RESUMEN_BESS,
)
from .utiles import ErrorEntrada, _texto_seguro, _tiene_valor, normalizar


# ============================================================
# LECTURA DE MEDIDAS SAE (A:I)
# ============================================================

def leer_medidas_sae(ruta):
    """Lee Medidas_SAE.xlsx y valida que traiga las 9 columnas."""

    df = pd.read_excel(ruta, sheet_name=HOJA_MEDIDAS_SAE)

    faltantes = [
        columna for columna in COLUMNAS_AI
        if columna not in df.columns
    ]

    if faltantes:
        raise ErrorEntrada(
            f"A {ARCHIVO_MEDIDAS_SAE} le faltan columnas:\n"
            f"  {faltantes}\n"
            f"Columnas encontradas:\n"
            f"  {list(df.columns)}"
        )

    df = df[COLUMNAS_AI].copy()

    df["intervalo"] = pd.to_datetime(df["intervalo"])

    return df


# ============================================================
# LECTURA DEL MAESTRO
# ============================================================

def _leer_hoja_con_encabezado(
    ruta, nombre_hoja, columnas_buscadas, filas_a_revisar=15
):
    """
    Lee una hoja detectando la fila de encabezados en vez de asumir
    que es la primera: las hojas reales suelen traer un titulo arriba
    ("Cuadro N° 1: Resumen BESS"). Mismo criterio que
    detectar_fila_nombres() para el SoC: nunca una posicion fija, se
    busca la fila que contiene los textos esperados.

    columnas_buscadas: lista de tuplas de fragmentos; una fila sirve
    como encabezado si, para CADA tupla, alguna de sus celdas contiene
    todos los fragmentos de esa tupla. Ej.
    [("nombre", "activ"), ("barra",)] exige una fila con una celda
    tipo "Nombre activo" y otra tipo "Barra inyección".
    """

    crudo = pd.read_excel(
        ruta, sheet_name=nombre_hoja, header=None, nrows=filas_a_revisar
    )

    fila_encabezado = None

    for indice in range(len(crudo)):

        textos = [normalizar(v) for v in crudo.iloc[indice].tolist()]

        if all(
            any(all(f in t for f in fragmentos) for t in textos)
            for fragmentos in columnas_buscadas
        ):
            fila_encabezado = indice
            break

    if fila_encabezado is None:
        esperadas = ", ".join(
            "+".join(fragmentos) for fragmentos in columnas_buscadas
        )
        raise ErrorEntrada(
            f"No se encontro, en las primeras {filas_a_revisar} filas de "
            f"la hoja '{nombre_hoja}' de {ARCHIVO_CENTRALES}, una fila de "
            f"encabezados con las columnas esperadas ({esperadas})."
        )

    return pd.read_excel(ruta, sheet_name=nombre_hoja, header=fila_encabezado)


def _leer_resumen_bess(ruta, nombre_hoja, filas_a_revisar=15):
    """
    "Resumen BESS": la fila de encabezados es la que trae
    'Nombre activo' y 'Barra inyección'.
    """

    return _leer_hoja_con_encabezado(
        ruta, nombre_hoja, [("nombre", "activ"), ("barra",)], filas_a_revisar
    )


def leer_centrales(ruta):
    """Devuelve (resumen_bess, diccionario) como DataFrames."""

    excel = pd.ExcelFile(ruta)

    def buscar_hoja(nombre_buscado):
        for hoja in excel.sheet_names:
            if normalizar(hoja) == normalizar(nombre_buscado):
                return hoja
        raise ErrorEntrada(
            f"{ARCHIVO_CENTRALES} no tiene la hoja "
            f"'{nombre_buscado}'. Hojas: {excel.sheet_names}"
        )

    resumen = _leer_resumen_bess(ruta, buscar_hoja(HOJA_RESUMEN_BESS))

    diccionario = pd.read_excel(
        ruta,
        sheet_name=buscar_hoja(HOJA_DICCIONARIO),
        header=None,
    )

    return resumen, diccionario


def _bloques_columnas_diccionario(diccionario):
    """
    Detecta los bloques de columnas de la hoja Diccionario: son
    VARIAS tablas de equivalencia independientes puestas una al lado
    de la otra (encontrado con un Diccionario real: encabezados "FD"
    en A1, "Subastas" en E1, "ofertas" en G1, con las columnas C:D
    completamente vacias separando el primer bloque del segundo).

    Una columna separa dos bloques cuando esta VACIA EN TODAS LAS
    FILAS del archivo (no alcanza con mirar una sola fila: la fila de
    encabezados, por ejemplo, suele tener texto solo en la primera
    columna de cada bloque). Devuelve una lista de listas de indices
    de columna, un grupo por bloque.
    """

    vacia = []
    for col in range(diccionario.shape[1]):
        serie = diccionario.iloc[:, col]
        vacia.append(
            serie.map(lambda v: not _tiene_valor(v)).all()
        )

    bloques = []
    actual = []

    for col, es_vacia in enumerate(vacia):
        if es_vacia:
            if actual:
                bloques.append(actual)
                actual = []
        else:
            actual.append(col)

    if actual:
        bloques.append(actual)

    return bloques


# ============================================================
# HOJA "Diccionario": los dos formatos
# ============================================================
#
# FORMATO VIEJO (el que trae Centrales.xlsx hasta hoy): varias tablas
# de equivalencia independientes puestas lado a lado, separadas por
# columnas completamente vacias, con el titulo de cada una en la
# primera fila y sin titulo en el resto de sus columnas:
#
#     A: FD           E: Subastas              G: ofertas
#     -----------     ---------------------    --------------
#     SAE-TOCOPILLA   SAE-TOCOPILLA            (vacio)
#     SAE TOCOPILLA   BAT_TOCOPILLA
#
# Ese formato tiene dos agujeros que el usuario ya sufrio:
#   - no hay lugar para la nomenclatura de FMA CPF (el bloque
#     "FMA CPF" nunca existio), asi que el FMA de las filas CPF se
#     buscaba con la Configuración tal cual y no encontraba nada; y
#   - el bloque "FD" arranca en la columna del nombre Balance_BESS,
#     asi que homologar la Configuración de Subastas (que para
#     Tocopilla es BAT_TOCOPILLA) contra el nombre del archivo de
#     desempeño tampoco encontraba.
#
# FORMATO NUEVO (propuesta del usuario, Centrales_Propuesta_de_mejora):
# UNA sola tabla, con una fila de encabezados que nombra cada columna
# y una fila por central:
#
#     Balance_BESS   FD                 Subastas        Ofertas   FMA_CPF
#     ------------   ----------------   -------------   -------   -----------------
#     SAE-TOCOPILLA  SAE TOCOPILLA      BAT_TOCOPILLA             Tocopilla - BESS
#
# La primera columna es el nombre canonico (el que usa el resto del
# programa) y cada una de las demas dice como se llama esa misma
# central en cada origen.
#
# Se aceptan los dos: si la hoja tiene la fila de encabezados del
# formato nuevo se usa esa, y si no se cae al recorrido por bloques de
# siempre. Ninguna de las dos lecturas cambia de resultado para las
# centrales que ya funcionaban.

ROL_BALANCE_BESS = "balance_bess"
ROL_FD = "fd"
ROL_SUBASTAS = "subastas"
ROL_OFERTAS = "ofertas"
ROL_FMA_CPF = "fma_cpf"

# Encabezado (normalizado, con "_" tratado como espacio) -> rol.
TITULOS_DICCIONARIO = {
    "balance bess": ROL_BALANCE_BESS,
    "fd": ROL_FD,
    "subastas": ROL_SUBASTAS,
    "ofertas": ROL_OFERTAS,
    "fma cpf": ROL_FMA_CPF,
}

# Cuantas filas de arriba se miran buscando la fila de encabezados: el
# archivo real empieza con una fila vacia, asi que nunca es la 0 fija.
FILAS_ENCABEZADO_DICCIONARIO = 10


def _clave_titulo(valor):
    """'FMA_CPF' -> 'fma cpf' (para comparar encabezados)."""

    return normalizar(valor).replace("_", " ")


def encabezado_diccionario(
    diccionario, filas_a_revisar=FILAS_ENCABEZADO_DICCIONARIO
):
    """
    Ubica la fila de encabezados del FORMATO NUEVO de la hoja
    Diccionario y devuelve (indice_de_esa_fila, {rol: indice_columna}).

    Devuelve (None, {}) si la hoja esta en el formato viejo de bloques:
    ahi no hay ninguna fila que nombre la columna "Balance_BESS", que
    es justamente lo que distingue un formato del otro.
    """

    limite = min(filas_a_revisar, len(diccionario))

    for indice in range(limite):

        columnas = {}

        for col in range(diccionario.shape[1]):
            rol = TITULOS_DICCIONARIO.get(
                _clave_titulo(diccionario.iloc[indice, col])
            )
            if rol is not None and rol not in columnas:
                columnas[rol] = col

        # Hace falta la columna canonica y al menos un origen: una fila
        # suelta que diga "FD" (el titulo de bloque del formato viejo)
        # no alcanza para confundirla con un encabezado.
        if ROL_BALANCE_BESS in columnas and len(columnas) >= 2:
            return indice, columnas

    return None, {}


def filas_diccionario(diccionario):
    """
    Las filas de DATOS de la hoja Diccionario en el formato nuevo, como
    lista de dicts {rol: texto}. Lista vacia si la hoja esta en el
    formato viejo.
    """

    fila_encabezado, columnas = encabezado_diccionario(diccionario)

    if fila_encabezado is None:
        return []

    filas = []

    for indice in range(fila_encabezado + 1, len(diccionario)):

        fila = {
            rol: _texto_seguro(diccionario.iloc[indice, col])
            for rol, col in columnas.items()
        }

        if not _tiene_valor(fila.get(ROL_BALANCE_BESS)):
            continue

        filas.append(fila)

    return filas


def mapa_diccionario(diccionario, rol_clave, rol_valor):
    """
    {nombre normalizado de rol_clave -> texto de rol_valor} en el
    formato nuevo. Gana la primera aparicion, igual que el resto de las
    lecturas de esta hoja. {} si la hoja esta en el formato viejo o si
    alguno de los dos roles no esta.
    """

    mapa = {}

    for fila in filas_diccionario(diccionario):

        clave = normalizar(fila.get(rol_clave))
        valor = fila.get(rol_valor)

        if clave and _tiene_valor(valor) and clave not in mapa:
            mapa[clave] = valor

    return mapa


def construir_homologacion(diccionario):
    """
    Arma un mapa nombre_origen -> nombre_canonico a partir de la
    hoja Diccionario.

    TRAMPA REAL (encontrada con un Diccionario real, no en los
    sinteticos): la hoja NO es "una fila = todos los sinonimos de una
    central". Son VARIAS tablas independientes de equivalencia
    puestas lado a lado por columnas (ej. "FD" en A:B, "Subastas" en
    E:F:G -- ver _bloques_columnas_diccionario()), y el orden de
    filas de una tabla NO tiene por que coincidir con el de la de al
    lado (se vio con datos reales: para la mayoria de las centrales
    las dos tablas coinciden fila a fila por casualidad, pero para las
    ultimas 2-3 centrales el orden se corre). Tratar la fila entera
    como un solo grupo de sinonimos (como hacia esta funcion antes)
    mezclaba centrales de una tabla con las de la otra en esas filas
    corridas -- por ejemplo, homologaba una central hacia la central
    de la fila de al lado, no hacia si misma.

    Ahora cada BLOQUE de columnas se procesa por separado: dentro de
    un bloque, cada fila SI aporta equivalencias entre todos sus
    textos no vacios (esa parte de la estrategia original era
    correcta), pero un bloque nunca contribuye equivalencias con
    otro. Si dos bloques distintos terminan dando canonicos distintos
    para el mismo texto normalizado, gana el primero encontrado
    (mismo criterio que ya usaba esta funcion dentro de una fila).
    """

    mapa = {}

    # FORMATO NUEVO: cada fila ya dice cual es el nombre canonico (la
    # columna Balance_BESS), asi que no hay que adivinarlo con "el
    # primer valor no vacio de la fila", y la fila de encabezados no se
    # confunde con una central (era el riesgo de leerla como bloque:
    # "Balance_BESS", "FD", "Subastas"... habrian quedado registrados
    # como sinonimos entre si).
    filas = filas_diccionario(diccionario)

    if filas:

        for fila in filas:

            canonico = _texto_seguro(fila.get(ROL_BALANCE_BESS))

            for valor in fila.values():
                if _tiene_valor(valor):
                    mapa.setdefault(normalizar(valor), canonico)

        return mapa

    for columnas_bloque in _bloques_columnas_diccionario(diccionario):

        sub = diccionario.iloc[:, columnas_bloque]

        for _, fila in sub.iterrows():

            valores = [
                str(v).strip()
                for v in fila.tolist()
                if v is not None
                and str(v).strip() != ""
                and str(v).strip().lower() != "nan"
            ]

            if len(valores) < 2:
                continue

            canonico = valores[0]

            for valor in valores:
                mapa.setdefault(normalizar(valor), canonico)

    return mapa
