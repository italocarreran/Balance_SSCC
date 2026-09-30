# -*- coding: utf-8 -*-
"""
El SoC por bloques, desde el archivo del SCADA.
"""

import pandas as pd

from .utiles import ErrorEntrada, normalizar


# ============================================================
# EXTRACCION DEL SoC POR BLOQUES
# ============================================================

def detectar_fila_nombres(df_crudo, maximo_filas=30):
    """
    Busca la fila que contiene los nombres de centrales.

    Criterio: se sube desde la fila de encabezados 'Time Stamp'/
    'Value' hasta encontrar una fila con contenido util, y desde ahi
    se sigue subiendo mientras las filas sigan teniendo contenido
    util (sin saltar un hueco en blanco) -- se devuelve la MAS
    ARRIBA de ese bloque contiguo, no la primera que se encuentra.

    Hace falta este segundo paso porque un archivo real (visto con
    datos reales) puede traer DOS filas de metadata pegadas justo
    arriba del hueco en blanco que precede a los encabezados: la fila
    con el nombre limpio de la central, y debajo (mas cerca de
    'Time Stamp'/'Value') una fila con informacion adicional (ej. la
    ruta SCADA completa del punto). Quedarse con "la primera fila util
    subiendo" agarra la fila equivocada (la de mas informacion, no la
    del nombre limpio) en ese caso -- subir hasta el tope del bloque
    contiguo trae la de mas arriba, que es la que tiene el nombre.
    """

    fila_encabezados = None

    for indice in range(min(maximo_filas, len(df_crudo))):

        textos = {
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        }

        if "time stamp" in textos and "value" in textos:
            fila_encabezados = indice
            break

    if fila_encabezados is None:
        raise ErrorEntrada(
            "No se encontro ninguna fila con los encabezados "
            "'Time Stamp' y 'Value' en el archivo SOC."
        )

    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    def _fila_tiene_contenido_util(indice):
        textos = [
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        ]
        return any(t not in encabezados_bloque for t in textos)

    fila_nombres = None

    for indice in range(fila_encabezados - 1, -1, -1):

        if _fila_tiene_contenido_util(indice):
            # Sigue siendo candidata mientras haya contenido util;
            # se actualiza en cada vuelta para quedarse con la MAS
            # ARRIBA del bloque contiguo, no la primera encontrada.
            fila_nombres = indice
            continue

        # Fila en blanco: si ya se encontro una candidata, el bloque
        # contiguo termino aca -- se corta la busqueda.
        if fila_nombres is not None:
            break

    if fila_nombres is None:
        raise ErrorEntrada(
            "Se encontraron los encabezados 'Time Stamp'/'Value' "
            "pero no una fila de nombres de centrales arriba."
        )

    return fila_nombres, fila_encabezados


def detectar_bloques(df_crudo, fila_nombres, fila_encabezados):
    """
    Para cada nombre de la fila de nombres, define su bloque
    horizontal y ubica dentro sus columnas Time Stamp y Value.

    No se usan offsets fijos: los encabezados se buscan por
    texto dentro del rango de cada central.
    """

    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    fila_n = df_crudo.iloc[fila_nombres].tolist()

    posiciones = []

    for indice, valor in enumerate(fila_n):
        texto = normalizar(valor)
        if texto and texto not in encabezados_bloque:
            posiciones.append((indice, str(valor).strip()))

    if not posiciones:
        raise ErrorEntrada(
            "La fila de nombres no contiene ninguna central."
        )

    fila_e = [
        normalizar(v)
        for v in df_crudo.iloc[fila_encabezados].tolist()
    ]

    bloques = []
    incidencias = []

    for orden, (columna_inicio, nombre) in enumerate(posiciones):

        if orden + 1 < len(posiciones):
            columna_fin = posiciones[orden + 1][0] - 1
        else:
            columna_fin = len(fila_e) - 1

        columnas_ts = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "time stamp"
        ]

        columnas_val = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "value"
        ]

        if len(columnas_ts) != 1 or len(columnas_val) != 1:
            incidencias.append(
                f"{nombre}: se esperaba exactamente un "
                f"'Time Stamp' y un 'Value' entre las columnas "
                f"{columna_inicio + 1} y {columna_fin + 1}; "
                f"se encontraron {len(columnas_ts)} y "
                f"{len(columnas_val)}. Bloque excluido."
            )
            continue

        bloques.append(
            {
                "nombre_bess_origen": nombre,
                "columna_inicio_bloque": columna_inicio,
                "columna_fin_bloque": columna_fin,
                "columna_timestamp": columnas_ts[0],
                "columna_value": columnas_val[0],
            }
        )

    return bloques, incidencias


def _extraer_nombre_desde_ruta_scada(texto):
    r"""
    Si el nombre de un bloque de SoC viene como una ruta SCADA (visto
    con datos reales: exportaciones tipo PI traen el nombre de la
    central como
        \\SERVIDOR\SEN\Generación\SEN\<region>\<central>|<sufijo>
    en vez de solo "<central>"), devuelve unicamente "<central>": el
    ultimo tramo de la ruta (separado por barras invertidas), sin el
    sufijo despues de "|".

    Este docstring es CRUDO (lleva una r adelante) justamente por esas
    barras invertidas: si no, la ruta del ejemplo no seria texto sino
    escapes, y Python avisa "SyntaxWarning: invalid escape sequence"
    al importar el modulo.

    No es una reinterpretacion de datos: es separar una ESTRUCTURA
    conocida (ruta + sufijo) que ya viene asi en el archivo, no una
    suposicion sobre a que central corresponde. Si el texto no tiene
    ese formato (ninguna barra invertida), se devuelve tal cual -- no
    se inventa nada quitando texto de un nombre que no es una ruta.
    """

    texto = str(texto).strip()

    if "\\" not in texto:
        return texto

    return texto.split("\\")[-1].split("|")[0].strip()


def extraer_soc(ruta_soc, mapa_homologacion=None):
    """
    Lee el archivo SOC y devuelve (df_soc, incidencias).

    df_soc: central | timestamp | soc | nombre_scada_original
    """

    df_crudo = pd.read_excel(
        ruta_soc,
        sheet_name=0,
        header=None,
    )

    fila_nombres, fila_encabezados = detectar_fila_nombres(
        df_crudo
    )

    bloques, incidencias = detectar_bloques(
        df_crudo,
        fila_nombres,
        fila_encabezados,
    )

    mapa_homologacion = mapa_homologacion or {}

    partes = []

    for bloque in bloques:

        nombre_origen = bloque["nombre_bess_origen"]

        # Primero se prueba el texto literal (compatibilidad con
        # cualquier archivo de SoC "limpio", sin ruta SCADA). Si no
        # hay match, se prueba de nuevo con el nombre extraido de la
        # ruta (ver _extraer_nombre_desde_ruta_scada) -- el
        # Diccionario puede tener registrada cualquiera de las dos
        # formas. Si ninguna tiene match, se usa igual el nombre
        # extraido (no la ruta completa) como "canonico": aunque no
        # homologue, es mucho mas legible en avisos/incidencias que
        # la ruta cruda, y no cambia el comportamiento (sigue sin
        # cruzar contra Medidores).
        clave_directa = normalizar(nombre_origen)
        nombre_limpio = _extraer_nombre_desde_ruta_scada(nombre_origen)

        if clave_directa in mapa_homologacion:
            canonico = mapa_homologacion[clave_directa]
        else:
            canonico = mapa_homologacion.get(
                normalizar(nombre_limpio),
                nombre_limpio,
            )

        sub = df_crudo.iloc[
            fila_encabezados + 1:,
            [
                bloque["columna_timestamp"],
                bloque["columna_value"],
            ],
        ].copy()

        sub.columns = ["timestamp", "soc"]

        sub["timestamp"] = pd.to_datetime(
            sub["timestamp"],
            errors="coerce",
        )

        sub["soc"] = pd.to_numeric(
            sub["soc"],
            errors="coerce",
        )

        sub = sub.dropna(subset=["timestamp"])

        if sub.empty:
            incidencias.append(
                f"{nombre_origen}: bloque sin registros validos."
            )
            continue

        if sub["soc"].isna().any():
            incidencias.append(
                f"{nombre_origen}: "
                f"{int(sub['soc'].isna().sum())} valores de SoC "
                f"no numericos."
            )

        duplicados = sub["timestamp"].duplicated().sum()

        if duplicados:
            incidencias.append(
                f"{nombre_origen}: {duplicados} timestamps "
                f"duplicados."
            )

        sub["central"] = canonico
        sub["nombre_scada_original"] = nombre_origen

        partes.append(sub)

    if not partes:
        raise ErrorEntrada(
            "No se pudo extraer SoC de ningun bloque.\n"
            + "\n".join(incidencias)
        )

    df_soc = pd.concat(partes, ignore_index=True)

    df_soc = df_soc[
        [
            "central",
            "timestamp",
            "soc",
            "nombre_scada_original",
        ]
    ]

    return df_soc, incidencias
