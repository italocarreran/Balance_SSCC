# -*- coding: utf-8 -*-
"""
RE545: los tres bloques de reservas (AC:AT) y AU.
"""

import pandas as pd

from .alertas import ALTA, Alerta, anotar_muchas
from .utiles import (
    ErrorEntrada, _normaliza_valor_vba, _texto_seguro, _tiene_valor,
)


# ------------------------------------------------------------
# CALCULO RE545 (etapa 2): AC:AU -- reservas por subasta
#
# Tres bloques de 6 columnas con los MISMOS 6 encabezados
# (CPF(-), CSF(-), CTF(-), CPF(+), CSF(+), CTF(+)), que se
# distinguen por el titulo de grupo de la fila 2: "Subastas"
# (AC:AH), "FD" (AI:AN) y "FMA" (AO:AT). Los tres son el mismo
# SUMIFS contra Subastas, cambiando la columna que se suma:
#
#   AC4 = SUMIFS(Subastas!$O:$O, Subastas!$K:$K, $G4,
#                Subastas!$J:$J, $D4, Subastas!$B:$B, AC$3)
#   AI4 = idem sobre Subastas!$P:$P
#   AO4 = idem sobre Subastas!$Q:$Q
#
# Criterios (homologados por NOMBRE contra nuestra hoja Subastas,
# igual que en calcular_l y en los umbrales de E Costos):
#   central       -> Configuración
#   hora del mes  -> Hora_mes
#   tipo          -> Concepto (la columna con las etiquetas
#                    CPF(-)/CSF(+)/etc ya armadas; "Control" -- una
#                    columna DISTINTA -- solo tiene el tipo SIN
#                    direccion, CSF/CTF/CPF, no sirve para esto)
#
# CONFIRMADO (ver BITACORA, sesion de correccion con Libro1.xlsx):
# la formula real del libro (seccion 5.3 del documento de
# trazabilidad) usa Subastas!$B:$B como rango de coincidencia contra
# el encabezado de columna de RE545 (ej. AC$3="CPF(-)") -- y
# Subastas!B es "Concepto" en el archivo real, no "Control". Las
# tres columnas que se suman SI son O, P, Q por posicion (Energía
# SSCC, FD, FMA en el archivo real -- los titulos de grupo de RE545
# dicen Subastas/FD/FMA para esas mismas tres, que es coherente:
# "Subastas" como titulo de grupo generico para la primera, que la
# formula real llama "Energía SSCC").
# ------------------------------------------------------------

# Los 6 encabezados que la formula usa como criterio (AC$3 y sus
# equivalentes). Van en este orden en los tres bloques.
TIPOS_RESERVA_RE545 = (
    "CPF(-)", "CSF(-)", "CTF(-)", "CPF(+)", "CSF(+)", "CTF(+)",
)

# AC:AH, AI:AN, AO:AT -- las tres claves internas de cada bloque.
_BLOQUES_RESERVA_RE545 = (
    ("AC", "AD", "AE", "AF", "AG", "AH"),
    ("AI", "AJ", "AK", "AL", "AM", "AN"),
    ("AO", "AP", "AQ", "AR", "AS", "AT"),
)


def construir_dic_reservas_subastas(df_subastas):
    """
    Arma los tres diccionarios (central, hora del mes, tipo) -> suma,
    uno por cada columna de Subastas que suman los tres bloques de
    RE545 (O, P y Q por posicion; ver el comentario de seccion).

    Un SUMIFS sin coincidencias da 0, asi que el valor por defecto de
    los tres diccionarios es 0, no blanco.
    """

    columnas = list(df_subastas.columns)

    if len(columnas) < 16:
        raise ErrorEntrada(
            f"La hoja Subastas tiene {len(columnas)} columna(s); hacen "
            f"falta al menos 16 (B:Q) para las reservas de Calculo "
            f"RE545."
        )

    # B=0 ... N=12, O=13, P=14, Q=15.
    columnas_suma = (columnas[13], columnas[14], columnas[15])

    claves = [
        (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(hora_mes),
            _normaliza_valor_vba(tipo),
        )
        for central, hora_mes, tipo in zip(
            df_subastas["Configuración"],
            df_subastas["Hora_mes"],
            df_subastas["Concepto"],
        )
    ]

    diccionarios = []

    for columna in columnas_suma:

        valores = pd.to_numeric(df_subastas[columna], errors="coerce")

        acumulado = {}

        for clave, valor in zip(claves, valores):
            if pd.isna(valor):
                continue
            acumulado[clave] = acumulado.get(clave, 0.0) + float(valor)

        diccionarios.append(acumulado)

    return tuple(diccionarios)


def calcular_reservas_re545(df_re545, dics_reservas, registrar=print):
    """
    Replica AC:AT (los tres bloques de 6 columnas) y AU:

        AU = SUMPRODUCT(AC:AH, AI:AN, AO:AT) / 4 * 1000

    o sea, la suma de los 6 productos "reserva x FD x FMA", dividida
    por 4 y por mil.

    TRAMPA REAL (encontrada comparando fila a fila contra "Pagos_BESS
    real" -- planilla 11 -- que el usuario pego en la hoja "RE545 P11":
    a diferencia de los otros 15 valores de los tres bloques, las
    columnas CPF(+)/CSF(+)/CTF(+) del bloque "FMA" (el tercero, AR:AT
    en el archivo real) NO son un SUMIFS -- son la CONSTANTE 1 en TODAS
    las filas. Confirmado contra las formulas guardadas del archivo
    real (docs/Calculo_RE545_reducido_para_IA.xlsx, hoja
    Mapa_Formulas): AO/AP/AQ (CPF(-)/CSF(-)/CTF(-) del mismo bloque)
    son SUMIFS igual que los otros dos bloques, pero AR/AS/AT aparecen
    como valor literal "1", sin formula. Antes de esta correccion se
    les aplicaba el mismo SUMIFS (que da 0 casi siempre), lo que
    tambien arrastraba un error a AU (el SUMPRODUCT las multiplica).
    """

    df = df_re545

    centrales = df["clave"].map(_normaliza_valor_vba)
    horas_mes = df["Hora Mes"].map(_normaliza_valor_vba)

    columnas = {}

    indice_bloque_fma = 2
    posiciones_constante_uno = (3, 4, 5)

    for indice_bloque, (bloque, dic) in enumerate(
        zip(_BLOQUES_RESERVA_RE545, dics_reservas)
    ):

        for posicion, (interno, tipo) in enumerate(
            zip(bloque, TIPOS_RESERVA_RE545)
        ):

            if (
                indice_bloque == indice_bloque_fma
                and posicion in posiciones_constante_uno
            ):
                columnas[interno] = pd.Series(1.0, index=df.index)
                continue

            tipo_normalizado = _normaliza_valor_vba(tipo)

            columnas[interno] = pd.Series(
                [
                    dic.get((central, hora_mes, tipo_normalizado), 0.0)
                    for central, hora_mes in zip(centrales, horas_mes)
                ],
                index=df.index,
            )

    # Un SUMIFS sin coincidencias da 0 y eso es legitimo hora a hora.
    # Lo que no es legitimo es una central que no aparece en NINGUNA
    # clave de Subastas: ahi las 18 reservas quedan en 0, AU sale 0 y
    # la central se paga como si no hubiera tenido reservas.
    centrales_en_subastas = {
        clave[0] for dic in dics_reservas for clave in dic
    }

    sin_subastas = sorted({
        _texto_seguro(central)
        for central, normalizada in zip(df["clave"], centrales)
        if _tiene_valor(central)
        and normalizada not in centrales_en_subastas
    })

    if sin_subastas:
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "SUB-011", ALTA, "Calculo RE545",
                    "Central sin ninguna fila en la hoja Subastas.",
                    central=central,
                    valor_esperado="al menos una fila en Subastas",
                    accion="AC:AT y AU quedan en 0: se paga como si no "
                           "hubiera tenido reservas",
                    hoja="Subastas",
                    origen_control="CONTROL NUEVO",
                )
                for central in sin_subastas
            ],
            f"  [{ALTA}] SUB-011: Calculo RE545: {len(sin_subastas):,} "
            f"central(es) no aparecen en ninguna fila de la hoja Subastas "
            f"({', '.join(repr(v) for v in sin_subastas[:15])}); sus "
            f"reservas (AC:AT) y AU quedan en 0.",
        )

    au = pd.Series(0.0, index=df.index)

    for posicion in range(6):
        au = au + (
            columnas[_BLOQUES_RESERVA_RE545[0][posicion]]
            * columnas[_BLOQUES_RESERVA_RE545[1][posicion]]
            * columnas[_BLOQUES_RESERVA_RE545[2][posicion]]
        )

    columnas["AU"] = au / 4.0 * 1000.0

    return columnas
