# -*- coding: utf-8 -*-
"""
Las hojas de entrada del CRA cuya carga ya esta definida en la
trazabilidad: FP (6.4.1), CO (6.4.2) y SC y CO (6.6), y ENERGIA (desde
Medidas_CRA.xlsx, ya no del Formato_Solicitud de la traz. 6.1).

Cada construir_*() recibe la ruta del archivo origen y devuelve el
DataFrame de la hoja, sin escribir nada. Se replica la carga tal cual
esta documentada; lo que la trazabilidad deja PENDIENTE
(CO_Barra_Propia) no se calcula aca.
"""

import re
from pathlib import Path

import pandas as pd

from ..nucleo.utiles import ErrorEntrada, normalizar
from . import parametros as p
from .lectura import (
    a_fecha_hora,
    dia_excel,
    leer_columnas,
    numero_o_nada,
    texto_excel,
)


def _nada(*_args, **_kwargs):
    pass


# ============================================================
# ENERGIA: desde Medidas_CRA.xlsx
#
# Antes salia del Formato_Solicitud_SSAA_SSCC_Hidro_MesAAAA.xlsx (traz.
# 6.1). El usuario lo saco (2026-11-26: "Formato solicitud sacalo, ya no
# se necesita. Solo deja Medidas_CRA"): ahora la energia viene de PRMTE,
# en Medidas_CRA.xlsx, con el formato de Medidas_SAE (una fila por clave
# = Unidad y cuarto de hora, Gen_Unidad ya con el signo del Flujo).
#
#   Unidad Generadora/Central = clave
#   Año / Mes / DIA           = de 'intervalo' (hora local)
#   HORADIA                   = Hora - 1      (la HORA() de Excel, 0..23)
#   PERIODO DE CALCULO        = (Hora - 1) * 4 + Minutos / 15 + 1
#                               (= J*4 + (MINUTO+15)/15 de la traz.;
#                               llega a 100 el dia que se atrasa la hora)
#   Cuarto de Hora            = el del mes (cruza con el resto del balance)
#   Neto                      = Gen_Unidad
#
# El Formato_Solicitud traia kWhD y kWhR por separado y el libro armaba
# el Neto con dos formulas (K, L, M); Medidas_CRA ya trae el neto, asi
# que esas columnas no estan.
# ============================================================

def construir_energia(ruta, anio, mes, registrar=_nada):
    """Medidas_CRA.xlsx -> hoja ENERGIA (ver arriba)."""

    try:
        origen = pd.read_excel(ruta, sheet_name=p.HOJA_MEDIDAS_CRA)
    except ValueError as error:
        raise ErrorEntrada(
            f"{Path(ruta).name} no tiene la hoja '{p.HOJA_MEDIDAS_CRA}': "
            f"generalo de nuevo con 'Generar'."
        ) from error

    faltan = [c for c in p.COLUMNAS_MEDIDAS_ORIGEN if c not in origen.columns]
    if faltan:
        raise ErrorEntrada(
            f"{Path(ruta).name} no trae las columnas {faltan}: generalo de "
            f"nuevo con 'Generar'."
        )

    fechas = pd.to_datetime(origen["intervalo"], errors="coerce")
    hora = pd.to_numeric(origen["Hora"], errors="coerce")
    minutos = pd.to_numeric(origen["Minutos"], errors="coerce")

    df = pd.DataFrame({
        p.CAMPO_UNIDAD: origen["clave"].astype(str),
        p.CAMPO_ANIO: fechas.dt.year,
        p.CAMPO_MES: fechas.dt.month,
        p.CAMPO_DIA: fechas.dt.day,
        p.CAMPO_HORADIA: hora - 1,
        p.CAMPO_PERIODO: (hora - 1) * 4 + minutos / 15 + 1,
        p.CAMPO_CUARTO: pd.to_numeric(origen["Cuarto de Hora"], errors="coerce"),
        p.CAMPO_NETO: pd.to_numeric(origen["Gen_Unidad"], errors="coerce"),
    })

    sin_fecha = int(fechas.isna().sum())
    if sin_fecha:
        registrar(f"  AVISO ENERGIA: {sin_fecha} fila(s) sin fecha en 'intervalo'.")
    for campo in (p.CAMPO_HORADIA, p.CAMPO_PERIODO, p.CAMPO_CUARTO):
        df[campo] = df[campo].astype("Int64")

    _validar_periodo(df.dropna(subset=[p.CAMPO_ANIO]), anio, mes, ruta)

    return (
        df.sort_values([p.CAMPO_UNIDAD, p.CAMPO_CUARTO])
        [p.COLUMNAS_ENERGIA].reset_index(drop=True)
    )


def _validar_periodo(df, anio, mes, ruta):
    anios = df[p.CAMPO_ANIO].map(numero_o_nada)
    meses = df[p.CAMPO_MES].map(numero_o_nada)
    distintas = (anios != anio) | (meses != mes)

    if distintas.any():
        ejemplos = (
            df.loc[distintas, [p.CAMPO_ANIO, p.CAMPO_MES]]
            .astype(str).drop_duplicates().head(5)
            .apply(lambda f: f"{f.iloc[0]}/{f.iloc[1]}", axis=1)
        )
        raise ErrorEntrada(
            f"{Path(ruta).name}: {int(distintas.sum())} fila(s) con Año/Mes "
            f"distinto del periodo {anio}/{mes} (por ejemplo: "
            f"{', '.join(ejemplos)}). Revisar que sea el archivo del mes."
        )


# ============================================================
# FP y CO (traz. 6.4)
# ============================================================

def _copia_directa(ruta, hoja, fila_inicio, copia):
    origen = leer_columnas(ruta, hoja, fila_inicio, list(copia))

    return pd.DataFrame(
        {campo: origen[letra] for letra, campo in copia.items()},
        dtype=object,
    ).reset_index(drop=True)


def construir_fp(ruta, registrar=_nada):
    """fp_*.xlsx, bloque A2:D -> BarNom, Hora, FP, dia."""

    return _copia_directa(
        ruta, p.HOJA_ORIGEN_FP, p.FILA_INICIO_FP, p.COPIA_FP
    )


def construir_co(ruta, registrar=_nada):
    """
    cvar_cra_AAMM_*.xlsx, bloque A2:D -> nombre_configuración, Día
    (de D), Hora (de B), Costos_Operación (de C).
    """

    return _copia_directa(
        ruta, p.HOJA_ORIGEN_CO, p.FILA_INICIO_CO, p.COPIA_CO
    )


# ============================================================
# SC y CO (traz. 6.6)
# ============================================================

def construir_sc_co_desde_reporte(ruta, registrar=_nada, periodo=None):
    """
    Reporte_CRA*.csv (registros CO), desde la fila 2.
    Clave_Bloque = DIA(C) & "#" & E.

    periodo: (anio, mes) para validar que todas las fechas sean del mes.
    """

    letras = list(p.COPIA_REPORTE_CRA) + [
        p.LETRA_FECHA_REPORTE_CRA, p.LETRA_BLOQUE_REPORTE_CRA,
    ]
    origen = leer_columnas(
        ruta, p.HOJA_ORIGEN_REPORTE_CRA, p.FILA_INICIO_REPORTE_CRA, letras
    )

    df = pd.DataFrame(
        {campo: origen[letra] for letra, campo in p.COPIA_REPORTE_CRA.items()},
        dtype=object,
    )

    fechas = origen[p.LETRA_FECHA_REPORTE_CRA]
    _validar_mes(fechas, periodo, "Reporte_CRA",
                     p.LETRA_FECHA_REPORTE_CRA, registrar)
    dias = fechas.map(dia_excel)
    df[p.CAMPO_CLAVE_BLOQUE] = [
        _clave_bloque(dia, bloque)
        for dia, bloque in zip(dias, origen[p.LETRA_BLOQUE_REPORTE_CRA])
    ]

    _avisar_sin_dia(dias, "Reporte_CRA", p.LETRA_FECHA_REPORTE_CRA, registrar)

    return filtrar_embalses(df[p.COLUMNAS_SC_CO], "Reporte_CRA", registrar)


def construir_sc_co_desde_sobrecostos(ruta, registrar=_nada, periodo=None):
    """
    Cálculo_SobrecostosSSCC_*.xlsm (registros SC), desde la fila 7.

    Cada CPF/CSF/CTF(+/-) es la suma de tres columnas (vacio cuenta 0,
    como la suma de Excel).
    Clave_Bloque = DIA(A) & "#" & (R - (DIA(A)-1)*96).
    """

    columnas_suma = [
        letra for letras in p.SUMAS_SOBRECOSTOS.values() for letra in letras
    ]
    letras = list(p.COPIA_SOBRECOSTOS) + columnas_suma + [
        p.LETRA_FECHA_SOBRECOSTOS, p.LETRA_PERIODO_SOBRECOSTOS,
    ]
    origen = leer_columnas(
        ruta, p.HOJA_ORIGEN_SOBRECOSTOS, p.FILA_INICIO_SOBRECOSTOS, letras
    )

    df = pd.DataFrame(
        {campo: origen[letra] for letra, campo in p.COPIA_SOBRECOSTOS.items()},
        dtype=object,
    )

    no_numericos = 0
    for campo, letras_suma in p.SUMAS_SOBRECOSTOS.items():
        total = pd.Series(0.0, index=origen.index)
        for letra in letras_suma:
            valores = origen[letra]
            numeros = valores.map(numero_o_nada)
            no_numericos += int(
                (numeros.isna() & valores.map(texto_excel).ne("")).sum()
            )
            total = total + numeros.fillna(0.0)
        df[campo] = total

    if no_numericos:
        registrar(
            f"  AVISO SC y CO: {no_numericos} celda(s) con texto no "
            f"numerico en las columnas que se suman del archivo de "
            f"Sobrecostos (en Excel darian #¡VALOR!); se contaron como 0."
        )

    fechas = origen[p.LETRA_FECHA_SOBRECOSTOS]
    _validar_mes(fechas, periodo, "Sobrecostos",
                     p.LETRA_FECHA_SOBRECOSTOS, registrar)
    dias = fechas.map(dia_excel)
    periodos = origen[p.LETRA_PERIODO_SOBRECOSTOS].map(numero_o_nada)
    df[p.CAMPO_CLAVE_BLOQUE] = [
        None if dia is None or periodo is None
        else _clave_bloque(
            dia,
            periodo - (dia - 1) * p.PERIODOS_POR_DIA_FORMULA_SC,
        )
        for dia, periodo in zip(dias, periodos)
    ]

    _avisar_sin_dia(
        dias, "Sobrecostos", p.LETRA_FECHA_SOBRECOSTOS, registrar
    )

    return filtrar_embalses(df[p.COLUMNAS_SC_CO], "Sobrecostos", registrar)


def construir_sc_co(ruta_reporte, ruta_sobrecostos, registrar=_nada,
                    periodo=None):
    """
    La tabla SC y CO: primero los registros SC (Sobrecostos) y abajo
    los CO (Reporte_CRA), apilados con el mismo esquema (traz. 6.6), solo
    de las centrales de embalse. El orden SC arriba / CO abajo es el de la
    hoja real (Actualiza_SC_CO.py del usuario).
    """

    sc = construir_sc_co_desde_sobrecostos(ruta_sobrecostos, registrar, periodo)
    co = construir_sc_co_desde_reporte(ruta_reporte, registrar, periodo)

    registrar(
        f"  SC y CO: {len(sc)} registro(s) SC + {len(co)} registro(s) CO."
    )

    tabla = pd.concat([sc, co], ignore_index=True)
    _avisar_duplicados(tabla, registrar)

    return tabla


def _avisar_duplicados(tabla, registrar):
    """
    Control del libro (SC y CO!AI, traz. v5 6.6): la misma Unidad con la
    misma Clave_Bloque en dos filas. Se avisa, no se corrige: CÁLCULO_CRA
    no esta todavia y no se sabe cual de las dos usaria.
    """

    claves = [p.CAMPO_TIPO, p.CAMPO_UNIDAD_SC, p.CAMPO_CLAVE_BLOQUE]
    repetidas = tabla[tabla.duplicated(claves, keep=False)]

    if not repetidas.empty:
        ejemplos = (
            repetidas[claves].astype(str).drop_duplicates().head(5)
            .apply(lambda f: " ".join(f), axis=1)
        )
        registrar(
            f"  AVISO SC y CO: {len(repetidas)} fila(s) con Tipo + Unidad + "
            f"Clave_Bloque repetidos (por ejemplo: {', '.join(ejemplos)})."
        )


# Una central que termina en "-numero" es una unidad: si no esta en la
# lista de embalses, puede ser una unidad de embalse nueva.
_RE_UNIDAD = re.compile(r"-\d+\s*$")


def filtrar_embalses(df, origen, registrar=_nada):
    """
    Solo las filas de centrales de embalse (p.CENTRALES_EMBALSE). Avisa
    que centrales de la lista no tienen ninguna fila y cuales terminan
    en "-numero" sin estar en la lista.
    """

    permitidas = {normalizar(c) for c in p.CENTRALES_EMBALSE}
    nombres = df[p.CAMPO_UNIDAD_SC].map(texto_excel)
    claves = nombres.map(normalizar)
    quedan = claves.isin(permitidas)

    descartadas = sorted(set(nombres[~quedan & nombres.ne("")]))
    registrar(
        f"  {origen}: {int(quedan.sum())} fila(s) de embalse de "
        f"{len(df)}; se descartan {len(descartadas)} central(es) que no "
        f"son embalse"
        + (f" ({', '.join(descartadas[:8])}"
           + (", ..." if len(descartadas) > 8 else "") + ")"
           if descartadas else "")
        + "."
    )

    vistas = set(claves[quedan])
    faltan = [c for c in p.CENTRALES_EMBALSE if normalizar(c) not in vistas]
    if faltan:
        registrar(
            f"  AVISO {origen}: sin filas para {len(faltan)} central(es) de "
            f"embalse: {', '.join(faltan)}."
        )

    sospechosas = [n for n in descartadas if _RE_UNIDAD.search(n)]
    if sospechosas:
        registrar(
            f"  AVISO {origen}: terminan en '-numero' y NO estan en la lista "
            f"de embalses (¿unidad nueva?): {', '.join(sospechosas)}. Si lo "
            f"son, agregarlas a CENTRALES_EMBALSE en Script/cra/parametros.py."
        )

    return df[quedan].reset_index(drop=True)


def participacion_por_servicio(sc_co):
    """
    CPF = CPF(+) + CPF(-), CSF = ..., CTF = ... (traz. 6.6.1). En
    memoria: estas tres no se escriben en la hoja.
    """

    salida = pd.DataFrame(index=sc_co.index)

    for servicio, (mas, menos) in p.SERVICIOS_SC_CO.items():
        salida[servicio] = (
            sc_co[mas].map(numero_o_nada).fillna(0.0)
            + sc_co[menos].map(numero_o_nada).fillna(0.0)
        )

    return salida


def _clave_bloque(dia, bloque):
    if dia is None:
        return None

    return f"{dia}#{texto_excel(bloque)}"


def _avisar_sin_dia(dias, origen, letra, registrar):
    sin_dia = int(dias.isna().sum())

    if sin_dia:
        registrar(
            f"  AVISO SC y CO: {sin_dia} fila(s) de {origen} sin fecha "
            f"en la columna {letra}: quedan sin Clave_Bloque."
        )


def _validar_mes(valores, periodo, origen, letra, registrar):
    """
    Cada archivo de SC y CO trae un solo mes (confirmado por el usuario):
    una fecha de otro mes quiere decir que el archivo no es el del
    periodo (o que una fecha de texto se leyo con dia y mes cruzados).
    """

    if periodo is None:
        return

    fechas = valores.map(a_fecha_hora).dropna()
    otras = [f for f in fechas if (f.year, f.month) != tuple(periodo)]

    if otras:
        meses = sorted({f"{f.year}-{f.month:02d}" for f in otras})
        raise ErrorEntrada(
            f"{origen}: {len(otras)} fila(s) con fecha (columna {letra}) "
            f"fuera del periodo {periodo[0]}-{periodo[1]:02d} "
            f"({', '.join(meses[:5])}). Revisar que sea el archivo del mes."
        )
