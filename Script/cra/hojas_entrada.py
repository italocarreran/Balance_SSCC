# -*- coding: utf-8 -*-
"""
Las hojas de entrada del CRA cuya carga ya esta definida en la
trazabilidad: ENERGIA (6.1), FP (6.4.1), CO (6.4.2) y SC y CO (6.6).

Cada construir_*() recibe la ruta del archivo origen y devuelve el
DataFrame de la hoja, sin escribir nada. Se replica la carga tal cual
esta documentada; lo que la trazabilidad deja PENDIENTE (el Neto de
ENERGIA, CO_Barra_Propia) no se calcula aca.
"""

from pathlib import Path

import pandas as pd

from ..nucleo.utiles import ErrorEntrada
from . import parametros as p
from .lectura import (
    a_fecha_hora,
    dia_excel,
    entero_si_se_puede,
    leer_columnas,
    numero_o_nada,
    texto_excel,
)


def _nada(*_args, **_kwargs):
    pass


# ============================================================
# ENERGIA (traz. 6.1)
# ============================================================

def construir_energia(ruta, anio, mes, registrar=_nada):
    """
    Formato_Solicitud_SSAA_SSCC_Hidro_MesAAAA.xlsx -> hoja ENERGIA.

      B:F -> copia directa (Unidad, Punto de Medida, Año, Mes, DIA)
      HORADIA            = HORA(G)
      PERIODO DE CALCULO = J*4 + (MINUTO(G)+15)/15
      H -> kWhD, I -> kWhR

    Se valida que Año/Mes de todas las filas sean los del periodo: el
    archivo es mensual y el nombre no se usa para buscarlo.
    """

    letras = list(p.COPIA_ENERGIA) + [
        p.LETRA_FECHA_HORA_ENERGIA, p.LETRA_HORA_ENERGIA,
    ]
    origen = leer_columnas(
        ruta, p.HOJA_ORIGEN_ENERGIA, p.FILA_INICIO_ENERGIA, letras
    )

    df = pd.DataFrame(
        {campo: origen[letra] for letra, campo in p.COPIA_ENERGIA.items()},
        dtype=object,
    )

    fechas = origen[p.LETRA_FECHA_HORA_ENERGIA].map(a_fecha_hora)
    horas = origen[p.LETRA_HORA_ENERGIA].map(numero_o_nada)

    df[p.CAMPO_HORADIA] = [
        None if f is None else f.hour for f in fechas
    ]
    df[p.CAMPO_PERIODO] = [
        None if f is None or h is None
        else entero_si_se_puede(h * 4 + (f.minute + 15) / 15)
        for f, h in zip(fechas, horas)
    ]

    sin_fecha = int(fechas.isna().sum())
    sin_hora = int(horas.isna().sum())
    if sin_fecha or sin_hora:
        registrar(
            f"  AVISO ENERGIA: {sin_fecha} fila(s) sin fecha/hora en "
            f"{p.LETRA_FECHA_HORA_ENERGIA} y {sin_hora} sin numero en "
            f"{p.LETRA_HORA_ENERGIA}: quedan sin HORADIA/PERIODO."
        )

    _validar_periodo(df, anio, mes, ruta)

    return df[p.COLUMNAS_ENERGIA].reset_index(drop=True)


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

def construir_sc_co_desde_reporte(ruta, registrar=_nada):
    """
    Reporte_CRA*.xlsx (registros CO), desde la fila 2.
    Clave_Bloque = DIA(C) & "#" & E.
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

    dias = origen[p.LETRA_FECHA_REPORTE_CRA].map(dia_excel)
    df[p.CAMPO_CLAVE_BLOQUE] = [
        _clave_bloque(dia, bloque)
        for dia, bloque in zip(dias, origen[p.LETRA_BLOQUE_REPORTE_CRA])
    ]

    _avisar_sin_dia(dias, "Reporte_CRA", p.LETRA_FECHA_REPORTE_CRA, registrar)

    return df[p.COLUMNAS_SC_CO].reset_index(drop=True)


def construir_sc_co_desde_sobrecostos(ruta, registrar=_nada):
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

    dias = origen[p.LETRA_FECHA_SOBRECOSTOS].map(dia_excel)
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

    return df[p.COLUMNAS_SC_CO].reset_index(drop=True)


def construir_sc_co(ruta_reporte, ruta_sobrecostos, registrar=_nada):
    """
    La tabla SC y CO: primero los registros CO (Reporte_CRA) y abajo
    los SC (Sobrecostos), apilados con el mismo esquema (traz. 6.6).
    """

    co = construir_sc_co_desde_reporte(ruta_reporte, registrar)
    sc = construir_sc_co_desde_sobrecostos(ruta_sobrecostos, registrar)

    registrar(
        f"  SC y CO: {len(co)} registro(s) CO + {len(sc)} registro(s) SC."
    )

    return pd.concat([co, sc], ignore_index=True)


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
