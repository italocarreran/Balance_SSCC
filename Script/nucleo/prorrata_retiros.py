# -*- coding: utf-8 -*-
"""Reparto de la compensacion entre los retiros, cuarto de hora a cuarto de hora.

La idea, en una linea: cada cuarto de hora del mes tiene un monto a
compensar (lo que suman "Calculo E Costos" y "Calculo RE545" en ese
mismo cuarto) y ese monto se reparte entre las empresas que retiraron
dentro de ese cuarto, segun el peso que trae la prorrata oficial.

    Monto a pagar(cuarto, empresa) =
        Monto a compensar(cuarto) * Prorrata(cuarto, empresa)
                                  / suma de prorratas de ese cuarto

Dividir por la suma del cuarto es lo que hace que el reparto sea
"segun peso" y no "segun el numero escrito": si la prorrata ya viene
normalizada (suma 1) la division no cambia nada, y si no lo viene el
cuarto igual se reparte entero. Antes esto se validaba a la entrada y
la corrida se caia (RET-001/004/005/006) en vez de repartir; ahora se
reparte y lo raro se informa por el log.
"""

from pathlib import Path

import numpy as np
import pandas as pd

from .utiles import ErrorEntrada, normalizar


HOJA_ORIGEN = "Prorrata 15min"
# La fuente se lee por POSICION (A, B, C), no por nombre: el archivo
# del CEN cambia el texto del encabezado de un periodo a otro
# ("Cuarto de Hora" / "Cuarto hora" / "CUARTO DE HORA"), pero el orden
# de las tres columnas es siempre el mismo.
COLUMNAS_ORIGEN = ("Cuarto de Hora", "Suministrador", "Prorrata")

COL_CUARTO = "Cuarto de Hora"
COL_SUMINISTRADOR = "Suministrador"
COL_PRORRATA = "Prorrata"
COL_MONTO_CUARTO = "Monto a compensar [$]"
COL_PAGO = "Monto a pagar [$]"
COL_TOTAL_PAGO = "Total a pagar [$]"


def buscar_archivo_prorrata(carpeta, aamm=None):
    """Encuentra un unico Prorrata_Retiros_AAMM_pre/def en su carpeta."""

    carpeta = Path(carpeta)
    if not carpeta.is_dir():
        return None
    candidatos = []
    for archivo in carpeta.iterdir():
        nombre = normalizar(archivo.stem).replace(" ", "_")
        if (
            archivo.is_file()
            and not archivo.name.startswith("~$")
            and archivo.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}
            and nombre.startswith("prorrata_retiros_")
            and (nombre.endswith("_pre") or nombre.endswith("_def"))
            and (not aamm or f"_{aamm}_" in nombre)
        ):
            candidatos.append(archivo)
    if len(candidatos) > 1:
        nombres = "\n".join(f"- {p.name}" for p in sorted(candidatos))
        raise ErrorEntrada(
            "Hay mas de un archivo de prorrata para el periodo. Deja solo "
            f"el que corresponda en {carpeta}:\n{nombres}"
        )
    return candidatos[0] if candidatos else None


def leer_prorrata_retiros(ruta, registrar=None):
    """Lee la hoja ``Prorrata 15min``: A cuarto, B suministrador, C prorrata.

    Lo unico que sigue siendo un error de entrada es que el archivo no
    se pueda abrir, que le falte la hoja o que no quede ninguna fila
    utilizable. Todo lo demas (prorratas negativas, cuartos que no
    suman 1, filas repetidas) se informa y se sigue: la validacion la
    hace el usuario mirando la hoja, no el script cortando la corrida.
    """

    registrar = registrar or (lambda _mensaje: None)
    ruta = Path(ruta)
    try:
        with pd.ExcelFile(ruta) as excel:
            hojas = list(excel.sheet_names)
    except Exception as error:
        raise ErrorEntrada(f"No se pudo abrir la prorrata {ruta}: {error}") from error
    if HOJA_ORIGEN not in hojas:
        raise ErrorEntrada(
            f"{ruta.name} no tiene la hoja obligatoria '{HOJA_ORIGEN}'."
        )
    # header=0: la primera fila es encabezado. usecols/names: se toman
    # las tres primeras columnas por posicion y se les pone el nombre
    # interno, sin depender del texto del encabezado real.
    df = pd.read_excel(
        ruta, sheet_name=HOJA_ORIGEN, header=0, usecols=[0, 1, 2],
        names=list(COLUMNAS_ORIGEN),
    )

    cuarto = pd.to_numeric(df[COL_CUARTO], errors="coerce")
    prorrata = pd.to_numeric(df[COL_PRORRATA], errors="coerce")
    suministrador = df[COL_SUMINISTRADOR].fillna("").astype(str).str.strip()

    # Fila utilizable = tiene cuarto y tiene empresa. Una prorrata
    # vacia se lee como 0 (esa empresa no retiro en ese cuarto).
    utilizable = cuarto.notna() & suministrador.ne("")
    descartadas = int((~utilizable).sum())
    if descartadas:
        registrar(
            f"  Prorrata: {descartadas:,} fila(s) sin cuarto de hora o sin "
            "suministrador; se ignoran."
        )
    df = pd.DataFrame({
        COL_CUARTO: cuarto[utilizable].astype(int),
        COL_SUMINISTRADOR: suministrador[utilizable],
        COL_PRORRATA: prorrata[utilizable].fillna(0.0).astype(float),
    }).reset_index(drop=True)
    if df.empty:
        raise ErrorEntrada(
            f"{ruta.name}/{HOJA_ORIGEN} no tiene ninguna fila utilizable "
            "(se esperan A: cuarto de hora, B: suministrador, C: prorrata)."
        )

    # Una empresa repetida dentro del mismo cuarto no es un error: se
    # suman sus pesos, que es lo que significan dos retiros en el mismo
    # cuarto de hora.
    repetidas = df.duplicated([COL_CUARTO, COL_SUMINISTRADOR]).sum()
    if repetidas:
        registrar(
            f"  Prorrata: {int(repetidas):,} fila(s) con el mismo cuarto y "
            "suministrador; se suman sus pesos."
        )
        df = (
            df.groupby([COL_CUARTO, COL_SUMINISTRADOR], as_index=False)[COL_PRORRATA]
            .sum()
        )

    negativas = int(df[COL_PRORRATA].lt(0).sum())
    if negativas:
        registrar(
            f"  Prorrata: {negativas:,} peso(s) negativo(s); se reparten tal "
            "como vienen."
        )
    sumas = df.groupby(COL_CUARTO)[COL_PRORRATA].sum()
    distintas = sumas[~np.isclose(sumas, 1.0, atol=1e-6, rtol=0)]
    if not distintas.empty:
        registrar(
            f"  Prorrata: {len(distintas):,} cuarto(s) cuyos pesos no suman 1; "
            "se reparte cada cuarto segun el peso relativo."
        )
    return df.sort_values(
        [COL_CUARTO, COL_SUMINISTRADOR], kind="stable"
    ).reset_index(drop=True)


def _compensacion_por_cuarto(df_ecostos, df_re545):
    """Suma el ``Monto a compensar`` de las dos hojas por cuarto de hora.

    Las dos hojas traen el cuarto de hora cronologico del mes en
    ``Bloque horario`` (es "Cuarto de Hora" renombrado). E Costos
    tambien tiene ``Bloque Mes Descarga``, pero ese es el orden del
    bloque dentro de la curva monotona, NO el cuarto del mes: usarlo
    para repartir -como se hacia antes- mezclaba dinero de cuartos que
    no tenian nada que ver entre si.
    """

    def monto_por_cuarto(df):
        if df is None or df.empty:
            return pd.Series(dtype=float)
        columna = "Bloque horario" if "Bloque horario" in df.columns else None
        if columna is None:
            raise ErrorEntrada(
                "La hoja de calculo no trae la columna 'Bloque horario' "
                "(cuarto de hora del mes) para repartir la compensacion."
            )
        q = pd.to_numeric(df[columna], errors="coerce")
        monto = pd.to_numeric(df["Monto a compensar"], errors="coerce").fillna(0.0)
        valido = q.notna()
        return monto[valido].groupby(q[valido].astype(int)).sum()

    ec = monto_por_cuarto(df_ecostos)
    re = monto_por_cuarto(df_re545)
    return ec.add(re, fill_value=0.0).sort_index()


def construir_prorrata_retiros(df_prorrata, df_ecostos, df_re545, registrar=None):
    """Arma los tres cuadros de la hoja PRORRATA_RETIROS.

    Devuelve, en este orden:

    1. ``por_cuarto``: cuarto de hora y monto a compensar (dos columnas).
    2. ``detalle``: la prorrata leida (cuarto, suministrador, peso) con
       el monto que le toca pagar a esa empresa en ese cuarto.
    3. ``pagos``: el total a pagar de cada empresa en el mes.
    """

    registrar = registrar or (lambda _mensaje: None)

    compensacion = _compensacion_por_cuarto(df_ecostos, df_re545)
    por_cuarto = pd.DataFrame({
        COL_CUARTO: compensacion.index.astype(int),
        COL_MONTO_CUARTO: compensacion.to_numpy(dtype=float),
    }).reset_index(drop=True)

    detalle = df_prorrata.copy()
    detalle[COL_MONTO_CUARTO] = (
        detalle[COL_CUARTO].map(compensacion).astype(float).fillna(0.0)
    )
    # El peso relativo dentro del cuarto: asi el cuarto se reparte
    # entero aunque la columna C no venga normalizada.
    suma_cuarto = detalle.groupby(COL_CUARTO)[COL_PRORRATA].transform("sum")
    peso = np.where(
        np.isclose(suma_cuarto, 0.0),
        0.0,
        detalle[COL_PRORRATA] / suma_cuarto.replace(0.0, np.nan),
    )
    detalle[COL_PAGO] = detalle[COL_MONTO_CUARTO] * np.nan_to_num(peso)
    detalle = detalle.loc[
        :, [COL_CUARTO, COL_SUMINISTRADOR, COL_PRORRATA, COL_PAGO]
    ].reset_index(drop=True)

    # Plata que no se pudo repartir: cuartos con monto pero sin
    # ninguna fila de prorrata (o con pesos que suman 0).
    cuartos_prorrata = set(detalle[COL_CUARTO])
    sin_prorrata = [
        int(q) for q, monto in compensacion.items()
        if not np.isclose(monto, 0.0) and q not in cuartos_prorrata
    ]
    if sin_prorrata:
        registrar(
            f"  {len(sin_prorrata):,} cuarto(s) con compensacion no tienen "
            f"prorrata y quedan sin repartir: {sin_prorrata[:10]}"
        )
    repartido = float(detalle[COL_PAGO].sum())
    total = float(compensacion.sum())
    if not np.isclose(repartido, total, atol=1e-4, rtol=1e-9):
        registrar(
            f"  Compensacion del mes: ${total:,.0f}; repartida: "
            f"${repartido:,.0f} (diferencia ${total - repartido:,.0f})."
        )

    pagos = (
        detalle.groupby(COL_SUMINISTRADOR, as_index=False)[COL_PAGO].sum()
        .rename(columns={COL_PAGO: COL_TOTAL_PAGO})
        .sort_values(COL_SUMINISTRADOR, kind="stable").reset_index(drop=True)
    )
    return por_cuarto, detalle, pagos
