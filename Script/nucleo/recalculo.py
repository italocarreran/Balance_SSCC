# -*- coding: utf-8 -*-
"""
Recalcular las hojas de calculo de un 11_PAGOS_BESS_AAMM.xlsx con el
archivo como UNICO dato.

Pedido del usuario (2026-11-10): junto con el 11_ se entrega a los
coordinados un codigo que "al darle play se recalculen las hojas de
calculo". Es el mismo calculo del programa (generar_pagos_bess y
generar_csv_sae), corrido sobre una COPIA del archivo:

  - el 11_ entregado no se toca: el resultado queda al lado, como
    11_PAGOS_BESS_AAMM_calculado.xlsx (la ventana de los coordinados no
    genera el SAE_AAMM.csv: es un reporte interno);
  - todo lo que el calculo necesita sale del propio 11_: las hojas de
    entrada (Centrales, Medidores, Ofertas SSCC, CMg, FD, FMA, Subastas)
    y, para PRORRATA_RETIROS y el Resumen, la prorrata de retiros, que
    se toma del cuadro 2 de la hoja PRORRATA_RETIROS (cuarto,
    suministrador, prorrata) porque el Excel de la prorrata no va en el
    entregable;
  - lo que el programa deja al lado (control de la corrida, sellos de
    las hojas) queda en una carpeta temporal y se borra.
"""

import os
import re
import shutil
import tempfile
from pathlib import Path

import pandas as pd

from .lector import abrir_excel, leer_excel
from .parametros import (
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CENTRALES,
    HOJA_COMPENSACION_CENTRAL, HOJA_PRORRATA_RETIROS, HOJA_RESUMEN,
)
from .prorrata_retiros import (
    COL_CUARTO, COL_PRORRATA, COL_SUMINISTRADOR, HOJA_ORIGEN,
)
from .rutas import nombre_csv_sae, nombre_salida
from .utiles import ErrorEntrada

# Lo que se puede recalcular, en el orden del calculo: (id de seccion de
# generar_pagos_bess, hoja). El CSV va aparte (con_csv).
SECCIONES_RECALCULO = (
    ("ecostos", HOJA_CALCULO_ECOSTOS),
    ("re545", HOJA_CALCULO_RE545),
    ("compensacion_central", HOJA_COMPENSACION_CENTRAL),
    ("prorrata_retiros", HOJA_PRORRATA_RETIROS),
    ("resumen", HOJA_RESUMEN),
)

# La copia se llama ..._calculado (pedido del usuario: "Calcular", no
# "Recalcular").
SUFIJO_RECALCULADO = "_calculado"
_PATRON_11 = re.compile(r"^11_PAGOS_BESS_(\d{4})\.xlsx$", re.I)


def aamm_del_archivo(ruta):
    """'11_PAGOS_BESS_2609.xlsx' -> '2609'."""

    hallado = _PATRON_11.match(Path(ruta).name)
    if not hallado:
        raise ErrorEntrada(
            f"{Path(ruta).name} no se llama 11_PAGOS_BESS_AAMM.xlsx: no se "
            f"sabe de que periodo es."
        )
    return hallado.group(1)


def buscar_11(carpetas):
    """El unico 11_PAGOS_BESS_AAMM.xlsx de la primera carpeta que tenga uno."""

    for carpeta in carpetas:
        try:
            hallados = sorted(
                p for p in Path(carpeta).iterdir()
                if p.is_file() and _PATRON_11.match(p.name)
            )
        except OSError:
            continue
        if len(hallados) == 1:
            return hallados[0]
    return None


def prorrata_desde_hoja(ruta_11):
    """
    La prorrata de retiros (cuarto, suministrador, prorrata) del cuadro 2
    de la hoja PRORRATA_RETIROS, con los nombres del Excel de origen.
    """

    try:
        crudo = leer_excel(ruta_11, sheet_name=HOJA_PRORRATA_RETIROS, header=None)
    except ValueError as error:
        raise ErrorEntrada(
            f"{Path(ruta_11).name} no tiene la hoja '{HOJA_PRORRATA_RETIROS}': "
            f"de ahi sale la prorrata de retiros."
        ) from error

    for fila in range(min(len(crudo), 10)):
        valores = [str(v).strip() if pd.notna(v) else "" for v in crudo.iloc[fila]]
        for col in range(len(valores) - 2):
            if (
                valores[col] == COL_CUARTO
                and valores[col + 1] == COL_SUMINISTRADOR
                and valores[col + 2].startswith(COL_PRORRATA)
            ):
                tabla = crudo.iloc[fila + 1:, col:col + 3]
                tabla.columns = [COL_CUARTO, COL_SUMINISTRADOR, COL_PRORRATA]
                tabla = tabla[pd.to_numeric(tabla[COL_CUARTO], errors="coerce").notna()]
                if tabla.empty:
                    break
                return tabla.reset_index(drop=True)
    raise ErrorEntrada(
        f"La hoja '{HOJA_PRORRATA_RETIROS}' de {Path(ruta_11).name} no trae "
        f"el cuadro de la prorrata (Cuarto de Hora, Suministrador, Prorrata)."
    )


def recalcular(ruta_11, secciones, con_csv=True, registrar=print):
    """
    Recalcula las secciones pedidas (ids de SECCIONES_RECALCULO) sobre una
    copia del 11_ y, si con_csv, el SAE_AAMM.csv. Devuelve
    (ruta de la copia recalculada o None, ruta del CSV o None).
    """

    from .csv_sae import generar_csv_sae
    from .proceso import generar_pagos_bess

    ruta_11 = Path(ruta_11)
    if not ruta_11.is_file():
        raise ErrorEntrada(f"No se encontro {ruta_11}")
    aamm = aamm_del_archivo(ruta_11)
    secciones = [s for s, _ in SECCIONES_RECALCULO if s in set(secciones)]
    if not secciones and not con_csv:
        raise ErrorEntrada("No hay nada seleccionado para recalcular.")

    with abrir_excel(ruta_11) as libro:
        hojas = list(libro.sheet_names)
    if HOJA_CENTRALES not in hojas:
        raise ErrorEntrada(
            f"{ruta_11.name} no trae la hoja '{HOJA_CENTRALES}', que es de "
            f"donde el calculo saca las centrales."
        )

    with tempfile.TemporaryDirectory(prefix="recalculo_11_") as temporal:
        temporal = Path(temporal)
        # Lo que el programa guarda aparte (sellos) queda en la temporal.
        os.environ.setdefault("BESS_CARPETA_CONFIG", str(temporal / "config"))
        caso = temporal / "caso"
        caso.mkdir()
        copia_caso = caso / nombre_salida(aamm)
        registrar(f"Copiando {ruta_11.name} (el original no se modifica)...")
        shutil.copyfile(ruta_11, copia_caso)

        if {"prorrata_retiros", "resumen"} & set(secciones):
            registrar(f"Tomando la prorrata de retiros de la hoja '{HOJA_PRORRATA_RETIROS}'...")
            prorrata = prorrata_desde_hoja(ruta_11)
            carpeta = caso / "Prorrata retiros"
            carpeta.mkdir()
            prorrata.to_excel(
                carpeta / f"Prorrata_Retiros_{aamm}_def.xlsx",
                sheet_name=HOJA_ORIGEN, index=False,
            )

        salida = None
        if secciones:
            generar_pagos_bess(
                caso, set(secciones), registrar=registrar, aamm=aamm
            )
            salida = ruta_11.with_name(f"{ruta_11.stem}{SUFIJO_RECALCULADO}.xlsx")
            shutil.copyfile(copia_caso, salida)
            registrar(f"Listo: {salida}")

        csv = None
        if con_csv:
            generado = generar_csv_sae(caso, aamm, registrar=registrar)
            csv = ruta_11.with_name(
                f"{Path(nombre_csv_sae(aamm)).stem}{SUFIJO_RECALCULADO}.csv"
            )
            shutil.copyfile(generado, csv)
            registrar(f"Listo: {csv}")

    return salida, csv
