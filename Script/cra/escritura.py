# -*- coding: utf-8 -*-
"""
Escritura de la planilla del CRA (5_REMUNERACION_CRA_AAMM.xlsx): se
reemplazan solo las hojas que se regeneran y el resto del libro se
conserva tal cual estaba (mismo criterio que la planilla del BESS).

Con las mismas convenciones de la planilla del BESS
(Script/nucleo/escritura.py), para que las dos se lean igual:

  - fila 1 vacia, "Cuadro N° 1 — <titulo>" en la fila 2, nombres de
    columna en la 3 y datos desde la 4;
  - la columna A vacia, de margen (el cuadro empieza en la B);
  - el formato de nucleo.formato (encabezado azul, marco, anchos,
    separador de miles, panel inmovilizado);
  - escritura rapida: las hojas grandes (FD_CTF pasa de 300 mil filas)
    se escriben con openpyxl solo hasta FILAS_MUESTRA y el resto de las
    filas las agrega relleno_rapido directo en el XML; las hojas que no
    se tocan quedan en el zip tal cual (xlsx_rapido.pegar_hojas).
"""

import time
from pathlib import Path

import pandas as pd

from ..nucleo.escritura import (
    COLUMNA_MARGEN, _escribir_tabla_con_titulo, _rellenos_de,
)
from ..nucleo.formato import formatear_hoja
from ..nucleo.lectura import titulo_de_cuadro
from ..nucleo.utiles import ErrorEntrada
from ..nucleo.xlsx_rapido import pegar_hojas
from .parametros import ORDEN_HOJAS_SALIDA, TITULOS_HOJAS

# Filas (1-indexadas) de los titulos de cuadro y de los nombres de
# columna en todas las hojas del CRA.
FILA_TITULO_CUADRO = 2
FILA_ENCABEZADO = 3


def escribir_hojas(ruta_salida, hojas, registrar=None):
    """
    hojas: {nombre de hoja: DataFrame}. Reemplaza esas hojas en el
    libro (lo crea si no existe), deja las demas como estaban y ordena
    segun ORDEN_HOJAS_SALIDA (las que no esten ahi van al final).
    """

    ruta_salida = Path(ruta_salida)
    temporal = ruta_salida.with_name(f".~nuevas_{ruta_salida.name}")
    inicio = time.time()

    try:
        with pd.ExcelWriter(temporal, engine="openpyxl") as writer:
            for nombre, df in hojas.items():
                _escribir_hoja(writer, nombre, df)
                formatear_hoja(
                    writer.sheets[nombre], FILA_ENCABEZADO, FILA_TITULO_CUADRO
                )
            rellenos = _rellenos_de(writer.book)

        pegar_hojas(
            ruta_salida, temporal, ORDEN_HOJAS_SALIDA, rellenos=rellenos
        )
    except PermissionError as error:
        raise ErrorEntrada(
            f"No se pudo guardar {ruta_salida.name}: esta abierto en "
            f"Excel. Cerralo y volve a intentar."
        ) from error
    finally:
        temporal.unlink(missing_ok=True)

    if registrar:
        registrar(
            f"  escritura: {time.time() - inicio:.1f} s".replace(".", ",")
        )


def _escribir_hoja(writer, nombre, df):
    """Fila 1 vacia, titulo de cuadro en la 2, la tabla desde la 3, en la B."""

    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    titulo = TITULOS_HOJAS.get(nombre, nombre)

    _escribir_tabla_con_titulo(
        writer, nombre, df, titulo_de_cuadro(1, titulo),
        fila_inicio=FILA_TITULO_CUADRO - 1, columna_inicio=COLUMNA_MARGEN,
    )
