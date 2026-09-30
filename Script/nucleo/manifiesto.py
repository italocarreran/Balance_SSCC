# -*- coding: utf-8 -*-
"""
Manifiesto de entradas: que archivo exacto alimento esta corrida.

Sin esto, "APROBADA" no es reproducible: dentro de tres meses nadie
puede demostrar CUAL SOC_2607.xlsx se uso. Para un proceso que termina
en pagos, eso vale mas que varias de las alertas.

De cada archivo se guarda nombre, ruta, tamano, fecha de modificacion
y el sha256 del contenido. Dos corridas con el mismo manifiesto
tuvieron las mismas entradas; si un total cambio y el manifiesto no,
el cambio esta en el codigo, no en los datos.
"""

import datetime
import hashlib
from pathlib import Path

import pandas as pd

COLUMNAS_MANIFIESTO = (
    "entrada", "archivo", "ruta", "bytes", "modificado", "sha256",
)

# Leer de a pedazos para no cargar en memoria un Excel grande solo
# para hashearlo.
_PEDAZO = 1024 * 1024


def sha256_de(ruta):
    """El sha256 del archivo, o '' si no se pudo leer."""

    digestor = hashlib.sha256()

    try:
        with open(ruta, "rb") as archivo:
            for pedazo in iter(lambda: archivo.read(_PEDAZO), b""):
                digestor.update(pedazo)
    except OSError:
        return ""

    return digestor.hexdigest()


def _fila(etiqueta, ruta):
    try:
        estado = ruta.stat()
    except OSError:
        return None

    modificado = datetime.datetime.fromtimestamp(estado.st_mtime)

    return {
        "entrada": etiqueta,
        "archivo": ruta.name,
        "ruta": str(ruta),
        "bytes": estado.st_size,
        "modificado": modificado.strftime("%Y-%m-%d %H:%M:%S"),
        "sha256": sha256_de(ruta),
    }


def construir_manifiesto(entradas):
    """
    entradas: iterable de (etiqueta, ruta). Las rutas que no existen o
    no son archivos se saltan; las repetidas se cuentan una vez.

    Devuelve un DataFrame con una fila por archivo realmente leido.
    """

    filas = []
    vistas = set()

    for etiqueta, ruta in entradas:
        if ruta is None:
            continue

        ruta = _como_ruta(ruta)
        if ruta is None or not ruta.is_file():
            continue

        clave = str(ruta)
        if clave in vistas:
            continue
        vistas.add(clave)

        fila = _fila(etiqueta, ruta)
        if fila is not None:
            filas.append(fila)

    return pd.DataFrame(filas, columns=list(COLUMNAS_MANIFIESTO))


def _como_ruta(valor):
    try:
        return Path(valor)
    except TypeError:
        return None
