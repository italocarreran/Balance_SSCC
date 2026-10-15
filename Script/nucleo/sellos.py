# -*- coding: utf-8 -*-
"""
Cuando se escribio cada hoja de Balance_BESS.xlsx.

Las hojas viven todas en el mismo archivo, asi que la fecha del
archivo no dice nada de una hoja en particular (cambia con cualquiera).
Cada vez que se escribe una hoja se anota aca la hora, en un archivo
chico al lado del libro. Con eso la ventana sabe si una hoja quedo
vieja respecto de lo que usa (ver vigencia.py).

Si el archivo no esta (un libro de antes) o no se puede leer, no pasa
nada: las hojas sin sello usan la fecha del libro.
"""

import json
import time
from pathlib import Path


def ruta_sellos(ruta_libro):
    """<carpeta>/.sellos_Balance_BESS.json, al lado del libro."""

    ruta_libro = Path(ruta_libro)
    return ruta_libro.with_name(f".sellos_{ruta_libro.stem}.json")


def leer_sellos(ruta_libro):
    """{hoja: instante (segundos, como st_mtime)}; {} si no hay."""

    try:
        datos = json.loads(ruta_sellos(ruta_libro).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return {
        str(hoja): float(valor)
        for hoja, valor in datos.items()
        if isinstance(valor, (int, float))
    }


def sellar(ruta_libro, hojas, instante=None):
    """Anota que las hojas se acaban de escribir. Nunca levanta error."""

    hojas = [hoja for hoja in hojas if hoja]
    if not hojas:
        return
    instante = time.time() if instante is None else instante
    sellos = leer_sellos(ruta_libro)
    for hoja in hojas:
        sellos[hoja] = instante
    try:
        ruta_sellos(ruta_libro).write_text(
            json.dumps(sellos, ensure_ascii=False, indent=1, sort_keys=True),
            encoding="utf-8",
        )
    except OSError:
        pass
