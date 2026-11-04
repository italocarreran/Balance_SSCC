# -*- coding: utf-8 -*-
"""
Cuando se escribio cada hoja de 11_PAGOS_BESS_AAMM.xlsx.

Las hojas viven todas en el mismo archivo, asi que la fecha del
archivo no dice nada de una hoja en particular (cambia con cualquiera).
Cada vez que se escribe una hoja se anota aca la hora, en un archivo
chico al lado del libro. Con eso la ventana sabe si una hoja quedo
vieja respecto de lo que usa (ver vigencia.py). El archivo va en la
carpeta Log/ del caso (ruta_sellos).

Si el archivo no esta (un libro de antes) o no se puede leer, no pasa
nada: las hojas sin sello usan la fecha del libro.
"""

import json
import time
from pathlib import Path


# Van en la carpeta Log/ del caso, con lo demas de cada corrida (pedido
# del usuario 2026-11-04: el .json suelto al lado de la planilla no se
# entendia que era).
CARPETA_SELLOS = "Log"


def ruta_sellos(ruta_libro):
    """<caso>/Log/sellos_11_PAGOS_BESS_AAMM.json."""

    ruta_libro = Path(ruta_libro)
    return ruta_libro.parent / CARPETA_SELLOS / f"sellos_{ruta_libro.stem}.json"


def ruta_sellos_antigua(ruta_libro):
    """Donde estaban antes: <caso>/.sellos_<libro>.json, al lado del libro."""

    ruta_libro = Path(ruta_libro)
    return ruta_libro.with_name(f".sellos_{ruta_libro.stem}.json")


def leer_sellos(ruta_libro):
    """{hoja: instante (segundos, como st_mtime)}; {} si no hay."""

    datos = None
    for ruta in (ruta_sellos(ruta_libro), ruta_sellos_antigua(ruta_libro)):
        try:
            datos = json.loads(ruta.read_text(encoding="utf-8"))
            break
        except (OSError, ValueError):
            continue
    if not isinstance(datos, dict):
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
        ruta_sellos(ruta_libro).parent.mkdir(parents=True, exist_ok=True)
        ruta_sellos(ruta_libro).write_text(
            json.dumps(sellos, ensure_ascii=False, indent=1, sort_keys=True),
            encoding="utf-8",
        )
    except OSError:
        pass
