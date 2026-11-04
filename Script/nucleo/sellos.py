# -*- coding: utf-8 -*-
"""
Cuando se escribio cada hoja de 11_PAGOS_BESS_AAMM.xlsx.

Las hojas viven todas en el mismo archivo, asi que la fecha del
archivo no dice nada de una hoja en particular (cambia con cualquiera).
Cada vez que se escribe una hoja se anota aca la hora, en un archivo
chico al lado del libro. Con eso la ventana sabe si una hoja quedo
vieja respecto de lo que usa (ver vigencia.py). El archivo va en
__config__/AAAA/MM Mes/Balance BESS/<caso>/ (ruta_sellos), no en la
carpeta del caso.

Si el archivo no esta (un libro de antes) o no se puede leer, no pasa
nada: las hojas sin sello usan la fecha del libro.
"""

import hashlib
import json
import re
import time
from pathlib import Path


# Van en __config__/AAAA/MM Mes/Balance BESS/<caso>/ (regla del usuario,
# 2026-11-04: los json no van en las carpetas de trabajo). <caso> es el
# nombre de la carpeta del caso mas una marca corta de su ruta completa:
# un mismo mes puede tener mas de un caso (R01D, R02D...) y las carpetas
# suelen llamarse igual ("SSCC", "02 CASO RELIQUIDACION").
CARPETA_LOG_ANTIGUA = "Log"


def _config():
    try:
        from .. import config
    except ImportError:  # pragma: no cover - depende de como se importe
        import config
    return config


def carpeta_del_caso(carpeta_caso, aamm):
    """__config__/AAAA/MM Mes/Balance BESS/<nombre del caso>_<marca>."""

    carpeta_caso = Path(carpeta_caso)
    try:
        completa = str(carpeta_caso.resolve())
    except OSError:
        completa = str(carpeta_caso)
    marca = hashlib.sha1(completa.casefold().encode("utf-8")).hexdigest()[:8]
    nombre = re.sub(r'[<>:"/\\|?*]', "_", carpeta_caso.name or "caso")
    return _config().carpeta_del_mes(aamm) / f"{nombre}_{marca}"


def _aamm_del_libro(ruta_libro):
    """'11_PAGOS_BESS_2609.xlsx' -> '2609'; None si el nombre no lo dice."""

    hallado = re.search(r"_(\d{4})$", Path(ruta_libro).stem)
    return hallado.group(1) if hallado else None


def ruta_sellos(ruta_libro):
    """__config__/AAAA/MM Mes/Balance BESS/<caso>/sellos_<libro>.json."""

    ruta_libro = Path(ruta_libro)
    return (
        carpeta_del_caso(ruta_libro.parent, _aamm_del_libro(ruta_libro))
        / f"sellos_{ruta_libro.stem}.json"
    )


def rutas_sellos_antiguas(ruta_libro):
    """Donde estuvieron antes: <caso>/Log/ y, mas antes, al lado del libro."""

    ruta_libro = Path(ruta_libro)
    return [
        ruta_libro.parent / CARPETA_LOG_ANTIGUA / f"sellos_{ruta_libro.stem}.json",
        ruta_libro.with_name(f".sellos_{ruta_libro.stem}.json"),
    ]


def ruta_sellos_antigua(ruta_libro):
    """La mas antigua: <caso>/.sellos_<libro>.json, al lado del libro."""

    return rutas_sellos_antiguas(ruta_libro)[-1]


def mover_sellos_antiguos(ruta_libro, ruta_libro_antigua=None):
    """
    Pasa a __config__ los sellos que hayan quedado en la carpeta del caso
    (de ruta_libro o, si se renombro, de ruta_libro_antigua). Nunca pisa
    unos que ya esten; si no se puede mover, quedan donde estaban.
    """

    destino = ruta_sellos(ruta_libro)
    origenes = rutas_sellos_antiguas(ruta_libro)
    if ruta_libro_antigua is not None:
        origenes += rutas_sellos_antiguas(ruta_libro_antigua)
    for origen in origenes:
        if not origen.is_file():
            continue
        try:
            if not destino.exists():
                destino.parent.mkdir(parents=True, exist_ok=True)
                origen.replace(destino)
            else:
                origen.unlink()
        except OSError:
            pass
    try:
        carpeta_log = Path(ruta_libro).parent / CARPETA_LOG_ANTIGUA
        if carpeta_log.is_dir() and not any(carpeta_log.iterdir()):
            carpeta_log.rmdir()
    except OSError:
        pass


def leer_sellos(ruta_libro):
    """{hoja: instante (segundos, como st_mtime)}; {} si no hay."""

    datos = None
    for ruta in [ruta_sellos(ruta_libro)] + rutas_sellos_antiguas(ruta_libro):
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
