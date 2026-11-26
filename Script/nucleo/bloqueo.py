# -*- coding: utf-8 -*-
"""
Caso bloqueado: un archivo en la carpeta del BESS del caso.

Pedido del usuario (2026-11-18): "un boton de Bloquear/Desbloquear caso,
para no correr un caso anterior por error". El bloqueo es un archivo
(ARCHIVO_BLOQUEO) en <mes>/BESS/ (o en la carpeta del caso si es de antes
de BESS/), asi viaja con la carpeta: lo ve cualquiera que la abra, en
cualquier PC, y no depende de config.json. Adentro dice quien y cuando
lo bloqueo (texto, para que se entienda si alguien lo abre).

Lo respeta la ventana del BESS: con el caso bloqueado no corre ningun
proceso (botones del arbol, Ejecutar todo, antecedentes SAE). Desbloquear
es borrar el archivo.

El CRA usa el mismo mecanismo con su propio archivo en <mes>/CRA/ (pedido
del usuario 2026-11-26, la ventana del CRA igual a la del BESS): cada
funcion recibe `raiz`, la carpeta donde va el archivo (por defecto la del
BESS), y la ventana del CRA pasa la suya (cra.raiz_cra). Bloquear uno no
bloquea al otro.
"""

import time
from pathlib import Path

from .rutas import raiz_bess

ARCHIVO_BLOQUEO = "CASO_BLOQUEADO.txt"


def ruta_bloqueo(carpeta_base, raiz=None):
    raiz = Path(raiz) if raiz is not None else raiz_bess(Path(carpeta_base))
    return raiz / ARCHIVO_BLOQUEO


def info_bloqueo(carpeta_base, raiz=None):
    """
    {"quien": ..., "cuando": ..., "motivo": ...} si el caso esta
    bloqueado; None si no (o si no hay carpeta).
    """

    if not carpeta_base:
        return None
    ruta = ruta_bloqueo(carpeta_base, raiz)
    try:
        if not ruta.is_file():
            return None
        texto = ruta.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    datos = {"quien": "", "cuando": "", "motivo": ""}
    for linea in texto.splitlines():
        clave, _, valor = linea.partition(":")
        clave = clave.strip().lower()
        if clave in datos:
            datos[clave] = valor.strip()
    return datos


def esta_bloqueado(carpeta_base, raiz=None):
    return info_bloqueo(carpeta_base, raiz) is not None


def bloquear(carpeta_base, quien, motivo="", raiz=None, herramienta="Balance BESS"):
    """Deja el archivo de bloqueo. Devuelve su ruta."""

    ruta = ruta_bloqueo(carpeta_base, raiz)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(
        f"Caso bloqueado: la ventana del {herramienta} no corre ningun proceso.\n"
        "Para desbloquearlo, usar el boton 'Desbloquear caso' (o borrar este archivo).\n"
        f"quien: {quien}\n"
        f"cuando: {time.strftime('%d-%m-%Y %H:%M')}\n"
        f"motivo: {motivo}\n",
        encoding="utf-8",
    )
    return ruta


def desbloquear(carpeta_base, raiz=None):
    """Borra el archivo de bloqueo. True si habia uno."""

    ruta = ruta_bloqueo(carpeta_base, raiz)
    try:
        ruta.unlink()
    except FileNotFoundError:
        return False
    return True
