# -*- coding: utf-8 -*-
"""
Como se reparte la carpeta de un mes entre las dos herramientas.

Pedido del usuario (2026-11-16): el BESS y el CRA de un mes trabajan
sobre LA MISMA carpeta. Cada uno guarda lo suyo en su propia subcarpeta
y lo que usan los dos queda en la raiz:

    <carpeta del mes>/
        BESS/                       (Balance_BESS.py)
            Auxiliares/  Medidas/  Ofertas/  Cmg/  Subastas/  Log/
            11_PAGOS_BESS_AAMM.xlsx, SAE_AAMM.csv
        CRA/                        (Balance_CRA.py)
            Auxiliares/  Energia/  FP/  CO/  SC y CO/  Log/
            5_REMUNERACION_CRA_AAMM.xlsx
        FD y FMA/                   (comun: SSCC_Desempeño_*; los fma_* son del BESS)
        Prorrata retiros/           (comun)

Un caso armado ANTES de esta division (las carpetas del BESS sueltas en
la raiz) sigue andando sin mover nada: si no existe <carpeta>/BESS/ y
si hay alguna carpeta propia del BESS suelta, la raiz del BESS es la
carpeta misma (ver raiz_herramienta). Lo mismo para el CRA. Nunca se
mezcla: o todo lo propio va adentro de BESS/, o todo suelto.
"""

from pathlib import Path


CARPETA_BESS = "BESS"
CARPETA_CRA = "CRA"

# Las carpetas que comparten las dos herramientas, en la raiz del mes.
# "FD y FMA" se llamaba antes "SSCC_Desempeño": se sigue aceptando si
# la nueva no existe (ver carpeta_fd).
CARPETA_FD_FMA = "FD y FMA"
CARPETA_FD_FMA_ANTIGUA = "SSCC_Desempeño"
CARPETA_PRORRATA_RETIROS = "Prorrata retiros"

CARPETAS_COMUNES = (CARPETA_FD_FMA, CARPETA_PRORRATA_RETIROS)


def es_caso_antiguo(base, herramienta, distintivas, archivos=()):
    """
    True si 'base' es un caso de antes de la division: no tiene la
    subcarpeta de la herramienta y si tiene suelta en la raiz alguna de
    sus carpetas propias que la delatan (`distintivas`: las que la otra
    herramienta no tiene) o su planilla (`archivos`: patrones glob, p.
    ej. "11_PAGOS_BESS_*.xlsx").
    """

    base = Path(base)
    if (base / herramienta).is_dir() or not base.is_dir():
        return False
    if any((base / nombre).is_dir() for nombre in distintivas):
        return True
    try:
        return any(
            ruta.is_file() for patron in archivos for ruta in base.glob(patron)
        )
    except OSError:
        return False


def raiz_herramienta(base, herramienta, distintivas, archivos=()):
    """
    Donde van las carpetas propias de la herramienta y su planilla:
    base/<herramienta>, o base misma en un caso de antes (ver
    es_caso_antiguo).
    """

    base = Path(base)
    if es_caso_antiguo(base, herramienta, distintivas, archivos):
        return base
    return base / herramienta


def carpeta_fd(base, antiguas=()):
    """
    La carpeta comun del FD: 'FD y FMA' en la raiz, salvo que no exista
    y si una de las de nombre anterior ('SSCC_Desempeño' y las que se
    pasen en 'antiguas', relativas a base).
    """

    base = Path(base)
    nueva = base / CARPETA_FD_FMA
    if nueva.is_dir():
        return nueva
    for nombre in (CARPETA_FD_FMA_ANTIGUA,) + tuple(antiguas):
        if (base / nombre).is_dir():
            return base / nombre
    return nueva


def relativa(ruta, base):
    """'BESS/Medidas' (para mostrar): la ruta relativa a base, con '/'."""

    try:
        return Path(ruta).relative_to(Path(base)).as_posix()
    except ValueError:
        return Path(ruta).name
