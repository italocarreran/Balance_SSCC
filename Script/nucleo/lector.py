# -*- coding: utf-8 -*-
"""
Leer Excel rapido.

Pedido del usuario (2026-11-03): "hacerlo mucho más rápido". Leer con
openpyxl es lo mas lento de cada paso: el archivo de Ofertas tardaba 36 s
y las hojas grandes de Balance_BESS.xlsx unos 12 s cada una. El motor
"calamine" de pandas (paquete python-calamine, escrito en Rust) lee lo
mismo entre 4 y 6 veces mas rapido.

Se usa SOLO donde se comprobo que entrega exactamente el mismo
DataFrame (valores y tipos) que openpyxl:

  - las hojas de Balance_BESS.xlsx, que escribe este mismo programa
    (numeros, textos y las fechas de Medidores);
  - el archivo de Ofertas SSCC y el de Prorrata de retiros (numeros y
    textos).

Los archivos con fechas y horas de otros sistemas (SOC, Medidas_SAE,
SSCC_Desempeño) siguen con openpyxl: ahi un lector distinto podria
redondear una hora de otra forma y no hay como comprobarlo con los
archivos reales.

Si python-calamine no esta instalado, todo se lee con openpyxl como
antes (mismo resultado, mas lento).
"""

import pandas as pd

try:  # pragma: no cover - depende de lo instalado
    import python_calamine  # noqa: F401
    MOTOR_RAPIDO = "calamine"
except ImportError:  # pragma: no cover
    MOTOR_RAPIDO = None


def leer_excel(fuente, *args, **opciones):
    """pd.read_excel con el motor rapido (si fuente es una ruta)."""

    if MOTOR_RAPIDO and not isinstance(fuente, pd.ExcelFile):
        opciones.setdefault("engine", MOTOR_RAPIDO)
    return pd.read_excel(fuente, *args, **opciones)


def abrir_excel(ruta):
    """pd.ExcelFile con el motor rapido (las lecturas desde el usan el mismo)."""

    if MOTOR_RAPIDO:
        return pd.ExcelFile(ruta, engine=MOTOR_RAPIDO)
    return pd.ExcelFile(ruta)
