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
  - los demas archivos que escribe este programa: Medidas_SAE.xlsx (con
    sus fechas; comprobado con un mes entero), Cmg/cmg.xlsx y
    Subastas/subastas_AAMM.xlsx (numeros y textos);
  - el archivo de Ofertas SSCC y el de Prorrata de retiros (numeros y
    textos);
  - el archivo de SoC: con el real (docs/SOC_real_2607.xlsx) da el mismo
    DataFrame, fechas incluidas.

SSCC_Desempeño (de otro sistema, con fechas y sin un archivo real a mano
para comprobarlo) sigue con openpyxl; se lee en otro proceso mientras se
hace lo demas.

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
