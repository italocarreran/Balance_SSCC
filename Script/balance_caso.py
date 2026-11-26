# -*- coding: utf-8 -*-
"""
El balance de un caso como lo eligen las ventanas: año _ mes _ proceso
("2026_Marzo_Definitivo").

Pedido del usuario (2026-11-18, ventana del BESS; 2026-11-26, la del CRA
igual): tres casillas -año, mes escrito y proceso- en vez del AAMM. Cada
balance (mes + proceso) tiene su propia carpeta. Por dentro el calculo
sigue usando el AAMM ("2603"), que sale del año y el mes.

Solo lo que no depende de cada ventana (sin config.json, sin tkinter):
lo usan Balance_BESS.py y Balance_CRA.py.
"""

import time
from pathlib import Path


ANO_INICIAL = 2025

MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)

# Los que estan siempre; los demas (R01P, R01D, ...) los agrega el
# usuario con el "+" de la lista y quedan en config.json.
PROCESOS_FIJOS = ("Preliminar", "Definitivo")
OPCION_AGREGAR_PROCESO = "+ Agregar otro..."
CLAVE_PROCESOS = "procesos"


def anos_disponibles(ano_actual=None):
    """['2025', '2026', ...] hasta el año de hoy."""

    actual = int(ano_actual or time.localtime().tm_year)
    return [str(ano) for ano in range(ANO_INICIAL, max(actual, ANO_INICIAL) + 1)]


def aamm_de(ano, mes):
    """('2026', 'Marzo') -> '2603'. '' si falta algo o no es valido."""

    try:
        numero = MESES.index(str(mes)) + 1
        ano = int(ano)
    except (ValueError, TypeError):
        return ""
    return f"{ano % 100:02d}{numero:02d}"


def ano_mes_de(aamm):
    """'2603' -> ('2026', 'Marzo'). ('', '') si no es un AAMM."""

    aamm = str(aamm or "").strip()
    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        return "", ""
    return str(2000 + int(aamm[:2])), MESES[int(aamm[2:]) - 1]


def nombre_balance(ano, mes, proceso):
    """'2026_Marzo_Definitivo' (lo que falte no se escribe)."""

    return "_".join(str(p) for p in (ano, mes, proceso) if p)


def clave_balance(aamm, proceso):
    """La clave de carpetas_por_periodo: '2603_Definitivo'."""

    return f"{aamm}_{proceso}" if proceso else str(aamm)


def limpiar_proceso(texto):
    """
    El nombre de un proceso agregado a mano: sin espacios ni
    caracteres que no van en un nombre de carpeta. '' si no queda nada.
    """

    texto = "".join(
        c for c in str(texto or "").strip()
        if c.isalnum() or c in "-"
    )
    return texto


def proceso_en_el_nombre(ruta, procesos=None):
    """
    El proceso que dice el nombre de la carpeta ('... 2603 Definitivo' ->
    'Definitivo'), o '' si no dice ninguno (o dice mas de uno).
    """

    nombre = Path(str(ruta or "")).name.lower()
    hallados = [
        proceso for proceso in (procesos or PROCESOS_FIJOS)
        if proceso.lower() in nombre
    ]
    return hallados[0] if len(hallados) == 1 else ""


def nombre_de_clave(clave):
    """'2602_Definitivo' -> '2026_Febrero_Definitivo'; '2602' -> '2026_Febrero'."""

    aamm, _, proceso = str(clave).partition("_")
    ano, mes = ano_mes_de(aamm)
    return nombre_balance(ano, mes, proceso) if ano else str(clave)
