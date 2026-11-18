# -*- coding: utf-8 -*-
"""
Las horas "No participó" del archivo de desempeno, POR DIRECCION: solo
para informar (Reporte), no cambian el pago.

Historia: el 2026-11-11 se puso en 0 el FD de la direccion "No participó"
(commit 1af3556). El 2026-11-18 el usuario lo devolvio: el FD lo calcula
el DCO y viene listo (docs/Trazabilidad_FD_a_DB_Y_planilla3.md, seccion
25: "No participó" con FD = 1 es un caso normal); en CPF es lo habitual
(85-95 % de las horas adjudicadas en oct) y en CSF ya descuenta el Vector
de Participacion del FMA. Ademas lo observaron los coordinados
(O2026-0052605/0052606, EMOAC). Hoy E Costos y RE545 usan el FD tal cual.

  Respuesta CPF+  -> CPF(+)        Respuesta CPF-  -> CPF(-)
  Respuesta CSF   -> CSF(+) y CSF(-)   (una sola respuesta para las dos)
"""

import pandas as pd

from .utiles import _texto_seguro, normalizar


def es_no_participo(valor):
    """'No Participó', 'No participo', ... -> True."""

    return normalizar(_texto_seguro(valor)).startswith("no particip")


def _direccion(nombre_columna):
    """'+', '-' o None (CSF: una respuesta para las dos direcciones)."""

    texto = normalizar(nombre_columna).replace(" ", "")
    if "cpf+" in texto:
        return "+"
    if "cpf-" in texto:
        return "-"
    return None


def mascaras_no_participo(df_fd, servicio):
    """
    {"+": filas "No participó" en subida, "-": en bajada} del bloque de la
    hoja FD del servicio ("CPF" o "CSF"). Sin columnas de respuesta, todo
    False.
    """

    falso = pd.Series(False, index=df_fd.index)
    mascaras = {"+": falso.copy(), "-": falso.copy()}
    for columna in df_fd.columns:
        if "respuesta" not in normalizar(columna):
            continue
        datos = df_fd[columna]
        if isinstance(datos, pd.DataFrame):
            datos = datos.iloc[:, 0]
        filas = datos.map(es_no_participo)
        direccion = _direccion(columna) if servicio == "CPF" else None
        for sentido in ((direccion,) if direccion else ("+", "-")):
            mascaras[sentido] = mascaras[sentido] | filas
    return mascaras


def sentido_de_concepto(concepto):
    """'CPF(+)' -> '+', 'CSF(-)' -> '-', otro -> None."""

    texto = _texto_seguro(concepto).replace(" ", "")
    if texto.endswith("(+)"):
        return "+"
    if texto.endswith("(-)"):
        return "-"
    return None
