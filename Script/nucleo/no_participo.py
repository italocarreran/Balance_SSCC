# -*- coding: utf-8 -*-
"""
Las horas "No participó" del FD no se pagan, POR DIRECCION.

Decision del usuario (2026-11-11), despues del experimento de la rama
claude/experimento-no-participo: "pasa al main la correccion [...] la que
va por direccion, porque ese es el que vale. Si tu estas adjudicado por
control de bajada y no prestaste control de subida o no participaste en
control de subida, el control de bajada te lo tienen que pagar".

En el archivo de desempeno una hora "No participó" puede venir con FD = 1
(docs/Trazabilidad_FD_a_DB_Y_planilla3.md, seccion 25): con esto, el FD de
ESA direccion pasa a 0. La otra direccion queda como venga.

  Respuesta CPF+  -> CPF(+)        Respuesta CPF-  -> CPF(-)
  Respuesta CSF   -> CSF(+) y CSF(-)   (una sola respuesta para las dos)

Donde se aplica:
  - Calculo E Costos (AM:AR): usa el FD por direccion de la hoja FD
    (fd_por_direccion, antes de armar los diccionarios del FD);
  - Calculo RE545 (el FD de cada fila de Subastas): cada fila mira la
    Respuesta de SU direccion (Concepto CPF(+), CPF(-), ...), ver
    factores_subastas.calcular_fd_subastas y claves_no_participo.

La hoja FD de la planilla no se toca: muestra el archivo tal cual vino.
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


def fd_por_direccion(df_fd, servicio):
    """
    Copia del bloque con el FD de cada direccion en 0 donde esa direccion
    "No participó". Devuelve (copia, cuantas direcciones-hora quedaron en 0).
    """

    if df_fd is None or df_fd.empty:
        return df_fd, 0
    copia = df_fd.copy()
    mascaras = mascaras_no_participo(df_fd, servicio)
    cambiadas = 0
    for sentido, filas in mascaras.items():
        columna = f"{servicio}({sentido})"
        if columna not in copia or not filas.any():
            continue
        valores = pd.to_numeric(copia[columna], errors="coerce")
        cambiadas += int((filas & valores.fillna(0).ne(0)).sum())
        copia.loc[filas, columna] = 0.0
    return copia, cambiadas


def claves_no_participo(df_fd, servicio, normalizar_unidad):
    """
    {(servicio en minusculas, sentido, unidad normalizada, hora del mes)}
    de las direcciones "No participó": para el FD de cada fila de Subastas
    en RE545 (que trae su direccion en Concepto).
    """

    if df_fd is None or df_fd.empty or "Unidad" not in df_fd:
        return set()
    mascaras = mascaras_no_participo(df_fd, servicio)
    unidades = df_fd["Unidad"].map(lambda u: normalizar_unidad(_texto_seguro(u)))
    horas = pd.to_numeric(df_fd.iloc[:, 1], errors="coerce")
    claves = set()
    for sentido, filas in mascaras.items():
        for unidad, hora in zip(unidades[filas], horas[filas]):
            if unidad and pd.notna(hora):
                claves.add((servicio.lower(), sentido, unidad, int(hora)))
    return claves


def sentido_de_concepto(concepto):
    """'CPF(+)' -> '+', 'CSF(-)' -> '-', otro -> None."""

    texto = _texto_seguro(concepto).replace(" ", "")
    if texto.endswith("(+)"):
        return "+"
    if texto.endswith("(-)"):
        return "-"
    return None
