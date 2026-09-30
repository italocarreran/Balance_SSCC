# -*- coding: utf-8 -*-
"""
Lo minimo compartido por los modulos de Medidas.

A proposito NO se importa nada de nucleo (ver Script/__init__.py):
`normalizar` esta duplicada -son diez lineas- para no crear un ciclo
de imports entre el nucleo y sus modulos de etapa.
"""

import unicodedata


class ErrorMedidas(Exception):
    """Error previsible al armar Medidas_SAE.xlsx."""


# ============================================================
# CLAVES DE LAS APIS DEL COORDINADOR
#
# Son DOS y son distintas entre si: la de medidas.coordinador.cl
# (PRMTE) no sirve para operacion.coordinador.cl (Gen real) ni al
# reves. Antes vivian escritas en el codigo; ahora salen de la seccion
# "claves_api" de config.json, que no se versiona.
#
# El valor no depende de quien corra el programa -es el mismo para
# todo el equipo-, pero como config.json es local, cada uno lo pega
# una vez en su copia. El formato para pegar esta en
# config.ejemplo.json, y el error de Script/config.py lo repite
# entero cuando falta.
#
# Se leen en el momento de usarlas, no al importar: asi se puede
# completar el config con el programa abierto.
# ============================================================

try:
    from ..config import (
        ErrorConfig, clave_api, CLAVE_PRMTE, CLAVE_GENERACION_REAL,
    )
except ImportError:  # pragma: no cover - depende de como se importe
    from config import (
        ErrorConfig, clave_api, CLAVE_PRMTE, CLAVE_GENERACION_REAL,
    )


def leer_clave_api(cual):
    """
    clave_api(), pero el error sale como ErrorMedidas: el resto de
    Medidas ya sabe traducir ese a ErrorEntrada, y la ventana lo
    muestra tal cual.
    """

    try:
        return clave_api(cual)
    except ErrorConfig as error:
        raise ErrorMedidas(str(error)) from error


EXTENSIONES_EXCEL = (".xlsx", ".xlsm", ".xlsb", ".xls")


def normalizar(texto):
    """Minusculas, sin tildes, espacios colapsados."""

    if texto is None:
        return ""

    try:
        import pandas as pd
        if pd.isna(texto):
            return ""
    except (ImportError, TypeError, ValueError):
        pass

    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return " ".join(texto.lower().split())


def columna_que_contenga(df, *fragmentos):
    """
    Primera columna de df cuyo nombre normalizado contenga TODOS los
    fragmentos dados. None si no hay ninguna. Se busca por nombre y no
    por posicion, igual que en el resto del proyecto.
    """

    for columna in df.columns:
        clave = normalizar(columna)
        if all(fragmento in clave for fragmento in fragmentos):
            return columna

    return None


__all__ = [
    "ErrorMedidas", "leer_clave_api", "CLAVE_PRMTE",
    "CLAVE_GENERACION_REAL", "EXTENSIONES_EXCEL", "normalizar",
    "columna_que_contenga",
]
