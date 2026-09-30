# -*- coding: utf-8 -*-
"""
Hojas del CRA que salen de las MISMAS fuentes que ya usa el BESS
(confirmado por el usuario):

  - FD_CPF, FD_CSF, FD_CTF: las tres hojas horarias del
    SSCC_Desempeño_* del DCO (CPF/CSF/CTF Horario), con los rangos que ya
    usa Script/Fd/Desempeno_Horario.py (encabezados en la fila 11).
  - PRORRATA_RETIROS, bloque 1 (traz. 8.1): la matriz periodo x empresa
    pagadora, desde la hoja "Prorrata 15min" de Prorrata_Retiros_<AAMM>,
    leida con el mismo lector del BESS (nucleo.prorrata_retiros).

Las hojas FD se copian enteras, con todas las unidades y sus
encabezados reales: el filtro por unidad candidata del CRA (su
InfoTecnica) se aplica recien en CÁLCULO_CRA, que no esta todavia.
"""

import calendar


from ..nucleo.prorrata_retiros import (
    COL_CUARTO,
    COL_PRORRATA,
    COL_SUMINISTRADOR,
    leer_prorrata_retiros,
)
from . import parametros as p
from .lectura import leer_columnas, leer_encabezados


def _nada(*_args, **_kwargs):
    pass


# ============================================================
# FD_CPF, FD_CSF, FD_CTF
# ============================================================

def construir_fd(ruta_sscc, seccion, registrar=_nada):
    """
    Una de las tres hojas horarias del SSCC_Desempeño_* tal cual: los
    encabezados de la fila 11 y los datos desde la 12.
    """

    _, hoja_origen, columnas = p.HOJAS_FD[seccion]
    letras = list(columnas)

    encabezados = leer_encabezados(
        ruta_sscc, hoja_origen, p.FILA_ENCABEZADO_FD, letras
    )
    datos = leer_columnas(
        ruta_sscc, hoja_origen, p.FILA_ENCABEZADO_FD + 1, letras
    )

    datos.columns = _nombres_unicos(letras, encabezados)
    registrar(f"  {hoja_origen}: {len(datos):,} fila(s).")

    return datos.reset_index(drop=True)


def _nombres_unicos(letras, encabezados):
    """El texto del encabezado; la letra si esta vacio o repetido."""

    nombres, vistos = [], set()

    for letra in letras:
        nombre = encabezados.get(letra) or letra
        if nombre in vistos:
            nombre = f"{nombre} ({letra})"
        vistos.add(nombre)
        nombres.append(nombre)

    return nombres


# ============================================================
# PRORRATA_RETIROS, bloque 1
# ============================================================

def construir_matriz_prorrata(ruta_prorrata, anio, mes, registrar=_nada):
    """
    Prorrata 15min (cuarto, suministrador, prorrata) -> una fila por
    periodo de 15 minutos del mes y una columna por empresa pagadora,
    en orden alfabetico (el que deja el lector del BESS). Donde una
    empresa no retiro en ese periodo, la celda queda vacia.
    """

    largo = leer_prorrata_retiros(ruta_prorrata, registrar=registrar)

    empresas = sorted(set(largo[COL_SUMINISTRADOR]))
    matriz = (
        largo.pivot(index=COL_CUARTO, columns=COL_SUMINISTRADOR,
                    values=COL_PRORRATA)
        .reindex(columns=empresas)
        .sort_index()
    )

    _avisar_cobertura(matriz.index, anio, mes, registrar)
    registrar(
        f"  Prorrata: {len(matriz):,} periodo(s) x {len(empresas)} empresa(s)."
    )

    matriz.index.name = p.COLUMNA_PERIODO_PRORRATA
    matriz.columns.name = None

    return matriz.reset_index()


def _avisar_cobertura(periodos, anio, mes, registrar):
    """
    Aviso si los periodos no son exactamente 1..dias*96. No es un error:
    en un mes con cambio de hora el total no es dias*96 y esa regla
    todavia no esta cerrada (traz. 0.1).
    """

    esperados = calendar.monthrange(anio, mes)[1] * p.PERIODOS_POR_DIA
    presentes = set(int(x) for x in periodos)
    faltan = sorted(set(range(1, esperados + 1)) - presentes)
    sobran = sorted(presentes - set(range(1, esperados + 1)))

    if faltan or sobran:
        registrar(
            f"  AVISO Prorrata: se esperaban los periodos 1..{esperados} "
            f"({esperados // p.PERIODOS_POR_DIA} dias x "
            f"{p.PERIODOS_POR_DIA}); faltan {len(faltan)} y sobran "
            f"{len(sobran)}"
            + (f" (primeros que faltan: {faltan[:5]})" if faltan else "")
            + (f" (primeros que sobran: {sobran[:5]})" if sobran else "")
            + ". Si el mes tiene cambio de hora, puede ser correcto."
        )
