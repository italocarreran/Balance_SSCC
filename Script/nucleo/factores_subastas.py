# -*- coding: utf-8 -*-
"""
FD y FMA de cada fila de Subastas, calculados en el momento de usarlos.

Pedido del usuario (2026-10-13): "el FD se aplica mas adelante, hay que
evitar que se genere un doble descuento: quitarlo de la hoja Subastas y
calcularlo y aplicarlo en los calculos cuando corresponda. Lo mismo con
el FMA".

Hasta ahora la hoja Subastas traia dos columnas mas (P = FD, Q = FMA)
calculadas al importarla, y RE545 las sumaba en sus bloques "FD"
(AI:AN) y "FMA" (AO:AT). Ahora:

  - la hoja Subastas queda solo con lo que dice el Access (B:O, hasta
    "Energía SSCC"): no se descuenta nada al importarla, y para
    importarla ya no se abre el SSCC_Desempeño_* ni las salidas de FMA;
  - el FD sale de la hoja 'FD' de la planilla (los bloques CSF y CPF ya
    importados) y el FMA de la hoja 'FMA' (las salidas fma_cpf/csf/cft
    importadas con su boton);
  - el unico lugar donde se aplican es Calculo RE545, que es donde la
    planilla original los multiplica (AU = reserva x FD x FMA). E Costos
    aplica el FD por su lado, desde la misma hoja FD (AM:AR), y nunca
    uso las columnas de Subastas.

El calculo fila por fila es EXACTAMENTE el de antes
(calcular_fd_subastas / calcular_fma_subastas de fma.py, con las mismas
claves Control + Unidad + Hora_mes y la misma equivalencia de nombres):
lo unico que cambia es de donde salen las tablas.
"""

from pathlib import Path

import pandas as pd

from .externos import desempeno_fd
from .fma import (
    TITULO_BLOQUE_FD, TITULO_BLOQUE_FMA_CPF, calcular_fd_subastas,
    calcular_fma_subastas, construir_dic_bloque_diccionario,
)
from .ecostos_prorratas import construir_dic_mapeo_diccionario
from .lector import leer_excel
from .parametros import FMA_CON_CTF, HOJA_FMA
from .lectura import fila_de_titulos, leer_tabla_bajo_titulo
from .hojas_entrada import NOMBRES_SUBASTAS
from .utiles import ErrorEntrada, _texto_seguro, _tiene_valor, normalizar

# Los nombres con que quedan las dos columnas en la copia EN MEMORIA de
# Subastas que usa RE545 (la hoja ya no las tiene).
COLUMNA_FD_SUBASTAS = "FD"
COLUMNA_FMA_SUBASTAS = "FMA"

# Las columnas de la hoja FD de la planilla que hacen falta para armar
# las tablas de FD por unidad y hora (ver NOMBRES_FD_CSF/NOMBRES_FD_CPF).
_COLUMNAS_FD = {
    "csf": {
        "fd": "Factor de Desempeño\n (Fd_CSF)",
        "respuesta": "Respuesta CSF\n (Fact_CSF)",
    },
    "cpf": {
        "fd": "Factor de Desempeño\n(Fd_CPF)",
        "respuesta": None,
    },
}


# ============================================================
# FD: las tablas por unidad y hora, desde la hoja FD
# ============================================================

def tablas_fd_desde_hoja(df_fd_csf, df_fd_cpf, registrar=print):
    """
    Lo mismo que desempeno_fd.construir_tablas_fd() devolvia leyendo el
    SSCC_Desempeño_*, pero armado con los dos bloques de la hoja FD de
    la planilla:

        {"cpf": {(unidad_norm, hora_mes): fd}, "csf": {...},
         "ctf": {}, "participacion_csf": {(unidad_norm, hora_mes): 0|1}}

    Mismas reglas: primera aparicion gana, el FD se toma tal cual, y el
    Vector de Participacion CSF es 0 solo si la unidad "No Participó" y
    su alternativa TG/TV tampoco (desempeno_fd._participacion_csf).

    CTF queda vacio: la hoja FD no trae 'CTF Horario' (el usuario
    confirmo que las BESS no tienen FD CTF; antes esas filas tambien
    quedaban sin dato).
    """

    tablas = {"cpf": {}, "csf": {}, "ctf": {}, "participacion_csf": {}}

    for control, df_fd in (("csf", df_fd_csf), ("cpf", df_fd_cpf)):

        if df_fd is None or df_fd.empty:
            continue

        columnas = _COLUMNAS_FD[control]
        hora_mes = pd.to_numeric(df_fd.iloc[:, 1], errors="coerce")

        tabla = pd.DataFrame({
            "Unidad": df_fd["Unidad"].map(_texto_seguro),
            "Hora_Mes": hora_mes,
            "FD": pd.to_numeric(df_fd[columnas["fd"]], errors="coerce"),
        })

        if columnas["respuesta"]:
            tabla["Respuesta"] = df_fd[columnas["respuesta"]].map(_texto_seguro)

        tabla = tabla[tabla["Unidad"] != ""].dropna(subset=["Hora_Mes"])
        tabla["Hora_Mes"] = tabla["Hora_Mes"].astype(int)
        tabla["Unidad_norm"] = tabla["Unidad"].map(desempeno_fd._normalizar)
        tabla = tabla.reset_index(drop=True)

        dic = {}
        for fila in tabla.itertuples(index=False):
            clave = (fila.Unidad_norm, fila.Hora_Mes)
            if clave not in dic and pd.notna(fila.FD):
                dic[clave] = float(fila.FD)
        tablas[control] = dic

        if control == "csf" and "Respuesta" in tabla.columns:
            participacion = desempeno_fd._participacion_csf(tabla)
            dic_participacion = {}
            for fila, valor in zip(tabla.itertuples(index=False), participacion):
                dic_participacion.setdefault(
                    (fila.Unidad_norm, fila.Hora_Mes), valor
                )
            tablas["participacion_csf"] = dic_participacion

    registrar(
        f"  FD de subastas desde la hoja FD: {len(tablas['cpf']):,} "
        f"unidad-hora CPF, {len(tablas['csf']):,} CSF."
    )

    return tablas


# ============================================================
# FMA: la hoja 'FMA' de la planilla
# ============================================================

# Una tabla por salida de FMA, una al lado de la otra con su titulo
# (mismo formato que 'Ofertas SSCC' y 'Centrales'). La clave es la que
# usa cargar_tablas_fma().
_NOMBRES_FMA = {
    "cpf": "FMA CPF (fma_cpf)",
    "csf": "FMA CSF (fma_csf)",
    "ctf_unidad": "FMA CTF por unidad (fma_cft)",
    "ctf_configuracion": "FMA CTF por configuracion (fma_cft)",
}
# Todas las tablas que puede tener la hoja (las claves que devuelve
# leer_fma_consolidado, aunque el CTF este apagado).
CLAVES_FMA = tuple(_NOMBRES_FMA)
# Las que se escriben, con su titulo de cuadro ("Cuadro N° 1 — ...",
# mismo formato que PRORRATA_RETIROS). Sin el CTF (FMA_CON_CTF) quedan
# solo CPF y CSF.
TITULOS_FMA = {
    clave: f"Cuadro N° {numero} — {nombre}"
    for numero, (clave, nombre) in enumerate(
        (
            (clave, nombre) for clave, nombre in _NOMBRES_FMA.items()
            if FMA_CON_CTF or not clave.startswith("ctf")
        ),
        start=1,
    )
}


def leer_fma_consolidado(ruta_salida, registrar=print, libro=None):
    """
    Las tablas de la hoja 'FMA' de la planilla, con la misma forma que
    devuelve cargar_tablas_fma() (None para la que no esta).
    ErrorEntrada si la planilla no tiene la hoja: sin ella, el FMA de
    RE545 quedaria en 0 sin que nadie se entere.
    """

    ruta_salida = Path(ruta_salida)
    fuente = ruta_salida if libro is None else libro

    try:
        crudo = leer_excel(fuente, sheet_name=HOJA_FMA, header=None)
    except ValueError as error:
        raise ErrorEntrada(
            f"{ruta_salida.name} no tiene la hoja '{HOJA_FMA}' todavia: "
            f"generala con su boton 'Actualizar' (el FMA de Calculo RE545 "
            f"sale de ahi)."
        ) from error

    tablas = {clave: None for clave in CLAVES_FMA}

    if crudo.empty:
        return tablas

    # Los titulos van en la fila 2 (la 1 queda vacia); un libro de
    # antes los tiene en la 1.
    fila = fila_de_titulos(crudo)
    titulos = {
        _texto_seguro(crudo.iloc[fila, col]): col
        for col in range(crudo.shape[1])
        if _tiene_valor(crudo.iloc[fila, col])
    }

    for clave, titulo in TITULOS_FMA.items():

        # El titulo de ahora o el de antes de los "Cuadro N°".
        columna = titulos.get(titulo, titulos.get(_NOMBRES_FMA[clave]))
        if columna is None:
            continue

        ancho = 0
        while (
            columna + ancho < crudo.shape[1]
            and fila + 1 < crudo.shape[0]
            and _tiene_valor(crudo.iloc[fila + 1, columna + ancho])
        ):
            ancho += 1

        if not ancho:
            continue

        tablas[clave] = leer_tabla_bajo_titulo(
            fuente, HOJA_FMA, crudo, columna, ancho, fila_titulo=fila
        )

    registrar(
        "  FMA desde la hoja FMA: "
        + ", ".join(
            f"{clave} {len(tabla):,} fila(s)" if tabla is not None
            else f"{clave} sin tabla"
            for clave, tabla in tablas.items()
        )
        + "."
    )

    return tablas


# ============================================================
# EL CALCULO
# ============================================================

def _completar_con_otros_nombres(
    df_subastas, diccionario, tablas_fd, tablas_fma, dic_unidad_fd, dic_cpf,
    registrar,
):
    """
    Configuración de Subastas -> unidad del FD y -> central de FMA CPF:
    si el nombre que da el Diccionario (o el nombre tal cual) no esta en
    la tabla, se prueba con los demas nombres de la central en el
    Diccionario; si alguno esta, se usa y se avisa (DIC-002). Completa
    dic_unidad_fd y dic_cpf en el lugar.
    """

    from .equivalencias import Equivalencias
    from .lectura import ROL_FD, ROL_FMA_CPF

    columna_config = NOMBRES_SUBASTAS["K"]
    if diccionario is None or columna_config not in df_subastas.columns:
        return

    norm_fd = desempeno_fd._normalizar
    unidades_fd = {
        unidad
        for clave in ("cpf", "csf", "participacion_csf")
        for unidad, _hora in (tablas_fd or {}).get(clave, {})
    }
    tabla_cpf = (tablas_fma or {}).get("cpf")
    centrales_cpf = (
        {normalizar(c) for c in tabla_cpf["Central"]}
        if tabla_cpf is not None and "Central" in tabla_cpf.columns
        else set()
    )

    eq_fd = Equivalencias(diccionario, normalizar_fn=norm_fd)
    eq_cpf = Equivalencias(diccionario)
    if not eq_fd:
        return

    for config in pd.unique(df_subastas[columna_config].dropna()):
        clave = normalizar(config)
        if unidades_fd:
            actual = dic_unidad_fd.get(clave, config)
            if norm_fd(actual) not in unidades_fd:
                unidad = eq_fd.buscar(
                    config, ROL_FD, unidades_fd, "las unidades de la hoja FD"
                )
                if unidad is not None:
                    dic_unidad_fd[clave] = unidad
        if centrales_cpf:
            actual = dic_cpf.get(clave, config)
            if normalizar(actual) not in centrales_cpf:
                central = eq_cpf.buscar(
                    config, ROL_FMA_CPF, centrales_cpf,
                    "las centrales de FMA CPF",
                )
                if central is not None:
                    dic_cpf[clave] = central

    eq_fd.informar(registrar)
    eq_cpf.informar(registrar)


def subastas_con_fd_y_fma(
    df_subastas, df_fd_csf, df_fd_cpf, tablas_fma, diccionario,
    registrar=print,
):
    """
    Copia de Subastas con COLUMNA_FD_SUBASTAS y COLUMNA_FMA_SUBASTAS
    calculadas en el momento (lo que antes eran Subastas!P y Q). Si la
    hoja es de antes y trae esas columnas, se reemplazan: el FD y el
    FMA tienen un solo origen, nunca dos.

    diccionario: la hoja Diccionario (de la hoja 'Centrales'), para la
    equivalencia Configuración -> unidad del FD y -> central de FMA CPF.
    """

    dic_unidad_fd = {}
    dic_cpf = {}

    if diccionario is not None:
        dic_unidad_fd = construir_dic_bloque_diccionario(
            diccionario, TITULO_BLOQUE_FD
        ) or construir_dic_mapeo_diccionario(diccionario)
        dic_cpf = construir_dic_bloque_diccionario(
            diccionario, TITULO_BLOQUE_FMA_CPF
        )

    registrar("  Calculo RE545: FD y FMA de cada fila de Subastas...")

    tablas_fd = tablas_fd_desde_hoja(df_fd_csf, df_fd_cpf, registrar=registrar)

    _completar_con_otros_nombres(
        df_subastas, diccionario, tablas_fd, tablas_fma, dic_unidad_fd,
        dic_cpf, registrar,
    )

    df = df_subastas.copy()
    df[COLUMNA_FD_SUBASTAS] = calcular_fd_subastas(
        df, tablas_fd, dic_unidad_fd=dic_unidad_fd, registrar=registrar
    )
    df[COLUMNA_FMA_SUBASTAS] = calcular_fma_subastas(
        df, tablas_fma, dic_cpf=dic_cpf, tablas_fd=tablas_fd,
        dic_unidad_fd=dic_unidad_fd, registrar=registrar,
    )

    return df
