# -*- coding: utf-8 -*-
"""
Matriz de control: cada central y si se encontro en cada origen de datos.

Pedido del usuario (2026-11-03): "una matriz al control de corrida que
muestre las centrales en las filas y en las columnas los origenes (FD,
Subastas, ofertas, cmg, medidas, etc) y que marque si se encontro o no".

Se arma con lo que el calculo ya tiene leido (las hojas de entrada de
Balance_BESS.xlsx y la hoja Centrales), con los MISMOS cruces que usa el
calculo (por el Diccionario), asi que un "No" es justo una central a la
que ese dato no le va a llegar. Cada celda dice "Sí ..." o "No ...";
la hoja de control las pinta de verde y rojo.
"""

import pandas as pd

from .lectura import (
    ROL_BALANCE_BESS, ROL_FD, ROL_FMA_CPF, ROL_OFERTAS, ROL_SUBASTAS,
    filas_diccionario,
)
from .utiles import _texto_seguro, normalizar

COLUMNA_CENTRAL_MATRIZ = "Central"


def _nombres_de(central, fila_dic):
    """Todos los nombres con que puede aparecer la central."""

    nombres = {normalizar(central)}
    for rol in (ROL_BALANCE_BESS, ROL_FD, ROL_SUBASTAS, ROL_OFERTAS, ROL_FMA_CPF):
        valor = (fila_dic or {}).get(rol)
        if _texto_seguro(valor):
            nombres.add(normalizar(valor))
    return nombres


def _si_no(cantidad, unidad, si_cero="No"):
    if cantidad:
        return f"Sí ({cantidad:,} {unidad})".replace(",", ".")
    return si_cero


def construir_matriz_centrales(
    resumen=None, diccionario=None, df_medidores=None, df_wxy=None,
    df_cmg=None, df_subastas=None, df_fd_csf=None, df_fd_cpf=None,
    tablas_fma=None,
):
    """
    DataFrame con una fila por central (las de 'Resumen BESS' y las de
    Medidores) y una columna por origen. Un origen que no se leyo en
    esta corrida no aparece como columna.
    """

    filas_dic = filas_diccionario(diccionario) if diccionario is not None else []
    fila_de = {
        normalizar(fila.get(ROL_BALANCE_BESS)): fila
        for fila in filas_dic if _texto_seguro(fila.get(ROL_BALANCE_BESS))
    }

    centrales = []
    if resumen is not None and len(resumen.columns):
        columna = next(
            (c for c in resumen.columns if "nombre" in normalizar(c)), resumen.columns[0]
        )
        centrales += [_texto_seguro(v) for v in resumen[columna] if _texto_seguro(v)]
    if df_medidores is not None and "clave" in df_medidores:
        centrales += [_texto_seguro(v) for v in pd.unique(df_medidores["clave"]) if _texto_seguro(v)]
    vistas = {}
    for central in centrales:
        vistas.setdefault(normalizar(central), central)
    centrales = sorted(vistas.values(), key=str.upper)

    columnas = {}

    if resumen is not None and len(resumen.columns):
        en_resumen = {
            normalizar(v) for v in resumen[next(
                (c for c in resumen.columns if "nombre" in normalizar(c)),
                resumen.columns[0],
            )]
        }
        columnas["Centrales (Resumen BESS)"] = [
            "Sí" if normalizar(c) in en_resumen else "No" for c in centrales
        ]

    if filas_dic:
        columnas["Diccionario"] = [
            "Sí" if normalizar(c) in fila_de else "No" for c in centrales
        ]

    if df_medidores is not None and "clave" in df_medidores:
        claves = df_medidores["clave"].map(normalizar)
        cuartos = claves.value_counts().to_dict()
        columnas["Medidores (Medidas_SAE)"] = [
            _si_no(cuartos.get(normalizar(c), 0), "cuartos") for c in centrales
        ]
        if "SoC" in df_medidores:
            con_soc = (
                pd.to_numeric(df_medidores["SoC"], errors="coerce").notna()
                .groupby(claves).sum().to_dict()
            )
            columnas["SoC"] = [
                _si_no(int(con_soc.get(normalizar(c), 0)), "cuartos") for c in centrales
            ]

    if resumen is not None and df_cmg is not None and df_cmg.shape[1] > 3:
        barra_col = next(
            (c for c in resumen.columns if "barra" in normalizar(c)), None
        )
        nombre_col = next(
            (c for c in resumen.columns if "nombre" in normalizar(c)), resumen.columns[0]
        )
        if barra_col is not None:
            barra_de = {
                normalizar(n): _texto_seguro(b)
                for n, b in zip(resumen[nombre_col], resumen[barra_col])
            }
            con_cmg = df_cmg.iloc[:, 3].map(normalizar).value_counts().to_dict()
            valores = []
            for c in centrales:
                barra = barra_de.get(normalizar(c), "")
                if not barra:
                    valores.append("No (sin barra)")
                else:
                    cantidad = con_cmg.get(normalizar(barra), 0)
                    valores.append(
                        f"Sí ({barra})" if cantidad else f"No ({barra} sin CMg)"
                    )
            columnas["CMg (por barra)"] = valores

    if df_wxy is not None and "Central" in df_wxy and "Oferta completa" in df_wxy:
        tabla = df_wxy.assign(_c=df_wxy["Central"].map(normalizar))
        valores = []
        for c in centrales:
            filas = tabla[tabla["_c"] == normalizar(c)]
            if filas.empty:
                valores.append("No")
                continue
            completos = int(pd.to_numeric(filas["Oferta completa"], errors="coerce").fillna(0).sum())
            en_archivo = (
                filas["Nombre"].map(normalizar).ne(normalizar(c)).any()
                or completos > 0
            )
            valores.append(
                f"Sí ({completos} de {len(filas)} días completos)"
                if en_archivo else "No (no está en el archivo; oferta 0)"
            )
        columnas["Ofertas SSCC"] = valores

    if df_subastas is not None and "Configuración" in df_subastas:
        por_nombre = df_subastas["Configuración"].map(normalizar).value_counts().to_dict()
        columnas["Subastas"] = [
            _si_no(
                sum(por_nombre.get(n, 0) for n in _nombres_de(c, fila_de.get(normalizar(c)))),
                "filas",
            )
            for c in centrales
        ]

    for etiqueta, bloque in (("FD CSF", df_fd_csf), ("FD CPF", df_fd_cpf)):
        if bloque is not None and "Central" in bloque:
            # Solo las horas que existen: la hora que se salta en el
            # cambio de hora queda en la hoja FD (en 0 y sin Hora Mes) y
            # no cuenta (septiembre: 719 horas, no 720).
            if "Hora Mes" in bloque:
                bloque = bloque[pd.to_numeric(bloque["Hora Mes"], errors="coerce").notna()]
            por_central = bloque["Central"].map(normalizar).value_counts().to_dict()
            columnas[etiqueta] = [
                _si_no(por_central.get(normalizar(c), 0), "horas") for c in centrales
            ]

    tabla_cpf = (tablas_fma or {}).get("cpf")
    if tabla_cpf is not None and "Central" in tabla_cpf:
        por_nombre = tabla_cpf["Central"].map(normalizar).value_counts().to_dict()
        columnas["FMA CPF"] = [
            _si_no(
                sum(por_nombre.get(n, 0) for n in _nombres_de(c, fila_de.get(normalizar(c)))),
                "horas",
            )
            for c in centrales
        ]

    return pd.DataFrame({COLUMNA_CENTRAL_MATRIZ: centrales, **columnas})
