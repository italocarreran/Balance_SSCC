# -*- coding: utf-8 -*-
"""
Matriz de control: cada central y si se encontro en cada origen de datos.

Pedido del usuario (2026-11-03): "una matriz al control de corrida que
muestre las centrales en las filas y en las columnas los origenes (FD,
Subastas, ofertas, cmg, medidas, etc) y que marque si se encontro o no".

Se arma con lo que hay en las hojas de entrada de Balance_BESS.xlsx y la
hoja Centrales, con los MISMOS cruces que usa el calculo (por el
Diccionario), asi que un "No" es justo una central a la que ese dato no
le va a llegar.

Pedido del usuario (2026-11-10): que salga tambien en una corrida
PARCIAL (por ejemplo Subastas antes que FD), con cada origen que aun no
se corre marcado como tal, y con cuatro colores:

  - verde (OK): se encontro;
  - amarillo (AVISO): no esta, pero puede no estar (una central que no
    se adjudico CSF no tiene por que salir en el FD de CSF);
  - rojo (FALTA): no esta y deberia (la central se adjudico CPF en
    Subastas y no tiene FD CPF; no tiene medidas; no esta en el
    Diccionario...);
  - gris (PENDIENTE): ese origen aun no se corre.

Lo que la central se adjudico en Subastas manda sobre los demas: si se
adjudico CPF tiene que tener FD CPF y FMA CPF (y ofertas); si se
adjudico CSF, FD CSF. Si presta solo CPF, no aparecer en el FD de CSF
es normal (amarillo).

El estado de cada celda va en matriz.attrs["estados"], {(central,
columna): estado}; la hoja de control lo usa para pintar.
"""

import pandas as pd

from .lectura import (
    ROL_BALANCE_BESS, ROL_FD, ROL_FMA_CPF, ROL_OFERTAS, ROL_SUBASTAS,
    filas_diccionario,
)
from .utiles import _texto_seguro, normalizar

OK, AVISO, FALTA, PENDIENTE = "ok", "aviso", "falta", "pendiente"
TEXTO_PENDIENTE = "Aún no se corre"

# Las columnas de la matriz, siempre todas y en este orden (las que no
# se corrieron salen como TEXTO_PENDIENTE).
COL_RESUMEN = "Centrales (Resumen BESS)"
COL_DICCIONARIO = "Diccionario"
COL_MEDIDORES = "Medidores (Medidas_SAE)"
COL_SOC = "SoC"
COL_CMG = "CMg (por barra)"
COL_OFERTAS = "Ofertas SSCC"
COL_SUBASTAS = "Subastas"
COL_FD_CSF = "FD CSF"
COL_FD_CPF = "FD CPF"
COL_FMA_CPF = "FMA CPF"
COLUMNAS_MATRIZ = (
    COL_RESUMEN, COL_DICCIONARIO, COL_MEDIDORES, COL_SOC, COL_CMG,
    COL_OFERTAS, COL_SUBASTAS, COL_FD_CSF, COL_FD_CPF, COL_FMA_CPF,
)

COLUMNA_CENTRAL_MATRIZ = "Central"


def _nombres_de(central, fila_dic):
    """Todos los nombres con que puede aparecer la central."""

    nombres = {normalizar(central)}
    for rol in (ROL_BALANCE_BESS, ROL_FD, ROL_SUBASTAS, ROL_OFERTAS, ROL_FMA_CPF):
        valor = (fila_dic or {}).get(rol)
        if _texto_seguro(valor):
            nombres.add(normalizar(valor))
    return nombres


def _col(df, nombre):
    """
    La columna 'nombre' como Serie, aunque venga repetida (la hoja FD
    trae "Hora Mes" dos veces a proposito: df["Hora Mes"] da una tabla
    de dos columnas y pd.to_numeric revienta). Se usa la primera.
    """

    columna = df[nombre]
    if isinstance(columna, pd.DataFrame):
        columna = columna.iloc[:, 0]
    return columna


def _si_no(cantidad, unidad, si_cero="No"):
    if cantidad:
        return f"Sí ({cantidad:,} {unidad})".replace(",", ".")
    return si_cero


def _servicios_adjudicados(df_subastas, centrales, fila_de):
    """
    central -> set de servicios ("CPF", "CSF", "CTF") que tiene en
    Subastas (por la columna Concepto: "CPF(+)" -> "CPF"). Sin columna
    Concepto, set vacio (no se sabe el servicio).
    """

    if df_subastas is None or "Configuración" not in df_subastas:
        return {c: set() for c in centrales}
    nombres = _col(df_subastas, "Configuración").map(normalizar)
    if "Concepto" in df_subastas:
        servicio = _col(df_subastas, "Concepto").map(
            lambda v: _texto_seguro(v)[:3].upper()
        )
    else:
        servicio = pd.Series([""] * len(df_subastas), index=df_subastas.index)
    por_nombre = {}
    for nombre, serv in zip(nombres, servicio):
        if serv:
            por_nombre.setdefault(nombre, set()).add(serv)
    return {
        c: set().union(*[
            por_nombre.get(n, set())
            for n in _nombres_de(c, fila_de.get(normalizar(c)))
        ])
        for c in centrales
    }


def construir_matriz_centrales(
    resumen=None, diccionario=None, df_medidores=None, df_wxy=None,
    df_cmg=None, df_subastas=None, df_fd_csf=None, df_fd_cpf=None,
    tablas_fma=None,
):
    """
    DataFrame con una fila por central (las de 'Resumen BESS' y las de
    Medidores) y una columna por origen (COLUMNAS_MATRIZ, siempre
    todas). Un origen en None es uno que aun no se corre: sus celdas
    dicen TEXTO_PENDIENTE. El estado de cada celda (OK, AVISO, FALTA,
    PENDIENTE) va en .attrs["estados"].
    """

    filas_dic = filas_diccionario(diccionario) if diccionario is not None else []
    fila_de = {
        normalizar(fila.get(ROL_BALANCE_BESS)): fila
        for fila in filas_dic if _texto_seguro(fila.get(ROL_BALANCE_BESS))
    }

    tiene_resumen = resumen is not None and len(resumen.columns) > 0
    nombre_col = None
    if tiene_resumen:
        nombre_col = next(
            (c for c in resumen.columns if "nombre" in normalizar(c)),
            resumen.columns[0],
        )

    centrales = []
    if tiene_resumen:
        centrales += [_texto_seguro(v) for v in _col(resumen, nombre_col) if _texto_seguro(v)]
    if df_medidores is not None and "clave" in df_medidores:
        centrales += [_texto_seguro(v) for v in pd.unique(_col(df_medidores, "clave")) if _texto_seguro(v)]
    vistas = {}
    for central in centrales:
        vistas.setdefault(normalizar(central), central)
    centrales = sorted(vistas.values(), key=str.upper)

    columnas = {}
    estados = {}

    def poner(columna, celdas):
        """celdas: [(texto, estado)] en el orden de 'centrales'."""
        columnas[columna] = [texto for texto, _ in celdas]
        for central, (_, estado) in zip(centrales, celdas):
            estados[(central, columna)] = estado

    def pendiente():
        return [(TEXTO_PENDIENTE, PENDIENTE)] * len(centrales)

    def si_o_falta(cantidad, unidad):
        if cantidad:
            return _si_no(cantidad, unidad), OK
        return "No", FALTA

    # --- maestro de centrales ----------------------------------------
    if tiene_resumen:
        en_resumen = {normalizar(v) for v in _col(resumen, nombre_col)}
        poner(COL_RESUMEN, [
            ("Sí", OK) if normalizar(c) in en_resumen else ("No", FALTA)
            for c in centrales
        ])
    else:
        poner(COL_RESUMEN, pendiente())

    if filas_dic:
        poner(COL_DICCIONARIO, [
            ("Sí", OK) if normalizar(c) in fila_de else ("No", FALTA)
            for c in centrales
        ])
    else:
        poner(COL_DICCIONARIO, pendiente())

    # --- medidas -------------------------------------------------------
    if df_medidores is not None and "clave" in df_medidores:
        claves = _col(df_medidores, "clave").map(normalizar)
        cuartos = claves.value_counts().to_dict()
        poner(COL_MEDIDORES, [
            si_o_falta(cuartos.get(normalizar(c), 0), "cuartos")
            for c in centrales
        ])
        if "SoC" in df_medidores:
            con_soc = (
                pd.to_numeric(_col(df_medidores, "SoC"), errors="coerce").notna()
                .groupby(claves).sum().to_dict()
            )
            poner(COL_SOC, [
                si_o_falta(int(con_soc.get(normalizar(c), 0)), "cuartos")
                for c in centrales
            ])
        else:
            poner(COL_SOC, pendiente())
    else:
        poner(COL_MEDIDORES, pendiente())
        poner(COL_SOC, pendiente())

    # --- CMg -----------------------------------------------------------
    if df_cmg is not None and df_cmg.shape[1] > 3:
        barra_col = (
            next((c for c in resumen.columns if "barra" in normalizar(c)), None)
            if tiene_resumen else None
        )
        barra_de = {}
        if barra_col is not None:
            barra_de = {
                normalizar(n): _texto_seguro(b)
                for n, b in zip(_col(resumen, nombre_col), _col(resumen, barra_col))
            }
        con_cmg = df_cmg.iloc[:, 3].map(normalizar).value_counts().to_dict()
        celdas = []
        for c in centrales:
            barra = barra_de.get(normalizar(c), "")
            if not barra:
                celdas.append(("No (sin barra)", FALTA))
            elif con_cmg.get(normalizar(barra), 0):
                celdas.append((f"Sí ({barra})", OK))
            else:
                celdas.append((f"No ({barra} sin CMg)", FALTA))
        poner(COL_CMG, celdas)
    else:
        poner(COL_CMG, pendiente())

    # --- Subastas: lo que se adjudico manda sobre FD, FMA y ofertas ----
    hay_subastas = (
        df_subastas is not None and "Configuración" in df_subastas
    )
    servicios = _servicios_adjudicados(
        df_subastas if hay_subastas else None, centrales, fila_de
    )

    filas_subastas = {c: 0 for c in centrales}
    if hay_subastas:
        por_nombre = _col(df_subastas, "Configuración").map(normalizar).value_counts().to_dict()
        filas_subastas = {
            c: sum(
                por_nombre.get(n, 0)
                for n in _nombres_de(c, fila_de.get(normalizar(c)))
            )
            for c in centrales
        }

    def sin_dato(central, servicio, que):
        """Texto y estado de una celda sin dato de 'que' ('FD', 'FMA')."""
        if not hay_subastas:
            return "No (Subastas aún no se corre)", AVISO
        if servicio in servicios[central]:
            return (
                f"No — adjudicada {servicio} en Subastas, debería tener {que}",
                FALTA,
            )
        return f"No (sin {servicio} en Subastas)", AVISO

    # --- ofertas -------------------------------------------------------
    if df_wxy is not None and "Central" in df_wxy and "Oferta completa" in df_wxy:
        tabla = df_wxy.assign(_c=_col(df_wxy, "Central").map(normalizar).to_numpy())
        celdas = []
        for c in centrales:
            filas = tabla[tabla["_c"] == normalizar(c)]
            completos = (
                int(pd.to_numeric(_col(filas, "Oferta completa"), errors="coerce").fillna(0).sum())
                if not filas.empty else 0
            )
            # Una central que no esta en el archivo de ofertas igual sale
            # en la tabla (con su propio nombre y oferta 0).
            en_archivo = not filas.empty and (
                completos > 0
                or ("Nombre" in filas
                    and _col(filas, "Nombre").map(normalizar).ne(normalizar(c)).any())
            )
            if en_archivo:
                celdas.append((
                    f"Sí ({completos} de {len(filas)} días completos)", OK
                ))
                continue
            texto = "No" if filas.empty else "No (no está en el archivo; oferta 0)"
            if hay_subastas and filas_subastas[c]:
                que = ", ".join(sorted(servicios[c])) or "algo"
                celdas.append((
                    f"{texto} — adjudicada {que} en Subastas", FALTA,
                ))
            elif hay_subastas:
                celdas.append((f"{texto} (sin adjudicación en Subastas)", AVISO))
            else:
                celdas.append((f"{texto} (Subastas aún no se corre)", AVISO))
        poner(COL_OFERTAS, celdas)
    else:
        poner(COL_OFERTAS, pendiente())

    if hay_subastas:
        celdas = []
        for c in centrales:
            filas = filas_subastas[c]
            if filas:
                detalle = ", ".join(sorted(servicios[c]))
                texto = _si_no(filas, "filas")
                if detalle:
                    texto = texto[:-1] + f"; {detalle})"
                celdas.append((texto, OK))
            else:
                # Una central puede no adjudicarse nada en el mes.
                celdas.append(("No (sin adjudicación)", AVISO))
        poner(COL_SUBASTAS, celdas)
    else:
        poner(COL_SUBASTAS, pendiente())

    # --- FD y FMA, por servicio ------------------------------------------
    for etiqueta, bloque, servicio in (
        (COL_FD_CSF, df_fd_csf, "CSF"), (COL_FD_CPF, df_fd_cpf, "CPF"),
    ):
        if bloque is None or "Central" not in bloque:
            poner(etiqueta, pendiente())
            continue
        # Solo las horas que existen: la hora que se salta en el
        # cambio de hora queda en la hoja FD (en 0 y sin Hora Mes) y
        # no cuenta (septiembre: 719 horas, no 720).
        if "Hora Mes" in bloque:
            bloque = bloque[pd.to_numeric(_col(bloque, "Hora Mes"), errors="coerce").notna().to_numpy()]
        por_central = _col(bloque, "Central").map(normalizar).value_counts().to_dict()
        poner(etiqueta, [
            (_si_no(por_central[normalizar(c)], "horas"), OK)
            if por_central.get(normalizar(c), 0)
            else sin_dato(c, servicio, "FD")
            for c in centrales
        ])

    tabla_cpf = None if tablas_fma is None else tablas_fma.get("cpf")
    if tabla_cpf is not None and "Central" in tabla_cpf:
        por_nombre = _col(tabla_cpf, "Central").map(normalizar).value_counts().to_dict()
        celdas = []
        for c in centrales:
            horas = sum(
                por_nombre.get(n, 0)
                for n in _nombres_de(c, fila_de.get(normalizar(c)))
            )
            celdas.append(
                (_si_no(horas, "horas"), OK) if horas
                else sin_dato(c, "CPF", "FMA")
            )
        poner(COL_FMA_CPF, celdas)
    elif tablas_fma is not None:
        # La hoja FMA esta pero sin tabla de CPF.
        poner(COL_FMA_CPF, [sin_dato(c, "CPF", "FMA") for c in centrales])
    else:
        poner(COL_FMA_CPF, pendiente())

    matriz = pd.DataFrame({
        COLUMNA_CENTRAL_MATRIZ: centrales,
        **{columna: columnas[columna] for columna in COLUMNAS_MATRIZ},
    })
    matriz.attrs["estados"] = estados
    return matriz
