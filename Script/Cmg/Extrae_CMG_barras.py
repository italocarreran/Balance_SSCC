# -*- coding: utf-8 -*-
"""
Extrae_CMG_barras — arma cmg.xlsx desde el CSV 15-minutal oficial.

Viene del script suelto que se corria a mano al lado del CSV (autor
original: Freddy.Arriagada), con tres cambios pedidos por el usuario:

  1) el CSV ya no se busca al lado del .py: se baja de la unidad de
     red a la carpeta Cmg/ del caso ("Traer cmg_15min") y de ahi se
     lee;
  2) las barras a filtrar no estan escritas en el codigo: se las pasa
     quien llama (nucleo las saca de "Resumen BESS" de
     Centrales.xlsx);
  3) no se escribe el Excel aca: este modulo devuelve DataFrames y es
     nucleo el que resuelve rutas y escribe.

No importa nada de nucleo (solo pandas): asi este modulo se puede
probar y correr suelto, y el dia que se saquen mas etapas a modulos
propios no hay ciclos de import. Los errores previsibles salen como
ErrorCmg; nucleo los traduce a su ErrorEntrada.

El nombre del archivo usa guiones bajos (no espacios) para que sea
importable como modulo normal.
"""

import re
import shutil
from pathlib import Path

import pandas as pd


# ============================================================
# DONDE VIVE EL CSV 15-MINUTAL
#
#   T:\CMgReales 15MIN\AAAA\AAMM\Mensual\CMg\Cmg para balance\
#       cmgAAMM_def_15minutal.csv
#
# (AAAA = año completo, AAMM = el periodo de 4 digitos de la ventana).
# Es la unica ruta del programa que apunta fuera de la carpeta base
# del caso: si la unidad T: cambia de letra, se cambia aca y nada mas.
# ============================================================

RAIZ_CMG_REALES = r"T:\CMgReales 15MIN"
SUBCARPETAS_CMG_REALES = ("Mensual", "CMg", "Cmg para balance")
PLANTILLA_CSV_CMG_15MIN = "cmg{aamm}_def_15minutal.csv"

SEPARADOR_CSV = ";"
CODIFICACION_CSV = "latin1"
COLUMNA_VALOR = "CMg[CLP/KWh]"
COLUMNA_PROMEDIO = "CMg_CLP_KWh_Promedio_Horario"
COLUMNA_CUARTO = "Cuarto de Hora"

COLUMNAS_REQUERIDAS = ("FECHA", "HORA", "MINUTO", "BARRA", COLUMNA_VALOR)

_PATRON_AAMM = re.compile(r"^\d{4}$")


class ErrorCmg(Exception):
    """Error previsible al traer o procesar el CSV de CMg."""


def _validar_aamm(aamm):
    aamm = str(aamm or "").strip()

    if not _PATRON_AAMM.match(aamm):
        raise ErrorCmg(
            f"Periodo invalido: '{aamm}'. Tienen que ser 4 digitos "
            f"(por ejemplo 2607 para julio de 2026)."
        )

    return aamm


def nombre_csv_15min(aamm):
    """'2608' -> 'cmg2608_def_15minutal.csv'"""

    return PLANTILLA_CSV_CMG_15MIN.format(aamm=_validar_aamm(aamm))


def ruta_csv_en_red(aamm, raiz=None):
    """
    Ruta del CSV del periodo en la unidad de red. No se valida que
    exista: la ventana quiere poder mostrarla igual cuando falta.
    """

    aamm = _validar_aamm(aamm)
    carpeta = Path(raiz or RAIZ_CMG_REALES) / str(2000 + int(aamm[:2])) / aamm

    for subcarpeta in SUBCARPETAS_CMG_REALES:
        carpeta = carpeta / subcarpeta

    return carpeta / nombre_csv_15min(aamm)


def carpeta_origen_csv(aamm, raiz=None):
    """
    La carpeta de la unidad de red de la que se trae el CSV del
    periodo. Es lo que la ventana muestra (y abre) como origen de la
    fila del CSV; por eso no revienta con un AAMM a medio escribir:
    ahi devuelve la raiz.
    """

    try:
        return ruta_csv_en_red(aamm, raiz).parent
    except ErrorCmg:
        return Path(raiz or RAIZ_CMG_REALES)


def ruta_csv_local(carpeta_cmg, aamm):
    """
    Donde se espera el CSV una vez traido: al lado de cmg.xlsx, en la
    carpeta Cmg/ del caso.
    """

    return Path(carpeta_cmg) / nombre_csv_15min(aamm)


def traer_csv_15min(carpeta_cmg, aamm, raiz=None, registrar=print):
    """
    Copia el CSV del periodo desde la unidad de red a la carpeta Cmg/
    del caso (boton "Traer cmg_15min"). Devuelve la ruta local.

    Se copia en vez de leerlo directo de la red para que el caso
    quede autocontenido: una vez traido, cmg.xlsx se puede regenerar
    sin la unidad conectada, y queda registrado con que archivo se
    trabajo.
    """

    aamm = _validar_aamm(aamm)
    origen = ruta_csv_en_red(aamm, raiz)

    if not origen.is_file():
        raise ErrorCmg(
            f"No se encontro el CSV de CMg del periodo {aamm}:\n"
            f"{origen}\n\n"
            f"Revisa que la unidad de red este conectada y que el "
            f"archivo del periodo ya este publicado."
        )

    carpeta_cmg = Path(carpeta_cmg)
    carpeta_cmg.mkdir(parents=True, exist_ok=True)
    destino = ruta_csv_local(carpeta_cmg, aamm)

    registrar(f"Copiando {origen}")
    registrar(f"  -> {destino}")

    shutil.copy2(origen, destino)

    registrar(f"  {destino.stat().st_size / 1024 / 1024:.1f} MB copiados")

    return destino


# ============================================================
# EL PROCESO EN SI
# ============================================================

def construir_cmg_desde_csv(ruta_csv, barras, registrar=print):
    """
    Lee el CSV 15-minutal, numera el "Cuarto de Hora" global segun los
    bloques que el archivo realmente trae (no asume 96 por dia: los
    dias de cambio de hora traen 92 o 100), filtra por las barras
    pedidas y agrega el promedio horario de CMg[CLP/KWh] por
    FECHA + HORA + BARRA.

    Devuelve (df_salida, resumen_dias). El orden de columnas del
    resultado es el del CSV + "Cuarto de Hora" + el promedio horario,
    que es justo el layout A:I que nucleo.leer_cmg() espera despues.
    """

    ruta_csv = Path(ruta_csv)

    df = pd.read_csv(ruta_csv, sep=SEPARADOR_CSV, encoding=CODIFICACION_CSV)

    registrar(f"  filas leidas del CSV: {len(df):,}")

    faltantes = [
        columna for columna in COLUMNAS_REQUERIDAS
        if columna not in df.columns
    ]

    if faltantes:
        raise ErrorCmg(
            f"{ruta_csv.name} no tiene la(s) columna(s) {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    df["FECHA_DT"] = pd.to_datetime(df["FECHA"].astype(str), format="%Y%m%d")
    df["HORA"] = pd.to_numeric(df["HORA"], errors="coerce").astype("Int64")
    df["MINUTO"] = pd.to_numeric(df["MINUTO"], errors="coerce").astype("Int64")

    # El CSV viene con coma decimal (es-CL).
    df[COLUMNA_VALOR] = pd.to_numeric(
        df[COLUMNA_VALOR].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )

    # Bloques reales del archivo (fecha + hora + minuto distintos),
    # numerados dentro de cada dia y despues acumulados: asi el
    # "Cuarto de Hora" global sale de lo que el CSV trae y no de una
    # cuenta teorica de 96 bloques diarios.
    bloques = (
        df[["FECHA_DT", "HORA", "MINUTO"]]
        .drop_duplicates()
        .sort_values(["FECHA_DT", "HORA", "MINUTO"])
        .reset_index(drop=True)
    )

    bloques["QH_DIA"] = bloques.groupby("FECHA_DT").cumcount() + 1

    resumen_dias = (
        bloques.groupby("FECHA_DT", as_index=False)
        .agg(QH_DEL_DIA=("QH_DIA", "max"))
    )
    resumen_dias["OFFSET_DIA"] = (
        resumen_dias["QH_DEL_DIA"].cumsum().shift(fill_value=0)
    )
    resumen_dias["HORAS_DEL_DIA"] = resumen_dias["QH_DEL_DIA"] / 4

    bloques = bloques.merge(resumen_dias, on="FECHA_DT", how="left")
    bloques[COLUMNA_CUARTO] = bloques["OFFSET_DIA"] + bloques["QH_DIA"]

    df = df.merge(
        bloques[["FECHA_DT", "HORA", "MINUTO", COLUMNA_CUARTO]],
        on=["FECHA_DT", "HORA", "MINUTO"],
        how="left",
    )

    df["BARRA"] = df["BARRA"].astype(str).str.strip()

    # Se compara en mayusculas (mismo criterio que nucleo._buscar_cmg),
    # pero se conserva el texto tal cual viene del CSV.
    buscadas = {str(barra).strip().upper() for barra in barras}
    df_filtrado = df[df["BARRA"].str.upper().isin(buscadas)].copy()

    encontradas = set(df_filtrado["BARRA"].str.upper().unique())
    sin_datos = [
        barra for barra in barras
        if str(barra).strip().upper() not in encontradas
    ]

    if sin_datos:
        registrar(
            f"  AVISO: {len(sin_datos)} barra(s) de Centrales.xlsx no "
            f"aparecen en el CSV: {', '.join(sin_datos)}"
        )

    if df_filtrado.empty:
        raise ErrorCmg(
            f"Ninguna de las {len(barras)} barras de Centrales.xlsx "
            f"aparece en {ruta_csv.name}. Revisa que las barras esten "
            f"escritas igual que en el CSV (ej. 'TOCOPILLA_____110')."
        )

    df_filtrado[COLUMNA_PROMEDIO] = (
        df_filtrado
        .groupby(["FECHA", "HORA", "BARRA"])[COLUMNA_VALOR]
        .transform("mean")
        .round(6)
    )

    df_filtrado = (
        df_filtrado
        .drop(columns=["FECHA_DT"])
        .sort_values([COLUMNA_CUARTO, "BARRA"])
        .reset_index(drop=True)
    )

    return df_filtrado, resumen_dias


def validar_layout(df, registrar=print):
    """
    cmg.xlsx lo vuelve a leer nucleo.leer_cmg() POR POSICION (D =
    Barra, F = valor de Q, H = Cuarto de Hora, I = CMg promedio), asi
    que un cambio de columnas en el CSV de origen romperia
    silenciosamente la etapa siguiente. Se avisa aca, donde todavia se
    entiende por que.
    """

    if df.shape[1] < 9:
        raise ErrorCmg(
            f"El resultado quedo con {df.shape[1]} columnas y cmg.xlsx "
            f"necesita al menos 9 (A:I): el CSV de origen debe haber "
            f"cambiado de formato."
        )

    for indice, nombre in {3: "BARRA", 7: COLUMNA_CUARTO}.items():

        real = str(df.columns[indice])

        if real.strip().lower() != nombre.strip().lower():
            registrar(
                f"  AVISO: se esperaba '{nombre}' en la columna "
                f"{chr(ord('A') + indice)} y quedo '{real}'. La lectura "
                f"posterior de cmg.xlsx es por posicion: revisa el "
                f"formato del CSV."
            )


def resumen_dias_anomalos(resumen_dias):
    """
    Dias que no tienen 24 h (cambio de hora), como lista de textos
    listos para el log.
    """

    anomalos = resumen_dias[resumen_dias["HORAS_DEL_DIA"] != 24]

    return [
        f"{fila['FECHA_DT']:%Y-%m-%d}: {fila['HORAS_DEL_DIA']:g} h "
        f"({int(fila['QH_DEL_DIA'])} cuartos)"
        for _, fila in anomalos.iterrows()
    ]
