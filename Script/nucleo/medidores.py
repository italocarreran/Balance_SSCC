# -*- coding: utf-8 -*-
"""
La hoja Medidores: columnas calculadas y armado.
"""

import pandas as pd
from pathlib import Path

from .ofertas_sscc import (
    calcular_r, calcular_s, calcular_t, cargar_resumen_en_medidores,
    construir_resumen_ofertas_sscc, construir_resumen_ventana_oferta,
)
from .parametros import (
    ARCHIVO_MEDIDAS_SAE, COLUMNAS_MEDIDORES_SALIDA, INICIO_VENTANA,
    UMBRAL_SOC,
)
from .utiles import normalizar


# ============================================================
# COLUMNAS CALCULADAS
# ============================================================

def calcular_ventana(clave, hora, inicio_ventana=INICIO_VENTANA):
    """
    Replica Medidores!L:

        =IF(G4<>G3, 0,
            IF(C4=C3, 0, IF(C4=$S$1, 1, 0)) + L3)

    Es un contador acumulado que:
      - suma 1 cuando la hora cambia y pasa a ser la de inicio;
      - se reinicia a 0 cuando cambia la central.

    El reinicio es por BLOQUE de filas consecutivas con la misma
    clave, no por groupby, para ser fiel a la comparacion fila a
    fila que hace Excel.
    """

    clave = pd.Series(clave).reset_index(drop=True)
    hora = pd.Series(hora).reset_index(drop=True)

    cambia_clave = clave.ne(clave.shift())
    cambia_hora = hora.ne(hora.shift())

    incremento = (
        cambia_hora & hora.eq(inicio_ventana)
    ).astype(int)

    # En la primera fila de cada bloque L vale 0 sin importar
    # el incremento, porque la formula corta antes de sumarlo.
    incremento = incremento.where(~cambia_clave, 0)

    bloque = cambia_clave.cumsum()

    return (
        incremento
        .groupby(bloque)
        .cumsum()
        .astype("int64")
    )


def calcular_indicador_soc(soc, umbral=UMBRAL_SOC):
    """
    Replica Medidores!O:  =1*(J3>6%)

    Un SoC vacio da 0, igual que en Excel donde una celda vacia
    no supera el umbral.
    """

    return (
        pd.Series(soc)
        .fillna(0)
        .gt(umbral)
        .astype("int64")
        .reset_index(drop=True)
    )


def calcular_clave_auxiliar(dia, hora_mes):
    """
    Replica Medidores!N:  =B3&"&"&E3

    Ya no se escribe en la hoja (ver COLUMNAS_AUXILIARES_MEDIDORES:
    no la lee nadie). Queda como la replica de esa columna de la
    planilla original, para poder reconstruirla si algun dia hace
    falta cruzar por dia+hora del mes.

    Excel concatena como texto. Se reproduce el formato entero
    para que no aparezcan '.0' que el Excel no tiene.
    """

    def texto(serie):
        serie = pd.Series(serie).reset_index(drop=True)
        numerica = pd.to_numeric(serie, errors="coerce")
        return numerica.map(
            lambda v: (
                ""
                if pd.isna(v)
                else (
                    str(int(v))
                    if float(v).is_integer()
                    else str(v)
                )
            )
        )

    return texto(dia) + "&" + texto(hora_mes)


# ============================================================
# CONSTRUCCION DE LA HOJA
# ============================================================

def construir_medidores(
    df_sae,
    df_soc,
    mes,
    registrar=print,
):
    """
    Arma la tabla equivalente a Medidores (A:Q + U).

    NO lee el archivo de Ofertas SSCC ni depende de el: las columnas R,
    S y T de la planilla original salen de ahi y se calculan aparte
    (ver COLUMNAS_OFERTAS_EN_MEDIDORES y
    completar_ofertas_en_medidores). Es el pedido del usuario de
    "independizar Medidas de ofertas": el boton "Actualizar" de
    Medidores anda aunque todavia no exista el *OfertasSSCC* del
    periodo.

    Devuelve (df_medidores, avisos).
    """

    avisos = []

    df = df_sae.copy()

    # --------------------------------------------------------
    # ORDEN
    # --------------------------------------------------------

    # Excel trae los datos ordenados por central y tiempo, y las
    # columnas recursivas dependen de ese orden.
    df = (
        df.sort_values(["clave", "intervalo"])
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # VALIDACION DE PERIODO
    # --------------------------------------------------------

    meses = sorted(df["Mes"].dropna().unique().tolist())

    if meses != [mes]:
        avisos.append(
            f"El periodo del SOC indica mes {mes} pero "
            f"{ARCHIVO_MEDIDAS_SAE} contiene {meses}."
        )

    # --------------------------------------------------------
    # J: CRUCE DEL SoC
    # --------------------------------------------------------

    soc = df_soc.copy()

    soc["timestamp"] = pd.to_datetime(soc["timestamp"])

    soc = (
        soc
        .dropna(subset=["timestamp"])
        .drop_duplicates(
            subset=["central", "timestamp"],
            keep="last",
        )
    )

    df = df.merge(
        soc[["central", "timestamp", "soc"]].rename(
            columns={
                "central": "clave",
                "timestamp": "intervalo",
                "soc": "SoC",
            }
        ),
        on=["clave", "intervalo"],
        how="left",
    )

    # Reporte de cobertura
    centrales_sae = set(df["clave"].unique())
    centrales_soc = set(soc["central"].unique())

    sin_soc = sorted(centrales_sae - centrales_soc)
    sobrantes = sorted(centrales_soc - centrales_sae)

    if sin_soc:
        # Para cada central sin bloque, se busca si algun nombre CRUDO
        # (antes de homologar, "nombre_scada_original") normaliza igual
        # a esa central -- si lo encuentra, es una pista fuerte de que
        # el bloque SI esta en el archivo de SoC pero la homologacion
        # (Centrales.xlsx!Diccionario) lo esta mandando a otro nombre.
        candidatos_por_normalizado = {}
        for origen in soc["nombre_scada_original"].unique():
            candidatos_por_normalizado.setdefault(
                normalizar(origen), []
            ).append(origen)

        detalle_sin_soc = []
        for central in sin_soc:
            candidatos = candidatos_por_normalizado.get(
                normalizar(central), []
            )
            if candidatos:
                detalle_sin_soc.append(
                    f"{central} (el SoC SI trae un bloque con nombre "
                    f"crudo {candidatos!r} -- revisar si "
                    f"Centrales.xlsx!Diccionario lo esta homologando "
                    f"a otro nombre distinto de '{central}')"
                )
            else:
                detalle_sin_soc.append(central)

        avisos.append(
            f"Centrales en {ARCHIVO_MEDIDAS_SAE} sin bloque de "
            f"SoC: {detalle_sin_soc}"
        )

    if sobrantes:
        avisos.append(
            f"Centrales en el SOC que no estan en "
            f"{ARCHIVO_MEDIDAS_SAE}: {sobrantes}"
        )

    for central in sorted(centrales_sae & centrales_soc):
        sub = df[df["clave"] == central]
        faltan = int(sub["SoC"].isna().sum())
        if faltan:
            avisos.append(
                f"{central}: {faltan:,} de {len(sub):,} "
                f"intervalos sin SoC."
            )

    # --------------------------------------------------------
    # COLUMNAS CALCULADAS
    # --------------------------------------------------------

    df["Ventana"] = calcular_ventana(
        df["clave"],
        df["Hora"],
    )

    df["Indicador_SoC"] = calcular_indicador_soc(df["SoC"])

    # --------------------------------------------------------
    # ORDEN FINAL DE COLUMNAS
    # --------------------------------------------------------
    #
    # El orden de la planilla original (LETRA_A_CAMPO) sin los
    # auxiliares: las cuatro columnas vacias (M, P, Q, U), la clave
    # auxiliar N y la copia K de la Ventana. Ver
    # COLUMNAS_AUXILIARES_MEDIDORES: ninguna se pierde de vista, K se
    # repone al leer la hoja (reponer_auxiliares_medidores) y N se
    # calcula con calcular_clave_auxiliar() cuando haga falta.

    df = df[COLUMNAS_MEDIDORES_SALIDA]

    registrar(
        f"Medidores construido: {len(df):,} filas x "
        f"{len(df.columns)} columnas (sin las columnas de Ofertas SSCC: "
        f"ver la hoja 'Ofertas SSCC')"
    )

    return df, avisos


def reponer_auxiliares_medidores(df_medidores):
    """
    Repone, sobre la hoja Medidores recien leida del consolidado, los
    auxiliares que ya no se escriben (ver
    COLUMNAS_AUXILIARES_MEDIDORES) y que las dos hojas de calculo si
    consumen:

      Copia_Ventana  copia de Ventana, que en Calculo E Costos y en
                     Calculo RE545 se llama "Ciclo de Carga del mes".

    No recalcula nada: Copia_Ventana ES Ventana fila a fila (plan,
    seccion 16.3), asi que el resultado del calculo es identico al de
    cuando la columna venia escrita en la hoja. Si el libro que se
    esta leyendo es de una corrida vieja y todavia la trae, se
    respeta la que trae.

    Devuelve el mismo DataFrame (modificado en el lugar), para poder
    encadenarlo en la lectura.
    """

    if "Copia_Ventana" not in df_medidores.columns:
        df_medidores["Copia_Ventana"] = df_medidores["Ventana"]

    return df_medidores


# ============================================================
# OFERTAS SSCC: LA HOJA PROPIA Y LAS TRES COLUMNAS DERIVADAS
# ============================================================

def construir_ofertas_sscc(
    df_medidores, ruta_ofertas, diccionario, anio, mes, registrar=print
):
    """
    Arma las dos tablas de la hoja "Ofertas SSCC" del consolidado:

      df_wxy             equivalente a Medidores!W:Y (central x dia)
      df_resumen_ventana equivalente a Medidores!AB:AE (central x ventana)

    Necesita Medidores ya construido -las centrales y las ventanas
    salen de ahi-, no al reves: esa es toda la inversion de dependencia
    que pidio el usuario. El resumen intermedio equivalente a la hoja
    "Resumen Ofertas SSCC" (Nombre/Año/Mes/Día/servicios/Oferta
    completa) es puramente auxiliar para calcular df_wxy: no se
    devuelve ni se persiste.

    Devuelve (df_wxy, df_resumen_ventana, avisos).
    """

    avisos = []

    registrar(f"  Leyendo {Path(ruta_ofertas).name}...")
    df_resumen_ofertas = construir_resumen_ofertas_sscc(
        ruta_ofertas, registrar=registrar
    )

    df_wxy, periodo_ofertas, avisos_wxy = cargar_resumen_en_medidores(
        df_resumen_ofertas,
        df_medidores["clave"].unique(),
        diccionario,
        registrar=registrar,
    )
    avisos.extend(avisos_wxy)

    if periodo_ofertas != (anio, mes):
        avisos.append(
            f"El resumen de Ofertas SSCC indica el periodo "
            f"{periodo_ofertas[0]}-{periodo_ofertas[1]:02d}, pero el "
            f"caso corresponde a {anio}-{mes:02d}."
        )

    r_valor, avisos_r = calcular_r(
        df_medidores, df_wxy, diccionario, registrar=registrar
    )
    avisos.extend(avisos_r)

    df_resumen_ventana = construir_resumen_ventana_oferta(
        df_medidores["clave"],
        df_medidores["Ventana"],
        r_valor,
        registrar=registrar,
    )

    return df_wxy, df_resumen_ventana, avisos


def completar_ofertas_en_medidores(
    df_medidores, df_wxy, df_resumen_ventana, diccionario, registrar=print
):
    """
    Devuelve una copia de Medidores con las tres columnas que salen de
    Ofertas SSCC (COLUMNAS_OFERTAS_EN_MEDIDORES), reconstruidas a
    partir de las dos tablas de la hoja "Ofertas SSCC" del consolidado:

      R (Oferta_Completa_Dia)      = VLOOKUP(dia + central) en df_wxy
      S (Indicador_Ventana_Oferta) = formula sobre la Ventana y R
      T (Ventana_No_Completa)      = 1 - Completa(central, ventana)

    Es exactamente el mismo calculo que antes vivia dentro de
    construir_medidores(); lo unico que cambio es de donde salen las
    dos tablas -de la hoja ya generada, en vez de releer el archivo de
    ofertas- y que el resultado ya no se persiste en la hoja Medidores.

    Devuelve (df_medidores_con_ofertas, avisos).
    """

    avisos = []
    df = df_medidores.copy()

    r_valor, avisos_r = calcular_r(
        df, df_wxy, diccionario, registrar=registrar
    )
    avisos.extend(avisos_r)
    df["Oferta_Completa_Dia"] = r_valor

    df["Indicador_Ventana_Oferta"] = calcular_s(
        df["Ventana"], df["Oferta_Completa_Dia"]
    )

    t_valor, sin_match_t = calcular_t(
        df["clave"], df["Ventana"], df_resumen_ventana
    )
    df["Ventana_No_Completa"] = t_valor

    if sin_match_t:
        avisos.append(
            f"{sin_match_t:,} fila(s) de Medidores no encontraron su "
            "grupo central+ventana en el resumen de Ofertas SSCC al "
            "calcular la columna T."
        )

    for aviso in avisos:
        registrar(f"  [AVISO] {aviso}")

    return df, avisos
