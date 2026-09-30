# -*- coding: utf-8 -*-
"""
Diccionarios del maestro y del CMg: los usan las dos hojas.
"""

import pandas as pd

from .parametros import ARCHIVO_CENTRALES, HOJA_RESUMEN_BESS
from .utiles import ErrorEntrada, normalizar


# ============================================================
# CALCULO E COSTOS (etapa base)
#
# El usuario pidio avanzar "por etapas: primero H + CMg + traspaso
# de Medidores". Lo que sigue replica solo esa parte de dos macros:
#
#   - Traspasar_Medidores_A_Calculos_Rapido (modulo
#     B_medidores_a_calculos): A:G (con D<->E invertidas), I/J
#     (energia de Medidores!I separada por signo, solo si
#     Ventana_No_Completa==1; si no, es de "Calculo RE545", fuera de
#     alcance), Medidores!J -> K, Medidores!K -> P. H NO la toca esta
#     macro (es formula, ver mas abajo).
#   - Asignar_CMg_a_Calculos_Turbo (modulo A_Carga_Cmg_a_Destino):
#     arma un diccionario CMg!D (Barra) + "|" + CMg!H (Cuarto de
#     Hora, normalizado con NormalizaCuarto) -> CMg!F, y lo vuelca en
#     la columna Q. Para "Calculo E Costos" la macro NO escribe R
#     (escribirR=False): eso solo aplica a "Calculo RE545".
#
# H (Barra), en la planilla original, es formula:
#     =VLOOKUP(G4, Resumen!B:G, 6, FALSE)
# Se homologa por NOMBRE de columna ("Nombre activo" / "Barra
# inyeccion" de Centrales.xlsx!Resumen BESS) en vez de por posicion,
# porque Centrales.xlsx no reproduce el layout Resumen!B:G del libro
# original.
#
# El resto de columnas de Actualizar_Calculos_Columnas (L, M, N, O,
# R, S, T, U, W, X, Y, AB:AF, AG:AX, AZ) queda para una etapa
# posterior (decision explicita del usuario).
#
# Nombres de columna: son PLACEHOLDERS derivados de los comentarios
# de la macro. Todavia no se pudo confirmar contra un archivo real
# con los encabezados de "Calculo E Costos" (las dos veces que el
# usuario adjunto un archivo para esto, solo traia las hojas FD y
# Subastas) -- se corrigen apenas se reciba ese archivo.
# ============================================================

def _normaliza_cuarto(valor):
    """
    Replica NormalizaCuarto:
        Error       -> ""
        Numerico    -> CStr(CLng(valor))  (texto del entero redondeado)
        Otro        -> Trim(CStr(valor))
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, str):
        texto = valor.strip()
        if texto == "":
            return ""
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            return texto
        return str(round(numero))

    if isinstance(valor, (int, float)):
        return str(round(float(valor)))

    return str(valor).strip()


def construir_dic_cmg(df_cmg):
    """
    Replica el paso 1) de Asignar_CMg_a_Calculos_Turbo: arma un
    diccionario clave -> (valor para Q, valor para R), a partir de la
    hoja CMg (columnas por posicion, sin renombrar - ver leer_cmg):
        D (indice 3) = Barra
        F (indice 5) = valor a asignar en Q
        H (indice 7) = Cuarto de Hora
        I (indice 8) = valor a asignar en R ("CMg Promedio")

    Los dos valores son los dos elementos del Array() que guarda el
    diccionario del VBA: dictCMg(clave)(0) va a Q en las dos hojas y
    dictCMg(clave)(1) va a R, pero SOLO en "Calculo RE545"
    (CompletarDestinoTurbo se llama con escribirR:=False para
    "Calculo E Costos" y escribirR:=True para "Calculo RE545").

    Si una clave se repite, gana la primera fila (igual que
    "If Not dictCMg.Exists(clave) Then Add" en VBA).
    """

    columna_d = df_cmg.columns[3]
    columna_f = df_cmg.columns[5]
    columna_h = df_cmg.columns[7]
    columna_i = df_cmg.columns[8]

    diccionario = {}

    for _, fila in df_cmg.iterrows():

        barra = fila[columna_d]
        barra = "" if pd.isna(barra) else str(barra).strip()

        cuarto_hora = _normaliza_cuarto(fila[columna_h])

        if barra == "" or cuarto_hora == "":
            continue

        clave = barra.upper() + "|" + cuarto_hora

        if clave not in diccionario:
            diccionario[clave] = (fila[columna_f], fila[columna_i])

    return diccionario


def construir_mapa_barra(resumen_bess):
    """
    Arma nombre_central -> barra de inyeccion, a partir de la hoja
    "Resumen BESS" de Centrales.xlsx (columnas "Nombre activo" y
    "Barra inyeccion", confirmadas en el plan de traspaso, seccion
    4.1). Se busca por nombre de columna normalizado, no por
    posicion.
    """

    return _mapa_resumen_bess_por_nombre(
        resumen_bess, "barra", "Barra inyección",
        convertir=_texto_o_vacio,
    )


def construir_dic_resumen_factor(resumen_bess):
    """
    Arma nombre_central -> factor (columna "Pmax (MW)") y el umbral
    global de SoC minimo, a partir de la MISMA hoja "Resumen BESS" de
    Centrales.xlsx que ya usa construir_mapa_barra().

    Replica Resumen!B:C (factor, usado en AE/AF) y Resumen!H8
    (umbral, usado en M) de Actualizar_Calculos_Columnas. El usuario
    confirmo con un archivo real que la hoja "Resumen" del libro
    original es la MISMA tabla que "Resumen BESS" (los mismos 9
    encabezados: Nombre activo...Eficiencia) -- no hace falta una
    hoja nueva ni un archivo aparte.

    El umbral (celda fija H8 en el original) es, en la practica, el
    valor de "% Energia sobre minima" de la PRIMERA fila de datos de
    la tabla -- aca se toma igual (primera fila con nombre de
    central, no una fila fija: el encabezado de Centrales.xlsx no
    esta siempre en la misma posicion, ver _leer_resumen_bess()).
    """

    dic_factor = _mapa_resumen_bess_por_nombre(
        resumen_bess, "pmax", "Pmax (MW)"
    )

    columna_nombre = _exigir_columna(
        resumen_bess, "Nombre activo", "nombre", "activ"
    )
    columna_umbral = _exigir_columna(
        resumen_bess, "% Energía sobre mínima (indicador nuevo ciclo)",
        "energia sobre",
    )

    filas_con_nombre = resumen_bess[columna_nombre].notna()

    if not filas_con_nombre.any():
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} no "
            f"tiene filas de datos para sacar el umbral de SoC minimo."
        )

    primer_indice = resumen_bess.index[filas_con_nombre][0]
    umbral_soc_minimo = float(resumen_bess.loc[primer_indice, columna_umbral])

    return dic_factor, umbral_soc_minimo


def _buscar_cmg(dic_cmg, barra, cuarto_hora):
    """
    Replica la busqueda de CompletarDestinoTurbo: clave
    UCase(Barra)+"|"+NormalizaCuarto(Cuarto de Hora). Devuelve
    siempre un par (valor para Q, valor para R); sin match, los dos
    en blanco (el VBA escribe vbNullString en las dos).
    """

    barra = "" if not barra else str(barra).strip()
    cuarto = _normaliza_cuarto(cuarto_hora)

    if barra == "" or cuarto == "":
        return (pd.NA, pd.NA)

    return dic_cmg.get(barra.upper() + "|" + cuarto, (pd.NA, pd.NA))


def construir_dic_resumen_capacidad(resumen_bess):
    """
    Arma nombre_central -> "Capacidad (MWh)", de la hoja
    "Resumen BESS" de Centrales.xlsx.

    Es el VLOOKUP(G, Resumen!$B$8:$J$26, 4, 0) que aparece en U y BC
    de RE545 (y en BC de Calculo E Costos). OJO: la 4ta columna del
    rango B:J NO es "Pmax (MW)" sino "Capacidad (MWh)" -- el orden
    real de la tabla es Nombre activo, Pmax (MW), Horas para descarga
    forzada, Capacidad (MWh), Energia minima, Barra inyeccion, %
    Energia sobre minima, Ciclos max diarios, Eficiencia (confirmado
    en el plan seccion 25.8, y consistente con que H use el indice 6
    para la barra de inyeccion).

    El "factor" de AE/AF de Calculo E Costos es otra cosa (Resumen!
    B:C, o sea el indice 2 = "Pmax (MW)"): eso lo da
    construir_dic_resumen_factor(), y es el mismo indice 2 que usa BN
    de RE545. No confundirlas.
    """

    return _mapa_resumen_bess_por_nombre(
        resumen_bess, "capacidad", "Capacidad (MWh)",
    )


def _columna_resumen_bess(resumen_bess, *textos):
    """
    La PRIMERA columna de "Resumen BESS" cuyo nombre normalizado
    contiene todos los textos dados, o None. El nombre de la central
    se pide como ("nombre", "activ"); el resto con un solo texto.

    Se busca por nombre de columna, no por posicion: el encabezado de
    Centrales.xlsx no esta siempre en la misma parte de la hoja (ver
    _leer_resumen_bess()).
    """

    for columna in resumen_bess.columns:
        clave = normalizar(columna)
        if all(texto in clave for texto in textos):
            return columna

    return None


def _exigir_columna(resumen_bess, etiqueta, *textos):
    """_columna_resumen_bess(), pero explota con un mensaje util."""

    columna = _columna_resumen_bess(resumen_bess, *textos)

    if columna is None:
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} debe "
            f"tener una columna '{etiqueta}'. Columnas encontradas: "
            f"{list(resumen_bess.columns)}"
        )

    return columna


def _numero_o_na(valor):
    return pd.NA if pd.isna(valor) else float(valor)


def _texto_o_vacio(valor):
    return "" if pd.isna(valor) else str(valor).strip()


def _mapa_resumen_bess_por_nombre(
    resumen_bess, texto_buscado, etiqueta, convertir=_numero_o_na
):
    """
    Helper comun de los cuatro diccionarios que salen de "Resumen
    BESS": central normalizada -> valor de la columna cuyo nombre
    normalizado contiene texto_buscado. Las filas sin nombre de
    central se saltan.
    """

    columna_nombre = _exigir_columna(
        resumen_bess, "Nombre activo", "nombre", "activ"
    )
    columna_valor = _exigir_columna(resumen_bess, etiqueta, texto_buscado)

    return {
        normalizar(nombre): convertir(valor)
        for nombre, valor in zip(
            resumen_bess[columna_nombre], resumen_bess[columna_valor]
        )
        if not pd.isna(nombre)
    }


def construir_dic_resumen_eficiencia(resumen_bess):
    """
    Arma nombre_central -> "Eficiencia", de la misma hoja
    "Resumen BESS" de Centrales.xlsx.

    Es el VLOOKUP(G, Resumen!$B$8:$J$26, 9, 0) de la columna V de
    RE545: la 9na columna del rango B:J es la ultima de las 9 de esa
    tabla, "Eficiencia".
    """

    return _mapa_resumen_bess_por_nombre(
        resumen_bess, "eficiencia", "Eficiencia",
    )
