# -*- coding: utf-8 -*-
"""
Nombres de archivo, carpeta y hoja del caso CRA, y el mapeo de columnas
de cada archivo de entrada.

Todo lo que viene de la trazabilidad
(docs/Trazabilidad_CRA_Periodo_Generico_v3.md) esta citado con su
seccion. Lo que NO viene de ahi es una propuesta de esta migracion y lo
dice al lado (PROPUESTA / PENDIENTE): se cambia aca y nada mas.
"""


# ============================================================
# SALIDA
# ============================================================

ARCHIVO_SALIDA = "Balance_CRA.xlsx"

HOJA_ENERGIA = "ENERGIA"
HOJA_FP = "FP"
HOJA_CO = "CO"
HOJA_SC_CO = "SC y CO"

# Orden de las hojas en el libro, de fin a inicio (mismo criterio que
# Balance_BESS.xlsx: arriba lo ultimo del calculo, abajo las entradas).
# Por ahora solo estan las entradas con su carga ya definida; las de
# calculo (CÁLCULO_CRA, PRORRATA_RETIROS, RESUMEN) y las entradas
# pendientes (TC, COTAS, RENDIMIENTOS, CONDICION_EMBALSE, FD_*,
# EMPRESAS) se agregan arriba cuando se cierren.
ORDEN_HOJAS_SALIDA = [HOJA_SC_CO, HOJA_CO, HOJA_FP, HOJA_ENERGIA]


# ============================================================
# CARPETAS DEL CASO -- PROPUESTA, sin confirmar con el usuario.
#
# Una subcarpeta por hoja de entrada, con el nombre de la hoja. Si se
# decide otra estructura, se cambia aca.
# ============================================================

CARPETA_ENERGIA = "Energia"
CARPETA_FP = "FP"
CARPETA_CO = "CO"
CARPETA_SC_CO = "SC y CO"

SUBCARPETAS_CASO = (CARPETA_ENERGIA, CARPETA_FP, CARPETA_CO, CARPETA_SC_CO)


# ============================================================
# PATRONES DE NOMBRE DE ARCHIVO
#
# Se comparan contra el nombre normalizado (minusculas, sin tildes), asi
# que "Cálculo" y "Calculo" son lo mismo. "{aamm}" se reemplaza por el
# periodo de la ventana.
#
# Si en la carpeta hay MAS DE UN archivo que cumple el patron, el
# programa se detiene y lo dice (mismo criterio que el SoC del BESS):
# no elige por fecha de modificacion, para no tomar en silencio el mes
# equivocado.
# ============================================================

EXTENSIONES_EXCEL = (".xlsx", ".xlsm")

# Traz. 6.1: Formato_Solicitud_SSAA_SSCC_Hidro_MesAAAA.xlsx. "MesAAAA"
# no se usa para buscar (no se sabe si el mes va con nombre o numero):
# el periodo se valida contra las columnas Año/Mes del propio archivo.
PREFIJO_ENERGIA = "formato_solicitud_ssaa_sscc_hidro"

# Traz. 6.4.1: el ejemplo es fp_2603*.xlsx. PENDIENTE confirmar si el
# nombre siempre lleva el AAMM; mientras tanto se busca solo "fp_".
PREFIJO_FP = "fp_"

# Traz. 6.4.2: cvar_cra_AAMM_*.xlsx.
PREFIJO_CO = "cvar_cra_{aamm}"

# Traz. 6.6.2 y 6.6.3.
PREFIJO_REPORTE_CRA = "reporte_cra"
PREFIJO_SOBRECOSTOS = "calculo_sobrecostossscc"


# ============================================================
# HOJA DE CADA ARCHIVO -- PENDIENTE: la trazabilidad no la dice.
#
# None = "no se sabe": si el libro tiene una sola hoja se usa esa; si
# tiene varias, el programa se detiene y las lista, en vez de adivinar
# cual es. Cuando el usuario confirme el nombre, se escribe aca.
# ============================================================

HOJA_ORIGEN_ENERGIA = None
HOJA_ORIGEN_FP = None
HOJA_ORIGEN_CO = None
HOJA_ORIGEN_REPORTE_CRA = None
HOJA_ORIGEN_SOBRECOSTOS = None


# ============================================================
# MAPEOS DE CARGA (letra del archivo origen -> campo)
#
# El orden de cada dict es el orden de las columnas en la hoja de
# salida.
# ============================================================

# --- ENERGIA (traz. 6.1) -----------------------------------
# Los datos empiezan en la fila 3 del archivo origen.
FILA_INICIO_ENERGIA = 3

CAMPO_UNIDAD = "Unidad Generadora/Central"
CAMPO_PUNTO = "Punto De Medida"
CAMPO_ANIO = "Año"
CAMPO_MES = "Mes"
CAMPO_DIA = "DIA"
CAMPO_HORADIA = "HORADIA"
CAMPO_PERIODO = "PERIODO DE CALCULO"
CAMPO_KWHD = "kWhD"
CAMPO_KWHR = "kWhR"

# Copias directas.
COPIA_ENERGIA = {
    "B": CAMPO_UNIDAD,
    "C": CAMPO_PUNTO,
    "D": CAMPO_ANIO,
    "E": CAMPO_MES,
    "F": CAMPO_DIA,
    "H": CAMPO_KWHD,
    "I": CAMPO_KWHR,
}

# Columnas del origen que solo se usan para calcular:
#   HORADIA            = HORA(G)
#   PERIODO DE CALCULO = J*4 + (MINUTO(G)+15)/15
LETRA_FECHA_HORA_ENERGIA = "G"
LETRA_HORA_ENERGIA = "J"

COLUMNAS_ENERGIA = [
    CAMPO_UNIDAD, CAMPO_PUNTO, CAMPO_ANIO, CAMPO_MES, CAMPO_DIA,
    CAMPO_HORADIA, CAMPO_PERIODO, CAMPO_KWHD, CAMPO_KWHR,
]

# --- FP (traz. 6.4.1): origen A2:D -------------------------
FILA_INICIO_FP = 2

COPIA_FP = {
    "A": "BarNom",
    "B": "Hora",
    "C": "FP",
    "D": "dia",
}

# --- CO (traz. 6.4.2): origen A2:D -------------------------
# El orden de salida es el de la hoja CO (N, P, Q, R), no el del
# origen: Dia viene de D y Hora de B.
FILA_INICIO_CO = 2

COPIA_CO = {
    "A": "nombre_configuración",
    "D": "Día",
    "B": "Hora",
    "C": "Costos_Operación",
}

# --- SC y CO (traz. 6.6.1) ---------------------------------
CAMPO_CLAVE_ANIO_MES = "Clave Año_Mes"
CAMPO_TIPO = "Tipo"
CAMPO_UNIDAD_SC = "Unidad"
CAMPO_CLAVE_BLOQUE = "Clave_Bloque"
CAMPO_COSTO = "Costo de Oportunidad y sobrecosto"
CAMPO_CPF_MAS = "CPF(+)"
CAMPO_CPF_MENOS = "CPF(-)"
CAMPO_CSF_MAS = "CSF(+)"
CAMPO_CSF_MENOS = "CSF(-)"
CAMPO_CTF_MAS = "CTF(+)"
CAMPO_CTF_MENOS = "CTF(-)"

COLUMNAS_SC_CO = [
    CAMPO_CLAVE_ANIO_MES, CAMPO_TIPO, CAMPO_UNIDAD_SC, CAMPO_CLAVE_BLOQUE,
    CAMPO_COSTO,
    CAMPO_CPF_MAS, CAMPO_CPF_MENOS,
    CAMPO_CSF_MAS, CAMPO_CSF_MENOS,
    CAMPO_CTF_MAS, CAMPO_CTF_MENOS,
]

# Participacion por servicio: se calcula en memoria, no se escribe
# (traz. 6.6.1: "no necesitan quedar almacenadas").
SERVICIOS_SC_CO = {
    "CPF": (CAMPO_CPF_MAS, CAMPO_CPF_MENOS),
    "CSF": (CAMPO_CSF_MAS, CAMPO_CSF_MENOS),
    "CTF": (CAMPO_CTF_MAS, CAMPO_CTF_MENOS),
}

# Origen 1: Reporte_CRA*.xlsx, desde la fila 2 (traz. 6.6.2).
FILA_INICIO_REPORTE_CRA = 2

COPIA_REPORTE_CRA = {
    "A": CAMPO_CLAVE_ANIO_MES,
    "B": CAMPO_TIPO,
    "H": CAMPO_UNIDAD_SC,
    "I": CAMPO_COSTO,
    "J": CAMPO_CPF_MAS,
    "K": CAMPO_CPF_MENOS,
    "L": CAMPO_CSF_MAS,
    "M": CAMPO_CSF_MENOS,
    "N": CAMPO_CTF_MAS,
    "O": CAMPO_CTF_MENOS,
}

# Clave_Bloque = DIA(C) & "#" & E
LETRA_FECHA_REPORTE_CRA = "C"
LETRA_BLOQUE_REPORTE_CRA = "E"

# Origen 2: Cálculo_SobrecostosSSCC_*.xlsm, desde la fila 7 (traz. 6.6.3).
FILA_INICIO_SOBRECOSTOS = 7

COPIA_SOBRECOSTOS = {
    "S": CAMPO_CLAVE_ANIO_MES,
    "T": CAMPO_TIPO,
    "U": CAMPO_UNIDAD_SC,
    "W": CAMPO_COSTO,
}

# Cada campo es la suma de tres columnas del origen.
SUMAS_SOBRECOSTOS = {
    CAMPO_CPF_MAS: ("AW", "BC", "BI"),
    CAMPO_CPF_MENOS: ("AX", "BD", "BJ"),
    CAMPO_CSF_MAS: ("AY", "BE", "BK"),
    CAMPO_CSF_MENOS: ("AZ", "BF", "BL"),
    CAMPO_CTF_MAS: ("BA", "BG", "BM"),
    CAMPO_CTF_MENOS: ("BB", "BH", "BN"),
}

# Clave_Bloque = DIA(A) & "#" & (R - (DIA(A)-1)*96)
LETRA_FECHA_SOBRECOSTOS = "A"
LETRA_PERIODO_SOBRECOSTOS = "R"

# Los 96 de la formula de Clave_Bloque SC. Se replica tal cual el
# Excel; la traz. 6.6.3 advierte que no sirve para dias de 92/100
# periodos (cambio de hora) y que eso se revisa aparte.
PERIODOS_POR_DIA_FORMULA_SC = 96
