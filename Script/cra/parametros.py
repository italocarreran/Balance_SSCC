# -*- coding: utf-8 -*-
"""
Nombres de archivo, carpeta y hoja del caso CRA, y el mapeo de columnas
de cada archivo de entrada.

Todo lo que viene de la trazabilidad
(docs/Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_Formulas.md) esta citado con su
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
HOJA_FD_CPF = "FD_CPF"
HOJA_FD_CSF = "FD_CSF"
HOJA_FD_CTF = "FD_CTF"
HOJA_PRORRATA = "PRORRATA_RETIROS"

# Orden de las hojas en el libro, de fin a inicio (mismo criterio que
# Balance_BESS.xlsx: arriba lo ultimo del calculo, abajo las entradas).
# Por ahora solo estan las entradas con su carga ya definida (de
# PRORRATA_RETIROS, solo el bloque 1: la matriz de prorratas); el
# calculo (CÁLCULO_CRA, cuadros 2 y 3 de PRORRATA_RETIROS, RESUMEN) y
# las entradas pendientes (TC, COTAS, RENDIMIENTOS, CONDICION_EMBALSE,
# EMPRESAS) se agregan cuando se cierren.
ORDEN_HOJAS_SALIDA = [
    HOJA_PRORRATA,
    HOJA_FD_CTF, HOJA_FD_CSF, HOJA_FD_CPF,
    HOJA_SC_CO, HOJA_CO, HOJA_FP, HOJA_ENERGIA,
]


# ============================================================
# CARPETAS DEL CASO -- una subcarpeta por hoja de entrada (confirmado
# por el usuario el 2026-09-30). FD y prorrata de retiros son las MISMAS
# fuentes que usa el BESS; "Prorrata retiros" lleva el mismo nombre que
# en el caso BESS.
# ============================================================

CARPETA_ENERGIA = "Energia"
CARPETA_FP = "FP"
CARPETA_CO = "CO"
CARPETA_SC_CO = "SC y CO"
CARPETA_FD = "FD"
CARPETA_PRORRATA = "Prorrata retiros"
CARPETA_AUXILIARES = "Auxiliares"

SUBCARPETAS_CASO = (
    CARPETA_ENERGIA, CARPETA_FP, CARPETA_CO, CARPETA_SC_CO,
    CARPETA_FD, CARPETA_PRORRATA, CARPETA_AUXILIARES,
)

# Los maestros del CRA, en un libro con tres hojas (ver maestros.py):
# "centrales_cra" (filtra cvar_cra), "empresas" y "diccionario" (que
# unidades del SSCC_Desempeño entran a cada FD_*). Nombre fijo, en
# Auxiliares/.
ARCHIVO_CENTRALES_CRA = "centrales_cra.xlsx"


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

# Reporte_CRA viene en .csv (confirmado por el usuario); se acepta
# tambien en Excel por si alguna vez llega guardado asi.
EXTENSIONES_REPORTE_CRA = (".csv",) + EXTENSIONES_EXCEL

# Traz. 6.1: Formato_Solicitud_SSAA_SSCC_Hidro_MesAAAA.xlsx, con el mes
# en palabras: ..._Agosto2026 (confirmado por el usuario). {mes} es el
# nombre del mes en minusculas y sin tildes; ademas se valida el
# periodo contra las columnas Año/Mes del propio archivo.
PREFIJO_ENERGIA = "formato_solicitud_ssaa_sscc_hidro_{mes}{anio}"

MESES = (
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre",
)
# "Setiembre" tambien se usa en Chile: se acepta cualquiera de las dos.
MESES_ALTERNATIVOS = {9: ("setiembre",)}

# Traz. 6.4.1: fp_AAMM*.xlsx (el usuario confirmo que siempre lleva el
# AAMM).
PREFIJO_FP = "fp_{aamm}"

# Traz. 6.4.2: cvar_cra_AAMM_*.xlsx.
PREFIJO_CO = "cvar_cra_{aamm}"

# Traz. 6.6.2 y 6.6.3.
PREFIJO_REPORTE_CRA = "reporte_cra"
PREFIJO_SOBRECOSTOS = "calculo_sobrecostossscc"


# ============================================================
# HOJA DE CADA ARCHIVO
#
#   "NOMBRE" = esa hoja (se para si no esta).
#   0        = la primera hoja, sea cual sea (confirmado por el usuario
#              para Reporte_CRA, que en Excel se llama
#              Reporte_CRA_15min_AAMM y es la unica).
#   None     = no confirmada: si el libro tiene una sola hoja se usa
#              esa; si tiene varias, se para y las lista.
#
# Un .csv no tiene hojas: este valor no se usa.
# ============================================================

HOJA_ORIGEN_ENERGIA = None
HOJA_ORIGEN_FP = None
HOJA_ORIGEN_CO = None
HOJA_ORIGEN_REPORTE_CRA = 0
HOJA_ORIGEN_SOBRECOSTOS = "SOBRECOSTOS"


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

# "Corregidas" (traz. v5, 6.1): K, L y M son formulas del libro y M
# (Neto) es lo que consume CÁLCULO_CRA. Ver corregir_energia().
CAMPO_KWHD_CORREGIDO = "kWhD corregido"
CAMPO_KWHR_CORREGIDO = "kWhR corregido"
CAMPO_NETO = "Neto"

COLUMNAS_ENERGIA = [
    CAMPO_UNIDAD, CAMPO_PUNTO, CAMPO_ANIO, CAMPO_MES, CAMPO_DIA,
    CAMPO_HORADIA, CAMPO_PERIODO, CAMPO_KWHD, CAMPO_KWHR,
    CAMPO_KWHD_CORREGIDO, CAMPO_KWHR_CORREGIDO, CAMPO_NETO,
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

# Centrales de embalse: SC y CO se filtran a estas (las demas, p. ej.
# los parques eolicos PE-*, se descartan). Van con el nombre EXACTO del
# origen (Reporte_CRA!H, SOBRECOSTOS!U); se comparan normalizados.
# Sale de la lista del Actualiza_SC_CO.py del usuario (2026-09-30).
# Si se agrega una unidad de embalse nueva, se agrega aca: una central
# que termina en "-numero" y no esta en la lista se avisa, por si es eso.
CENTRALES_EMBALSE = [
    "CANUTILLAR-1", "CANUTILLAR-2",
    "ELTORO-1", "ELTORO-2", "ELTORO-3", "ELTORO-4",
    "RALCO-1", "RALCO-2",
    "RAPEL-1", "RAPEL-2", "RAPEL-3", "RAPEL-4", "RAPEL-5",
    "PEHUENCHE-1", "PEHUENCHE-2",
    "COLBUN-1", "COLBUN-2",
    "CIPRESES-1", "CIPRESES-2", "CIPRESES-3",
    "PANGUE-1", "PANGUE-2",
    "ANTUCO-1", "ANTUCO-2",
    "ANGOSTURA-1", "ANGOSTURA-2", "ANGOSTURA-3",
]

# Los 96 de la formula de Clave_Bloque SC. Se replica tal cual el
# Excel; la traz. 6.6.3 advierte que no sirve para dias de 92/100
# periodos (cambio de hora) y que eso se revisa aparte.
PERIODOS_POR_DIA_FORMULA_SC = 96


# ============================================================
# FD_CPF, FD_CSF, FD_CTF -- del SSCC_Desempeño_* del DCO, la misma
# fuente del BESS (confirmado por el usuario). Se copian las tres hojas
# horarias con los rangos que ya usa Script/Fd/Desempeno_Horario.py:
# encabezados en la fila 11, datos desde la 12.
# ============================================================

FILA_ENCABEZADO_FD = 11

# Clave de las hojas FD del libro (traz. v5, 6.7):
#   A = D & (B + TIMEVALUE(C & " :00"))  -> Unidad + Fecha/Hora horaria.
# Se escribe como una columna "Fecha Hora" (Fecha + Hora horas) al
# principio de la hoja; la Unidad ya esta. El texto pegado del Excel (el
# numero de serie de la fecha como texto) no se reproduce: el cruce se
# hace por (Unidad, Fecha Hora).
CAMPO_FECHA_HORA_FD = "Fecha Hora"
LETRA_FECHA_FD = "B"
LETRA_HORA_FD = "C"
# La unidad (InfoTecnica): por esta columna se filtra contra la hoja
# "diccionario" de centrales_cra.xlsx (FD_CPF / FD_CSF / FD_CTF).
LETRA_UNIDAD_FD = "D"

HOJAS_FD = {
    # seccion: (hoja de salida, hoja del SSCC_Desempeño, columnas)
    "fd_cpf": (HOJA_FD_CPF, "CPF Horario", "BCDEFGHIJ"),
    "fd_csf": (HOJA_FD_CSF, "CSF Horario", "BCDEFGH"),
    "fd_ctf": (HOJA_FD_CTF, "CTF Horario", "BCDEFGHI"),
}


# ============================================================
# PRORRATA_RETIROS, bloque 1 (traz. 8.1): la matriz periodo x empresa
# pagadora. Sale de Prorrata_Retiros_<AAMM>_pre/def.xlsx, hoja
# "Prorrata 15min", la misma fuente del BESS (se usa su lector).
# ============================================================

COLUMNA_PERIODO_PRORRATA = "Cuarto de Hora"
PERIODOS_POR_DIA = 96
