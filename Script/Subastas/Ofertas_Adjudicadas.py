# -*- coding: utf-8 -*-
"""
Ofertas_Adjudicadas — las subastas leidas desde su origen real, los
Access OfertasSSCCAdj*.accdb.

Viene de la rutina `calc_subastas` del script suelto entradas_sscc.py
(autor original: Gerardo.Vieyra), que se corria a mano con un
archivo_de_configuracion.yaml al lado. Cambios respecto de ese script,
todos pedidos por el usuario:

  1) los .accdb no se leen desde la unidad de red: se copian primero a
     <CARPETA_BASE>/Subastas/DB subastas/ con el boton "Traer
     subastas" y de ahi se leen (mismo patron que el CSV de CMg);
  2) el periodo no sale de un yaml: es el AAMM que el usuario escribe
     en la ventana;
  3) no se escribe ningun Excel aca: este modulo devuelve DataFrames y
     es nucleo el que resuelve rutas y escribe.

No importa nada de nucleo (solo pandas): asi este modulo se puede
probar y correr suelto, igual que Script/Cmg/Extrae_CMG_barras.py. Los
errores previsibles salen como ErrorSubastas; nucleo los traduce a
su ErrorEntrada.

pyodbc se importa DENTRO de las funciones que leen, no arriba: es la
unica dependencia del proyecto que solo existe en Windows (necesita el
"Microsoft Access Database Engine"), y no tiene por que romper el
import de todo el programa en una maquina donde no este instalada.
"""

import calendar
import re
import shutil
from pathlib import Path

import pandas as pd


# ============================================================
# DONDE VIVEN LOS ACCESS DE SUBASTAS
#
#   \\nas-cen1\Estadisticas\progdiar_adjudicaSEN\
#       OfertasSSCCAdjAAAAMMDD.accdb        <- el del dia (PO)
#       OfertasSSCCAdjAAAAMMDD_HH.accdb     <- rehecho en la hora HH (PID)
#
# Es la segunda ruta del programa que apunta fuera de la carpeta base
# del caso (la otra es la del CSV de CMg): si el servidor cambia, se
# cambia aca y nada mas. Sale de `path_subastas_origen` del
# archivo_de_configuracion.yaml de entradas_sscc.py.
# ============================================================

RAIZ_SUBASTAS_ORIGEN = r"\\nas-cen1\Estadisticas\progdiar_adjudicaSEN"

# `archivo_subasta_prefix` del mismo yaml.
PREFIJO_ACCDB = "OfertasSSCCAdj"
EXTENSION_ACCDB = ".accdb"

# Subcarpeta de <CARPETA_BASE>/Subastas/ donde se guardan los .accdb
# copiados. El nombre lo eligio el usuario (con espacio, tal cual).
CARPETA_DB_SUBASTAS = "DB subastas"

# Las horas en que puede haber un PID (re-subasta). El archivo de la
# hora 0 es el del PO del dia y va sin sufijo.
HORAS_PID = range(0, 24)

# La consulta es la misma de entradas_sscc.py, sin tocar: une la
# tabla de datos (AASS_Data) con el nombre de la configuracion
# (Config_List.PO_Name) y el del servicio (AASS_List.NemoTecnico).
SQL_SUBASTAS = (
    "SELECT Config_List.PO_Name AS CONFIGURACIÓN, "
    "AASS_List.NemoTecnico AS SERVICIO, "
    "Year(AASS_Data.DateTime) AS AÑO, "
    "Month(AASS_Data.DateTime) AS MES, "
    "Day(AASS_Data.DateTime) AS DIA, "
    "AASS_Data.Period AS HORA, "
    "AASS_Data.[Band] AS BANDA, "
    "AASS_Data.Quantity AS [CANTIDAD MW], "
    "AASS_Data.Price AS [PRECIO USD/MW], "
    "AASS_Data.Quantity2 AS [CANTIDAD PONDERADA MW] "
    "FROM Config_List INNER JOIN (AASS_List INNER JOIN AASS_Data "
    "ON AASS_List.Id = AASS_Data.ServiceId) "
    "ON Config_List.Id = AASS_Data.ConfigId;"
)

CONTROLADOR_ACCESS = "Microsoft Access Driver (*.mdb, *.accdb)"

# Las 10 que trae la consulta, mas Hora_PID que se agrega al leer.
COLUMNAS_CRUDAS = [
    "CONFIGURACIÓN",
    "SERVICIO",
    "AÑO",
    "MES",
    "DIA",
    "HORA",
    "BANDA",
    "CANTIDAD MW",
    "PRECIO USD/MW",
    "CANTIDAD PONDERADA MW",
    "Hora_PID",
]

_PATRON_AAMM = re.compile(r"^\d{4}$")


class ErrorSubastas(Exception):
    """Error previsible al traer o leer los Access de subastas."""


# ============================================================
# NOMBRES DE ARCHIVO
# ============================================================

def _validar_aamm(aamm):
    aamm = str(aamm or "").strip()

    if not _PATRON_AAMM.match(aamm):
        raise ErrorSubastas(
            f"Periodo invalido: '{aamm}'. Tienen que ser 4 digitos "
            f"(por ejemplo 2603 para marzo de 2026)."
        )

    return aamm


def periodo_desde_aamm(aamm):
    """'2603' -> (2026, 3). Mismo criterio que nucleo."""

    aamm = _validar_aamm(aamm)
    anio = 2000 + int(aamm[:2])
    mes = int(aamm[2:])

    if not 1 <= mes <= 12:
        raise ErrorSubastas(f"El periodo '{aamm}' no tiene un mes valido.")

    return anio, mes


def nombre_accdb(aamm, dia, hora=0):
    """
    (2603, 1, 0)  -> 'OfertasSSCCAdj20260301.accdb'     (PO del dia)
    (2603, 1, 14) -> 'OfertasSSCCAdj20260301_14.accdb'  (PID de las 14)

    El sufijo de hora sale igual que en entradas_sscc.py: la hora 0 no
    lleva sufijo y el resto va con dos digitos.
    """

    anio, mes = periodo_desde_aamm(aamm)
    sufijo = "" if hora == 0 else f"_{hora:02d}"

    return f"{PREFIJO_ACCDB}{anio}{mes:02d}{dia:02d}{sufijo}{EXTENSION_ACCDB}"


def nombres_del_periodo(aamm):
    """
    Todos los nombres de archivo POSIBLES del mes, en orden
    (dia, hora): los del PO y los de cada PID. Son hasta 31 x 24; la
    enorme mayoria no existe y se saltea sin ruido.

    Devuelve una lista de tuplas (dia, hora, nombre).
    """

    anio, mes = periodo_desde_aamm(aamm)
    dias_del_mes = calendar.monthrange(anio, mes)[1]

    return [
        (dia, hora, nombre_accdb(aamm, dia, hora))
        for dia in range(1, dias_del_mes + 1)
        for hora in HORAS_PID
    ]


def ruta_origen(raiz=None):
    """Carpeta de red de donde se copian los .accdb."""

    return Path(raiz or RAIZ_SUBASTAS_ORIGEN)


def carpeta_db_subastas(carpeta_subastas):
    """<CARPETA_BASE>/Subastas/ -> <...>/Subastas/DB subastas/"""

    return Path(carpeta_subastas) / CARPETA_DB_SUBASTAS


def asegurar_carpeta_db(carpeta_subastas):
    """
    Crea <...>/Subastas/DB subastas/ si no existe (el usuario pidio
    que la cree el programa, no la persona) y devuelve su ruta. Si no
    se puede crear, no revienta: devuelve la ruta igual y quien llama
    decide -- asi un permiso raro no tumba el diagrama entero.
    """

    carpeta = carpeta_db_subastas(carpeta_subastas)

    try:
        carpeta.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

    return carpeta


def accdb_presentes(carpeta_subastas, aamm):
    """
    Los .accdb del periodo que YA estan copiados en 'DB subastas'.
    Lista de tuplas (dia, hora, ruta), ordenada por (dia, hora).
    """

    carpeta = carpeta_db_subastas(carpeta_subastas)

    if not carpeta.is_dir():
        return []

    presentes = []

    for dia, hora, nombre in nombres_del_periodo(aamm):
        ruta = carpeta / nombre
        if ruta.is_file():
            presentes.append((dia, hora, ruta))

    return presentes


# ============================================================
# TRAER (copiar de la red a la carpeta del caso)
# ============================================================

def traer_accdb(carpeta_subastas, aamm, raiz=None, registrar=print):
    """
    Copia a <CARPETA_BASE>/Subastas/DB subastas/ todos los .accdb del
    periodo que existan en la carpeta de red (boton "Traer subastas").

    Un archivo que ya esta copiado y no cambio (mismo tamaño y misma
    fecha de modificacion) se saltea: son hasta 744 archivos por mes y
    volver a copiarlos todos cada vez no tiene sentido.

    Devuelve (copiados, salteados, faltantes_en_origen).
    """

    aamm = _validar_aamm(aamm)
    origen = ruta_origen(raiz)

    if not origen.is_dir():
        raise ErrorSubastas(
            f"No se puede leer la carpeta de origen de las subastas:\n"
            f"  {origen}\n\n"
            f"Revisa que tengas conexion y permiso sobre ese servidor. "
            f"Si la ruta cambio, se cambia en RAIZ_SUBASTAS_ORIGEN "
            f"(Script/Subastas/Ofertas_Adjudicadas.py)."
        )

    destino = asegurar_carpeta_db(carpeta_subastas)

    if not destino.is_dir():
        raise ErrorSubastas(f"No se pudo crear la carpeta {destino}")

    copiados, salteados, faltantes = [], [], 0

    for _, _, nombre in nombres_del_periodo(aamm):

        ruta_origen_archivo = origen / nombre

        if not ruta_origen_archivo.is_file():
            faltantes += 1
            continue

        ruta_destino = destino / nombre

        if _ya_esta_copiado(ruta_origen_archivo, ruta_destino):
            salteados.append(nombre)
            continue

        registrar(f"  copiando {nombre}...")

        try:
            shutil.copy2(ruta_origen_archivo, ruta_destino)
        except OSError as error:
            raise ErrorSubastas(
                f"No se pudo copiar {nombre}: {error}"
            ) from error

        copiados.append(nombre)

    if not copiados and not salteados:
        raise ErrorSubastas(
            f"No se encontro ningun {PREFIJO_ACCDB}* del periodo {aamm} "
            f"en {origen}. Revisa que el periodo sea el correcto."
        )

    registrar(
        f"  {len(copiados)} archivo(s) copiado(s), "
        f"{len(salteados)} ya estaban al dia."
    )

    return copiados, salteados, faltantes


def _ya_esta_copiado(origen, destino):
    """Mismo tamaño y misma fecha de modificacion = no hay que copiar."""

    if not destino.is_file():
        return False

    info_origen = origen.stat()
    info_destino = destino.stat()

    return (
        info_origen.st_size == info_destino.st_size
        and int(info_origen.st_mtime) == int(info_destino.st_mtime)
    )


# ============================================================
# LEER (los .accdb ya copiados)
# ============================================================

def _conectar(ruta_accdb):
    """Abre un .accdb con el driver ODBC de Access."""

    try:
        import pyodbc
    except ImportError as error:
        raise ErrorSubastas(
            "Falta el paquete pyodbc, que es el que sabe leer los "
            "Access de las subastas.\n\n"
            "Se instala con:  pip install pyodbc\n"
            "y ademas necesita el 'Microsoft Access Database Engine' "
            "instalado en el equipo (Windows)."
        ) from error

    cadena = (
        "DRIVER={" + CONTROLADOR_ACCESS + "};"
        "DBQ=" + str(ruta_accdb) + ";"
    )

    try:
        return pyodbc.connect(cadena)
    except Exception as error:
        raise ErrorSubastas(
            f"No se pudo abrir {Path(ruta_accdb).name}.\n\n"
            f"Verifica que este instalado el 'Microsoft Access Database "
            f"Engine' (el driver '{CONTROLADOR_ACCESS}') y que la "
            f"arquitectura (32/64 bits) coincida con la de Python.\n\n"
            f"Detalle: {error}"
        ) from error


def leer_accdb(ruta_accdb):
    """Un .accdb -> DataFrame con las 10 columnas de SQL_SUBASTAS."""

    conexion = _conectar(ruta_accdb)

    try:
        return pd.read_sql_query(SQL_SUBASTAS, conexion)
    except Exception as error:
        raise ErrorSubastas(
            f"No se pudo leer la consulta de subastas en "
            f"{Path(ruta_accdb).name}: {error}"
        ) from error
    finally:
        conexion.close()


def construir_crudo(carpeta_subastas, aamm, registrar=print):
    """
    Recorre los .accdb del periodo que estan en 'DB subastas' y arma
    el equivalente exacto de subastas_AAMM.xlsx (la salida de
    entradas_sscc.py), con las mismas reglas:

      - por cada dia se parte del archivo del PO (hora 0) y despues,
        por cada PID de la hora HH que exista, se REEMPLAZAN las horas
        >= HH: lo anterior a HH queda como estaba y lo de HH en
        adelante pasa a ser lo que dice el PID. Queda registrado en
        'Hora_PID' de que archivo salio cada fila;
      - se descartan las filas con CANTIDAD MW = 0;
      - el precio solo vale en la banda 1: en las demas se pone 0;
      - CANTIDAD PONDERADA MW vacia se completa con CANTIDAD MW;
      - se eliminan filas duplicadas.

    Devuelve (df_crudo, resumen), donde resumen es un dict con la
    cuenta de archivos leidos y los dias sin archivo.
    """

    aamm = _validar_aamm(aamm)
    archivos = accdb_presentes(carpeta_subastas, aamm)

    if not archivos:
        raise ErrorSubastas(
            f"No hay ningun {PREFIJO_ACCDB}* del periodo {aamm} en "
            f"{carpeta_db_subastas(carpeta_subastas)}.\n\n"
            f"Usa el boton 'Traer subastas' para copiarlos desde la "
            f"carpeta de red."
        )

    por_dia = {}
    for dia, hora, ruta in archivos:
        por_dia.setdefault(dia, []).append((hora, ruta))

    partes = []
    leidos = 0

    for dia in sorted(por_dia):

        df_dia = pd.DataFrame()

        for hora, ruta in sorted(por_dia[dia]):

            registrar(f"  leyendo {ruta.name}...")
            df_archivo = leer_accdb(ruta)
            leidos += 1

            # El PID de la hora HH manda de HH en adelante; lo previo
            # queda como lo dejo el archivo anterior.
            if hora > 0 and not df_dia.empty:
                df_dia = df_dia[df_dia["HORA"] < hora]

            df_archivo = df_archivo[df_archivo["HORA"] >= hora].copy()
            df_archivo["Hora_PID"] = hora

            df_dia = pd.concat([df_dia, df_archivo], ignore_index=True)

        if not df_dia.empty:
            partes.append(df_dia)

    if not partes:
        raise ErrorSubastas(
            f"Los {leidos} archivo(s) de subastas del periodo {aamm} "
            f"no trajeron ninguna fila."
        )

    df = pd.concat(partes, ignore_index=True)

    df = df[df["CANTIDAD MW"] != 0]

    # El precio adjudicado solo aplica a la banda 1 (entradas_sscc.py).
    df["PRECIO USD/MW"] = df["PRECIO USD/MW"].where(df["BANDA"] == 1, 0)

    df["CANTIDAD PONDERADA MW"] = df["CANTIDAD PONDERADA MW"].fillna(
        df["CANTIDAD MW"]
    )

    df = df.drop_duplicates().reset_index(drop=True)
    df = df[COLUMNAS_CRUDAS]

    dias_esperados = {
        dia for dia, _, _ in nombres_del_periodo(aamm)
    }
    dias_sin_archivo = sorted(dias_esperados - set(por_dia))

    resumen = {
        "archivos_leidos": leidos,
        "dias_con_archivo": len(por_dia),
        "dias_sin_archivo": dias_sin_archivo,
        "filas": len(df),
    }

    registrar(
        f"  {leidos} Access leido(s), {len(por_dia)} dia(s) con datos, "
        f"{len(df):,} fila(s)."
    )

    return df, resumen
