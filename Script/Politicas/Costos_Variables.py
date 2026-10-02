# -*- coding: utf-8 -*-
"""
Costos_Variables — arma cvar_cra_<AAMM>_1_<N>.xlsx y fp_<AAMM>_1_<N>.xlsx.

Es el paso "costo_variable" de entradas_sscc.py (autor original:
Gerardo.Vieyra), reescrito para el proyecto. De ese script se tomaron
las rutas, los nombres de archivo y la logica; no el resto (subastas,
cotas, TC, CMg, FMA, RIO, que son otras cosas).

Por cada dia del mes:

  1. Costo variable, de la politica del dia:
         <RAIZ_POLITICAS>/<AA>/PO<AAMMDD>.csv
     Primera columna = configuracion (su encabezado es la fecha
     'DD-MM-AAAA'), una columna por hora (1..24) -> se "despivotea" a
     (Configuracion, hora, Cvar, dia) y se descartan las celdas vacias.

  2. Reprogramaciones (PID), hora 1..23: si para esa hora existe el
     programa
         <RAIZ_PRG>/<AAMM>/Politicas/PRG<AAMMDD>_<HH>.xlsx
     se lee la politica reprogramada
         <RAIZ_PID>/20AA/PID_20AAMM/PID_20AAMMDD/Publicacion/
             PID_CDC_<HH>/PO<AAMMDD>_<HH>.csv
     y reemplaza las horas >= HH. Como se recorre de la 1 a la 23, cada
     hora queda con la ULTIMA reprogramacion publicada que la cubre.
     Si hay programa pero no politica, se avisa y se sigue.

  3. Factores de penalizacion, del mismo dia:
         <RAIZ_POLITICAS>/<AA>/PO<AAMMDD>.xlsx
     tercera hoja, encabezado en la fila 2, columnas B:Z (BarNom + una
     columna por hora) -> (BarNom, Hora, FP, dia). Los FP salen siempre
     de la politica del dia, sin PID (igual que el script original).

Al final, el costo variable se filtra a las configuraciones del CRA
(hoja "centrales_cra" del maestro Auxiliares/centrales_cra.xlsx, que lee
Script/cra/maestros.py y llega aca como lista) -> cvar_cra.

No importa nada de nucleo (solo pandas). Los errores previsibles salen
como ErrorPoliticas.
"""

import calendar
from pathlib import Path

import pandas as pd


# ============================================================
# DONDE ESTA CADA COSA (rutas de entradas_sscc.py)
# ============================================================

RAIZ_POLITICAS = r"\\nas-cen1\Estadisticas\progdiar_SEN"
RAIZ_PID = r"\\nas-cen1\DPID\1-Productivo_PID"
RAIZ_PRG = r"\\nas-cen1\D. Transferencias\CMgReales"

# Hojas del xlsx de la politica: los FP estan en la tercera.
HOJA_FACTORES = 2
FILA_ENCABEZADO_FACTORES = 1        # 0-indexada: fila 2 de Excel
COLUMNAS_FACTORES = "B:Z"

HORAS_PID = range(1, 24)

COLUMNA_CONFIGURACION = "Configuracion"
COLUMNAS_CVAR = [COLUMNA_CONFIGURACION, "hora", "Cvar", "dia"]
COLUMNAS_FP = ["BarNom", "Hora", "FP", "dia"]


class ErrorPoliticas(Exception):
    """Error previsible al armar cvar_cra / fp."""


# ============================================================
# NOMBRES Y RUTAS
# ============================================================

def _partes(aamm):
    aamm = str(aamm or "").strip()

    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        raise ErrorPoliticas(f"Periodo invalido: '{aamm}' (se espera AAMM).")

    return aamm[:2], aamm[2:]


def dias_del_mes(aamm):
    aa, mm = _partes(aamm)
    return calendar.monthrange(2000 + int(aa), int(mm))[1]


def nombre_salida(prefijo, aamm):
    """('fp', '2603') -> 'fp_2603_1_31.xlsx' (el nombre de siempre)."""

    _partes(aamm)
    return f"{prefijo}_{aamm}_1_{dias_del_mes(aamm)}.xlsx"


def carpeta_politicas(aamm, raiz=None):
    # El script original arma <raiz>/<año> con el año de DOS digitos.
    aa, _ = _partes(aamm)
    return Path(raiz or RAIZ_POLITICAS) / aa


def ruta_politica(aamm, dia, extension, raiz=None):
    return carpeta_politicas(aamm, raiz) / f"PO{aamm}{dia:02d}.{extension}"


def ruta_programa_pid(aamm, dia, hora, raiz=None):
    return (
        Path(raiz or RAIZ_PRG) / aamm / "Politicas"
        / f"PRG{aamm}{dia:02d}_{hora:02d}.xlsx"
    )


def ruta_politica_pid(aamm, dia, hora, raiz=None):
    aa, mm = _partes(aamm)
    return (
        Path(raiz or RAIZ_PID) / f"20{aa}" / f"PID_20{aa}{mm}"
        / f"PID_20{aa}{mm}{dia:02d}" / "Publicacion" / f"PID_CDC_{hora:02d}"
        / f"PO{aamm}{dia:02d}_{hora:02d}.csv"
    )


# ============================================================
# LECTURA
# ============================================================

def leer_cvar(ruta, dia):
    """PO*.csv -> (Configuracion, hora, Cvar, dia)."""

    try:
        crudo = pd.read_csv(ruta)
    except Exception as error:
        raise ErrorPoliticas(f"No se pudo leer {ruta}: {error}") from error

    # La primera columna es la configuracion; su encabezado es la fecha
    # del dia, asi que se toma por posicion y no por nombre.
    id_col = crudo.columns[0]
    largo = crudo.melt(id_vars=[id_col], var_name="hora", value_name="Cvar")
    largo = largo.dropna()
    largo.columns = [COLUMNA_CONFIGURACION, "hora", "Cvar"]
    largo["dia"] = dia

    try:
        largo["hora"] = largo["hora"].astype(int)
    except ValueError as error:
        raise ErrorPoliticas(
            f"{Path(ruta).name}: hay columnas de hora que no son numeros "
            f"({sorted(set(largo['hora']))[:5]})."
        ) from error

    return largo[COLUMNAS_CVAR].reset_index(drop=True)


def leer_fp(ruta, dia):
    """PO*.xlsx, tercera hoja, B:Z -> (BarNom, Hora, FP, dia)."""

    try:
        crudo = pd.read_excel(
            ruta, sheet_name=HOJA_FACTORES, header=FILA_ENCABEZADO_FACTORES,
            usecols=COLUMNAS_FACTORES, engine="openpyxl",
        )
    except Exception as error:
        raise ErrorPoliticas(
            f"No se pudo leer la hoja de factores de {ruta}: {error}"
        ) from error

    # El encabezado real es 'BarNom' con espacios atras: por posicion.
    id_col = crudo.columns[0]
    largo = crudo.melt(id_vars=[id_col], var_name="Hora", value_name="FP")
    largo.columns = ["BarNom", "Hora", "FP"]
    largo["dia"] = dia

    return largo[COLUMNAS_FP].reset_index(drop=True)


def aplicar_pid(cvar_dia, aamm, dia, raiz_pid=None, raiz_prg=None,
                registrar=print):
    """
    Reemplaza las horas reprogramadas. Devuelve (cvar_dia, horas usadas).
    """

    usadas = []

    for hora in HORAS_PID:
        if not ruta_programa_pid(aamm, dia, hora, raiz_prg).is_file():
            continue

        politica = ruta_politica_pid(aamm, dia, hora, raiz_pid)
        if not politica.is_file():
            registrar(
                f"  AVISO dia {dia}: hay programa PID de la hora {hora} "
                f"pero no su politica ({politica.name}); se sigue con la "
                f"anterior."
            )
            continue

        nueva = leer_cvar(politica, dia)
        cvar_dia = pd.concat(
            [cvar_dia[cvar_dia["hora"] < hora], nueva[nueva["hora"] >= hora]],
            ignore_index=True,
        )
        usadas.append(hora)

    return cvar_dia, usadas


# ============================================================
# EL PROCESO
# ============================================================

def construir(aamm, que=("cvar", "fp"), configuraciones=None,
              raiz_politicas=None, raiz_pid=None, raiz_prg=None,
              registrar=print, progreso=None):
    """
    Lee el mes completo y devuelve {"cvar": DataFrame, "fp": DataFrame}
    con lo pedido en `que`. "cvar" ya viene filtrado a `configuraciones`
    (la lista de configuraciones del CRA; obligatoria para "cvar").

    Si falta el archivo de la politica de algun dia, se para y lista
    todos los que faltan (no arma un mes incompleto).
    """

    que = set(que)
    dias = range(1, dias_del_mes(aamm) + 1)

    extensiones = {"cvar": "csv", "fp": "xlsx"}
    faltan = [
        str(ruta_politica(aamm, dia, extensiones[cual], raiz_politicas))
        for cual in sorted(que) for dia in dias
        if not ruta_politica(aamm, dia, extensiones[cual], raiz_politicas).is_file()
    ]
    if faltan:
        raise ErrorPoliticas(
            f"Faltan {len(faltan)} archivo(s) de politica del periodo "
            f"{aamm} (¿unidad de red conectada?):\n"
            + "\n".join(f"  - {r}" for r in faltan[:10])
            + ("\n  ..." if len(faltan) > 10 else "")
        )

    if "cvar" in que:
        if not configuraciones:
            raise ErrorPoliticas(
                "Faltan las configuraciones del CRA para filtrar el costo "
                "variable."
            )
        configuraciones = pd.DataFrame(
            {COLUMNA_CONFIGURACION: list(dict.fromkeys(configuraciones))}
        )

    partes = {"cvar": [], "fp": []}

    for i, dia in enumerate(dias):
        if "cvar" in que:
            cvar_dia = leer_cvar(
                ruta_politica(aamm, dia, "csv", raiz_politicas), dia
            )
            cvar_dia, usadas = aplicar_pid(
                cvar_dia, aamm, dia, raiz_pid, raiz_prg, registrar
            )
            partes["cvar"].append(cvar_dia)
            registrar(
                f"  dia {dia}: costo variable"
                + (f" (PID horas {', '.join(map(str, usadas))})" if usadas else "")
            )

        if "fp" in que:
            partes["fp"].append(
                leer_fp(ruta_politica(aamm, dia, "xlsx", raiz_politicas), dia)
            )
            registrar(f"  dia {dia}: factores de penalizacion")

        if progreso:
            progreso(int(90 * (i + 1) / len(dias)))

    salida = {}

    if "fp" in que:
        salida["fp"] = pd.concat(partes["fp"], ignore_index=True)

    if "cvar" in que:
        cvar = pd.concat(partes["cvar"], ignore_index=True)
        cvar["hora"] = cvar["hora"].astype(int)
        cvar_cra = cvar.merge(
            configuraciones, on=COLUMNA_CONFIGURACION, how="inner"
        ).drop_duplicates()

        sin_datos = sorted(
            set(configuraciones[COLUMNA_CONFIGURACION].dropna())
            - set(cvar_cra[COLUMNA_CONFIGURACION])
        )
        if sin_datos:
            registrar(
                f"  AVISO: {len(sin_datos)} configuracion(es) de "
                f"centrales_cra sin costo variable en el mes: "
                f"{', '.join(map(str, sin_datos[:10]))}"
                + (" ..." if len(sin_datos) > 10 else "")
            )
        salida["cvar"] = cvar_cra.reset_index(drop=True)

    return salida
