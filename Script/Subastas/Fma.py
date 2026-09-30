# -*- coding: utf-8 -*-
"""
Fma — las tres salidas de FMA normalizadas, para alimentar Subastas!FMA.

En la planilla 3 la columna FMA (`DB!V`) es una formula larga que, segun
el Concepto de la fila (CPF/CSF/CTF, + o -), busca en una de tres hojas
(`FMA_CPF`, `FMA_CSF`, `FMA_CTF`). Esas tres hojas no son un origen: se
arman pegando las salidas de entradas_sscc.py

    fma_cpf_AAMM.xlsx
    fma_csf_AAMM.xlsx
    fma_cft_AAMM.xlsx  (o .csv -- "cft" es como lo escribe el script)

que el usuario guarda en <CARPETA_BASE>/FD y FMA/. Este modulo lee esos
tres archivos y devuelve tablas ya normalizadas; el cruce contra cada
fila de Subastas lo hace nucleo (calcular_fma_subastas).

Todo lo que hace aca sale del documento de trazabilidad que entrego el
usuario (Trazabilidad FMA -> columna DB!V de Planilla 3), no de una
suposicion:

  - CPF: Suma_Hace_CPF = suma de las 7 columnas de horas; si es >= 0.98
    pasa a ser 1. FMA CPF(+) = Suma_Hace_CPF * Tiempo f<49.975 y
    FMA CPF(-) = Suma_Hace_CPF * Tiempo f>50.025. Esos dos tiempos YA
    vienen como fraccion (0,33 = 33%): no se vuelven a dividir por 100.
  - CSF: la hora del archivo va de 0 a 23 y la de las subastas de 1 a
    24, asi que Hora_DB = "Hora Dia" + 1. Y se usan las columnas "-m" /
    "+m" (FMA CSF-m [%] y FMA CSF+m [%]) divididas por 100, NO las
    variantes sin "m".
  - CTF: la duracion de cada activacion es (tfin - t0) en horas, se
    reparte segun el signo de `variacion`, y las activaciones que caen
    en la misma hora se SUMAN. Se arman dos resumenes, uno por `unidad`
    y otro por `Configuracion`, porque la formula original busca en uno
    y despues en el otro.

No importa nada de nucleo (solo pandas), igual que los otros modulos
de Script/. Los errores previsibles salen como ErrorFma.
"""

import re
import unicodedata
from pathlib import Path

import pandas as pd


# ============================================================
# LOS TRES ARCHIVOS
# ============================================================

# Prefijos con los que entradas_sscc.py nombra sus salidas. El de CTF
# esta escrito "cft" en el script original (esta al reves); se aceptan
# los dos para no depender de que lo arreglen o no.
PREFIJOS_FMA = {
    "cpf": ("fma_cpf",),
    "csf": ("fma_csf",),
    "ctf": ("fma_cft", "fma_ctf"),
}

EXTENSIONES_FMA = (".xlsx", ".xlsm", ".xls", ".csv")

# Columnas de fma_cpf_AAMM.xlsx que se suman para "Suma Hace CPF".
COLUMNAS_SUMA_CPF = (
    "Hace CPF [hrs]",
    "Paráms. Fuera de rango",
    "Hace CPF con lím superior >Pmax",
    "Hace CPF con lím inferior <MT",
    "No hace CPF [hrs]",
    "Indef. [hrs]",
    "Otro [hrs]",
)

# Si la suma llega a este valor, la planilla la lleva a 1.
UMBRAL_SUMA_CPF = 0.98

COLUMNA_CPF_MAS = "Tiempo f<49.975 [%]"
COLUMNA_CPF_MENOS = "Tiempo f>50.025 [%]"

# Las dos columnas "-m"/"+m" de fma_csf (no las variantes sin "m").
COLUMNA_CSF_MENOS = "FMA CSF-m [%]"
COLUMNA_CSF_MAS = "FMA CSF+m [%]"


class ErrorFma(Exception):
    """Error previsible al leer o normalizar las salidas de FMA."""


def _normalizar(texto):
    """Minusculas, sin tildes, espacios colapsados (como en nucleo)."""

    if texto is None:
        return ""

    try:
        if pd.isna(texto):
            return ""
    except (TypeError, ValueError):
        pass

    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return re.sub(r"\s+", " ", texto).strip().lower()


def buscar_archivos_fma(carpeta, aamm):
    """
    Busca en la carpeta "FD y FMA" los tres archivos del periodo.

    No exige el nombre literal `fma_cpf_2603.xlsx`: alcanza con que el
    nombre empiece con el prefijo y contenga el AAMM (mismo criterio
    tolerante que se usa para el SoC). Si hay varios candidatos, gana
    el mas reciente.

    Devuelve un dict {"cpf": ruta|None, "csf": ruta|None, "ctf": ruta|None}.
    """

    carpeta = Path(carpeta)
    encontrados = {tipo: None for tipo in PREFIJOS_FMA}

    if not carpeta.is_dir():
        return encontrados

    archivos = [a for a in carpeta.iterdir() if a.is_file()]

    for tipo, prefijos in PREFIJOS_FMA.items():

        candidatos = [
            a for a in archivos
            if a.suffix.lower() in EXTENSIONES_FMA
            and any(_normalizar(a.stem).startswith(p) for p in prefijos)
            and str(aamm) in a.stem
        ]

        if candidatos:
            # Si estan el .xlsx y el .csv del mismo archivo (el CTF se
            # escribe en los dos formatos), gana el Excel: asi dos
            # corridas leen siempre el mismo y no el que quedo ultimo.
            encontrados[tipo] = max(
                candidatos,
                key=lambda a: (a.suffix.lower() != ".csv", a.stat().st_mtime),
            )

    return encontrados


def _leer(ruta):
    """
    Lee un .xlsx/.csv de FMA y le saca la columna indice que pandas
    deja al escribir con to_excel() (aparece como 'Unnamed: 0').
    """

    ruta = Path(ruta)

    try:
        if ruta.suffix.lower() == ".csv":
            df = pd.read_csv(ruta)
        else:
            df = pd.read_excel(ruta)
    except Exception as error:
        raise ErrorFma(f"No se pudo leer {ruta.name}: {error}") from error

    sobrantes = [
        c for c in df.columns if str(c).startswith("Unnamed:")
    ]

    return df.drop(columns=sobrantes)


def _columna(df, buscada, nombre_archivo):
    """
    Ubica una columna por nombre normalizado (tolera tildes, mayusculas
    y espacios de mas). Levanta ErrorFma con la lista real de columnas
    si no esta: es el error que de verdad se va a ver si el formato de
    la salida cambia.
    """

    objetivo = _normalizar(buscada)

    for columna in df.columns:
        if _normalizar(columna) == objetivo:
            return columna

    raise ErrorFma(
        f"A {nombre_archivo} le falta la columna '{buscada}'.\n\n"
        f"Columnas encontradas: {list(df.columns)}"
    )


def _columna_temporal(df, nombres, nombre_archivo):
    """Igual que _columna pero probando varios nombres alternativos."""

    for nombre in nombres:
        try:
            return _columna(df, nombre, nombre_archivo)
        except ErrorFma:
            continue

    raise ErrorFma(
        f"A {nombre_archivo} le falta una columna de "
        f"{' / '.join(nombres)}.\n\n"
        f"Columnas encontradas: {list(df.columns)}"
    )


# ============================================================
# CPF
# ============================================================

def tabla_cpf(ruta, registrar=print):
    """
    fma_cpf_AAMM.xlsx -> tabla normalizada con
    [Central, Año, Mes, Dia, Hora, FMA_CPF_mas, FMA_CPF_menos].

    La hora de este archivo ya viene en 1..24 (entradas_sscc.py la
    calcula como el indice de fila + 1), igual que la de las subastas:
    aca NO se corre en uno, a diferencia del CSF.
    """

    nombre = Path(ruta).name
    df = _leer(ruta)

    col_central = _columna(df, "Central", nombre)
    col_anio = _columna(df, "Año", nombre)
    col_mes = _columna(df, "Mes", nombre)
    col_dia = _columna_temporal(df, ("Día", "Dia"), nombre)
    col_hora = _columna(df, "Hora", nombre)

    columnas_suma = [_columna(df, c, nombre) for c in COLUMNAS_SUMA_CPF]
    col_mas = _columna(df, COLUMNA_CPF_MAS, nombre)
    col_menos = _columna(df, COLUMNA_CPF_MENOS, nombre)

    numeros = df[columnas_suma].apply(pd.to_numeric, errors="coerce")
    suma = numeros.sum(axis=1)

    # "=SI(SUMA(...)>=0,98; 1; SUMA(...))"
    suma = suma.where(suma < UMBRAL_SUMA_CPF, 1)

    tiempo_mas = pd.to_numeric(df[col_mas], errors="coerce").fillna(0)
    tiempo_menos = pd.to_numeric(df[col_menos], errors="coerce").fillna(0)

    salida = pd.DataFrame({
        "Central": df[col_central].map(_normalizar),
        "Año": pd.to_numeric(df[col_anio], errors="coerce"),
        "Mes": pd.to_numeric(df[col_mes], errors="coerce"),
        "Dia": pd.to_numeric(df[col_dia], errors="coerce"),
        "Hora": pd.to_numeric(df[col_hora], errors="coerce"),
        "FMA_CPF_mas": suma * tiempo_mas,
        "FMA_CPF_menos": suma * tiempo_menos,
    })

    salida = salida.dropna(subset=["Año", "Mes", "Dia", "Hora"])

    registrar(f"  FMA CPF: {len(salida):,} fila(s) de {nombre}")

    return salida


# ============================================================
# CSF
# ============================================================

def tabla_csf(ruta, registrar=print):
    """
    fma_csf_AAMM.xlsx -> tabla normalizada con
    [Año, Mes, Dia, Hora, FMA_CSF_mas_base, FMA_CSF_menos_base].

    Dos cosas que NO son obvias y salen del documento de trazabilidad:

      - la "Hora Día" de este archivo va de 0 a 23 y la de las subastas
        de 1 a 24, asi que se le suma 1;
      - se usan las columnas "-m"/"+m" (FMA CSF-m [%], FMA CSF+m [%]),
        no las que se llaman igual sin la "m".

    Los valores quedan como FRACCION (se dividen por 100). Son la BASE:
    el FMA final de CSF es esta base por el Vector de Participacion CSF,
    que todavia no tenemos (depende de FD -- ver nucleo).
    """

    nombre = Path(ruta).name
    df = _leer(ruta)

    col_anio = _columna(df, "Año", nombre)
    col_mes = _columna(df, "Mes", nombre)
    col_dia = _columna_temporal(df, ("Día", "Dia"), nombre)
    col_hora = _columna_temporal(
        df, ("Hora Día", "Hora Dia", "Hora_dia", "Hora"), nombre
    )
    col_menos = _columna(df, COLUMNA_CSF_MENOS, nombre)
    col_mas = _columna(df, COLUMNA_CSF_MAS, nombre)

    salida = pd.DataFrame({
        "Año": pd.to_numeric(df[col_anio], errors="coerce"),
        "Mes": pd.to_numeric(df[col_mes], errors="coerce"),
        "Dia": pd.to_numeric(df[col_dia], errors="coerce"),
        # 0..23 -> 1..24
        "Hora": pd.to_numeric(df[col_hora], errors="coerce") + 1,
        "FMA_CSF_menos_base": (
            pd.to_numeric(df[col_menos], errors="coerce") / 100
        ),
        "FMA_CSF_mas_base": (
            pd.to_numeric(df[col_mas], errors="coerce") / 100
        ),
    })

    salida = salida.dropna(subset=["Año", "Mes", "Dia", "Hora"])

    registrar(f"  FMA CSF: {len(salida):,} fila(s) de {nombre}")

    return salida


# ============================================================
# CTF
# ============================================================

def tablas_ctf(ruta, registrar=print):
    """
    fma_cft_AAMM.xlsx (o .csv) -> DOS tablas normalizadas, una por
    `unidad` y otra por `Configuracion`, cada una con
    [nombre, Año, Mes, Dia, Hora, FMA_CTF_mas, FMA_CTF_menos,
    Activaciones].

    Son dos porque la formula original busca primero con una
    nomenclatura y despues con la otra. La fecha/hora sale de `t0`
    (hora + 1, para pasar de 0..23 a 1..24) y la duracion de cada
    activacion es (tfin - t0) en horas, repartida segun el signo de
    `variacion`. Varias activaciones en la misma hora se SUMAN.
    """

    nombre = Path(ruta).name
    df = _leer(ruta)

    col_t0 = _columna(df, "t0", nombre)
    col_tfin = _columna(df, "tfin", nombre)
    col_unidad = _columna(df, "unidad", nombre)
    col_config = _columna_temporal(
        df, ("Configuracion", "Configuración"), nombre
    )
    col_variacion = _columna(df, "variacion", nombre)

    t0 = pd.to_datetime(df[col_t0], errors="coerce")
    tfin = pd.to_datetime(df[col_tfin], errors="coerce")

    duracion = (tfin - t0).dt.total_seconds() / 3600
    variacion = pd.to_numeric(df[col_variacion], errors="coerce")

    base = pd.DataFrame({
        "unidad": df[col_unidad].map(_normalizar),
        "configuracion": df[col_config].map(_normalizar),
        "Año": t0.dt.year,
        "Mes": t0.dt.month,
        "Dia": t0.dt.day,
        # 0..23 -> 1..24
        "Hora": t0.dt.hour + 1,
        "FMA_CTF_mas": duracion.where(variacion > 0, 0),
        "FMA_CTF_menos": duracion.where(variacion < 0, 0),
    })

    base = base.dropna(subset=["Año", "Mes", "Dia", "Hora"])

    def resumir(columna_nombre):
        resumen = base.groupby(
            [columna_nombre, "Año", "Mes", "Dia", "Hora"], as_index=False
        ).agg(
            FMA_CTF_mas=("FMA_CTF_mas", "sum"),
            FMA_CTF_menos=("FMA_CTF_menos", "sum"),
            Activaciones=("FMA_CTF_mas", "size"),
        )
        return resumen.rename(columns={columna_nombre: "nombre"})

    por_unidad = resumir("unidad")
    por_configuracion = resumir("configuracion")

    registrar(
        f"  FMA CTF: {len(base):,} activacion(es) de {nombre} -> "
        f"{len(por_unidad):,} hora(s) por unidad"
    )

    return por_unidad, por_configuracion
