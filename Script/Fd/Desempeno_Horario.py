# -*- coding: utf-8 -*-
"""
Desempeno_Horario — el FD por unidad y hora, desde SSCC_Desempeño_*.

De aca salen DOS cosas que hasta ahora estaban pendientes:

  1. **Subastas!FD** (la columna P, antes `DB!Y` de la planilla 3), y
  2. el **Vector de Participacion CSF** (antes `DB!AC`), que multiplica
     al FMA de las filas CSF y era lo unico que faltaba para cerrar
     Subastas!FMA.

Las dos salen del mismo archivo, `SSCC_Desempeño_<Mes>_<AAAA>_V2.xlsx`
-el que ya trae el boton "Traer FD"-, de sus tres hojas horarias. Todo
lo que hace este modulo sale del documento de trazabilidad que entrego
el usuario (Trazabilidad FD -> columna DB!Y de Planilla 3):

  | Hoja          | Columnas | El FD es |
  |---------------|----------|----------|
  | CPF Horario   | B:J      | I (Fd_CPF) |
  | CSF Horario   | B:H      | H (Fd_CSF) |
  | CTF Horario   | B:I      | I (Fd_CTF) |

En las tres, los encabezados estan en la fila 11 y los datos arrancan
en la 12 (mismo criterio que ya usaba nucleo.construir_fd para armar la
hoja FD del consolidado, que lee estas mismas hojas para otra cosa).

La clave de busqueda es **Control + Unidad + Hora_Mes**, y hay tres
detalles que no son obvios:

  - la `Hora` de estas hojas va de **0 a 23**, asi que
    `Hora_Mes = (dia - 1) * 24 + hora + 1` -- que es la misma escala
    1..24 por dia que usa Subastas;
  - **CPF prueba una segunda nomenclatura**: si no encuentra la unidad,
    intercambia el sufijo `TG` <-> `TV` y busca de nuevo;
  - el **Indicador de Participacion CSF** es 0 solo si la unidad figura
    como "No Participó" Y su alternativa TG/TV tampoco participo; en
    cualquier otro caso es 1.

El FD **no se recalcula** a partir de las respuestas: se toma tal cual
viene en el archivo (asi lo pide el documento, seccion 25 -- puede venir
`Respuesta = "No Participó"` con `FD = 1`).

No importa nada de nucleo (solo pandas). Los errores previsibles
salen como ErrorDesempeno.
"""

import re
import unicodedata
from pathlib import Path

import pandas as pd


# ============================================================
# LAS TRES HOJAS
# ============================================================

# Para cada familia: nombre de la hoja, rango de columnas (como indices
# de pandas, 0-indexados y con el fin excluido) y en que posicion DENTRO
# de ese bloque caen la unidad, el FD y la respuesta.
#
#   CPF Horario, B:J -> B=0 Fecha, C=1 Hora, D=2 Unidad, ..., I=7 Fd_CPF
#   CSF Horario, B:H -> B=0 Fecha, C=1 Hora, D=2 Unidad, E=3 Respuesta,
#                       ..., H=6 Fd_CSF
#   CTF Horario, B:I -> B=0 Fecha, C=1 Hora, D=2 Unidad, ..., I=7 Fd_CTF
HOJAS_DESEMPENO = {
    "cpf": {
        "hoja": "CPF Horario",
        "columnas": (1, 10),
        "unidad": 2,
        "fd": 7,
        "respuesta": None,
    },
    "csf": {
        "hoja": "CSF Horario",
        "columnas": (1, 8),
        "unidad": 2,
        "fd": 6,
        "respuesta": 3,
    },
    "ctf": {
        "hoja": "CTF Horario",
        "columnas": (1, 9),
        "unidad": 2,
        "fd": 7,
        "respuesta": None,
    },
}

# Fila 12 de Excel (1-indexada) = indice 11. Los encabezados estan en la
# 11 y arriba hay un titulo.
PRIMERA_FILA_DATOS = 11

POSICION_FECHA = 0
POSICION_HORA = 1

# El texto con el que el archivo marca que la unidad no participo. Se
# compara normalizado, asi que no importan tildes ni mayusculas.
TEXTO_NO_PARTICIPO = "no participo"

SUFIJOS_ALTERNATIVOS = (("tg", "TV"), ("tv", "TG"))


class ErrorDesempeno(Exception):
    """Error previsible al leer el archivo de desempeño."""


def _normalizar(texto):
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


def unidad_alternativa(unidad):
    """
    'CENTRAL TG' -> 'CENTRAL TV' y al reves. Cualquier otra cosa se
    devuelve igual.

    Es la segunda nomenclatura que prueba la busqueda de CPF: en el
    archivo de desempeño la misma maquina puede aparecer con el sufijo
    cambiado.
    """

    texto = str(unidad or "").strip()
    normalizado = _normalizar(texto)

    for sufijo, reemplazo in SUFIJOS_ALTERNATIVOS:
        if normalizado.endswith(sufijo):
            return texto[: -len(sufijo)] + reemplazo

    return texto


def calcular_hora_mes(fechas, horas, dia_cambio_hora=None, ajuste=1):
    """
    Hora_Mes = (dia - 1) * 24 + hora + 1

    El "+ 1" es porque la hora de estas hojas va de 0 a 23 y la de
    Subastas de 1 a 24. El ajuste por cambio de hora es el mismo que en
    Subastas (`=...+SI(DIA(Fecha)>SUBASTAS!F2;1;0)`), y por ahora viene
    apagado por omision: falta definir de donde sale ese dia (ver
    BITACORA.md, "Pendientes abiertos"). Las DOS puntas tienen que usar
    el mismo criterio o el cruce se corre entero.
    """

    dias = pd.to_datetime(fechas, errors="coerce").dt.day
    horas = pd.to_numeric(horas, errors="coerce")

    hora_mes = (dias - 1) * 24 + horas + 1

    if dia_cambio_hora is not None:
        hora_mes = hora_mes + (dias > int(dia_cambio_hora)).astype(int) * ajuste

    return hora_mes


def leer_hoja(ruta, control, dia_cambio_hora=None, registrar=print):
    """
    Una de las tres hojas horarias -> DataFrame normalizado con
    [Unidad, Unidad_norm, Hora_Mes, FD] y, solo para CSF, tambien
    [Respuesta, Participacion].
    """

    ruta = Path(ruta)
    config = HOJAS_DESEMPENO[control]

    try:
        crudo = pd.read_excel(ruta, sheet_name=config["hoja"], header=None)
    except ValueError as error:
        raise ErrorDesempeno(
            f"No existe la hoja '{config['hoja']}' en {ruta.name}: {error}"
        ) from error
    except Exception as error:
        raise ErrorDesempeno(
            f"No se pudo leer '{config['hoja']}' de {ruta.name}: {error}"
        ) from error

    inicio, fin = config["columnas"]
    bloque = crudo.iloc[PRIMERA_FILA_DATOS:, inicio:fin]

    if bloque.shape[1] < fin - inicio:
        raise ErrorDesempeno(
            f"La hoja '{config['hoja']}' de {ruta.name} trae "
            f"{bloque.shape[1]} columnas en el bloque esperado y se "
            f"necesitan {fin - inicio}."
        )

    unidades = bloque.iloc[:, config["unidad"]]

    tabla = pd.DataFrame({
        "Unidad": unidades.map(lambda v: str(v).strip() if pd.notna(v) else ""),
        "Hora_Mes": calcular_hora_mes(
            bloque.iloc[:, POSICION_FECHA],
            bloque.iloc[:, POSICION_HORA],
            dia_cambio_hora=dia_cambio_hora,
        ),
        "FD": pd.to_numeric(bloque.iloc[:, config["fd"]], errors="coerce"),
    })

    if config["respuesta"] is not None:
        tabla["Respuesta"] = bloque.iloc[:, config["respuesta"]].map(
            lambda v: str(v).strip() if pd.notna(v) else ""
        )

    tabla = tabla[tabla["Unidad"] != ""]
    tabla = tabla.dropna(subset=["Hora_Mes"])
    tabla["Hora_Mes"] = tabla["Hora_Mes"].astype(int)
    tabla["Unidad_norm"] = tabla["Unidad"].map(_normalizar)

    tabla = tabla.reset_index(drop=True)

    registrar(f"  FD {control.upper()}: {len(tabla):,} fila(s)")

    return tabla


def _participacion_csf(tabla):
    """
    Indicador de Participacion CSF (el "Vector Participacion CSF",
    `DB!AC`): 0 si la unidad figura como "No Participó" **y** su
    alternativa TG/TV tampoco participo; 1 en cualquier otro caso.

    Se calcula sobre la tabla ya armada porque necesita mirar OTRA fila
    (la de la unidad alternativa en la misma hora).
    """

    no_participo = tabla["Respuesta"].map(
        lambda v: _normalizar(v).startswith(TEXTO_NO_PARTICIPO)
    )

    # Respuesta de cada (unidad, hora) para poder consultar la de al lado.
    por_clave = {
        (fila.Unidad_norm, fila.Hora_Mes): fila.Respuesta
        for fila in tabla.itertuples(index=False)
    }

    participacion = []

    for fila, sin_participar in zip(tabla.itertuples(index=False), no_participo):

        if not sin_participar:
            participacion.append(1)
            continue

        alternativa = _normalizar(unidad_alternativa(fila.Unidad))
        respuesta_alternativa = por_clave.get((alternativa, fila.Hora_Mes))

        if (
            alternativa == fila.Unidad_norm
            or respuesta_alternativa is None
            or _normalizar(respuesta_alternativa).startswith(TEXTO_NO_PARTICIPO)
        ):
            participacion.append(0)
        else:
            participacion.append(1)

    return participacion


def construir_tablas_fd(ruta_sscc, dia_cambio_hora=None, registrar=print):
    """
    El archivo SSCC_Desempeño_* -> las tres tablas de FD normalizadas y
    el diccionario de participacion de CSF.

    Devuelve un dict:
        {"cpf": {(unidad_norm, hora_mes): fd, ...},
         "csf": {...},
         "ctf": {...},
         "participacion_csf": {(unidad_norm, hora_mes): 0|1, ...}}

    Una hoja que no este se saltea con un aviso: el FD de esa familia
    queda sin dato, pero las otras dos sirven igual.
    """

    ruta_sscc = Path(ruta_sscc)

    try:
        hojas_del_archivo = pd.ExcelFile(ruta_sscc).sheet_names
    except Exception as error:
        raise ErrorDesempeno(
            f"No se pudo abrir {ruta_sscc.name}: {error}"
        ) from error

    presentes = {_normalizar(h) for h in hojas_del_archivo}

    tablas = {"cpf": {}, "csf": {}, "ctf": {}, "participacion_csf": {}}

    for control, config in HOJAS_DESEMPENO.items():

        if _normalizar(config["hoja"]) not in presentes:
            registrar(
                f"  [AVISO] {ruta_sscc.name} no tiene la hoja "
                f"'{config['hoja']}': el FD de las filas "
                f"{control.upper()} queda sin dato."
            )
            continue

        tabla = leer_hoja(
            ruta_sscc, control,
            dia_cambio_hora=dia_cambio_hora, registrar=registrar,
        )

        dic = {}

        for fila in tabla.itertuples(index=False):
            clave = (fila.Unidad_norm, fila.Hora_Mes)
            if clave not in dic and pd.notna(fila.FD):
                dic[clave] = float(fila.FD)

        tablas[control] = dic

        if control == "csf":
            participacion = _participacion_csf(tabla)
            dic_participacion = {}

            for fila, valor in zip(tabla.itertuples(index=False), participacion):
                clave = (fila.Unidad_norm, fila.Hora_Mes)
                if clave not in dic_participacion:
                    dic_participacion[clave] = valor

            tablas["participacion_csf"] = dic_participacion

            sin_participar = sum(
                1 for v in dic_participacion.values() if v == 0
            )
            registrar(
                f"  Participacion CSF: {sin_participar:,} de "
                f"{len(dic_participacion):,} en 0 (no participo)"
            )

    return tablas


def buscar_fd(tablas, control, unidad, hora_mes):
    """
    El FD de una fila. Devuelve None si no esta.

    CPF prueba primero la unidad tal cual y despues la alternativa
    TG/TV; CSF y CTF buscan una sola vez (asi es la formula original).
    """

    control = str(control or "").strip().lower()[:3]

    dic = tablas.get(control)

    if not dic:
        return None

    try:
        hora_mes = int(hora_mes)
    except (TypeError, ValueError):
        return None

    valor = dic.get((_normalizar(unidad), hora_mes))

    if valor is None and control == "cpf":
        valor = dic.get((_normalizar(unidad_alternativa(unidad)), hora_mes))

    return valor


def buscar_participacion_csf(tablas, unidad, hora_mes):
    """
    El Vector de Participacion CSF de una fila. Devuelve None si no
    esta (quien llama decide que hacer; la formula original no tiene
    respaldo para ese caso).
    """

    try:
        hora_mes = int(hora_mes)
    except (TypeError, ValueError):
        return None

    return tablas.get("participacion_csf", {}).get(
        (_normalizar(unidad), hora_mes)
    )
