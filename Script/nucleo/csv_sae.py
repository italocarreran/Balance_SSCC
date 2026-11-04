# -*- coding: utf-8 -*-
"""
El CSV de pagos que se sube: SAE_AAMM.csv.

Pedido del usuario (2026-11-04): "debe haber un boton final que genere
el siguiente documento" (SAE_2609.csv). Lo armaba la macro
C_Exporta_pagos.Generar_CSV_SAE de 11_PAGOS_BESS (ver
docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md); aca se replica:

  - lee 'Calculo E Costos' y 'Calculo RE545' (la macro: A Mes, B Dia,
    C Hora, E Minuto, F Bloque, G Configuracion y el pago en AZ de E
    Costos y CE de RE545; aca son las mismas columnas por su nombre:
    Mes, Dia, Hora, Minuto, Bloque horario, Configuracion y "Monto a
    compensar" de las dos hojas);
  - descarta los pagos en cero (|pago| <= 1e-7) y suma los dos pagos de
    un mismo Bloque + Configuracion;
  - Fecha_Hora = dia + (Hora - 1) horas + Minuto, como "dd-mm-aaaa hh:mm"
    (en la hoja la hora va 1..24; en el CSV, 00..23);
  - columnas Fecha_Hora, CONFIGURACION, Central (vacia), Pago (dos
    decimales con punto), Tipo_pago ("BESS"), Bloque_15min; ordenado por
    bloque y configuracion;
  - UTF-8 con BOM y fin de linea CRLF, igual que el archivo de la macro.

Lo que la macro anotaba en su hoja 'Verificacion_CSV' (fecha distinta
para un mismo bloque, saltos de hora y el total por configuracion) va
al registro de la corrida.
"""

from datetime import datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd

from .lector import abrir_excel, leer_excel
from .lectura import fila_de_encabezado
from .parametros import HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545
from .rutas import resolver_rutas, validar_aamm
from .unidades import sin_unidades
from .utiles import ErrorEntrada

ENCABEZADO_CSV = "Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min"
TIPO_PAGO_SALIDA = "BESS"
EPS_CERO = 0.0000001

COLUMNA_PAGO = "Monto a compensar"
COLUMNAS_NECESARIAS = (
    "Mes", "Dia", "Hora", "Minuto", "Bloque horario", "Configuracion",
    COLUMNA_PAGO,
)

# Saltos entre cuartos seguidos de una configuracion, en minutos.
_SALTOS = {
    75: "Cambio hora +1 detectado",
    -45: "Cambio hora -1 detectado",
    0: "Fecha/hora duplicada",
}


def leer_hoja_calculo(ruta, nombre, libro=None):
    """
    Las columnas que usa el CSV de una hoja de calculo de la planilla,
    con sus nombres internos. La hoja se lee UNA vez (con su fila de
    encabezado buscada en las primeras filas, como el resto del programa).
    """

    try:
        crudo = leer_excel(libro if libro is not None else ruta,
                           sheet_name=nombre, header=None)
    except (FileNotFoundError, ValueError) as error:
        raise ErrorEntrada(
            f"Primero calcula la hoja '{nombre}' de {Path(ruta).name}."
        ) from error

    fila = fila_de_encabezado(crudo.head(4)) or 1
    nombres = [
        f"Unnamed: {i}" if pd.isna(v) else str(v)
        for i, v in enumerate(crudo.iloc[fila])
    ]
    df = sin_unidades(
        crudo.iloc[fila + 1:].set_axis(nombres, axis=1).reset_index(drop=True),
        nombre,
    )
    df = df.loc[:, ~df.columns.duplicated()]

    faltan = [c for c in COLUMNAS_NECESARIAS if c not in df.columns]
    if faltan:
        raise ErrorEntrada(
            f"La hoja '{nombre}' de {Path(ruta).name} no trae {faltan}: "
            f"vuelve a calcularla."
        )
    return df[list(COLUMNAS_NECESARIAS)]


def _pago_csv(valor):
    """Format$(valor, "0.00") con punto decimal."""

    return str(Decimal(repr(float(valor))).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    ))


def _valor_csv(texto):
    """CSVValue de la macro: entre comillas si trae coma, comilla o salto."""

    texto = str(texto).replace('"', '""')
    if any(c in texto for c in (",", '"', "\r", "\n")):
        return f'"{texto}"'
    return texto


def _acumular(df, nombre, anio, datos, verificaciones):
    """AcumularHojaCSV: suma los pagos de la hoja en datos."""

    vacio_mes = df["Mes"].map(
        lambda v: str(v).strip() in ("", "nan", "None", "NaT")
    ).astype(bool)
    df = df.loc[~vacio_mes.to_numpy()]

    # ValorNumerico de la macro: vacio o texto que no es numero -> 0.
    pagos = pd.to_numeric(df[COLUMNA_PAGO], errors="coerce").fillna(0.0)

    filas = []
    for mes, dia, hora, minuto, bloque, config, pago in zip(
        df["Mes"], df["Dia"], df["Hora"], df["Minuto"],
        df["Bloque horario"], df["Configuracion"], pagos.tolist(),
    ):
        fecha = datetime(anio, int(mes), int(dia)) + timedelta(
            hours=int(hora) - 1, minutes=int(minuto)
        )
        config = str(config).strip()
        bloque = int(bloque)
        filas.append((config, bloque, fecha))
        if abs(pago) <= EPS_CERO:
            continue
        texto_fecha = fecha.strftime("%d-%m-%Y %H:%M")
        clave = (bloque, config)
        if clave not in datos:
            datos[clave] = [texto_fecha, pago]
        else:
            datos[clave][1] += pago
            if datos[clave][0] != texto_fecha:
                verificaciones.append(
                    f"{nombre}: {config} bloque {bloque}: fecha distinta "
                    f"para el mismo bloque ({datos[clave][0]} / {texto_fecha})"
                )

    # Saltos de hora por configuracion, en el orden de los bloques.
    por_config = {}
    for config, bloque, fecha in filas:
        por_config.setdefault(config, []).append((bloque, fecha))
    for config in sorted(por_config):
        anterior = None
        for bloque, fecha in sorted(por_config[config]):
            if anterior is not None:
                minutos = int((fecha - anterior).total_seconds() // 60)
                if minutos != 15:
                    motivo = _SALTOS.get(minutos, "Salto inesperado")
                    verificaciones.append(
                        f"{nombre}: {config} bloque {bloque}: {motivo} "
                        f"({anterior:%d-%m-%Y %H:%M} -> {fecha:%d-%m-%Y %H:%M})"
                    )
            anterior = fecha


def construir_csv_sae(df_ecostos, df_re545, aamm):
    """
    (texto del CSV, totales {configuracion: pago}, verificaciones).
    texto es None si todos los pagos quedan en cero.
    """

    aamm = validar_aamm(aamm)
    anio = 2000 + int(aamm[:2])
    datos, verificaciones = {}, []
    _acumular(df_ecostos, HOJA_CALCULO_ECOSTOS, anio, datos, verificaciones)
    _acumular(df_re545, HOJA_CALCULO_RE545, anio, datos, verificaciones)

    claves = sorted(
        datos,
        key=lambda c: f"{c[0]:06d}|{c[1]}|{TIPO_PAGO_SALIDA}".casefold(),
    )
    lineas = [ENCABEZADO_CSV]
    totales = {}
    for bloque, config in claves:
        texto_fecha, pago = datos[(bloque, config)]
        if abs(pago) <= EPS_CERO:
            continue
        lineas.append(",".join((
            _valor_csv(texto_fecha), _valor_csv(config), "",
            _pago_csv(pago), _valor_csv(TIPO_PAGO_SALIDA), str(bloque),
        )))
        totales[config] = totales.get(config, 0.0) + pago

    if len(lineas) == 1:
        return None, totales, verificaciones
    return "\r\n".join(lineas) + "\r\n", totales, verificaciones


def generar_csv_sae(carpeta_base, aamm, registrar=print, progreso=None):
    """
    Escribe <caso>/SAE_AAMM.csv desde las hojas de calculo de
    11_PAGOS_BESS_AAMM.xlsx. Devuelve la ruta.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base, aamm)
    if not rutas["salida"].is_file():
        raise ErrorEntrada(
            f"No esta {rutas['salida'].name}: primero calcula la planilla."
        )

    registrar(f"Leyendo las hojas de calculo de {rutas['salida'].name}...")
    with abrir_excel(rutas["salida"]) as libro:
        df_ecostos = leer_hoja_calculo(rutas["salida"], HOJA_CALCULO_ECOSTOS, libro)
        df_re545 = leer_hoja_calculo(rutas["salida"], HOJA_CALCULO_RE545, libro)
    if progreso:
        progreso(50)

    texto, totales, verificaciones = construir_csv_sae(df_ecostos, df_re545, aamm)

    saltos = [v for v in verificaciones if "Cambio hora" in v]
    otros = [v for v in verificaciones if "Cambio hora" not in v]
    if saltos:
        registrar(
            f"  cambio de hora: {len(saltos)} salto(s) de hora esperados "
            f"(uno por configuracion y hoja)"
        )
    for aviso in otros[:20]:
        registrar(f"  [AVISO] {aviso}")
    if len(otros) > 20:
        registrar(f"  [AVISO] ... y {len(otros) - 20} mas")

    if texto is None:
        raise ErrorEntrada(
            "Luego de consolidar, todos los pagos quedaron en cero: no se "
            "genero el CSV."
        )

    destino = rutas["csv_sae"]
    with open(destino, "w", encoding="utf-8-sig", newline="") as archivo:
        archivo.write(texto)

    registrar("  total pagado por configuracion:")
    for config in sorted(totales, key=str.casefold):
        registrar(f"    {config}: {totales[config]:,.2f}")
    registrar(f"    TOTAL GENERAL: {sum(totales.values()):,.2f}")
    registrar(f"Listo: {texto.count(chr(10)) - 1:,} fila(s) en {destino}")
    if progreso:
        progreso(100)
    return destino
