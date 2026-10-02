# -*- coding: utf-8 -*-
"""
Hojas del CRA que salen de las MISMAS fuentes que ya usa el BESS
(confirmado por el usuario):

  - FD_CPF, FD_CSF, FD_CTF: las tres hojas horarias del
    SSCC_Desempeño_* del DCO (CPF/CSF/CTF Horario), con los rangos que ya
    usa Script/Fd/Desempeno_Horario.py (encabezados en la fila 11).
  - PRORRATA_RETIROS, bloque 1 (traz. 8.1): la matriz periodo x empresa
    pagadora, desde la hoja "Prorrata 15min" de Prorrata_Retiros_<AAMM>,
    leida con el mismo lector del BESS (nucleo.prorrata_retiros).

Las hojas FD se copian enteras, con todas las unidades y sus
encabezados reales: el filtro por unidad candidata del CRA (su
InfoTecnica) se aplica recien en CÁLCULO_CRA, que no esta todavia.
"""

import calendar
import datetime as dt


from ..nucleo.prorrata_retiros import (
    COL_CUARTO,
    COL_PRORRATA,
    COL_SUMINISTRADOR,
    leer_prorrata_retiros,
)
from ..nucleo.utiles import normalizar
from . import parametros as p
from .lectura import (
    a_fecha_hora,
    leer_columnas,
    leer_encabezados,
    numero_o_nada,
)


def _nada(*_args, **_kwargs):
    pass


# ============================================================
# FD_CPF, FD_CSF, FD_CTF
# ============================================================

def construir_fd(ruta_sscc, seccion, unidades, registrar=_nada):
    """
    Una de las tres hojas horarias del SSCC_Desempeño_*, solo con las
    filas de `unidades` (las de esa columna del diccionario de
    centrales_cra.xlsx; se comparan normalizadas): los encabezados de la
    fila 11 y los datos desde la 12.
    """

    _, hoja_origen, columnas = p.HOJAS_FD[seccion]
    letras = list(columnas)

    encabezados = leer_encabezados(
        ruta_sscc, hoja_origen, p.FILA_ENCABEZADO_FD, letras
    )
    datos = leer_columnas(
        ruta_sscc, hoja_origen, p.FILA_ENCABEZADO_FD + 1, letras
    )

    leidas = len(datos)
    datos = _filtrar_unidades(datos, unidades, hoja_origen, registrar)
    registrar(
        f"  {hoja_origen}: {len(datos):,} de {leidas:,} fila(s) son de "
        f"unidades del CRA."
    )

    fechas_horas = [
        _fecha_hora(fecha, hora)
        for fecha, hora in zip(datos[p.LETRA_FECHA_FD], datos[p.LETRA_HORA_FD])
    ]
    datos.columns = _nombres_unicos(letras, encabezados)
    datos.insert(0, p.CAMPO_FECHA_HORA_FD, fechas_horas)

    sin_clave = sum(1 for v in fechas_horas if v is None)
    if sin_clave:
        registrar(
            f"  AVISO {hoja_origen}: {sin_clave} fila(s) sin Fecha/Hora "
            f"valida: quedan sin '{p.CAMPO_FECHA_HORA_FD}'."
        )
    return datos.reset_index(drop=True)


def _filtrar_unidades(datos, unidades, hoja_origen, registrar):
    """Solo las filas cuya Unidad (columna D) esta en `unidades`."""

    buscadas = {normalizar(u): u for u in unidades}
    claves = datos[p.LETRA_UNIDAD_FD].map(normalizar)

    sin_filas = [u for k, u in buscadas.items() if k not in set(claves)]
    if sin_filas:
        registrar(
            f"  AVISO {hoja_origen}: {len(sin_filas)} unidad(es) del "
            f"diccionario sin ninguna fila: {', '.join(sin_filas)}."
        )

    return datos[claves.isin(buscadas)].reset_index(drop=True)


def _fecha_hora(fecha, hora):
    """B + TIMEVALUE(C & " :00"): la fecha del dia mas C horas (0..23)."""

    fecha = a_fecha_hora(fecha)
    hora = numero_o_nada(hora)

    if fecha is None or hora is None:
        return None

    return (
        dt.datetime(fecha.year, fecha.month, fecha.day)
        + dt.timedelta(hours=hora)
    )


def _nombres_unicos(letras, encabezados):
    """El texto del encabezado; la letra si esta vacio o repetido."""

    nombres, vistos = [], set()

    for letra in letras:
        nombre = encabezados.get(letra) or letra
        if nombre in vistos:
            nombre = f"{nombre} ({letra})"
        vistos.add(nombre)
        nombres.append(nombre)

    return nombres


# ============================================================
# PRORRATA_RETIROS, bloque 1
# ============================================================

def construir_matriz_prorrata(ruta_prorrata, anio, mes, registrar=_nada):
    """
    Prorrata 15min (cuarto, suministrador, prorrata) -> una fila por
    periodo de 15 minutos del mes y una columna por empresa pagadora,
    en orden alfabetico (el que deja el lector del BESS). Donde una
    empresa no retiro en ese periodo, la celda queda vacia.
    """

    largo = leer_prorrata_retiros(ruta_prorrata, registrar=registrar)

    empresas = sorted(set(largo[COL_SUMINISTRADOR]))
    matriz = (
        largo.pivot(index=COL_CUARTO, columns=COL_SUMINISTRADOR,
                    values=COL_PRORRATA)
        .reindex(columns=empresas)
        .sort_index()
    )

    _avisar_cobertura(matriz.index, anio, mes, registrar)
    registrar(
        f"  Prorrata: {len(matriz):,} periodo(s) x {len(empresas)} empresa(s)."
    )

    matriz.index.name = p.COLUMNA_PERIODO_PRORRATA
    matriz.columns.name = None

    return matriz.reset_index()


def _avisar_cobertura(periodos, anio, mes, registrar):
    """
    Aviso si los periodos no son exactamente 1..dias*96. No es un error:
    en un mes con cambio de hora el total no es dias*96 y esa regla
    todavia no esta cerrada (traz. 0.1).
    """

    esperados = calendar.monthrange(anio, mes)[1] * p.PERIODOS_POR_DIA
    presentes = set(int(x) for x in periodos)
    faltan = sorted(set(range(1, esperados + 1)) - presentes)
    sobran = sorted(presentes - set(range(1, esperados + 1)))

    if faltan or sobran:
        registrar(
            f"  AVISO Prorrata: se esperaban los periodos 1..{esperados} "
            f"({esperados // p.PERIODOS_POR_DIA} dias x "
            f"{p.PERIODOS_POR_DIA}); faltan {len(faltan)} y sobran "
            f"{len(sobran)}"
            + (f" (primeros que faltan: {faltan[:5]})" if faltan else "")
            + (f" (primeros que sobran: {sobran[:5]})" if sobran else "")
            + ". Si el mes tiene cambio de hora, puede ser correcto."
        )
