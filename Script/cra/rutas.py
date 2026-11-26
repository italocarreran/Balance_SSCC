# -*- coding: utf-8 -*-
"""
Rutas de un caso CRA: todo se deriva de la carpeta base y del periodo.
Nunca rutas absolutas ni relativas al .py (mismo criterio que el BESS).
"""

import re
from pathlib import Path

from ..nucleo.carpetas import CARPETA_CRA, carpeta_fd, raiz_herramienta
from ..nucleo.prorrata_retiros import buscar_archivo_prorrata
from ..nucleo.rutas import buscar_archivo_sscc_desempeno
from ..nucleo.utiles import ErrorEntrada
from . import parametros as p
from .lectura import buscar_por_prefijo


def validar_aamm(aamm):
    """'2608' -> (2026, 8). ErrorEntrada si no son 4 digitos validos."""

    aamm = str(aamm or "").strip()

    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        raise ErrorEntrada(
            f"Periodo invalido: '{aamm}'. Tienen que ser 4 digitos AAMM "
            f"(por ejemplo 2608 para agosto de 2026)."
        )

    return 2000 + int(aamm[:2]), int(aamm[2:])


def nombre_salida(aamm):
    """'2608' -> '5_REMUNERACION_CRA_2608.xlsx'."""

    return f"{p.PREFIJO_SALIDA}{aamm}.xlsx"


def raiz_cra(carpeta_base):
    """
    Donde van las carpetas propias del CRA y su planilla:
    <mes>/CRA/, o la carpeta misma en un caso de antes (todo suelto en
    la raiz; ver Script/nucleo/carpetas.py).
    """

    return raiz_herramienta(
        carpeta_base, CARPETA_CRA, p.CARPETAS_DISTINTIVAS,
        archivos=(f"{p.PREFIJO_SALIDA}*.xlsx", p.ARCHIVO_SALIDA_ANTIGUO),
    )


def _periodo_o_none(aamm):
    texto = str(aamm or "").strip()
    try:
        validar_aamm(texto)
    except ErrorEntrada:
        return None
    return texto


def _salida_sin_periodo(raiz):
    """La 5_REMUNERACION_CRA_AAMM.xlsx que haya, la de antes o una generica."""

    try:
        encontradas = sorted(
            r for r in raiz.glob(f"{p.PREFIJO_SALIDA}*.xlsx")
            if re.fullmatch(rf"{re.escape(p.PREFIJO_SALIDA)}\d{{4}}\.xlsx", r.name)
        )
    except OSError:
        encontradas = []
    if len(encontradas) == 1:
        return encontradas[0]
    if (raiz / p.ARCHIVO_SALIDA_ANTIGUO).is_file():
        return raiz / p.ARCHIVO_SALIDA_ANTIGUO
    return raiz / p.ARCHIVO_SALIDA


def _migrar_salida_antigua(raiz, salida):
    """
    Balance_CRA.xlsx pasa a 5_REMUNERACION_CRA_AAMM.xlsx (nunca pisa
    uno que ya este). Si no se puede (abierto en Excel) se sigue usando
    el de antes y se reintenta la proxima vez.
    """

    vieja = raiz / p.ARCHIVO_SALIDA_ANTIGUO
    if not vieja.is_file() or salida.exists():
        return salida
    try:
        vieja.replace(salida)
        return salida
    except OSError:
        return vieja


def resolver_rutas(carpeta_base, aamm=None):
    """
    Las carpetas del caso y el archivo de salida (existan o no).

    carpeta_base es la carpeta del MES (la misma del BESS): lo propio
    del CRA en <mes>/CRA/ y lo comun (FD, prorrata) en la raiz.

    aamm: el periodo, para el nombre de la planilla
    (5_REMUNERACION_CRA_AAMM.xlsx); sin uno valido se usa la que haya.
    """

    base = Path(carpeta_base)
    raiz = raiz_cra(base)
    periodo = _periodo_o_none(aamm)

    salida = raiz / nombre_salida(periodo) if periodo else _salida_sin_periodo(raiz)
    if periodo and raiz.is_dir():
        salida = _migrar_salida_antigua(raiz, salida)

    return {
        "base": base,
        "raiz_cra": raiz,
        "energia": raiz / p.CARPETA_ENERGIA,
        "fp": raiz / p.CARPETA_FP,
        "co": raiz / p.CARPETA_CO,
        "sc_co": raiz / p.CARPETA_SC_CO,
        "auxiliares": raiz / p.CARPETA_AUXILIARES,
        # Comunes con el BESS. Un caso CRA de antes tenia su propia FD/.
        "fd": carpeta_fd(base, antiguas=(p.CARPETA_FD_ANTIGUA,)),
        "prorrata": base / p.CARPETA_PRORRATA,
        "cmg": base / p.CARPETA_CMG,
        "salida": salida,
        "medidas_cra": raiz / p.CARPETA_ENERGIA / p.ARCHIVO_MEDIDAS_CRA,
        "puntos_fallidos": raiz / p.CARPETA_ENERGIA / p.ARCHIVO_PUNTOS_FALLIDOS_CRA,
        "trabajo_medidas": raiz / p.CARPETA_ENERGIA / p.CARPETA_TRABAJO_MEDIDAS,
        "log_dir": raiz / p.CARPETA_LOG,
        "registro": raiz / p.CARPETA_LOG / p.ARCHIVO_REGISTRO,
    }


# ============================================================
# LAS ENTRADAS
#
# Cada una: id, carpeta (clave de resolver_rutas), como se muestra en el
# arbol, que se busca (para el mensaje de error) y la funcion que la
# encuentra: buscar(carpeta, aamm) -> ruta o None.
# ============================================================

def prefijos_energia(aamm):
    """'2608' -> ('formato_solicitud_ssaa_sscc_hidro_agosto2026',)"""

    anio, mes = validar_aamm(aamm)
    meses = (p.MESES[mes - 1],) + p.MESES_ALTERNATIVOS.get(mes, ())

    return tuple(
        p.PREFIJO_ENERGIA.format(mes=nombre, anio=anio) for nombre in meses
    )


def _buscar_energia(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, prefijos_energia(aamm), "energia (Formato_Solicitud_*)"
    )


def _buscar_fp(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_FP.format(aamm=aamm), "factores de penalizacion"
    )


def _buscar_co(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_CO.format(aamm=aamm), "costos de operacion"
    )


def _buscar_reporte_cra(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_REPORTE_CRA, "Reporte_CRA",
        extensiones=p.EXTENSIONES_REPORTE_CRA,
    )


def _buscar_sobrecostos(carpeta, aamm):
    return buscar_por_prefijo(
        carpeta, p.PREFIJO_SOBRECOSTOS, "Cálculo_SobrecostosSSCC"
    )


def _buscar_sscc_desempeno(carpeta, aamm):
    # El mismo buscador del BESS: si hay varios, el mas reciente (asi
    # lo hace la macro original).
    return buscar_archivo_sscc_desempeno(carpeta)


def _buscar_prorrata(carpeta, aamm):
    return buscar_archivo_prorrata(carpeta, aamm)


def _buscar_centrales_cra(carpeta, aamm):
    ruta = Path(carpeta) / p.ARCHIVO_CENTRALES_CRA
    return ruta if ruta.is_file() else None


def _buscar_homologacion(carpeta, aamm):
    # El mismo buscador del BESS: cualquier Excel con "homologacion" en
    # el nombre (si hay varios, el mas reciente).
    from ..nucleo.externos import Homologacion

    return Homologacion.buscar_archivo_homologacion(carpeta)


def _buscar_medidas_cra(carpeta, aamm):
    ruta = Path(carpeta) / p.ARCHIVO_MEDIDAS_CRA
    return ruta if ruta.is_file() else None


ENTRADAS = [
    ("centrales_cra", "auxiliares", p.ARCHIVO_CENTRALES_CRA,
     _buscar_centrales_cra),
    ("homologacion", "auxiliares", "*Homologacion*.xlsx",
     _buscar_homologacion),
    ("energia", "energia",
     "Formato_Solicitud_*Hidro_<Mes><AAAA>.xlsx", _buscar_energia),
    # La genera el boton "Generar" de la fila de arriba (desde PRMTE).
    ("medidas_cra", "energia", p.ARCHIVO_MEDIDAS_CRA, _buscar_medidas_cra),
    ("fp", "fp", "fp_<AAMM>*.xlsx", _buscar_fp),
    ("co", "co", "cvar_cra_<AAMM>_*.xlsx", _buscar_co),
    ("reporte_cra", "sc_co", "Reporte_CRA*.csv", _buscar_reporte_cra),
    ("sobrecostos", "sc_co", "Cálculo_SobrecostosSSCC_*.xlsm",
     _buscar_sobrecostos),
    ("sscc_desempeno", "fd", "SSCC_Desempeño_*.xlsx", _buscar_sscc_desempeno),
    ("prorrata", "prorrata", "Prorrata_Retiros_<AAMM>_pre/_def.xlsx",
     _buscar_prorrata),
]


def _entrada(id_entrada):
    for entrada in ENTRADAS:
        if entrada[0] == id_entrada:
            return entrada

    raise KeyError(id_entrada)


def buscar_entrada(rutas, id_entrada, aamm):
    """La ruta del archivo de esa entrada, o None si no esta."""

    _, carpeta, _, buscar = _entrada(id_entrada)

    return buscar(rutas[carpeta], aamm)


def exigir_entrada(rutas, id_entrada, aamm):
    """Como buscar_entrada(), pero ErrorEntrada si no esta."""

    ruta = buscar_entrada(rutas, id_entrada, aamm)

    if ruta is None:
        _, carpeta, patron, _ = _entrada(id_entrada)
        raise ErrorEntrada(
            f"Falta el archivo {patron.replace('<AAMM>', aamm)} en "
            f"{rutas[carpeta]}."
        )

    return ruta


def subcarpetas_del_caso(carpeta_base):
    """
    Las subcarpetas que tiene que tener ESTE caso, relativas a la
    carpeta del mes: las de un caso nuevo (p.SUBCARPETAS_CASO) o, en uno
    de antes con todo suelto, las de antes (asi completarlo no le crea
    un CRA/ vacio al lado, que lo haria buscar todo ahi). La carpeta del
    FD es la que el caso ya use (ver nucleo.carpetas.carpeta_fd).
    """

    base = Path(carpeta_base)
    raiz = raiz_cra(base)
    rutas = resolver_rutas(base)
    propias = [
        nombre if raiz == base else f"{CARPETA_CRA}/{nombre}"
        for nombre in p.CARPETAS_PROPIAS
    ]
    return tuple(propias) + (
        p.CARPETA_CMG, rutas["fd"].name, p.CARPETA_PRORRATA,
    )


def crear_carpetas_caso(carpeta_base):
    """
    Crea las subcarpetas que falten (ver subcarpetas_del_caso). No toca
    lo que ya hay. Si el mes ya tiene BESS/, CRA/ queda a su lado.
    """

    base = Path(carpeta_base)
    creadas = []

    for nombre in subcarpetas_del_caso(base):
        ruta = base / nombre
        if not ruta.exists():
            ruta.mkdir(parents=True)
            creadas.append(nombre)

    return creadas
