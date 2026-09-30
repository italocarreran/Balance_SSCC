# -*- coding: utf-8 -*-
"""
Conciliacion de energia entre Medidores y las dos hojas de calculo.

Control TRA del catalogo. El traspaso reparte cada fila de Medidores a
UNA de las dos hojas segun Medidores!L (Ventana_No_Completa): si vale
1 la energia va a "Calculo E Costos", si no va a "Calculo RE545". Por
construccion entonces:

    Energia Medidores = Energia E Costos + Energia RE545

Es una resta, y es el control mas barato que existe para detectar una
fila perdida, una contada dos veces o un filtro que se comio algo. La
planilla original ya lo hacia; la version Python no lo tenia.
"""

import pandas as pd

from .alertas import CRITICA, INFO, Alerta, anotar
from .parametros import HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545


# Tolerancia de la comparacion. La energia se suma en float64 sobre
# ~100.000 filas, asi que la igualdad exacta no sirve: se acepta una
# diferencia relativa de 1e-9 y, cuando el total es casi cero, un piso
# absoluto. Los dos numeros quedan escritos en la hoja "Ejecucion"
# junto con la diferencia observada (TRA-008 del catalogo pide
# justamente eso).
TOLERANCIA_RELATIVA = 1e-9
TOLERANCIA_ABSOLUTA = 1e-6


# Las dos hojas se concilian SIEMPRE con los nombres internos de
# columna, no con los reales del Excel ("Descarga kWh"/"Carga kWh"):
# despues de renombrar, el indice de columnas tiene nombres repetidos a
# proposito y ya no se puede indexar por nombre sin ambiguedad.
COLUMNAS_ENERGIA = ("Energia_Positiva", "Energia_Negativa")


def _energia_total(df):
    """Energia de una hoja de calculo: positiva + negativa."""

    if df is None or df.empty:
        return 0.0

    faltantes = [c for c in COLUMNAS_ENERGIA if c not in df.columns]

    if faltantes:
        raise KeyError(
            f"La hoja que llego a conciliar no tiene {faltantes}: "
            f"se esperaban los nombres INTERNOS de columna, no los del "
            f"Excel ya renombrados. Columnas recibidas: "
            f"{list(df.columns)[:12]}..."
        )

    total = 0.0

    for columna in COLUMNAS_ENERGIA:
        total += float(
            pd.to_numeric(df[columna], errors="coerce").fillna(0.0).sum()
        )

    return total


def conciliar_energia(
    df_medidores, df_ecostos=None, df_re545=None, registrar=print,
):
    """
    Compara la energia de Medidores contra la de las dos hojas y anota
    las alertas que correspondan. Devuelve un dict con los totales,
    para la hoja "Ejecucion".

    Si solo se genero una de las dos hojas (el usuario puede pedir una
    sola desde la ventana), la igualdad completa no se puede exigir: en
    ese caso se concilia contra la parte que le toca a esa hoja y se
    deja dicho en el resultado.
    """

    energia = pd.to_numeric(
        df_medidores["Gen_Unidad"], errors="coerce"
    ).fillna(0.0)

    # El mismo criterio del traspaso: L = 1 -> E Costos; el resto -> RE545.
    va_a_ecostos = pd.to_numeric(
        df_medidores["Ventana_No_Completa"], errors="coerce"
    ).eq(1)

    total_medidores = float(energia.sum())
    esperado_ecostos = float(energia.where(va_a_ecostos, 0.0).sum())
    esperado_re545 = float(energia.where(~va_a_ecostos, 0.0).sum())

    real_ecostos = _energia_total(df_ecostos)
    real_re545 = _energia_total(df_re545)

    completa = df_ecostos is not None and df_re545 is not None

    resultado = {
        "filas_medidores": len(df_medidores),
        "energia_medidores": total_medidores,
        "energia_ecostos": real_ecostos if df_ecostos is not None else None,
        "energia_re545": real_re545 if df_re545 is not None else None,
        "esperado_ecostos": esperado_ecostos,
        "esperado_re545": esperado_re545,
        "conciliacion_completa": completa,
        "tolerancia_relativa": TOLERANCIA_RELATIVA,
        "tolerancia_absoluta": TOLERANCIA_ABSOLUTA,
        "diferencia": None,
    }

    tolerancia = max(
        TOLERANCIA_ABSOLUTA, abs(total_medidores) * TOLERANCIA_RELATIVA
    )

    # --- TRA-001: cada hoja recibe TODAS las filas de Medidores -----
    for df, nombre in ((df_ecostos, HOJA_CALCULO_ECOSTOS),
                       (df_re545, HOJA_CALCULO_RE545)):
        if df is None or len(df) == len(df_medidores):
            continue
        anotar(registrar, Alerta(
            "TRA-001", CRITICA, nombre,
            f"La hoja tiene {len(df):,} fila(s) y Medidores "
            f"{len(df_medidores):,}: el traspaso perdio o duplico filas.",
            valor_encontrado=f"{len(df):,} filas",
            valor_esperado=f"{len(df_medidores):,} filas",
            accion="no se puede conciliar la energia de esa hoja",
            origen_control="CATALOGO TRA-001",
        ))

    # --- TRA-005 / TRA-006: cada hoja contra lo que le toca ---------
    for real, esperado, nombre, control in (
        (real_ecostos, esperado_ecostos, HOJA_CALCULO_ECOSTOS, "TRA-005"),
        (real_re545, esperado_re545, HOJA_CALCULO_RE545, "TRA-006"),
    ):
        if (nombre == HOJA_CALCULO_ECOSTOS and df_ecostos is None) or \
           (nombre == HOJA_CALCULO_RE545 and df_re545 is None):
            continue
        if abs(real - esperado) <= tolerancia:
            continue
        anotar(registrar, Alerta(
            control, CRITICA, nombre,
            f"La energia de la hoja no coincide con la que le asigna "
            f"Medidores!L: diferencia {real - esperado:,.6f}.",
            valor_encontrado=f"{real:,.6f}",
            valor_esperado=f"{esperado:,.6f}",
            accion="corrida NO APROBADA",
            origen_control=f"CATALOGO {control}",
        ))

    # --- TRA-007 / TRA-008: la suma de las dos contra Medidores -----
    if completa:
        diferencia = total_medidores - (real_ecostos + real_re545)
        resultado["diferencia"] = diferencia

        if abs(diferencia) > tolerancia:
            anotar(registrar, Alerta(
                "TRA-007", CRITICA, "Traspaso",
                f"La energia no se conserva: Medidores "
                f"{total_medidores:,.6f} != E Costos {real_ecostos:,.6f} "
                f"+ RE545 {real_re545:,.6f} (diferencia "
                f"{diferencia:,.6f}, tolerancia {tolerancia:,.9f}).",
                valor_encontrado=f"{real_ecostos + real_re545:,.6f}",
                valor_esperado=f"{total_medidores:,.6f}",
                accion="corrida NO APROBADA",
                origen_control="CATALOGO TRA-007/TRA-008",
            ))
        else:
            registrar(
                f"  Conciliacion de energia OK: Medidores "
                f"{total_medidores:,.3f} = E Costos {real_ecostos:,.3f} + "
                f"RE545 {real_re545:,.3f} (diferencia {diferencia:,.9f})."
            )
    else:
        anotar(registrar, Alerta(
            "TRA-009", INFO, "Traspaso",
            "Solo se genero una de las dos hojas: se concilio esa contra "
            "la parte que le asigna Medidores!L, pero NO la igualdad "
            "Medidores = E Costos + RE545.",
            accion="conciliacion parcial",
            origen_control="CATALOGO TRA-009",
        ))

    return resultado
