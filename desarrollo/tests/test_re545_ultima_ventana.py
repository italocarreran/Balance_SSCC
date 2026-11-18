# -*- coding: utf-8 -*-
"""
La ultima ventana del mes queda pendiente para el mes siguiente
(observacion O2026-0052718, acogida el 2026-11-18): Edisp_T = 0 en RE545 y
monto 0 en E Costos, sea cual sea su numero (30 o 31). Entre el 2026-11-02
y el 2026-11-18 se pagaba parcial en el mes.
"""

import unittest
from unittest import mock

import pandas as pd

from Script.nucleo.diccionarios import _normaliza_cuarto
from Script.nucleo.re545 import completar_calculo_re545, construir_calculo_re545
from Script.nucleo.re545_resumen import construir_resumen_ventanas_re545
from Script.nucleo.utiles import normalizar

CENTRAL, BARRA = "SAE-DEMO", "B-DEMO"
silencio = lambda *a, **kw: None  # noqa: E731


def _edisp(ventana):
    filas, dic_cmg = [], {}
    for h in (10, 11):
        for q in range(4):
            cuarto = (24 * 30 + h - 1) * 4 + q + 1
            filas.append({
                "Mes": 3, "Dia": 31, "Hora": h, "Minutos": q * 15,
                "Hora Mes": 24 * 30 + h, "Cuarto de Hora": cuarto,
                "clave": CENTRAL, "Gen_Unidad": -1000.0 if h == 10 else 500.0,
                "SoC": 0.2, "Ventana": ventana, "Copia_Ventana": ventana,
                "Ventana_No_Completa": 0,
            })
            dic_cmg[BARRA.upper() + "|" + _normaliza_cuarto(cuarto)] = (50.0, 50.0)
    k = normalizar(CENTRAL)
    re = construir_calculo_re545(pd.DataFrame(filas), {k: BARRA}, dic_cmg, registrar=silencio)
    subastas = pd.DataFrame([
        {"Concepto": "CPF(+)", "Control": "CPF", "Sub_Baj": "SUBIDA", "Mes": 3,
         "Dia": 31, "Hora_dia": h, "Hora_mes": 24 * 30 + h,
         "Configuración": CENTRAL, "Energía SSCC": 3.0, "FD": 1.0, "FMA": 1.0}
        for h in (10, 11)
    ])
    re = completar_calculo_re545(re, subastas, {k: 40.0}, {k: 0.85}, registrar=silencio)
    oferta = pd.DataFrame([{"Central": CENTRAL, "Ventana T": ventana, "Oferta": 8, "Completa": 1}])
    res = construir_resumen_ventanas_re545(re, oferta, {k: 40.0}, registrar=silencio)
    return float(pd.to_numeric(res.iloc[0]["Edisp_T"]))


class UltimaVentana(unittest.TestCase):

    def test_la_ultima_ventana_queda_pendiente(self):
        # Sola en la hoja, la ventana es la ultima del mes: Edisp_T = 0.
        self.assertEqual(_edisp(31), 0.0)
        self.assertEqual(_edisp(30), 0.0)

    def test_sin_la_regla_se_pagaria_igual_que_las_demas(self):
        from Script.nucleo import re545_resumen

        with mock.patch.object(re545_resumen, "VENTANA_FINAL_PENDIENTE", False):
            self.assertGreater(_edisp(31), 0.0)
            self.assertEqual(_edisp(31), _edisp(30))


class UltimaVentanaECostos(unittest.TestCase):

    def test_el_monto_de_la_ultima_ventana_es_cero(self):
        from Script.nucleo import ecostos_valorizacion as ev

        df = pd.DataFrame({
            "clave": ["SAE-X"] * 4, "Copia_Ventana": [29, 29, 30, 30],
            "BA": [100.0, 100.0, 100.0, 100.0], "U": [0.0] * 4,
        })
        monto = ev.calcular_monto_a_compensar(df, registrar=silencio)
        self.assertEqual(list(monto), [100.0, 100.0, 0.0, 0.0])
        with mock.patch.object(ev, "VENTANA_FINAL_PENDIENTE", False):
            self.assertEqual(
                list(ev.calcular_monto_a_compensar(df, registrar=silencio)),
                [100.0] * 4,
            )


if __name__ == "__main__":
    unittest.main()
