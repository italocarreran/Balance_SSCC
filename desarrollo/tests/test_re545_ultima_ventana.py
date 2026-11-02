# -*- coding: utf-8 -*-
"""
La ultima ventana del mes se paga (parcial) aunque sea la 31: el .xlsm
anulaba la Energia Disponible con (AX<>31), y solo en meses de 31 dias.
"""

import unittest

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
    re = completar_calculo_re545(re, subastas, 0.1, {k: 40.0}, {k: 0.85}, registrar=silencio)
    oferta = pd.DataFrame([{"Central": CENTRAL, "Ventana T": ventana, "Oferta": 8, "Completa": 1}])
    res = construir_resumen_ventanas_re545(re, oferta, {k: 40.0}, registrar=silencio)
    return float(pd.to_numeric(res.iloc[0]["Edisp_T"]))


class UltimaVentana(unittest.TestCase):

    def test_la_ventana_31_se_paga_igual_que_las_demas(self):
        self.assertGreater(_edisp(31), 0.0)
        self.assertEqual(_edisp(31), _edisp(30))


if __name__ == "__main__":
    unittest.main()
