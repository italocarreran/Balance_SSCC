# -*- coding: utf-8 -*-
"""Generacion Real: el aviso de filas fuera del calendario dice cuales son.

En la corrida real de 2609 salieron "8 fila(s) ... quedaron fuera del
calendario": 2 centrales x 4 cuartos de la hora 1 del 6-09, que la API
de operacion real numera como las 00:00 (que ese dia no existen). El
aviso ahora dice la hora y si traian energia, porque esa energia no
entra al balance.
"""

import unittest

import pandas as pd

from Script.Medidas import Generacion_Real as gr


def _calendario():
    local = list(pd.date_range("2026-09-05 23:00", periods=4, freq="15min")) + list(
        pd.date_range("2026-09-06 01:00", periods=4, freq="15min")
    )
    return pd.DataFrame({
        "intervalo": local,
        "Cuarto de Hora": range(477, 485),
        "Mes": 9, "Dia": [5] * 4 + [6] * 4, "Hora": [24] * 4 + [2] * 4,
        "Minutos": [0, 15, 30, 45] * 2, "Hora Mes": [120] * 4 + [121] * 4,
    })


class TestGenRealFueraDelCalendario(unittest.TestCase):
    def test_el_aviso_dice_la_hora_y_la_energia(self):
        cuartos = pd.DataFrame({
            "intervalo": list(pd.date_range("2026-09-06 00:00", periods=4, freq="15min"))
            + list(pd.date_range("2026-09-06 01:00", periods=4, freq="15min")),
            "clave": "SAE-X",
            "Gen_Unidad": [250.0] * 4 + [100.0] * 4,
        })
        log = []
        df = gr.pegar_calendario(cuartos, _calendario(), registrar=log.append)
        self.assertEqual(len(df), 4)
        aviso = next(l for l in log if "fuera del calendario" in l)
        self.assertIn("06-09 00:00", aviso)
        self.assertIn("SAE-X", aviso)
        self.assertIn("1,000.000 kWh", aviso)
        self.assertIn("NO entran", aviso)


if __name__ == "__main__":
    unittest.main()
