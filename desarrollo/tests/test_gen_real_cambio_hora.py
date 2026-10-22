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


def _horario(fechas_horas, valor):
    filas = []
    for fecha, horas in fechas_horas:
        for h in horas:
            filas.append({
                "fecha_consulta": fecha, gr.COLUMNA_TOPOLOGY: "BESS X (inyeccion)",
                gr.COLUMNA_HORA: h, gr.COLUMNA_VALOR: valor(fecha, h),
                gr.COLUMNA_ID: 1,
            })
    return pd.DataFrame(filas)


CENTRALES = [{"topologyName": "BESS X (inyeccion)", "clave": "SAE-X",
              "factor": 1, "unidad": "kwh"}]


class TestGenRealComoVieneLaApi(unittest.TestCase):
    """Con la forma real de la API (prueba_opreal_horas.xlsx, 2026)."""

    def _calendario(self, *args, **kwargs):
        from Script.Medidas import Claves_Balance as cb
        from desarrollo.tests.test_claves_balance import descarga
        df = cb.normalizar_fechas(descarga(*args, **kwargs))
        df, _ = cb.descartar_hora_inexistente(df)
        return cb.construir_calendario(df)

    def test_abril_la_hora_25_es_la_segunda_pasada_de_las_23(self):
        cal = self._calendario("2026-04-01", "2026-05-01", "2026-04-05 00:00",
                               ofset_antes=-3, ofset_despues=-4)
        valor = lambda f, h: 100.0 * h + (1000 if f == "2026-04-05" else 0)
        horario = _horario(
            [("2026-04-04", range(1, 26)), ("2026-04-05", range(1, 25))], valor
        )
        log = []
        cuartos = gr.expandir_a_cuartos(horario, CENTRALES, registrar=log.append)
        df = gr.pegar_calendario(cuartos, cal, registrar=log.append)
        sab = df[df["Dia"] == 4]
        self.assertEqual(sorted(sab["Hora"].unique()), list(range(1, 26)))
        # hora 24 (primera 23:00) = 2400/4 por cuarto, hora 25 (segunda) = 2500/4
        self.assertEqual(sab[sab["Hora"] == 24]["Gen_Unidad"].tolist(), [600.0] * 4)
        self.assertEqual(sab[sab["Hora"] == 25]["Gen_Unidad"].tolist(), [625.0] * 4)
        # el domingo 00:00 tiene solo su propia hora 1 (antes se le sumaba la 25)
        dom = df[(df["Dia"] == 5) & (df["Hora"] == 1)]
        self.assertEqual(dom["Gen_Unidad"].tolist(), [275.0] * 4)
        self.assertFalse(any("fuera del calendario" in l for l in log), log)
        self.assertAlmostEqual(df["Gen_Unidad"].sum(), horario[gr.COLUMNA_VALOR].sum())

    def test_septiembre_la_hora_1_viene_en_cero_y_queda_fuera(self):
        cal = self._calendario("2026-09-01", "2026-10-01", "2026-09-06 00:00")
        valor = lambda f, h: 0.0 if (f == "2026-09-06" and h == 1) else 10.0 * h
        horario = _horario(
            [("2026-09-05", range(1, 25)), ("2026-09-06", range(1, 25))], valor
        )
        log = []
        cuartos = gr.expandir_a_cuartos(horario, CENTRALES, registrar=log.append)
        df = gr.pegar_calendario(cuartos, cal, registrar=log.append)
        dom = df[df["Dia"] == 6]
        self.assertEqual(sorted(dom["Hora"].unique()), list(range(2, 25)))
        # la 01:00 del 6-09 (hora 2 local) trae la hora 2 de la API
        self.assertEqual(dom[dom["Hora"] == 2]["Gen_Unidad"].tolist(), [5.0] * 4)
        aviso = next(l for l in log if "fuera del calendario" in l)
        self.assertIn("06-09 00:00", aviso)
        self.assertIn("venian en cero", aviso)
