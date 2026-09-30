"""La hoja COMPENSACION_CENTRAL y el cruce del Resumen.

Dos cosas distintas: que la compensacion se agrupe por central y por el
ciclo/ventana que le corresponde a cada hoja de calculo, y que el
Resumen cruce RECIBE y PAGA en una sola fila por empresa aunque los dos
lados escriban el nombre distinto.
"""

import unittest

import pandas as pd

from Script import nucleo


RESUMEN_BESS = pd.DataFrame({
    "Nombre activo": ["BESS UNO", "BESS DOS", "BESS TRES"],
    "Propietario": ["COLBUN S.A.", "ENGIE", ""],
})


class TestCompensacionCentral(unittest.TestCase):
    def test_agrupa_por_central_y_por_ciclo_o_ventana(self):
        ecostos = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS UNO", "BESS UNO", "BESS DOS"],
            "Ciclo de Carga del mes": [1, 1, 2, 1],
            "Monto a compensar": [50.0, 50.0, 30.0, 90.0],
        })
        # RE545 agrupa por "Ventana de valorizacion", no por el ciclo.
        re545 = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS UNO"],
            "Ventana de valorizacion": [3, 3],
            "Ciclo de Carga del mes": [9, 8],
            "Monto a compensar": [20.0, 5.0],
        })

        por_ecostos, por_re545, total = nucleo.construir_compensacion_central(
            ecostos, re545, RESUMEN_BESS
        )

        self.assertEqual(
            list(por_ecostos.columns),
            ["Central", "Empresa", "Ciclo / Ventana", "Compensación [$]"],
        )
        # BESS UNO ciclo 1 junta las dos filas de 50.
        uno_ciclo_1 = por_ecostos[
            (por_ecostos["Central"] == "BESS UNO")
            & (por_ecostos["Ciclo / Ventana"] == 1)
        ]
        self.assertEqual(len(uno_ciclo_1), 1)
        self.assertAlmostEqual(uno_ciclo_1["Compensación [$]"].iloc[0], 100.0)
        self.assertEqual(len(por_re545), 1)  # una sola ventana
        self.assertAlmostEqual(por_re545["Compensación [$]"].iloc[0], 25.0)

        por_empresa = dict(zip(total["Empresa"], total["Compensación Total [$]"]))
        self.assertAlmostEqual(por_empresa["COLBUN S.A."], 155.0)  # 130 + 25
        self.assertAlmostEqual(por_empresa["ENGIE"], 90.0)

    def test_central_sin_propietario_queda_a_su_propio_nombre(self):
        ecostos = pd.DataFrame({
            "Configuracion": ["BESS TRES"],
            "Ciclo de Carga del mes": [1],
            "Monto a compensar": [10.0],
        })
        avisos = []
        _, _, total = nucleo.construir_compensacion_central(
            ecostos, pd.DataFrame(), RESUMEN_BESS, registrar=avisos.append
        )
        self.assertEqual(total["Empresa"].tolist(), ["BESS TRES"])
        self.assertTrue(any("sin Propietario" in m for m in avisos), avisos)

    def test_el_total_por_empresa_no_pierde_ni_inventa_plata(self):
        ecostos = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS DOS", "BESS TRES"],
            "Ciclo de Carga del mes": [1, 1, 1],
            "Monto a compensar": [100.0, 90.0, 10.0],
        })
        re545 = pd.DataFrame({
            "Configuracion": ["BESS UNO"],
            "Ventana de valorizacion": [1],
            "Monto a compensar": [20.0],
        })
        _, _, total = nucleo.construir_compensacion_central(
            ecostos, re545, RESUMEN_BESS
        )
        self.assertAlmostEqual(total["Compensación Total [$]"].sum(), 220.0)


class TestResumen(unittest.TestCase):
    def test_cruza_por_nombre_normalizado(self):
        compensacion = pd.DataFrame({
            "Empresa": ["COLBUN S.A."],
            "Compensación Total [$]": [150.0],
        })
        # La prorrata del CEN escribe el nombre a su manera.
        pagos = pd.DataFrame({
            "Suministrador": ["colbun  s.a.", "OTRA EMPRESA"],
            "Total a pagar [$]": [200.0, 55.0],
        })
        resumen = nucleo.construir_resumen(compensacion, pagos)

        self.assertEqual(len(resumen), 2)  # no se parte Colbun en dos filas
        colbun = resumen.set_index("NOMBRE").loc["COLBUN S.A."]
        self.assertAlmostEqual(colbun["RECIBE"], 150.0)
        self.assertAlmostEqual(colbun["PAGA"], 200.0)
        self.assertAlmostEqual(colbun["NETO"], -50.0)

    def test_quien_solo_paga_o_solo_recibe_igual_aparece(self):
        compensacion = pd.DataFrame({
            "Empresa": ["A", "C"], "Compensación Total [$]": [150.0, 20.0],
        })
        pagos = pd.DataFrame({
            "Suministrador": ["A", "B"], "Total a pagar [$]": [100.0, 70.0],
        })
        resumen = nucleo.construir_resumen(compensacion, pagos).set_index("NOMBRE")
        self.assertEqual(set(resumen.index), {"A", "B", "C"})
        self.assertAlmostEqual(resumen.loc["A", "NETO"], 50.0)
        self.assertAlmostEqual(resumen.loc["B", "NETO"], -70.0)
        self.assertAlmostEqual(resumen.loc["C", "NETO"], 20.0)
        self.assertEqual(
            list(resumen.columns), ["RECIBE", "PAGA", "NETO"]
        )

    def test_avisa_cuando_lo_que_se_recibe_no_es_lo_que_se_paga(self):
        compensacion = pd.DataFrame({
            "Empresa": ["A"], "Compensación Total [$]": [1000.0],
        })
        pagos = pd.DataFrame({
            "Suministrador": ["B"], "Total a pagar [$]": [400.0],
        })
        avisos = []
        nucleo.construir_resumen(compensacion, pagos, registrar=avisos.append)
        self.assertTrue(
            any("no se repartio" in m for m in avisos), avisos
        )


if __name__ == "__main__":
    unittest.main()
