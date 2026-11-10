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


    def test_total_por_central_suma_los_ciclos_de_cada_central(self):
        ecostos = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS UNO", "BESS UNO", "BESS DOS"],
            "Ciclo de Carga del mes": [1, 1, 2, 1],
            "Monto a compensar": [50.0, 50.0, 30.0, 90.0],
        })
        por_ecostos, _, total = nucleo.construir_compensacion_central(
            ecostos, pd.DataFrame(), RESUMEN_BESS
        )
        por_central = nucleo.total_por_central(por_ecostos)

        self.assertEqual(
            list(por_central.columns),
            ["Central", "Empresa", "Compensación [$]"],
        )
        montos = dict(zip(por_central["Central"], por_central["Compensación [$]"]))
        self.assertEqual(montos, {"BESS DOS": 90.0, "BESS UNO": 130.0})
        self.assertEqual(
            dict(zip(por_central["Central"], por_central["Empresa"])),
            {"BESS DOS": "ENGIE", "BESS UNO": "COLBUN S.A."},
        )
        # El cuadro intermedio cierra con el total por empresa.
        self.assertAlmostEqual(
            por_central["Compensación [$]"].sum(),
            total["Compensación Total [$]"].sum(),
        )

    def test_total_por_central_sin_detalle_queda_vacio(self):
        vacio = nucleo.total_por_central(pd.DataFrame())
        self.assertTrue(vacio.empty)
        self.assertEqual(
            list(vacio.columns), ["Central", "Empresa", "Compensación [$]"]
        )


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


class TestCuadrosCombinados(unittest.TestCase):
    """E Costos y RE545 lado a lado (antes cuadros 1+3 y 2+4 separados)."""

    def setUp(self):
        ecostos = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS UNO", "BESS DOS"],
            "Ciclo de Carga del mes": [1, 2, 1],
            "Monto a compensar": [100.0, 0.0, 90.0],
        })
        re545 = pd.DataFrame({
            "Configuracion": ["BESS UNO", "BESS UNO"],
            "Ventana de valorizacion": [1, 3],
            "Monto a compensar": [20.0, 5.0],
        })
        self.por_ecostos, self.por_re545, self.total = (
            nucleo.construir_compensacion_central(ecostos, re545, RESUMEN_BESS)
        )

    def test_detalle_sin_empresa_y_ordenado_por_central(self):
        detalle = nucleo.detalle_combinado(self.por_ecostos, self.por_re545)
        self.assertEqual(list(detalle.columns), [
            "Central", "Ciclo / Ventana", "Compensación Ecostos [$]",
            "Compensación RE545 [$]", "Compensación Total [$]",
        ])
        filas = {
            (c, v): (e, r) for c, v, e, r, _ in detalle.itertuples(index=False)
        }
        self.assertEqual(filas[("BESS UNO", 1)], (100.0, 20.0))
        self.assertEqual(filas[("BESS UNO", 3)], (0.0, 5.0))
        # Las ventanas de cada central juntas, con o sin compensacion.
        self.assertEqual(
            list(zip(detalle["Central"], detalle["Ciclo / Ventana"])),
            [("BESS DOS", 1), ("BESS UNO", 1), ("BESS UNO", 2),
             ("BESS UNO", 3)],
        )
        self.assertEqual(filas[("BESS UNO", 2)], (0.0, 0.0))

    def test_por_central_con_empresa_y_total_que_cierra(self):
        central = nucleo.por_central_combinado(self.por_ecostos, self.por_re545)
        self.assertEqual(list(central.columns), [
            "Central", "Empresa", "Compensación Ecostos [$]",
            "Compensación RE545 [$]", "Compensación Total [$]",
        ])
        uno = central.set_index("Central").loc["BESS UNO"]
        self.assertEqual(uno["Empresa"], "COLBUN S.A.")
        self.assertAlmostEqual(uno["Compensación Total [$]"], 125.0)
        self.assertAlmostEqual(
            central["Compensación Total [$]"].sum(),
            self.total["Compensación Total [$]"].sum(),
        )

    def test_un_lado_vacio_queda_en_cero(self):
        central = nucleo.por_central_combinado(self.por_ecostos, pd.DataFrame())
        self.assertTrue(central["Compensación RE545 [$]"].eq(0.0).all())
        vacio = nucleo.detalle_combinado(pd.DataFrame(), None)
        self.assertTrue(vacio.empty)
        self.assertIn("Compensación Total [$]", vacio.columns)


class TestFormatoDeLasHojas(unittest.TestCase):
    """Titulos de cuadro, filtro y comentarios en el libro escrito."""

    def _escribir(self, ruta, detalle, central, empresa):
        cuarto = pd.DataFrame({
            "Cuarto de Hora": [1, 2], "Monto a compensar [$]": [10.0, 20.0],
        })
        reparto = pd.DataFrame({
            "Cuarto de Hora": [1, 2], "Suministrador": ["A", "A"],
            "Prorrata": [1.0, 1.0], "Monto a pagar [$]": [10.0, 20.0],
        })
        pagos = pd.DataFrame({"Suministrador": ["A"], "Total a pagar [$]": [30.0]})
        nucleo.escribir_pagos_bess(
            ruta, df_compensacion_detalle=detalle,
            df_compensacion_central=central, df_compensacion_empresa=empresa,
            df_compensacion_cuarto=cuarto, df_prorrata_retiros=reparto,
            df_pagos_suministrador=pagos,
            hojas_regenerar={"COMPENSACION_CENTRAL", "PRORRATA_RETIROS"},
            registrar=lambda _m: None,
        )

    def test_titulos_y_comentarios_sin_filtro(self):
        import tempfile
        import zipfile
        from pathlib import Path

        import openpyxl

        ecostos = pd.DataFrame({
            "Configuracion": ["BESS UNO"] * 4 + ["BESS DOS"] * 2,
            "Ciclo de Carga del mes": [1, 2, 3, 4, 1, 2],
            "Monto a compensar": [10.0, 20.0, 0.0, 0.0, 5.0, 0.0],
        })
        por_ecostos, por_re545, empresa = nucleo.construir_compensacion_central(
            ecostos, pd.DataFrame(), RESUMEN_BESS
        )
        detalle = nucleo.detalle_combinado(por_ecostos, por_re545)
        central = nucleo.por_central_combinado(por_ecostos, por_re545)

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "salida.xlsx"
            # Dos veces: la segunda reemplaza las hojas (y sus
            # comentarios) en un libro que ya existe.
            self._escribir(ruta, detalle, central, empresa)
            self._escribir(ruta, detalle, central, empresa)

            with zipfile.ZipFile(ruta) as z:
                comentarios = [n for n in z.namelist() if "comments" in n]
            self.assertEqual(len(comentarios), 1, comentarios)

            libro = openpyxl.load_workbook(ruta)
            prorrata = libro["PRORRATA_RETIROS"]
            self.assertIsNone(prorrata["B1"].value)
            self.assertEqual(prorrata.row_dimensions[1].height, 14.5)
            self.assertEqual(prorrata.row_dimensions[2].height, 14.5)
            combinadas = {str(r) for r in prorrata.merged_cells.ranges}
            # Una columna de separacion entre cuadros.
            self.assertEqual(combinadas, {"B2:C2", "E2:H2", "J2:K2"})
            titulo = prorrata["B2"]
            self.assertEqual(titulo.font.color.rgb[-6:], "244062")
            self.assertIsNone(titulo.fill.fill_type)
            self.assertEqual(prorrata["B3"].fill.fill_type, "solid")
            self.assertIn("Pago(e,q)", prorrata["H3"].comment.text)
            self.assertIn("Total(e)", prorrata["K3"].comment.text)
            # La columna del cuarto de hora queda al ancho de su nombre.
            self.assertLess(prorrata.column_dimensions["B"].width, 20)

            compensacion = libro["COMPENSACION_CENTRAL"]
            self.assertIsNone(compensacion["B1"].value)
            # Totales del mes en la fila 1 sobre cada columna de plata,
            # cada cuadro con su largo (pedido del usuario).
            nombres = {
                c.column_letter: str(c.value) for c in compensacion[3] if c.value
            }
            plata = [l for l, n in nombres.items() if n.startswith("Compensación")]
            self.assertTrue(plata)
            filas = {
                l: max(
                    r for r in range(4, compensacion.max_row + 1)
                    if compensacion[f"{l}{r}"].value is not None
                )
                for l in plata
            }
            for letra in plata:
                self.assertEqual(
                    compensacion[f"{letra}1"].value,
                    f"=SUM({letra}4:{letra}{filas[letra]})", letra,
                )
                self.assertEqual(
                    compensacion[f"{letra}1"].number_format, '"Total: "#,##0'
                )
            ultima_empresa = max(plata, key=lambda l: compensacion[f"{l}3"].column)
            self.assertEqual(
                compensacion[f"{ultima_empresa}1"].font.color.rgb[-6:], "375623"
            )
            self.assertEqual(
                compensacion[f"{plata[0]}1"].font.color.rgb[-6:], "244062"
            )
            # Sin plata, sin total (Central, Empresa, Ciclo).
            self.assertIsNone(compensacion["I1"].value)
            self.assertEqual(compensacion["B3"].value, "Central")
            self.assertEqual(compensacion["B3"].fill.fill_type, "solid")
            self.assertEqual(compensacion["H3"].value, "Central")
            self.assertEqual(compensacion["I3"].value, "Empresa")
            self.assertEqual(compensacion["N3"].value, "Empresa")
            self.assertEqual(compensacion["F3"].value, "Compensación Total [CLP]")
            self.assertEqual(prorrata["H3"].value, "Monto a pagar [CLP]")
            self.assertIsNone(compensacion.auto_filter.ref)
            self.assertFalse(any(
                d.hidden for d in compensacion.row_dimensions.values()
            ))
            self.assertEqual(compensacion["B2"].alignment.horizontal, "left")
