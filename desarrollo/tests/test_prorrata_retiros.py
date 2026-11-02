"""La hoja PRORRATA_RETIROS: reparto del monto de cada cuarto de hora.

Lo que se prueba aca es el nucleo de la hoja: el monto de un cuarto se
reparte entre las empresas que retiraron en ESE cuarto segun su peso, y
nada raro de la fuente (pesos negativos, cuartos que no suman 1, filas
repetidas) corta la corrida.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo


def _fuente(ruta, filas, encabezados=("Cuarto de Hora", "Suministrador", "Prorrata")):
    df = pd.DataFrame(filas, columns=list(encabezados))
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Prorrata 15min", index=False)
    return ruta


class TestProrrataRetiros(unittest.TestCase):
    def test_reparte_el_monto_de_cada_cuarto_segun_el_peso(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = _fuente(Path(tmp) / "Prorrata_Retiros_2607_def.xlsx", [
                (1, "A", 0.6), (1, "B", 0.4),
                (2, "A", 0.25), (2, "B", 0.75),
            ])
            cargada = nucleo.leer_prorrata_retiros(ruta)

            # El cuarto de hora del mes vive en "Bloque horario" en las
            # dos hojas de calculo.
            ecostos = pd.DataFrame({
                "Bloque horario": [1, 2],
                "Bloque Mes Descarga": [7, 9],  # orden monotono: no se usa
                "Monto a compensar": [100.0, 40.0],
            })
            re545 = pd.DataFrame({
                "Bloque horario": [1, 2],
                "Monto a compensar": [20.0, 60.0],
            })
            por_cuarto, detalle, pagos = nucleo.construir_prorrata_retiros(
                cargada, ecostos, re545
            )

            self.assertEqual(
                por_cuarto["Monto a compensar [$]"].tolist(), [120.0, 100.0]
            )
            self.assertAlmostEqual(detalle["Monto a pagar [$]"].sum(), 220.0)
            por_empresa = dict(zip(
                pagos["Suministrador"], pagos["Total a pagar [$]"]
            ))
            self.assertAlmostEqual(por_empresa["A"], 97.0)   # 72 + 25
            self.assertAlmostEqual(por_empresa["B"], 123.0)  # 48 + 75

    def test_pesos_que_no_suman_uno_igual_reparten_el_cuarto_entero(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = _fuente(Path(tmp) / "Prorrata_Retiros_2607_pre.xlsx", [
                (1, "A", 30.0), (1, "B", 10.0),
            ])
            avisos = []
            cargada = nucleo.leer_prorrata_retiros(ruta, registrar=avisos.append)
            ecostos = pd.DataFrame({
                "Bloque horario": [1], "Monto a compensar": [400.0],
            })
            _, detalle, pagos = nucleo.construir_prorrata_retiros(
                cargada, ecostos, pd.DataFrame()
            )

            self.assertTrue(any("no suman 1" in m for m in avisos), avisos)
            self.assertAlmostEqual(detalle["Monto a pagar [$]"].sum(), 400.0)
            por_empresa = dict(zip(
                pagos["Suministrador"], pagos["Total a pagar [$]"]
            ))
            self.assertAlmostEqual(por_empresa["A"], 300.0)
            self.assertAlmostEqual(por_empresa["B"], 100.0)

    def test_prorrata_negativa_avisa_pero_no_corta_la_corrida(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = _fuente(Path(tmp) / "Prorrata_Retiros_2607_def.xlsx", [
                (1, "A", 1.2), (1, "B", -0.2),
            ])
            avisos = []
            cargada = nucleo.leer_prorrata_retiros(ruta, registrar=avisos.append)

            self.assertTrue(any("negativo" in m for m in avisos), avisos)
            self.assertEqual(len(cargada), 2)

    def test_lee_por_posicion_aunque_cambie_el_texto_del_encabezado(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = _fuente(
                Path(tmp) / "Prorrata_Retiros_2607_def.xlsx",
                [(1, "A", 1.0)],
                encabezados=("CUARTO HORA", "EMPRESA", "PESO"),
            )
            cargada = nucleo.leer_prorrata_retiros(ruta)
            self.assertEqual(
                list(cargada.columns),
                ["Cuarto de Hora", "Suministrador", "Prorrata"],
            )
            self.assertEqual(cargada.loc[0, "Suministrador"], "A")

    def test_cuarto_con_monto_y_sin_prorrata_queda_sin_repartir(self):
        cargada = pd.DataFrame({
            "Cuarto de Hora": [1], "Suministrador": ["A"], "Prorrata": [1.0],
        })
        ecostos = pd.DataFrame({
            "Bloque horario": [1, 2], "Monto a compensar": [100.0, 50.0],
        })
        avisos = []
        por_cuarto, detalle, _ = nucleo.construir_prorrata_retiros(
            cargada, ecostos, pd.DataFrame(), registrar=avisos.append
        )
        self.assertEqual(len(por_cuarto), 2)
        self.assertAlmostEqual(detalle["Monto a pagar [$]"].sum(), 100.0)
        self.assertTrue(any("sin repartir" in m for m in avisos), avisos)

    def test_resumen_une_quien_recibe_y_quien_paga(self):
        compensacion = pd.DataFrame({
            "Empresa": ["A", "C"],
            "Compensación Total [$]": [150.0, 20.0],
        })
        pagos = pd.DataFrame({
            "Suministrador": ["A", "B"], "Total a pagar [$]": [100.0, 70.0]
        })
        resumen = nucleo.construir_resumen(compensacion, pagos).set_index("NOMBRE")
        self.assertEqual(set(resumen.index), {"A", "B", "C"})
        self.assertAlmostEqual(resumen.loc["A", "NETO"], 50.0)
        self.assertAlmostEqual(resumen.loc["B", "NETO"], -70.0)
        self.assertAlmostEqual(resumen.loc["C", "NETO"], 20.0)

    def test_busqueda_exige_un_archivo_del_periodo(self):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp)
            (carpeta / "Prorrata_Retiros_2606_def.xlsx").touch()
            esperado = carpeta / "Prorrata_Retiros_2607_pre.xlsx"
            esperado.touch()
            self.assertEqual(
                nucleo.buscar_archivo_prorrata(carpeta, "2607"), esperado
            )


if __name__ == "__main__":
    unittest.main()
