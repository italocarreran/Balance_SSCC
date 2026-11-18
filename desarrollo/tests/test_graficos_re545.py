# -*- coding: utf-8 -*-
"""
Graficos RE545 (Script/nucleo/graficos_re545.py): la pagina se arma con lo
que trae la planilla, aunque las columnas con formula vengan sin valor
guardado (planilla escrita por Python y no abierta en Excel).
"""

import importlib.util
import json
import re
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script.nucleo import graficos_re545
from Script.nucleo.escritura import escribir_pagos_bess
from Script.nucleo.re545 import renombrar_calculo_re545

RAIZ = Path(__file__).resolve().parents[2]


def _ejemplo():
    """El caso de explicacion/generar_ejemplo.py (SAE-DEMO, 12 bloques)."""

    ruta = RAIZ / "explicacion" / "generar_ejemplo.py"
    spec = importlib.util.spec_from_file_location("generar_ejemplo", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo.calcular_re545()


def _planilla(carpeta):
    re545, resumen = _ejemplo()
    ruta = Path(carpeta) / "11_PAGOS_BESS_2607.xlsx"
    escribir_pagos_bess(ruta, df_re545=renombrar_calculo_re545(re545),
                        df_resumen_re545=resumen, registrar=lambda *a: None)
    from openpyxl import load_workbook
    libro = load_workbook(ruta)
    hoja = libro.create_sheet("Centrales")
    hoja.append([])
    hoja.append(["Cuadro N° 1 — Resumen BESS", None, None, None, None, None, None,
                 "Cuadro N° 2 — Diccionario"])
    hoja.append(["Nombre activo", "Propietario", "Pmax (MW)", "Capacidad (MWh)",
                 "Barra inyección", "Eficiencia", None, "Nombre"])
    hoja.append(["SAE-DEMO", "DEMO SpA", 10, 40, "B-DEMO", 0.85, None, "SAE-DEMO"])
    libro.save(ruta)
    return ruta, re545


class TestGraficosRE545(unittest.TestCase):

    def test_datos_de_la_ventana(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta, re545 = _planilla(carpeta)
            datos = graficos_re545.armar_datos(ruta, registrar=lambda *a: None)
        self.assertEqual(datos["periodo"], "2607")
        [central] = datos["centrales"]
        self.assertEqual(central["nombre"], "SAE-DEMO")
        self.assertEqual((central["pmax"], central["capacidad"]), (10, 40))
        [ventana] = central["ventanas"]
        self.assertEqual(ventana["n"], 12)
        self.assertEqual((ventana["eini"], ventana["ealm"]), (8000, 5100))
        self.assertEqual(ventana["edisp"], 13100)
        self.assertEqual(ventana["reserva"], 13230)
        # C1 y C2 son formulas sin valor guardado: se calculan igual.
        self.assertAlmostEqual(ventana["c1"], 1_167_500)
        self.assertAlmostEqual(ventana["c2"], 1_005_500)
        self.assertAlmostEqual(ventana["monto"], 162_000)
        self.assertEqual(ventana["t"]["au"], [1102.5] * 12)
        self.assertEqual(ventana["t"]["bv"], [0] * 12)
        self.assertEqual(sorted(ventana["srv"]), ["CPF(+)", "CSF(−)"])
        self.assertEqual(ventana["srv"]["CPF(+)"]["mw"], [3] * 12)
        self.assertAlmostEqual(
            sum(ventana["srv"]["CPF(+)"]["kwh"]) + sum(ventana["srv"]["CSF(−)"]["kwh"]),
            13230)

    def test_pagina(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta, _ = _planilla(carpeta)
            destino = graficos_re545.generar_html(
                ruta, registrar=lambda *a: None, plotly_js="/* plotly */")
            self.assertEqual(destino, Path(carpeta) / "Log" / "Graficos_RE545_2607.html")
            html = destino.read_text(encoding="utf-8")
        self.assertNotIn("/*__DATOS__*/", html)
        self.assertNotIn("/*__PLOTLY__*/", html)
        self.assertIn("<script>/* plotly */</script>", html)
        datos = json.loads(re.search(r"const D = (\{.*?\});\n", html, re.S).group(1))
        self.assertEqual(datos["centrales"][0]["ventanas"][0]["monto"], 162000)

    def test_sin_centrales_se_estiman_pmax_y_capacidad(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta, _ = _planilla(carpeta)
            from openpyxl import load_workbook
            libro = load_workbook(ruta)
            del libro["Centrales"]
            libro.save(ruta)
            datos = graficos_re545.armar_datos(ruta, registrar=lambda *a: None)
        [central] = datos["centrales"]
        self.assertEqual((central["pmax"], central["capacidad"]), (10, 40))
        self.assertEqual(central["estimado"], ["Pmax", "capacidad"])
        self.assertAlmostEqual(central["ventanas"][0]["monto"], 162_000)

    def test_sin_la_hoja(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "11_PAGOS_BESS_2607.xlsx"
            pd.DataFrame({"a": [1]}).to_excel(ruta, index=False)
            with self.assertRaises(graficos_re545.ErrorGraficos):
                graficos_re545.armar_datos(ruta, registrar=lambda *a: None)


if __name__ == "__main__":
    unittest.main()
