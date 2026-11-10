# -*- coding: utf-8 -*-
"""
La unidad de cada columna entre corchetes, solo en lo que se ve.

Pedido del usuario: "agregar unidad de medida [...] entre corchetes
[]. Si es CLP o USD igual". La hoja muestra "Gen_Unidad [kWh]"; el
programa sigue leyendo "Gen_Unidad".
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo.unidades import NOMBRES_CON_UNIDAD, sin_unidad


class TestUnidades(unittest.TestCase):
    def test_medidores_cmg_y_subastas_muestran_la_unidad_y_se_leen_igual(self):
        medidores = pd.DataFrame({
            "Mes": [7], "Gen_Unidad": [5.0], "SoC": [0.5],
        })
        cmg = pd.DataFrame({"D_Barra": ["B"], "F_CMg": [50.0]})
        subastas = pd.DataFrame({
            "Concepto": ["CPF(+)"], "Configuración": ["X"],
            "Energía SSCC": [1.0],
        })
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            nucleo.escribir_salida(
                medidores, ruta, [], [], df_cmg=cmg, df_subastas=subastas,
                hojas_regenerar={
                    nucleo.HOJA_MEDIDORES, nucleo.HOJA_CMG,
                    nucleo.HOJA_SUBASTAS,
                },
                registrar=lambda *a: None,
            )
            libro = openpyxl.load_workbook(ruta)
            vistos = {
                hoja: [c.value for c in libro[hoja][3]]
                for hoja in (nucleo.HOJA_MEDIDORES, nucleo.HOJA_CMG,
                             nucleo.HOJA_SUBASTAS)
            }
            leida = nucleo.leer_hoja_con_fila_vacia(ruta, nucleo.HOJA_MEDIDORES)

        self.assertIn("Gen_Unidad [kWh]", vistos[nucleo.HOJA_MEDIDORES])
        self.assertIn("SoC [p.u.]", vistos[nucleo.HOJA_MEDIDORES])
        self.assertIn("F_CMg [CLP/kWh]", vistos[nucleo.HOJA_CMG])
        self.assertIn("Energía SSCC [MW]", vistos[nucleo.HOJA_SUBASTAS])
        self.assertEqual(list(leida.columns), ["Mes", "Gen_Unidad", "SoC"])

    def test_cada_nombre_visible_vuelve_a_su_nombre_interno(self):
        for hoja, nombres in NOMBRES_CON_UNIDAD.items():
            visibles = list(nombres.values())
            self.assertEqual(len(visibles), len(set(visibles)), hoja)
            for interno, visible in nombres.items():
                self.assertTrue(visible.endswith("]"), visible)
                self.assertEqual(sin_unidad(visible, hoja), interno)
                # Un libro de antes (sin unidad) se lee igual.
                self.assertEqual(sin_unidad(interno, hoja), interno)
            self.assertFalse(
                any("[$]" in v for v in visibles), hoja
            )  # pesos como CLP


class TestCentralesSinEnergiaMinima(unittest.TestCase):
    def test_la_hoja_centrales_no_trae_energia_minima(self):
        resumen = pd.DataFrame({
            "Nombre activo": ["SAE UNO"], "Pmax (MW)": [10],
            "Capacidad (MWh)": [40], "Energía mínima": [5],
        })
        diccionario = pd.DataFrame([["Balance_BESS", "FD"], ["SAE UNO", "U"]])
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            nucleo.escribir_salida(
                None, ruta, [], [], hojas_regenerar={nucleo.HOJA_CENTRALES},
                centrales=(resumen, diccionario), registrar=lambda *a: None,
            )
            encabezado = [
                c.value for c in openpyxl.load_workbook(ruta)[
                    nucleo.HOJA_CENTRALES][3]
            ]
            leido, _ = nucleo.leer_centrales_planilla(ruta)

        self.assertNotIn("Energía mínima", encabezado)
        self.assertIn("Capacidad [MWh]", encabezado)
        self.assertEqual(
            list(leido.columns), ["Nombre activo", "Pmax (MW)", "Capacidad (MWh)"]
        )


    def test_tampoco_trae_nup(self):
        """
        Pedido del usuario (2026-11-10): agrego la columna NUP en
        'Resumen BESS' de Centrales.xlsx; se descarta al importar a la 11_.
        """

        resumen = pd.DataFrame({
            "Nombre activo": ["SAE UNO"], "NUP": [12345], "Pmax (MW)": [10],
            "Capacidad (MWh)": [40],
        })
        diccionario = pd.DataFrame([["Balance_BESS", "FD"], ["SAE UNO", "U"]])
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            nucleo.escribir_salida(
                None, ruta, [], [], hojas_regenerar={nucleo.HOJA_CENTRALES},
                centrales=(resumen, diccionario), registrar=lambda *a: None,
            )
            encabezado = [
                c.value for c in openpyxl.load_workbook(ruta)[
                    nucleo.HOJA_CENTRALES][3]
            ]
            leido, _ = nucleo.leer_centrales_planilla(ruta)

        self.assertNotIn("NUP", encabezado)
        self.assertEqual(
            list(leido.columns), ["Nombre activo", "Pmax (MW)", "Capacidad (MWh)"]
        )


if __name__ == "__main__":
    unittest.main()
