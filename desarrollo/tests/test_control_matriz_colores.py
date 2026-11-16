"""La matriz de Control_corrida.xlsx conserva sus colores cuando otra
pasada (Log, Reporte) reescribe el control sin ella."""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from Script.nucleo.escritura import escribir_control
from Script.nucleo.matriz_centrales import construir_matriz_centrales


def _color(celda):
    return celda.fill.start_color.rgb if celda.fill.fill_type else None


class MatrizConservaColores(unittest.TestCase):

    def test_log_despues_de_la_matriz_no_borra_colores(self):
        medidores = pd.DataFrame({
            "clave": ["SAE-A"] * 4, "Gen_Unidad": [1.0] * 4,
            "SoC": [0.5] * 4,
        })
        matriz = construir_matriz_centrales(df_medidores=medidores)

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "Control_corrida.xlsx"
            escribir_control(ruta, df_matriz_centrales=matriz)
            antes = openpyxl.load_workbook(ruta)
            hoja = next(n for n in antes.sheetnames)
            colores_antes = [_color(c) for c in antes[hoja][4]]
            panel = antes[hoja].freeze_panes

            escribir_control(ruta, df_log=pd.DataFrame({"aviso": ["x"]}))

            despues = openpyxl.load_workbook(ruta)
            self.assertIn("Log", despues.sheetnames)
            self.assertEqual([_color(c) for c in despues[hoja][4]], colores_antes)
            self.assertTrue(any(colores_antes))
            self.assertEqual(despues[hoja].freeze_panes, panel)


if __name__ == "__main__":
    unittest.main()
