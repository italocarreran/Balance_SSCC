# -*- coding: utf-8 -*-
"""
La matriz de control: cada central y si se encontro en cada origen.

Pedido del usuario (2026-11-03): una matriz en el control de corrida con
las centrales en las filas y los origenes en las columnas.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo.matriz_centrales import construir_matriz_centrales

DICCIONARIO = pd.DataFrame([
    [None] * 5,
    ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
    ["SAE-UNO", "SAE-UNO", "BAT_UNO", "BESS UNO OF", "Uno - BESS"],
])
RESUMEN = pd.DataFrame({
    "Nombre activo": ["SAE-UNO", "SAE-DOS"], "Barra inyección": ["B1", ""],
})


def matriz():
    return construir_matriz_centrales(
        resumen=RESUMEN, diccionario=DICCIONARIO,
        df_medidores=pd.DataFrame({
            "clave": ["SAE-UNO"] * 3 + ["SAE-DOS"], "SoC": [0.5, 0.5, None, None],
        }),
        df_wxy=pd.DataFrame({
            "Nombre": ["BESS UNO OF", "BESS UNO OF", "SAE-DOS", "SAE-DOS"],
            "Central": ["SAE-UNO", "SAE-UNO", "SAE-DOS", "SAE-DOS"],
            "Dia": [1, 2, 1, 2], "Oferta completa": [1, 0, 0, 0],
        }),
        df_cmg=pd.DataFrame({"A": [""], "B": [""], "C": [""], "D": ["B1"]}),
        df_subastas=pd.DataFrame({"Configuración": ["BAT_UNO", "BAT_UNO"]}),
        # La segunda fila es la hora que no existe (cambio de hora): no cuenta.
        df_fd_csf=pd.DataFrame({"Central": ["SAE-UNO"] * 2, "Hora Mes": [1, None]}),
        df_fd_cpf=pd.DataFrame({"Central": []}),
        tablas_fma={"cpf": pd.DataFrame({"Central": ["Uno - BESS"]})},
    ).set_index("Central")


class TestMatrizCentrales(unittest.TestCase):
    def test_marca_lo_encontrado_por_los_mismos_cruces(self):
        m = matriz()
        uno, dos = m.loc["SAE-UNO"], m.loc["SAE-DOS"]
        self.assertEqual(uno["Diccionario"], "Sí")
        self.assertEqual(dos["Diccionario"], "No")
        self.assertEqual(uno["Medidores (Medidas_SAE)"], "Sí (3 cuartos)")
        self.assertEqual(uno["SoC"], "Sí (2 cuartos)")
        self.assertEqual(dos["SoC"], "No")
        self.assertEqual(uno["CMg (por barra)"], "Sí (B1)")
        self.assertEqual(dos["CMg (por barra)"], "No (sin barra)")
        # Subastas y FMA CPF se encuentran por el nombre del Diccionario.
        self.assertEqual(uno["Subastas"], "Sí (2 filas)")
        self.assertEqual(uno["FMA CPF"], "Sí (1 horas)")
        self.assertEqual(uno["FD CSF"], "Sí (1 horas)")
        self.assertEqual(uno["FD CPF"], "No")
        self.assertEqual(uno["Ofertas SSCC"], "Sí (1 de 2 días completos)")
        self.assertTrue(dos["Ofertas SSCC"].startswith("No"))

    def test_va_al_control_pintada(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.CARPETA_LOG / nucleo.ARCHIVO_CONTROL
            nucleo.escribir_control(
                ruta, df_matriz_centrales=matriz().reset_index()
            )
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_MATRIZ_CENTRALES]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        fila = next(r for r in range(4, 6) if ws[f"B{r}"].value == "SAE-DOS")
        celda = ws[f"{nombres['Diccionario']}{fila}"]
        self.assertEqual(celda.value, "No")
        self.assertEqual(celda.fill.start_color.rgb[-6:], "F8CBAD")
        self.assertTrue(str(ws["B2"].value).startswith("Cuadro N° 1"))


if __name__ == "__main__":
    unittest.main()
