# -*- coding: utf-8 -*-
"""Herramientas/Buscar_Barras_CMg.py: encontrar la barra exacta con un nombre aproximado."""

import tempfile
import unittest
from pathlib import Path

import sys

import pandas as pd

# Las herramientas viven en Herramientas/, hermana de Script/.
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "Herramientas"))
import Buscar_Barras_CMg as bb  # noqa: E402

REALES = [
    "ARENA_________220", "TOCOPILLA_____110", "PEQ___________220",
    "PVHFAMI_______033", "DONHUMBERTO___033", "LA_CABANA_____220",
    "ANDES_SNG_____220", "FUTURO________033", "VICTOR_JARA___220",
    "GABRIELA______220", "ESTEPA________220", "PVSOLDESIERTO_033",
]
RELLENO = [f"BARRA_{i:04d}____{220 if i % 2 else 110}" for i in range(1000)]
BARRAS = REALES + RELLENO + ["ARENAL________110", "TOCOPILLA_____220"]


class TestBuscar(unittest.TestCase):

    def primera(self, texto):
        return bb.buscar_barras(BARRAS, texto)[0][0]

    def test_nombres_aproximados(self):
        for texto, esperado in (
            ("arena 220", "ARENA_________220"),
            ("Tocopilla 110", "TOCOPILLA_____110"),
            ("tocopila 110", "TOCOPILLA_____110"),       # mal escrito
            ("victor jara", "VICTOR_JARA___220"),
            ("jara victor", "VICTOR_JARA___220"),        # otro orden
            ("donhumberto", "DONHUMBERTO___033"),
            ("cabaña", "LA_CABANA_____220"),             # con tilde/ñ
            ("desierto", "PVSOLDESIERTO_033"),
            ("gabriela", "GABRIELA______220"),
            ("barra 12 110", "BARRA_0012____110"),       # 12 = 0012
        ):
            self.assertEqual(self.primera(texto), esperado, texto)

    def test_vacio_y_sin_parecido(self):
        self.assertEqual(bb.buscar_barras(BARRAS, "  "), [])
        self.assertEqual(bb.buscar_barras(BARRAS, "zzzzqqq"), [])

    def test_lee_el_csv_del_coordinador(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "cmg2609_def_15minutal.csv"
            pd.DataFrame({
                "FECHA": [20260901] * 3, "HORA": [1] * 3, "MINUTO": [0] * 3,
                "BARRA": ["ARENA_________220 ", "TOCOPILLA_____110", "ARENA_________220 "],
                "CMg[CLP/KWh]": ["1,5", "2,5", "3,0"],
            }).to_csv(ruta, sep=";", index=False, encoding="latin1")
            self.assertEqual(
                bb.leer_barras(ruta),
                {"ARENA_________220": 2, "TOCOPILLA_____110": 1},
            )


if __name__ == "__main__":
    unittest.main()
