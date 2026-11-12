# -*- coding: utf-8 -*-
"""Herramientas/Buscar_Rutas_PI.py: busqueda aproximada y Centrales.xlsx."""

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "Herramientas"))
import Buscar_Rutas_PI as br  # noqa: E402

BASE = "\\\\SRV-SCADA-AF2\\SEN\\Generación\\SEN\\"
ELEMENTOS = [
    {"nombre": "SAE-PFV El Manzano", "pcp": "SAE-CRCA-PFV-MANZANO",
     "soc": True, "ruta": BASE + "07 Región RM\\SAE-PFV El Manzano"},
    {"nombre": "SAE-Arica II", "pcp": "SAE-ARICA2", "soc": True,
     "ruta": BASE + "01 Región XV\\SAE-Arica II"},
    {"nombre": "SAE-Lile", "pcp": "SAE-LILE", "soc": True,
     "ruta": BASE + "02 Región I\\SAE-Lile"},
    {"nombre": "SAE-PFV Andes Solar II-B", "pcp": "SAE-CRCA-PFV-ANDES2B",
     "soc": True, "ruta": BASE + "03 Región II\\SAE-PFV Andes Solar II-B"},
    {"nombre": "SAE-PFV Maria Elena", "pcp": "SAE-CRCA-PFV-ELENA",
     "soc": True, "ruta": BASE + "03 Región II\\SAE-PFV Maria Elena"},
] + [
    {"nombre": f"SAE-OTRA {i}", "pcp": f"SAE-OTRA-{i}", "soc": False,
     "ruta": BASE + f"X\\SAE-OTRA {i}"} for i in range(300)
]


class TestBuscar(unittest.TestCase):

    def primero(self, texto):
        return br.buscar(ELEMENTOS, texto)[0][0]["nombre"]

    def test_aproximado(self):
        for texto, esperado in (
            ("manzano", "SAE-PFV El Manzano"),
            ("SAE-CRCA-PFV-MANZANO", "SAE-PFV El Manzano"),
            ("arica 2", "SAE-Arica II"),
            ("SAE-ARICA2", "SAE-Arica II"),
            ("lile", "SAE-Lile"),
            ("andes 2b", "SAE-PFV Andes Solar II-B"),
            ("maria elena", "SAE-PFV Maria Elena"),
            ("SAE-CRCA-PFV-ELENA", "SAE-PFV Maria Elena"),
            ("manzanno", "SAE-PFV El Manzano"),
        ):
            self.assertEqual(self.primero(texto), esperado, texto)

    def test_vacio(self):
        self.assertEqual(br.buscar(ELEMENTOS, "  "), [])

    def test_partes_de_ruta(self):
        self.assertEqual(
            br.partes_de_ruta(BASE + "02 Región I\\SAE-Lile|Nombre PCP/PID"),
            ["Generación", "SEN", "02 Región I", "SAE-Lile"],
        )


class TestCentrales(unittest.TestCase):

    def test_lee_las_rutas_del_diccionario(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "Centrales.xlsx"
            libro = Workbook()
            libro.active.title = "Resumen BESS"
            dic = libro.create_sheet("Diccionario")
            dic.append([])
            dic.append(["Balance_BESS", "FD", "SoC_PI"])
            dic.append(["SAE-LILE", "SAE LILE", None])
            dic.append(["SAE-ARENA", "SAE ARENA", BASE + "SAE-ARENA"])
            dic.append([None, None, None])
            libro.save(ruta)

            self.assertEqual(
                br.centrales_del_diccionario(ruta),
                [("SAE-LILE", ""), ("SAE-ARENA", BASE + "SAE-ARENA")],
            )


if __name__ == "__main__":
    unittest.main()
