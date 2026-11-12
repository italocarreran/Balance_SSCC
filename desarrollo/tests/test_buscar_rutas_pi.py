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


class TestGuardarEnCentrales(unittest.TestCase):

    def armar(self, tmp, con_columna):
        ruta = Path(tmp) / "Centrales.xlsx"
        libro = Workbook()
        libro.active.title = "Resumen BESS"
        libro.active.append(["Nombre activo", "Barra inyección"])
        dic = libro.create_sheet("Diccionario")
        dic.append([])
        dic.append(["Balance_BESS", "FD"] + (["SoC_PI"] if con_columna else []))
        dic.append(["SAE-LILE", "SAE LILE"])
        dic.append(["SAE-ARENA", "SAE ARENA"]
                   + ([BASE + "VIEJA"] if con_columna else []))
        libro.save(ruta)
        return ruta

    def test_escribe_en_la_columna_que_ya_esta(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = self.armar(tmp, True)
            anterior = br.guardar_ruta_en_centrales(
                ruta, "sae-arena", BASE + "SAE-ARENA|Nombre PCP/PID")
            self.assertEqual(anterior, BASE + "VIEJA")
            self.assertEqual(
                br.centrales_del_diccionario(ruta),
                [("SAE-LILE", ""), ("SAE-ARENA", BASE + "SAE-ARENA")],
            )

    def test_agrega_la_columna_si_no_esta(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = self.armar(tmp, False)
            br.guardar_ruta_en_centrales(ruta, "SAE-LILE", BASE + "SAE-Lile")
            self.assertEqual(
                br.centrales_del_diccionario(ruta),
                [("SAE-LILE", BASE + "SAE-Lile"), ("SAE-ARENA", "")],
            )
            # El programa la lee igual (y no la toma como un nombre).
            sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
            import Script.nucleo as nucleo
            resumen, dic = nucleo.leer_centrales(ruta)
            mapa = nucleo.construir_homologacion(dic)
            self.assertNotIn(nucleo.normalizar(BASE + "SAE-Lile"), mapa)
            self.assertEqual(
                nucleo.mapa_diccionario(dic, nucleo.ROL_BALANCE_BESS,
                                        nucleo.ROL_SOC_PI),
                {"sae-lile": BASE + "SAE-Lile"},
            )

    def test_central_que_no_esta(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = self.armar(tmp, True)
            with self.assertRaises(ValueError):
                br.guardar_ruta_en_centrales(ruta, "SAE-NADA", BASE + "X")


if __name__ == "__main__":
    unittest.main()
