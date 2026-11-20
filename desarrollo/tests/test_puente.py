# -*- coding: utf-8 -*-
"""
Por que cambio lo que recibe cada central, por causas, recalculando
(Script/nucleo/puente.py). Pedido del usuario (2026-11-20): "como el precio
de una cuerda: la bencina 200 mas cara, el hilo 100 mas caro [...] en total
1000". La suma de las causas tiene que dar exacto la diferencia.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl

from Script.nucleo import puente


def _libro(ruta, hojas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, filas in hojas.items():
        ws = wb.create_sheet(nombre)
        for fila in filas:
            ws.append(fila)
    wb.save(ruta)
    return ruta


class TestPuente(unittest.TestCase):

    def test_las_causas_suman_la_diferencia_y_salta_lo_igual(self):
        # Montos de cada estado: anterior, recalculada, (sin reemplazos: no
        # hay), cambia el CMg (+30 en E Costos), cambian Subastas (-10 en
        # RE545), con los reemplazos de la actual (-5), la guardada (+1).
        anterior = {"A": {"E Costos": 100.0, "RE545": 50.0}}
        recalculada = {"A": {"E Costos": 102.0, "RE545": 50.0}}
        con_cmg = {"A": {"E Costos": 132.0, "RE545": 50.0}}
        con_subastas = {"A": {"E Costos": 132.0, "RE545": 40.0}}
        con_reemplazos = {"A": {"E Costos": 132.0, "RE545": 35.0}}
        guardada = {"A": {"E Costos": 133.0, "RE545": 35.0}}
        recalculos = iter([recalculada, con_cmg, con_subastas, con_reemplazos])
        cambian = {("CMg",), ("Subastas",)}

        with tempfile.TemporaryDirectory() as carpeta:
            a = Path(carpeta) / "11_PAGOS_BESS_2610_Def.xlsx"
            b = Path(carpeta) / "11_PAGOS_BESS_2610_Pre.xlsx"
            a.write_bytes(b"x")
            b.write_bytes(b"x")
            with mock.patch.object(puente, "montos_por_central",
                                   side_effect=lambda r: guardada if "Def" in str(r) else anterior), \
                    mock.patch.object(puente, "_recalcular", side_effect=lambda *x: next(recalculos)), \
                    mock.patch.object(puente, "_pasar_hojas"), \
                    mock.patch.object(puente, "_hojas_iguales",
                                      side_effect=lambda x, y, hojas: tuple(hojas) not in cambian), \
                    mock.patch.object(puente, "_tiene_reemplazos",
                                      side_effect=lambda r: "Def" in str(r)):
                por_central, _, _, _ = puente.explicar_por_causas(a, b, registrar=lambda *_: None)

        causas = {c: sum(v.values()) for c, v in por_central["A"].items()}
        self.assertAlmostEqual(causas[puente.CALCULO], 2.0)
        self.assertAlmostEqual(causas["CMg"], 30.0)
        self.assertAlmostEqual(causas["Subastas adjudicadas"], -10.0)
        self.assertAlmostEqual(causas[puente.REEMPLAZOS], -5.0)
        self.assertAlmostEqual(causas[puente.OTRA_VERSION], 1.0)
        self.assertAlmostEqual(causas["Medidas (energía)"], 0.0)
        self.assertAlmostEqual(sum(causas.values()), 168.0 - 150.0)

        arbol = puente.arbol_por_central(por_central, anterior, guardada)
        self.assertTrue(arbol[puente.COLUMNA_ARBOL].iloc[0].startswith("A  ·  recibe"))
        self.assertEqual(list(arbol[puente.COLUMNA_ARBOL])[1], "CMg")
        mes = puente.tabla_mes(por_central)
        self.assertAlmostEqual(mes["Aporte [CLP]"].iloc[-1], 18.0)

    def test_hojas_iguales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            a = _libro(Path(carpeta) / "a.xlsx", {"CMg": [[1, 2]], "FD": [[1]]})
            b = _libro(Path(carpeta) / "b.xlsx", {"CMg": [[1, 2]], "FD": [[2]]})
            self.assertTrue(puente._hojas_iguales(a, b, ("CMg",)))
            self.assertFalse(puente._hojas_iguales(a, b, ("FD",)))
            self.assertFalse(puente._hojas_iguales(a, b, ("Subastas",)))


if __name__ == "__main__":
    unittest.main()
