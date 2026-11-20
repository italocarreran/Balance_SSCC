# -*- coding: utf-8 -*-
"""
Detalle horario de la hoja Ofertas SSCC (Cuadro N° 3). Pedido del usuario
(2026-11-20): jerarquico, mostrando como un cero en una hora puede tocar dos
ventanas (la oferta completa es por dia y la ventana usa dos dias).
"""

import unittest

import pandas as pd

from Script.nucleo.ofertas_sscc import (
    COLUMNAS_DETALLE_HORARIO, construir_detalle_horario, _rangos,
)


class TestDetalleHorario(unittest.TestCase):

    def setUp(self):
        # Octubre: el dia 16 le falta CSF(+) en la hora 15 (periodo 15 >= 10:
        # es de la ventana 16), pero el dia entero queda sin oferta completa
        # y caen las ventanas 15 y 16.
        dias = range(1, 32)
        self.wxy = pd.DataFrame({
            "Nombre": ["BAT_X"] * 31, "Central": ["SAE-X"] * 31, "Dia": list(dias),
            "Oferta completa": [0 if d == 16 else 1 for d in dias],
        })
        self.resumen = pd.DataFrame({
            "Central": ["SAE-X"] * 31, "Ventana T": list(range(0, 31)),
            "Oferta": [0] * 31,
            "Completa": [0 if v in (15, 16) else 1 for v in range(0, 31)],
        })
        self.faltantes = [{"Nombre": "BAT_X", "Dia": 16, "Servicio": "CSF(+)_RS", "Horas": [15]}]

    def test_un_cero_toca_dos_ventanas(self):
        df = construir_detalle_horario(self.wxy, self.resumen, self.faltantes, 2026, 10)
        self.assertEqual(list(df.columns), list(COLUMNAS_DETALLE_HORARIO))
        texto = df.to_string()
        self.assertIn("29 de 31 ventanas con oferta completa", texto)
        # Las dos ventanas incompletas muestran el dia 16 y la hora que falta.
        v15 = df.index[df["Ventana"] == "Ventana 15"][0]
        v16 = df.index[df["Ventana"] == "Ventana 16"][0]
        self.assertIn("INCOMPLETA", df.loc[v15, "Detalle"])
        self.assertIn("el día 16", df.loc[v15, "Detalle"])
        self.assertIn("el día 16", df.loc[v16, "Detalle"])
        bloque = df.loc[v15:v16]
        dia16 = bloque[bloque["Día"] == "Día 16"].iloc[0]
        self.assertIn("ventanas 15 y 16", dia16["Detalle"])
        hora = bloque[bloque[COLUMNAS_DETALLE_HORARIO[3]] == "15"].iloc[0]
        self.assertIn("falta CSF(+)_RS", hora["Detalle"])
        self.assertIn("ventana 16", hora["Detalle"])
        # Una ventana completa va en una sola fila.
        self.assertIn("completa", df.loc[df["Ventana"] == "Ventana 3", "Detalle"].iloc[0])

    def test_rangos(self):
        self.assertEqual(_rangos([1, 2, 3, 7, 9, 10]), "01–03, 07, 09–10")


if __name__ == "__main__":
    unittest.main()
