# -*- coding: utf-8 -*-
"""
"Ofertas SSCC por dia": una fila por central y dia, sin centrales vacias
ni repetidas.

Pedido del usuario (2026-11-03): en la hoja de ofertas habia nombres sin
central y centrales repetidas. Cada nombre del archivo de ofertas se
busca primero en la columna Ofertas del Diccionario y despues en la de
Subastas; si no esta en ninguna no se agrega.
"""

import unittest

import pandas as pd

from Script.nucleo.ofertas_sscc import (
    agregar_central_ofertas, cargar_resumen_en_medidores,
)

DICCIONARIO = pd.DataFrame([
    [None] * 5,
    ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
    ["SAE-UNO", "BESS UNO", "SAE-UNO-SUB", "BESS UNO OFERTA", "Uno - BESS"],
    ["SAE-DOS", "BESS DOS", "BESS DOS SUBASTA", None, "Dos - BESS"],
    ["SAE-TRES", "BESS TRES", "SAE-TRES", "BESS TRES OF", "Tres - BESS"],
])


def resumen(*filas):
    return pd.DataFrame(
        [(n, 2026, 9, d, o) for n, d, o in filas],
        columns=["Nombre", "Año", "Mes", "Día", "Oferta completa"],
    )


class TestOfertasPorCentral(unittest.TestCase):
    def setUp(self):
        self.df, periodo, self.avisos = cargar_resumen_en_medidores(
            resumen(
                ("BESS UNO OFERTA", 1, 1),     # por la columna Ofertas
                ("BESS DOS SUBASTA", 1, 1),    # no esta en Ofertas: por Subastas
                ("BESS DESCONOCIDA", 1, 1),    # en ninguna: no se agrega
            ),
            ["SAE-UNO", "SAE-DOS", "SAE-TRES"],  # claves de Medidores
            DICCIONARIO, registrar=lambda *_: None,
        )
        self.assertEqual(periodo, (2026, 9))

    def test_una_fila_por_central_y_dia_con_su_central(self):
        por_central = self.df.groupby("Central").size().to_dict()
        self.assertEqual(por_central, {"SAE-UNO": 30, "SAE-DOS": 30, "SAE-TRES": 30})
        self.assertFalse(self.df["Central"].eq("").any())
        nombres = dict(zip(self.df["Central"], self.df["Nombre"]))
        self.assertEqual(nombres["SAE-UNO"], "BESS UNO OFERTA")
        self.assertEqual(nombres["SAE-DOS"], "BESS DOS SUBASTA")
        # Sin ofertas en el archivo: con oferta 0, una sola vez.
        self.assertEqual(nombres["SAE-TRES"], "SAE-TRES")
        tres = self.df[self.df["Central"] == "SAE-TRES"]
        self.assertEqual(int(tres["Oferta completa"].sum()), 0)

    def test_el_nombre_desconocido_no_se_agrega_y_se_avisa(self):
        self.assertNotIn("BESS DESCONOCIDA", set(self.df["Nombre"]))
        self.assertTrue(any("BESS DESCONOCIDA" in a for a in self.avisos))

    def test_la_oferta_queda_en_su_central_y_dia(self):
        uno = self.df[(self.df["Central"] == "SAE-UNO") & (self.df["Dia"] == 1)]
        self.assertEqual(int(uno["Oferta completa"].iloc[0]), 1)

    def test_la_central_ya_viene_puesta_en_la_hoja(self):
        con_central = agregar_central_ofertas(self.df, DICCIONARIO)
        self.assertEqual(
            list(con_central.columns), ["Nombre", "Central", "Dia", "Oferta completa"]
        )


if __name__ == "__main__":
    unittest.main()
