# -*- coding: utf-8 -*-
"""
Detalle horario de la hoja Ofertas SSCC. Pedido del usuario (2026-11-20):
"1 y cero no mas, todas las horas del mes de cada central, en la primera
tabla, para que los coordinados vean el detalle".
"""

import unittest

import pandas as pd

from Script.nucleo.ofertas_sscc import agregar_horas, columna_hora


class TestDetalleHorario(unittest.TestCase):

    def _resumen(self, filas, faltantes):
        df = pd.DataFrame(filas, columns=["Nombre", "Año", "Mes", "Día", "Oferta completa"])
        df.attrs["faltantes"] = faltantes
        return df

    def test_una_hora_que_falta_y_un_dia_que_no_esta(self):
        wxy = pd.DataFrame({
            "Nombre": ["BAT_X"] * 3, "Central": ["SAE-X"] * 3, "Dia": [15, 16, 17],
            "Oferta completa": [1, 0, 0],
        })
        resumen = self._resumen(
            [["BAT_X", 2026, 10, 15, 1], ["BAT_X", 2026, 10, 16, 0]],
            [{"Nombre": "BAT_X", "Dia": 16, "Servicio": "CSF(+)_RS", "Horas": [15]}],
        )
        df = agregar_horas(wxy, resumen, 2026, 10)
        horas = [columna_hora(h) for h in range(1, 25)]
        self.assertEqual(list(df.columns), ["Nombre", "Central", "Dia", *horas, "Oferta completa"])
        self.assertEqual(df.loc[0, horas].sum(), 24)                 # dia 15 completo
        self.assertEqual(df.loc[1, "H15"], 0)                         # la hora que falta
        self.assertEqual(df.loc[1, horas].sum(), 23)
        self.assertEqual(df.loc[2, horas].sum(), 0)                   # no esta en el archivo

    def test_el_dia_del_cambio_de_hora_deja_vacia_la_que_no_existe(self):
        from Script.nucleo.cambio_hora import horas_del_dia

        # Septiembre 2026: un dia de 23 horas.
        dia = next(d for d in range(1, 31) if len(horas_del_dia(2026, 9, d)) == 23)
        falta = (set(range(1, 25)) - set(horas_del_dia(2026, 9, dia))).pop()
        wxy = pd.DataFrame({"Nombre": ["BAT_X"], "Central": ["SAE-X"], "Dia": [dia],
                            "Oferta completa": [1]})
        resumen = self._resumen([["BAT_X", 2026, 9, dia, 1]], [])
        df = agregar_horas(wxy, resumen, 2026, 9)
        self.assertTrue(pd.isna(df.loc[0, columna_hora(falta)]))


if __name__ == "__main__":
    unittest.main()
