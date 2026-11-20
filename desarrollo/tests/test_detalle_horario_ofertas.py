# -*- coding: utf-8 -*-
"""
Listado horario de la hoja Ofertas SSCC (Cuadro N° 1). Pedido del usuario
(2026-11-20): "solo el listado horario por central": Central, Dia, Ventana,
Hora y Oferta completa, para que los coordinados vean el detalle.
"""

import unittest

import pandas as pd

from Script.nucleo.ofertas_sscc import COLUMNAS_POR_HORA, listado_horario


class TestListadoHorario(unittest.TestCase):

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
        df = listado_horario(wxy, resumen, 2026, 10)
        self.assertEqual(list(df.columns), list(COLUMNAS_POR_HORA))
        self.assertEqual(len(df), 3 * 24)
        por = df.set_index(["Dia", "Hora"])
        self.assertEqual(por.loc[(16, 15), "Oferta completa"], 0)
        self.assertEqual(por.loc[(16, 14), "Oferta completa"], 1)
        self.assertEqual(df[df["Dia"] == 17]["Oferta completa"].sum(), 0)   # no esta en el archivo
        # Ventana: la del dia desde el periodo 10, la del dia anterior antes.
        self.assertEqual(por.loc[(16, 9), "Ventana"], 15)
        self.assertEqual(por.loc[(16, 10), "Ventana"], 16)
        self.assertEqual(df["Central"].unique().tolist(), ["SAE-X"])

    def test_el_dia_del_cambio_de_hora_no_trae_la_hora_que_no_existe(self):
        from Script.nucleo.cambio_hora import horas_del_dia

        dia = next(d for d in range(1, 31) if len(horas_del_dia(2026, 9, d)) == 23)
        wxy = pd.DataFrame({"Nombre": ["BAT_X"], "Central": ["SAE-X"], "Dia": [dia],
                            "Oferta completa": [1]})
        df = listado_horario(wxy, self._resumen([["BAT_X", 2026, 9, dia, 1]], []), 2026, 9)
        self.assertEqual(len(df), 23)


if __name__ == "__main__":
    unittest.main()
