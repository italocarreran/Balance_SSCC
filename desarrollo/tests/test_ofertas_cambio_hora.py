# -*- coding: utf-8 -*-
"""Oferta completa el dia del cambio a horario de verano.

El 6-09-2026 el reloj salta de 00:00 a 01:00: la hora 1 no existe y las
ofertas no la traen. Exigirla dejaba la oferta del dia como incompleta.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script.nucleo import ofertas_sscc


def _filas(dia, horas, servicio="CPF_RS"):
    return [
        ["BESS UNO", 2026, 9, dia, hora, None, None, servicio, "Sí"]
        for hora in horas
    ]


def _resumen(filas):
    with tempfile.TemporaryDirectory() as carpeta:
        ruta = Path(carpeta) / "ofertas.xlsx"
        encabezado = [["Nombre", "Año", "Mes", "Día", "Hora",
                       "x", "y", "Servicio", "Ofertado"]]
        pd.DataFrame(encabezado + filas).to_excel(
            ruta, index=False, header=False
        )
        df = ofertas_sscc.construir_resumen_ofertas_sscc(
            ruta, registrar=lambda *_: None
        )
    return dict(zip(df["Día"], df["Oferta completa"]))


class TestHorasDelDia(unittest.TestCase):
    def test_el_dia_del_cambio_no_tiene_hora_1(self):
        self.assertEqual(
            ofertas_sscc.horas_del_dia(2026, 9, 6), frozenset(range(2, 25))
        )

    def test_los_demas_dias_tienen_24(self):
        for fecha in ((2026, 9, 5), (2026, 9, 7), (2026, 4, 5),
                      (2026, 7, 15)):
            self.assertEqual(
                ofertas_sscc.horas_del_dia(*fecha), frozenset(range(1, 25)),
                fecha,
            )


class TestOfertaCompletaCambioDeHora(unittest.TestCase):
    def test_sin_la_hora_1_el_dia_6_la_oferta_es_completa(self):
        oferta = _resumen(_filas(6, range(2, 25)) + _filas(7, range(1, 25)))
        self.assertEqual(oferta[6], 1)
        self.assertEqual(oferta[7], 1)

    def test_otro_dia_sin_la_hora_1_sigue_incompleto(self):
        oferta = _resumen(_filas(5, range(2, 25)))
        self.assertEqual(oferta[5], 0)

    def test_el_dia_6_le_falta_otra_hora_es_incompleto(self):
        oferta = _resumen(_filas(6, range(3, 25)))
        self.assertEqual(oferta[6], 0)


if __name__ == "__main__":
    unittest.main()
