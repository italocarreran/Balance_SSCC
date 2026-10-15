# -*- coding: utf-8 -*-
"""Deteccion automatica de los cambios de hora a partir del AAMM."""

import datetime
import unittest
from unittest import mock

from Script.nucleo import cambio_hora


class TestCambiosDeHora(unittest.TestCase):
    def tearDown(self):
        cambio_hora._cambios_del_mes.cache_clear()

    def test_septiembre_adelanta_y_no_existe_la_hora_1(self):
        (cambio,) = cambio_hora.cambios_del_periodo("2609")
        self.assertEqual(cambio.fecha, datetime.date(2026, 9, 6))
        self.assertEqual(cambio.tipo, "adelanta")
        self.assertEqual(cambio.horas_del_dia, 23)
        self.assertEqual(cambio.periodos_inexistentes, (1,))
        self.assertEqual(cambio.utc_anterior, datetime.timedelta(hours=-4))
        self.assertEqual(cambio.utc_nuevo, datetime.timedelta(hours=-3))

    def test_abril_atrasa_y_se_repite_la_hora_24_del_sabado(self):
        (cambio,) = cambio_hora.cambios_del_periodo("2604")
        self.assertEqual(cambio.fecha, datetime.date(2026, 4, 4))
        self.assertEqual(cambio.tipo, "atrasa")
        self.assertEqual(cambio.horas_del_dia, 25)
        self.assertEqual(cambio.periodos_repetidos, (24,))

    def test_un_mes_sin_cambio(self):
        self.assertEqual(cambio_hora.cambios_del_periodo("2607"), [])
        self.assertIn("Sin cambio", cambio_hora.describir_periodo("2607"))

    def test_detecta_otros_anios_sin_fechas_escritas(self):
        self.assertEqual(
            [c.fecha for c in cambio_hora.cambios_de_hora(2027, 9)],
            [datetime.date(2027, 9, 5)],
        )
        self.assertEqual(
            [c.fecha for c in cambio_hora.cambios_de_hora(2029, 4)],
            [datetime.date(2029, 4, 7)],
        )

    def test_horas_del_dia(self):
        self.assertEqual(
            cambio_hora.horas_del_dia(2026, 9, 6), frozenset(range(2, 25))
        )
        self.assertEqual(
            cambio_hora.horas_del_dia(2026, 4, 4), frozenset(range(1, 25))
        )

    def test_sin_tzdata_la_regla_da_lo_mismo(self):
        con_zona = {
            aamm: cambio_hora.cambios_del_periodo(aamm)
            for aamm in (
                f"{anio:02d}{mes:02d}" for anio in range(26, 36)
                for mes in (4, 7, 9)
            )
        }
        cambio_hora._cambios_del_mes.cache_clear()
        with mock.patch.object(cambio_hora, "_zona", lambda: None):
            for aamm, esperado in con_zona.items():
                obtenido = cambio_hora.cambios_del_periodo(aamm)
                self.assertEqual(
                    [(c.fecha, c.periodos_inexistentes, c.periodos_repetidos)
                     for c in obtenido],
                    [(c.fecha, c.periodos_inexistentes, c.periodos_repetidos)
                     for c in esperado],
                    aamm,
                )
            self.assertIn("tzdata", cambio_hora.describir_periodo("2609"))


if __name__ == "__main__":
    unittest.main()
