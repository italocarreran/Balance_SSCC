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


def _filas(dia, horas, servicio="CPF_RS", mes=9):
    return [
        ["BESS UNO", 2026, mes, dia, hora, None, None, servicio, "Sí"]
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


class TestOfertaCompletaAbril(unittest.TestCase):
    """El sabado 4-04-2026 tiene 25 horas: la columna E va de 1 a 25."""

    def test_el_sabado_tiene_25_horas(self):
        self.assertEqual(
            ofertas_sscc.horas_del_dia(2026, 4, 4), frozenset(range(1, 26))
        )

    def test_con_la_hora_25_la_oferta_es_completa(self):
        oferta = _resumen(
            _filas(4, range(1, 26), mes=4) + _filas(5, range(1, 25), mes=4)
        )
        self.assertEqual(oferta[4], 1)
        self.assertEqual(oferta[5], 1)

    def test_sin_la_hora_25_el_sabado_es_incompleto(self):
        oferta = _resumen(_filas(4, range(1, 25), mes=4))
        self.assertEqual(oferta[4], 0)

    def test_la_25_no_cuenta_por_otra_hora(self):
        oferta = _resumen(_filas(4, list(range(1, 24)) + [25], mes=4))
        self.assertEqual(oferta[4], 0)


def _ventanas(anio, mes, dias, sin=(), extra=()):
    """Medidores de una central con toda la oferta: (Dia, Hora) x 4."""

    from Script.nucleo.medidores import calcular_ventana

    filas = [
        (d, h) for d in range(1, dias + 1) for h in range(1, 25)
        if (d, h) not in sin
    ]
    filas += list(extra)
    filas.sort()
    df = pd.DataFrame(
        [f for f in filas for _ in range(4)], columns=["Dia", "Hora"]
    )
    df["clave"] = "SAE A"
    ventana = calcular_ventana(df["clave"], df["Hora"])
    resumen = ofertas_sscc.construir_resumen_ventana_oferta(
        df["clave"], ventana, [1] * len(df), registrar=lambda *_: None,
        anio=anio, mes=mes,
    )
    return resumen.set_index("Ventana T")


class TestVentanaDelCambioDeHora(unittest.TestCase):
    """La ventana que cruza el cambio de hora no tiene 96 cuartos."""

    def test_septiembre_la_ventana_5_tiene_92_cuartos(self):
        resumen = _ventanas(2026, 9, 30, sin={(6, 1)})
        self.assertEqual(resumen.loc[5, "Oferta"], 92)
        self.assertTrue((resumen["Completa"] == 1).all())

    def test_abril_la_ventana_4_tiene_100_cuartos(self):
        resumen = _ventanas(2026, 4, 30, extra=[(4, 25)])
        self.assertEqual(resumen.loc[4, "Oferta"], 100)
        self.assertTrue((resumen["Completa"] == 1).all())

    def test_sin_una_hora_ofertada_sigue_incompleta(self):
        resumen = _ventanas(2026, 9, 30, sin={(6, 1), (6, 3)})
        self.assertEqual(resumen.loc[5, "Completa"], 0)
        self.assertEqual(int((resumen["Completa"] == 0).sum()), 1)

    def test_mes_sin_cambio_igual_que_antes(self):
        resumen = _ventanas(2026, 7, 31)
        self.assertTrue((resumen["Completa"] == 1).all())


if __name__ == "__main__":
    unittest.main()
