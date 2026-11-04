# -*- coding: utf-8 -*-
"""
RE545 en otro proceso (en paralelo con E Costos) da lo mismo que en
serie: el mismo libro y el mismo registro, en el mismo orden. Y si el
otro proceso no puede correr, se calcula en serie sin que la corrida
falle.

Pedido del usuario (2026-11-03): "si se corren hojas independientes que
se puedan hacer en paralelo que se haga [...] mismos resultados".
"""

import tempfile
import unittest
from concurrent.futures.process import BrokenProcessPool
from pathlib import Path
from unittest import mock

import pandas as pd

from Script import nucleo
from Script.nucleo import proceso
from desarrollo.tests.test_control_corrida import PagosBessEndToEndTest


def _correr(carpeta, paralelo):
    base, _ = PagosBessEndToEndTest.armar_caso(None, carpeta)
    lineas = []
    with mock.patch.object(proceso, "FILAS_PARA_PARALELO", 0 if paralelo else 10**9):
        nucleo.generar_pagos_bess(
            base, {"ecostos", "re545"}, registrar=lineas.append
        )
    hojas = pd.read_excel(base / nucleo.ARCHIVO_SALIDA, sheet_name=None, header=None)
    sin_hora = [l for l in lineas if "escritura:" not in l and "Corrida:" not in l]
    return hojas, sin_hora


class TestParalelo(unittest.TestCase):
    def test_re545_aparte_da_lo_mismo_que_en_serie(self):
        with tempfile.TemporaryDirectory() as uno, tempfile.TemporaryDirectory() as dos:
            serie, registro_serie = _correr(uno, paralelo=False)
            paralelo, registro_paralelo = _correr(dos, paralelo=True)
        self.assertEqual(list(serie), list(paralelo))
        for nombre in serie:
            pd.testing.assert_frame_equal(serie[nombre], paralelo[nombre], obj=nombre)
        self.assertEqual(
            [l for l in registro_serie if not str(l).startswith(("Archivo generado", "Listo:"))],
            [l for l in registro_paralelo if not str(l).startswith(("Archivo generado", "Listo:"))],
        )

    def test_si_el_otro_proceso_se_cae_se_hace_en_serie(self):
        class Roto:
            def result(self):
                raise BrokenProcessPool("no arranco")

        aparte = proceso._Aparte(Roto(), lambda x, registrar: (registrar("hecho aca"), x * 2)[1], (21,))
        lineas = []
        self.assertEqual(proceso._resultado_aparte(aparte, lineas.append), 42)
        self.assertEqual(lineas, ["hecho aca"])

    def test_un_error_de_los_datos_se_levanta_igual(self):
        class ConError:
            def result(self):
                raise nucleo.ErrorEntrada("dato malo")

        aparte = proceso._Aparte(ConError(), None, ())
        with self.assertRaises(nucleo.ErrorEntrada):
            proceso._resultado_aparte(aparte, lambda *_: None)

    def test_bitacora_repite_lineas_y_alertas_en_orden(self):
        from Script.nucleo.alertas import ALTA, Alerta, Registro

        bitacora = proceso.Bitacora()
        alerta = Alerta("SUB-011", ALTA, "Calculo RE545", "texto")
        bitacora("uno")
        nucleo.anotar(bitacora, alerta)
        registro = Registro(salida=lambda *_: None)
        proceso.Bitacora.repetir(bitacora.eventos, registro)
        self.assertEqual(len(registro.alertas), 1)
        self.assertEqual(registro.lineas[:1], ["uno"])


if __name__ == "__main__":
    unittest.main()
