# -*- coding: utf-8 -*-
"""Menos idas a config.json desde la ventana (plan de la ventana, Fase 5.5).

config.json es compartido con los otros programas y puede estar en red:
recordar una carpeta que ya estaba anotada no escribe nada, y lo que se
escribe junto va en una sola escritura.
"""

import unittest
from unittest import mock

from desarrollo.tests.test_ventana_rueda import tk


@unittest.skipUnless(tk is not None, "sin tkinter")
class TestRecordarCarpeta(unittest.TestCase):
    def setUp(self):
        import Balance_BESS as app

        self.app = app
        self.cfg = {}
        self.escrituras = []

        def guardar(datos):
            self.escrituras.append(dict(datos))
            self.cfg.update(datos)

        self.parches = [
            mock.patch.object(app, "leer_config", lambda: dict(self.cfg)),
            mock.patch.object(app, "guardar_config", guardar),
        ]
        for parche in self.parches:
            parche.start()

    def tearDown(self):
        for parche in self.parches:
            parche.stop()

    def test_la_primera_vez_escribe(self):
        self.app.recordar_carpeta("2609", "D:/Caso 2609")
        self.assertEqual(
            self.escrituras,
            [{self.app.CLAVE_CARPETAS_PERIODO: {"2609": "D:/Caso 2609"}}],
        )

    def test_si_ya_estaba_anotada_no_escribe(self):
        self.app.recordar_carpeta("2609", "D:/Caso 2609")
        self.escrituras.clear()
        self.app.recordar_carpeta("2609", "D:/Caso 2609")
        self.assertEqual(self.escrituras, [])

    def test_una_carpeta_es_de_un_solo_periodo(self):
        self.app.recordar_carpeta("2608", "D:/Caso")
        self.app.recordar_carpeta("2609", "D:/Caso")
        self.assertEqual(
            self.cfg[self.app.CLAVE_CARPETAS_PERIODO], {"2609": "D:/Caso"}
        )

    def test_lo_de_ademas_va_en_la_misma_escritura(self):
        self.app.recordar_carpeta(
            "2609", "D:/Caso 2609", ademas={"carpeta_base": "D:/Caso 2609"}
        )
        self.assertEqual(len(self.escrituras), 1)
        self.assertEqual(self.escrituras[0]["carpeta_base"], "D:/Caso 2609")
        # Ya anotada: solo va lo de ademas.
        self.escrituras.clear()
        self.app.recordar_carpeta(
            "2609", "D:/Caso 2609", ademas={"carpeta_base": "D:/Caso 2609"}
        )
        self.assertEqual(self.escrituras, [{"carpeta_base": "D:/Caso 2609"}])

    def test_sin_periodo_solo_lo_de_ademas(self):
        self.app.recordar_carpeta("", "D:/Caso", ademas={"carpeta_base": "D:/Caso"})
        self.assertEqual(self.escrituras, [{"carpeta_base": "D:/Caso"}])


if __name__ == "__main__":
    unittest.main()
