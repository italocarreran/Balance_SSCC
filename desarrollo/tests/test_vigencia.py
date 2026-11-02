# -*- coding: utf-8 -*-
"""Lo que quedo viejo pasa de "ok" a "pendiente", en cadena.

Pedido del usuario: si se modifica una entrada (las medidas, el CMg...),
todo lo que esta "ok" y depende de ella tiene que pasar a pendiente en
la ventana principal, y lo que sale de eso tambien.
"""

import json
import os
import tempfile
import time
import unittest

from Script import nucleo
from Script.nucleo.sellos import leer_sellos, ruta_sellos
from desarrollo.tests.test_pasos_sueltos import armar_caso, correr_pagos


def _estados(base):
    _, filas = nucleo.revisar_estructura(base, "2607")
    return {f["id"]: f for f in filas}


class TestVigencia(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = armar_caso(self.tmp.name)
        correr_pagos(self.base, {"ecostos", "re545", "compensacion_central"}, [])
        # Punto de partida al dia: entradas en T0, cmg.xlsx en T0+10 y
        # las hojas escritas en T0+20.
        self.t0 = time.time() - 1000
        for ruta in self.base.rglob("*"):
            if ruta.is_file():
                os.utime(ruta, (self.t0, self.t0))
        self.cmg = self.base / "Cmg" / nucleo.ARCHIVO_CMG
        os.utime(self.cmg, (self.t0 + 10, self.t0 + 10))
        libro = self.base / nucleo.ARCHIVO_SALIDA
        sellos = {h: self.t0 + 20 for h in nucleo.ORDEN_HOJAS_SALIDA}
        ruta_sellos(libro).write_text(json.dumps(sellos), encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def test_se_sellan_las_hojas_escritas(self):
        antes = time.time()
        correr_pagos(self.base, {"ecostos"}, [])
        sellos = leer_sellos(self.base / nucleo.ARCHIVO_SALIDA)
        self.assertGreaterEqual(sellos[nucleo.HOJA_CALCULO_ECOSTOS], antes)
        # la que no se escribio conserva su sello
        self.assertEqual(sellos[nucleo.HOJA_CALCULO_RE545], self.t0 + 20)

    def test_una_hoja_sin_sello_no_avisa(self):
        # Un libro de antes de los sellos: no se puede saber, no se avisa.
        ruta_sellos(self.base / nucleo.ARCHIVO_SALIDA).unlink()
        os.utime(self.cmg, (self.t0 + 50, self.t0 + 50))
        filas = _estados(self.base)
        self.assertFalse([i for i, f in filas.items() if f.get("desactualizada")])

    def test_todo_al_dia_no_marca_nada(self):
        filas = _estados(self.base)
        self.assertFalse([i for i, f in filas.items() if f.get("desactualizada")])
        self.assertEqual(filas["pagos:ecostos"]["estado"], "ok")

    def test_cambiar_el_cmg_deja_viejo_todo_lo_que_sale_de_el(self):
        os.utime(self.cmg, (self.t0 + 50, self.t0 + 50))

        filas = _estados(self.base)
        viejas = {i for i, f in filas.items() if f.get("desactualizada")}
        self.assertEqual(viejas, {
            "consolidado:cmg", "pagos:ecostos", "pagos:re545",
            "pagos:compensacion_central",
        })
        self.assertEqual(filas["consolidado:cmg"]["estado"], "pendiente")
        self.assertIn("cmg.xlsx", filas["consolidado:cmg"]["detalle"])
        self.assertIn("hoja 'CMg'", filas["pagos:ecostos"]["detalle"])
        # lo que no depende del CMg sigue ok
        self.assertEqual(filas["consolidado:subastas"]["estado"], "ok")
        self.assertEqual(filas["consolidado:medidores"]["estado"], "ok")

        # El plan de "Ejecutar todo" ya lo propone rehacer
        plan = nucleo.planificar(self.base, "2607")
        propuestas = {t["id"] for t in plan["tareas"] if t["seleccionada"]}
        self.assertTrue(viejas <= propuestas, viejas - propuestas)

        # Se rehace la hoja CMg: queda ok, pero el calculo sigue viejo
        nucleo.generar_consolidado(
            self.base, "2607", {"cmg"}, registrar=lambda *_: None
        )
        filas = _estados(self.base)
        self.assertEqual(filas["consolidado:cmg"]["estado"], "ok")
        self.assertIn("cambio hoja 'CMg'", filas["pagos:ecostos"]["detalle"])

        # Y se rehace el calculo: todo al dia otra vez
        correr_pagos(self.base, {"ecostos", "re545", "compensacion_central"}, [])
        filas = _estados(self.base)
        self.assertFalse([i for i, f in filas.items() if f.get("desactualizada")])


if __name__ == "__main__":
    unittest.main()
