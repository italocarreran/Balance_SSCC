# -*- coding: utf-8 -*-
"""
La carpeta del mes compartida por el BESS y el CRA (pedido del usuario
2026-11-16): lo propio de cada uno en BESS/ y CRA/, lo comun (FD y FMA,
Prorrata retiros) en la raiz, y los casos de antes siguen andando sin
mover nada.
"""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from Script import cra  # noqa: E402
import Script.nucleo as nucleo  # noqa: E402


def _nada(*_a, **_k):
    pass


class TestMesNuevo(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.mes = Path(self._tmp.name) / "2608"

    def tearDown(self):
        self._tmp.cleanup()

    def test_los_dos_en_la_misma_carpeta(self):
        nucleo.crear_estructura_caso(self.mes, registrar=_nada)
        cra.crear_carpetas_caso(self.mes)

        for sub in ("BESS/Medidas", "BESS/Auxiliares", "BESS/Subastas/DB subastas",
                    "CRA/Auxiliares", "CRA/Energia", "CRA/SC y CO",
                    "Cmg", "FD y FMA", "Prorrata retiros"):
            self.assertTrue((self.mes / sub).is_dir(), sub)

        # Nada propio suelto en la raiz.
        for suelta in ("Medidas", "Auxiliares", "Energia", "FD", "Log",
                       "BESS/Cmg"):
            self.assertFalse((self.mes / suelta).exists(), suelta)

        bess = nucleo.resolver_rutas(self.mes, "2608")
        rutas_cra = cra.resolver_rutas(self.mes, "2608")

        # Las comunes son las mismas carpetas para los dos.
        self.assertEqual(bess["sscc_desempeno_dir"], rutas_cra["fd"])
        self.assertEqual(bess["prorrata_retiros_dir"], rutas_cra["prorrata"])
        self.assertEqual(rutas_cra["fd"], self.mes / "FD y FMA")
        self.assertEqual(bess["cmg"], self.mes / "Cmg" / "cmg.xlsx")
        self.assertEqual(rutas_cra["cmg"], bess["cmg_dir"])

        # Planillas y Log, cada una en lo suyo.
        self.assertEqual(bess["salida"], self.mes / "BESS" / "11_PAGOS_BESS_2608.xlsx")
        self.assertEqual(bess["control"].parent, self.mes / "BESS" / "Log")
        self.assertEqual(
            rutas_cra["salida"], self.mes / "CRA" / "5_REMUNERACION_CRA_2608.xlsx"
        )
        self.assertEqual(rutas_cra["registro"].parent, self.mes / "CRA" / "Log")

    def test_el_cra_primero_y_despues_el_bess(self):
        # Un mes que empezo el CRA: el BESS no lo toma por un caso de antes
        # (CRA/Auxiliares no es una carpeta del BESS) y se pone al lado.
        cra.crear_carpetas_caso(self.mes)
        self.assertEqual(nucleo.raiz_bess(self.mes), self.mes / "BESS")
        nucleo.crear_estructura_caso(self.mes, registrar=_nada)
        self.assertTrue((self.mes / "BESS" / "Medidas").is_dir())
        self.assertEqual(nucleo.faltan_subcarpetas(self.mes), [])

    def test_el_arbol_muestra_de_quien_es_cada_carpeta(self):
        nucleo.crear_estructura_caso(self.mes, registrar=_nada)
        cra.crear_carpetas_caso(self.mes)

        _, filas = nucleo.revisar_estructura(str(self.mes), "2608")
        textos = {f["id"]: f["etiqueta"] for f in filas}
        self.assertEqual(textos["medidas_dir"], "BESS/Medidas/")
        self.assertEqual(textos["sscc_dir"], "FD y FMA/")
        self.assertEqual(textos["cmg_dir"], "Cmg/")

        textos_cra = [f["texto"] for f in cra.revisar_estructura(self.mes, "2608")]
        self.assertIn("CRA/Energia/", textos_cra)
        self.assertIn("FD y FMA/", textos_cra)
        self.assertIn("CRA/5_REMUNERACION_CRA_2608.xlsx", textos_cra)


class TestCasosDeAntes(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.base = Path(self._tmp.name) / "Balance BESS 2607"

    def tearDown(self):
        self._tmp.cleanup()

    def test_bess_con_todo_suelto_sigue_igual(self):
        for sub in ("Medidas", "Auxiliares", "Ofertas", "Cmg", "Subastas",
                    "SSCC_Desempeño", "Prorrata retiros"):
            (self.base / sub).mkdir(parents=True)

        rutas = nucleo.resolver_rutas(self.base, "2607")
        self.assertEqual(rutas["medidas_dir"], self.base / "Medidas")
        self.assertEqual(rutas["centrales"], self.base / "Auxiliares" / "Centrales.xlsx")
        self.assertEqual(rutas["sscc_desempeno_dir"], self.base / "SSCC_Desempeño")
        self.assertEqual(rutas["salida"], self.base / "11_PAGOS_BESS_2607.xlsx")
        self.assertEqual(rutas["log_dir"], self.base / "Log")
        self.assertEqual(rutas["cmg_dir"], self.base / "Cmg")

        # Completarlo no le crea BESS/ ni "FD y FMA" vacias al lado.
        creadas = nucleo.crear_estructura_caso(self.base, registrar=_nada)
        self.assertFalse((self.base / "BESS").exists())
        self.assertFalse((self.base / "FD y FMA").exists())
        self.assertEqual(
            [c.relative_to(self.base).as_posix() for c in creadas],
            ["Subastas/DB subastas"],
        )

    def test_bess_con_solo_la_planilla_suelta(self):
        self.base.mkdir(parents=True)
        (self.base / "Balance_BESS.xlsx").write_bytes(b"")
        self.assertEqual(nucleo.raiz_bess(self.base), self.base)

    def test_cra_con_todo_suelto_y_su_fd(self):
        for sub in ("Auxiliares", "Energia", "FP", "CO", "SC y CO", "FD",
                    "Prorrata retiros"):
            (self.base / sub).mkdir(parents=True)
        (self.base / "Balance_CRA.xlsx").write_bytes(b"planilla")

        rutas = cra.resolver_rutas(self.base, "2608")
        self.assertEqual(rutas["energia"], self.base / "Energia")
        self.assertEqual(rutas["fd"], self.base / "FD")
        # La planilla de antes se renombra con el periodo, en su lugar.
        self.assertEqual(rutas["salida"], self.base / "5_REMUNERACION_CRA_2608.xlsx")
        self.assertEqual(rutas["salida"].read_bytes(), b"planilla")
        self.assertFalse((self.base / "Balance_CRA.xlsx").exists())

        # Solo le falta la comun Cmg/ (en la raiz, como en un mes nuevo).
        self.assertEqual(cra.crear_carpetas_caso(self.base), ["Cmg"])
        self.assertFalse((self.base / "CRA").exists())

        # Energia/ suelta es del CRA: el BESS de ese mes va en BESS/.
        self.assertEqual(nucleo.raiz_bess(self.base), self.base / "BESS")


if __name__ == "__main__":
    unittest.main()
