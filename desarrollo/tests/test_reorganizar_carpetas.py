# -*- coding: utf-8 -*-
"""Herramientas/Reorganizar_Carpetas.py: un mes de antes pasa a BESS/ y CRA/."""

import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "Herramientas"))

import Reorganizar_Carpetas as rc  # noqa: E402
from Script import cra  # noqa: E402
import Script.nucleo as nucleo  # noqa: E402
from Script.nucleo import sellos  # noqa: E402


def _archivo(ruta, texto="x"):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(texto, encoding="utf-8")
    return ruta


class TestReorganizar(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.mes = Path(self._tmp.name) / "Balance 2609"

    def tearDown(self):
        self._tmp.cleanup()

    def _bess_de_antes(self):
        _archivo(self.mes / "Medidas" / "SOC_2609.xlsx")
        _archivo(self.mes / "Auxiliares" / "Centrales.xlsx")
        _archivo(self.mes / "Ofertas" / "OfertasSSCC.xlsx")
        _archivo(self.mes / "Subastas" / "DB subastas" / "a.accdb")
        _archivo(self.mes / "Cmg" / "cmg.xlsx")
        _archivo(self.mes / "SSCC_Desempeño" / "SSCC_Desempeño_x.xlsx")
        _archivo(self.mes / "Prorrata retiros" / "Prorrata_Retiros_2609_def.xlsx")
        _archivo(self.mes / "Log" / "Control_corrida.xlsx")
        _archivo(self.mes / "11_PAGOS_BESS_2609.xlsx", "planilla")
        _archivo(self.mes / "SAE_2609.csv")

    def test_bess_de_antes_queda_como_uno_nuevo(self):
        self._bess_de_antes()
        antes = nucleo.resolver_rutas(self.mes, "2609")
        sellos_antes = sellos.ruta_sellos(antes["salida"])

        plan = rc.planear(self.mes)
        self.assertFalse(any(m.conflicto for m in plan))
        rc.ejecutar(plan, registrar=lambda *_: None)

        rutas = nucleo.resolver_rutas(self.mes, "2609")
        self.assertEqual(rutas["raiz_bess"], self.mes / "BESS")
        self.assertEqual(rutas["salida"].read_text(encoding="utf-8"), "planilla")
        self.assertTrue(rutas["centrales"].is_file())
        self.assertTrue((rutas["db_subastas_dir"] / "a.accdb").is_file())
        self.assertTrue(rutas["control"].is_file())
        self.assertTrue(rutas["csv_sae"].is_file())
        # Comunes en la raiz; el FD con su nombre nuevo.
        self.assertEqual(rutas["sscc_desempeno_dir"], self.mes / "FD y FMA")
        self.assertTrue((self.mes / "FD y FMA" / "SSCC_Desempeño_x.xlsx").is_file())
        self.assertTrue(rutas["cmg"].is_file())
        self.assertFalse((self.mes / "SSCC_Desempeño").exists())
        # Los sellos siguen siendo los mismos.
        self.assertEqual(sellos.ruta_sellos(rutas["salida"]), sellos_antes)
        # Una segunda pasada no tiene nada que hacer.
        self.assertEqual(rc.planear(self.mes), [])

    def test_cra_de_antes_con_su_fd(self):
        _archivo(self.mes / "Auxiliares" / "centrales_cra.xlsx")
        _archivo(self.mes / "Energia" / "Formato.xlsx")
        _archivo(self.mes / "FP" / "fp_2609.xlsx")
        _archivo(self.mes / "FD" / "SSCC_Desempeño_y.xlsx")
        _archivo(self.mes / "Balance_CRA.xlsx")

        rc.ejecutar(rc.planear(self.mes), registrar=lambda *_: None)

        rutas = cra.resolver_rutas(self.mes, "2609")
        self.assertEqual(rutas["raiz_cra"], self.mes / "CRA")
        self.assertTrue((rutas["auxiliares"] / "centrales_cra.xlsx").is_file())
        self.assertTrue((rutas["fp"] / "fp_2609.xlsx").is_file())
        self.assertEqual(rutas["fd"], self.mes / "FD y FMA")
        self.assertTrue((rutas["fd"] / "SSCC_Desempeño_y.xlsx").is_file())
        self.assertFalse((self.mes / "FD").exists())
        # Balance_CRA.xlsx se movio y despues el CRA lo renombra con el AAMM.
        self.assertEqual(rutas["salida"], self.mes / "CRA" / "5_REMUNERACION_CRA_2609.xlsx")
        self.assertTrue(rutas["salida"].is_file())

    def test_los_dos_de_antes_reparten_auxiliares(self):
        self._bess_de_antes()
        _archivo(self.mes / "Auxiliares" / "centrales_cra.xlsx")
        _archivo(self.mes / "Energia" / "Formato.xlsx")

        rc.ejecutar(rc.planear(self.mes), registrar=lambda *_: None)

        self.assertTrue((self.mes / "BESS" / "Auxiliares" / "Centrales.xlsx").is_file())
        self.assertTrue((self.mes / "CRA" / "Auxiliares" / "centrales_cra.xlsx").is_file())
        self.assertFalse((self.mes / "Auxiliares").exists())
        self.assertFalse((self.mes / "BESS" / "Auxiliares" / "centrales_cra.xlsx").exists())

    def test_con_conflicto_no_se_mueve_nada(self):
        self._bess_de_antes()
        _archivo(self.mes / "FD y FMA" / "otro.xlsx")    # ya hay una FD nueva con algo

        plan = rc.planear(self.mes)
        self.assertTrue(any(m.conflicto for m in plan))
        self.assertIn("CONFLICTO", rc.describir(plan, self.mes))
        with self.assertRaises(ValueError):
            rc.ejecutar(plan)
        self.assertTrue((self.mes / "Medidas").is_dir())
        self.assertFalse((self.mes / "BESS").exists())

    def test_mes_nuevo_no_tiene_nada_que_hacer(self):
        nucleo.crear_estructura_caso(self.mes, registrar=lambda *_: None)
        cra.crear_carpetas_caso(self.mes)
        self.assertEqual(rc.planear(self.mes), [])


if __name__ == "__main__":
    unittest.main()
