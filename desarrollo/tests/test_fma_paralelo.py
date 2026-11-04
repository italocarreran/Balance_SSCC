# -*- coding: utf-8 -*-
"""
FMA CPF y CSF leidos en paralelo (pedido del usuario 2026-11-03): un mes
entero de reportes diarios da EXACTAMENTE lo mismo, con los mismos avisos
y en el mismo orden, que leyendolos uno por uno.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from Script.Fd import Indices_FMA as fma
from desarrollo.tests.test_fma_cambio_hora import _reporte


class TestFmaEnParalelo(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.carpeta = Path(self.tmp.name)
        fma.limpiar_indice_cpf()

    def tearDown(self):
        fma.limpiar_indice_cpf()
        self.tmp.cleanup()

    def _dos_veces(self, funcion, *argumentos):
        resultados = []
        for paralelo in (False, True):
            fma.limpiar_indice_cpf()
            log = []
            with mock.patch.object(fma, "LEER_EN_PARALELO", paralelo):
                df = funcion(*argumentos, registrar=log.append)
            resultados.append((df, log))
        return resultados

    def test_cpf_septiembre(self):
        for dia in range(1, 31):
            if dia == 12:
                continue  # un dia sin reporte
            nombre = f"tabla_resumen_{dia}_9_2026"
            sub = self.carpeta / f"Reporte diario {nombre}"
            sub.mkdir()
            _reporte(sub / f"{nombre}.xlsx", dia * 100)
        (serie, log_s), (paralelo, log_p) = self._dos_veces(
            fma.construir_fma_cpf, self.carpeta, "2609"
        )
        pd.testing.assert_frame_equal(serie, paralelo)
        self.assertEqual(log_s, log_p)
        self.assertTrue(log_s)  # el aviso del 6-09 y el del dia sin reporte

    def test_csf_mes(self):
        for dia in range(1, 31):
            nombre = fma.PLANTILLA_CSF_DIARIO.format(anio=2026, mes=9, dia=dia)
            pd.DataFrame({"Unidad": ["A", "B"], "Valor": [dia, dia + 0.5]}).to_excel(
                self.carpeta / f"{nombre}.xlsx", index=False
            )
        (serie, log_s), (paralelo, log_p) = self._dos_veces(
            fma.construir_fma_csf, self.carpeta, "2609"
        )
        pd.testing.assert_frame_equal(serie, paralelo)
        self.assertEqual(log_s, log_p)
        self.assertEqual(len(serie), 60)


if __name__ == "__main__":
    unittest.main()
