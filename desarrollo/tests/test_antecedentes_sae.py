# -*- coding: utf-8 -*-
"""
Antecedentes SAE (pedido del usuario 2026-11-10): la carpeta
'03 Antecedentes SAE' junto al 11_ con UN codigo que recalcula las hojas
de calculo y el SAE_AAMM.csv usando solo el 11_, sobre una copia.

La prueba de punta a punta corre el codigo generado en otro Python
aislado (-I), desde otra carpeta y sin el repo: tiene que dar las hojas
identicas a las del 11_.
"""

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo
from Script.nucleo import antecedentes, recalculo
from desarrollo.tests.test_pasos_sueltos import armar_caso, correr_pagos, dejar_prorrata

TODAS = {"ecostos", "re545", "compensacion_central", "prorrata_retiros", "resumen"}


def caso_completo(carpeta):
    base = armar_caso(carpeta)
    dejar_prorrata(base)
    nucleo.generar_consolidado(base, "2607", {"centrales"}, registrar=lambda *_: None)
    correr_pagos(base, TODAS, [])
    return base


def hojas(ruta):
    with pd.ExcelFile(ruta) as libro:
        return {h: pd.read_excel(libro, sheet_name=h, header=None) for h in libro.sheet_names}


class TestRecalculador(unittest.TestCase):

    def test_el_codigo_es_uno_solo_y_recortado(self):
        texto = antecedentes.texto_recalculador("2609")
        compile(texto, "recalculador.py", "exec")
        for de_descarga in ("Descarga_PRMTE", "Generacion_Real", "Extrae_CMG_barras",
                            "Indices_FMA", "Indicadores_DCO", "Claves_Balance"):
            self.assertNotIn(f"FUENTES['Medidas/{de_descarga}.py']", texto)
            self.assertNotIn(f"/{de_descarga}.py'] = r", texto)
        for calculo in ("nucleo/ecostos.py", "nucleo/re545.py", "nucleo/csv_sae.py",
                        "nucleo/recalculo.py", "ventana_recalculo.py"):
            self.assertIn(f"FUENTES[{calculo!r}]", texto)

    def test_prorrata_desde_la_hoja(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = caso_completo(carpeta)
            prorrata = recalculo.prorrata_desde_hoja(base / nucleo.nombre_salida("2607"))
            self.assertEqual(list(prorrata.columns), ["Cuarto de Hora", "Suministrador", "Prorrata"])
            self.assertEqual(sorted(prorrata["Suministrador"]), ["EMPRESA DOS", "EMPRESA UNO"])

    def test_de_punta_a_punta_solo_con_el_11(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = caso_completo(carpeta)
            # (este caso chico no tiene pagos: va sin CSV)
            ant = nucleo.generar_antecedentes(base, "2607", incluir_csv=False,
                                              registrar=lambda *_: None)
            self.assertEqual(ant.name, "03 Antecedentes SAE")
            codigo = ant / "Calcular_11_PAGOS_BESS_2607.py"
            self.assertTrue(codigo.is_file())

            # Lo que recibe un coordinado: el 11_ y la carpeta, nada mas.
            entrega = Path(carpeta) / "entrega"
            entrega.mkdir()
            shutil.copy(base / nucleo.nombre_salida("2607"), entrega)
            shutil.copytree(ant, entrega / ant.name)
            entorno = {k: v for k, v in os.environ.items()
                       if k not in ("BESS_CARPETA_CONFIG", "PYTHONPATH")}
            proceso = subprocess.run(
                [sys.executable, "-I", str(entrega / ant.name / codigo.name), "--sin-ventana"],
                cwd=tempfile.gettempdir(), env=entorno, capture_output=True, text=True,
                timeout=600,
            )
            recalculado = entrega / "11_PAGOS_BESS_2607_calculado.xlsx"
            self.assertEqual(proceso.returncode, 0, proceso.stderr[-2000:])
            self.assertTrue(recalculado.is_file(), proceso.stderr[-2000:])
            original = hojas(entrega / nucleo.nombre_salida("2607"))
            nuevo = hojas(recalculado)
            self.assertEqual(list(original), list(nuevo))
            for hoja in original:
                pd.testing.assert_frame_equal(original[hoja], nuevo[hoja], obj=hoja)
            # El entregable no se toco y no quedo nada mas al lado.
            self.assertEqual(
                sorted(p.name for p in entrega.iterdir()),
                ["03 Antecedentes SAE", "11_PAGOS_BESS_2607.xlsx",
                 "11_PAGOS_BESS_2607_calculado.xlsx"],
            )

    def test_busca_el_11_en_la_carpeta_de_arriba(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            (raiz / "03 Antecedentes SAE").mkdir()
            (raiz / "11_PAGOS_BESS_2609.xlsx").write_bytes(b"x")
            (raiz / "11_PAGOS_BESS_2609_calculado.xlsx").write_bytes(b"x")
            self.assertEqual(
                recalculo.buscar_11([raiz / "03 Antecedentes SAE" / "..", raiz]).name,
                "11_PAGOS_BESS_2609.xlsx",
            )
            self.assertIsNone(recalculo.buscar_11([raiz / "03 Antecedentes SAE"]))


if __name__ == "__main__":
    unittest.main()
