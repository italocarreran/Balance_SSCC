# -*- coding: utf-8 -*-
"""
Los reportes en palabras (Script/nucleo/reporte.py).

Pedido del usuario (2026-11-10):
  - en el control de corrida, un boton "Generar reporte": si una central
    tiene cero ingreso, que diga por que ("segun FD no tuvo participacion,
    o FD=0");
  - en Herramientas, comparar dos planillas (mes o version anterior) y
    explicar los cambios ("tal central aumento su ingreso por RE545 dado
    que las ofertas aumentaron tanto"), con el resultado en Log/.
"""

import math
import tempfile
import unittest
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo import reporte as r


def fila(**valores):
    """Una central con todo "normal" (recibe por los dos lados)."""

    base = {
        r.C_CENTRAL: "SAE-UNO", r.C_EMPRESA: "EMPRESA",
        r.C_ECOSTOS: 100.0, r.C_RE545: 200.0, r.C_TOTAL: 300.0,
        r.C_DESCARGA: 500.0, r.C_CARGA: -600.0, r.C_CMG: 50.0,
        r.C_HORAS_ADJ: 100.0, r.C_ADJ_CPF: 10.0, r.C_ADJ_CSF: 5.0,
        r.C_RESERVA: 1000.0, r.C_FD_CPF: 0.9, r.C_FD_CSF: 0.8,
        r.C_OFERTA_DIAS: 20.0,
    }
    base.update(valores)
    return base


class TestPorQueNoRecibe(unittest.TestCase):

    def test_recibe_por_los_dos_lados_no_dice_nada(self):
        self.assertEqual(r.explicar_central(fila()), "")

    def test_fd_cero(self):
        texto = r.explicar_central(fila(
            **{r.C_ECOSTOS: 0.0, r.C_RE545: 0.0, r.C_TOTAL: 0.0,
               r.C_FD_CPF: 0.0, r.C_ADJ_CSF: 0.0}
        ))
        self.assertEqual(
            texto, "Según FD no tuvo participación (FD CPF = 0 en todo el mes)."
        )

    def test_sin_adjudicacion(self):
        texto = r.explicar_central(fila(
            **{r.C_TOTAL: 0.0, r.C_ADJ_CPF: 0.0, r.C_ADJ_CSF: 0.0}
        ))
        self.assertEqual(texto, "No se adjudicó nada en Subastas en el mes.")

    def test_sin_energia_manda(self):
        texto = r.explicar_central(fila(
            **{r.C_TOTAL: 0.0, r.C_DESCARGA: 0.0, r.C_CARGA: 0.0,
               r.C_FD_CPF: 0.0}
        ))
        self.assertTrue(texto.startswith("Sin energía medida"), texto)

    def test_adjudicada_sin_cruce(self):
        texto = r.explicar_central(fila(**{r.C_TOTAL: 0.0, r.C_HORAS_ADJ: 0.0}))
        self.assertIn("Diccionario", texto)

    def test_sin_motivo_en_las_entradas(self):
        texto = r.explicar_central(fila(**{r.C_TOTAL: 0.0}))
        self.assertIn("Total C1 ≤ Total C2", texto)

    def test_solo_e_costos(self):
        texto = r.explicar_central(fila(
            **{r.C_RE545: 0.0, r.C_TOTAL: 100.0, r.C_FD_CPF: 0.0}
        ))
        self.assertTrue(texto.startswith("Solo recibe por E Costos: según FD"), texto)

    def test_resumen_de_la_corrida(self):
        resumen = pd.DataFrame([
            fila(), fila(**{r.C_CENTRAL: "SAE-DOS", r.C_ECOSTOS: 0.0,
                            r.C_RE545: 0.0, r.C_TOTAL: 0.0}),
        ])
        lineas, tabla = r.reporte_corrida(resumen)
        self.assertEqual(lineas[0], "Total del mes: $300 entre 2 centrales.")
        self.assertIn("SAE-DOS", lineas[-1])
        self.assertEqual(tabla["Observación"].iloc[0], "")


class TestComparacion(unittest.TestCase):

    def test_explica_con_lo_que_se_movio_en_el_mismo_sentido(self):
        # Cambios de menos de $1.000 no se comentan: montos realistas.
        anterior = pd.DataFrame([
            fila(**{r.C_RE545: 200000.0, r.C_TOTAL: 200100.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        actual = pd.DataFrame([
            fila(**{r.C_RE545: 400000.0, r.C_TOTAL: 400100.0, r.C_OFERTA_DIAS: 28.0,
                    r.C_ADJ_CPF: 15.0, r.C_CMG: 10.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        lineas, tabla = r.comparar_resumenes(actual, anterior)
        uno = tabla.set_index(r.C_CENTRAL).loc["SAE-UNO", "Explicación"]
        self.assertTrue(uno.startswith("RE545 sube $200.000 (+100 %):"), uno)
        self.assertIn("lo adjudicado en CPF subió (10 → 15 MW)", uno)
        self.assertIn("los días con oferta completa subieron (20 → 28 días)", uno)
        # El CMg bajo, pero el RE545 sube: no lo explica, no se nombra.
        self.assertNotIn("CMg", uno)
        # Lo que no cambio va en una sola linea, no en su fila.
        self.assertEqual(tabla.set_index(r.C_CENTRAL).loc["SAE-DOS", "Explicación"], "")
        self.assertIn("Sin cambios relevantes: SAE-DOS.", lineas)
        # Lo que mas cambio, arriba.
        self.assertEqual(tabla[r.C_CENTRAL].iloc[0], "SAE-UNO")

    def test_central_nueva_y_que_ya_no_esta(self):
        anterior = pd.DataFrame([fila(**{r.C_CENTRAL: "VIEJA"})])
        actual = pd.DataFrame([fila(**{r.C_CENTRAL: "NUEVA"})])
        _, tabla = r.comparar_resumenes(actual, anterior)
        textos = dict(zip(tabla[r.C_CENTRAL], tabla["Explicación"]))
        self.assertTrue(textos["NUEVA"].startswith("Nueva"))
        self.assertTrue(textos["VIEJA"].startswith("Ya no está"))

    def test_plata_a_la_chilena(self):
        self.assertEqual(r.plata(1234567.6), "$1.234.568")
        self.assertEqual(r.plata(-5), "-$5")


def _planilla(carpeta, nombre, multiplicar=1.0):
    from desarrollo.tests import test_relleno_rapido as t

    d = t._datos(60)
    for clave in ("ec", "re545"):
        d[clave] = d[clave].copy()
        d[clave]["Monto a compensar"] = (
            np.abs(d[clave]["Monto a compensar"]) * 1e5 * multiplicar
        )
    Path(carpeta).mkdir(parents=True, exist_ok=True)
    ruta = Path(t._escribir(carpeta, True, d))
    destino = ruta.with_name(nombre)
    ruta.rename(destino)
    return destino


class TestArchivos(unittest.TestCase):

    def test_la_comparacion_queda_en_el_log_de_la_actual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            anterior = _planilla(Path(carpeta) / "sep", "11_PAGOS_BESS_2609.xlsx")
            actual = _planilla(Path(carpeta) / "oct", "11_PAGOS_BESS_2610.xlsx", 1.5)
            destino, lineas, tabla = nucleo.comparar_planillas(
                actual, anterior, registrar=lambda *a: None
            )
            self.assertEqual(destino, actual.parent / "Log" / "Comparacion_2610_vs_2609.xlsx")
            libro = openpyxl.load_workbook(destino)
            self.assertEqual(libro.sheetnames, ["Comparacion", "Detalle"])
            ws = libro["Comparacion"]
            self.assertTrue(str(ws["B2"].value).startswith("Cuadro N° 1"))
            self.assertEqual(ws["B3"].value, "Resumen")
            self.assertTrue(lineas[0].startswith("Total del mes:"))
            self.assertTrue(any("RE545 sube" in e for e in tabla["Explicación"]))

    def test_la_hoja_reporte_del_control_y_el_resto_se_preserva(self):
        from Script.nucleo.rutas import resolver_rutas

        with tempfile.TemporaryDirectory() as carpeta:
            rutas = resolver_rutas(carpeta, "2610")
            _planilla(carpeta, rutas["salida"].name)
            # Un control que ya tiene otra hoja: no se pierde.
            nucleo.escribir_control(
                rutas["control"],
                df_log=pd.DataFrame({"mensaje": ["algo"]}),
            )
            nucleo.generar_reporte_corrida(carpeta, "2610", registrar=lambda *a: None)
            libro = openpyxl.load_workbook(rutas["control"])
        self.assertIn(nucleo.HOJA_REPORTE, libro.sheetnames)
        self.assertIn(nucleo.HOJA_LOG, libro.sheetnames)
        ws = libro[nucleo.HOJA_REPORTE]
        nombres = [c.value for c in ws[3] if c.value]
        self.assertEqual(nombres[0], "Resumen")
        self.assertIn("Observación", nombres)
        self.assertTrue(str(ws["B4"].value).startswith("Total del mes:"))

    def test_sin_planilla_dice_que_hay_que_calcular(self):
        with tempfile.TemporaryDirectory() as carpeta:
            with self.assertRaises(nucleo.ErrorEntrada):
                nucleo.generar_reporte_corrida(carpeta, "2610", registrar=lambda *a: None)


class TestHerramienta(unittest.TestCase):

    def test_sugiere_la_planilla_del_mes_anterior(self):
        import importlib.util

        ruta = Path(__file__).resolve().parents[2] / "Herramientas" / "Comparar_Meses.py"
        spec = importlib.util.spec_from_file_location("comparar_meses", ruta)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)

        self.assertEqual(modulo.aamm_anterior("2601"), "2512")
        self.assertEqual(modulo.aamm_anterior("2610"), "2609")
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "2609" / "SSCC").mkdir(parents=True)
            (base / "2610" / "SSCC").mkdir(parents=True)
            anterior = base / "2609" / "SSCC" / "11_PAGOS_BESS_2609.xlsx"
            anterior.write_bytes(b"")
            actual = base / "2610" / "SSCC" / "11_PAGOS_BESS_2610.xlsx"
            self.assertEqual(modulo.sugerir_anterior(actual), anterior)


if __name__ == "__main__":
    unittest.main()
