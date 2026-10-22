"""Los cuatro diccionarios que salen de la hoja "Resumen BESS".

Los cuatro pasan ahora por el mismo helper. Estas pruebas fijan lo que
cada uno tiene de propio: la columna que busca, el tipo del valor y el
trato de las filas y celdas vacias.
"""

import unittest

import numpy as np
import pandas as pd

from Script import nucleo


def resumen_de_prueba(columnas_al_reves=False):
    resumen = pd.DataFrame({
        "Nombre activo": ["SAE Uno", "SAE Dos", None, "SAE Tres"],
        "Pmax (MW)": [10.0, np.nan, 5.0, 3],
        "Horas para descarga forzada": [1, 1, 1, 1],
        "Capacidad (MWh)": [40.0, 20.0, 1.0, np.nan],
        "Energía mínima": [0, 0, 0, 0],
        "Barra inyección": ["  BARRA A ", np.nan, "X", "BARRA C"],
        "% Energía sobre mínima (indicador nuevo ciclo)":
            [0.06, 0.07, 0.0, 0.08],
        "Ciclos max diarios": [2, 2, 2, 2],
        "Eficiencia": [0.9, np.nan, 0.8, 0.95],
    })
    if columnas_al_reves:
        resumen = resumen[list(reversed(resumen.columns))]
    return resumen


class DiccionariosResumenTest(unittest.TestCase):

    def test_barra_es_texto_recortado_y_vacio_si_falta(self):
        mapa = nucleo.construir_mapa_barra(resumen_de_prueba())

        self.assertEqual(mapa["sae uno"], "BARRA A")
        self.assertEqual(mapa["sae dos"], "")
        self.assertNotIn("", mapa)          # la fila sin nombre se salta

    def test_capacidad_y_eficiencia_son_numeros_o_NA(self):
        capacidad = nucleo.construir_dic_resumen_capacidad(resumen_de_prueba())
        eficiencia = nucleo.construir_dic_resumen_eficiencia(
            resumen_de_prueba()
        )

        self.assertEqual(capacidad["sae uno"], 40.0)
        self.assertTrue(pd.isna(capacidad["sae tres"]))
        self.assertEqual(eficiencia["sae tres"], 0.95)
        self.assertTrue(pd.isna(eficiencia["sae dos"]))

    def test_factor_devuelve_pmax_y_el_umbral_de_la_primera_fila(self):
        factor, umbral = nucleo.construir_dic_resumen_factor(
            resumen_de_prueba()
        )

        self.assertEqual(factor["sae uno"], 10.0)
        self.assertTrue(pd.isna(factor["sae dos"]))
        self.assertEqual(factor["sae tres"], 3.0)
        self.assertEqual(umbral, 0.06)

    def test_no_depende_del_orden_de_las_columnas(self):
        derecho = nucleo.construir_mapa_barra(resumen_de_prueba())
        al_reves = nucleo.construir_mapa_barra(resumen_de_prueba(True))

        self.assertEqual(derecho, al_reves)

    def test_falta_una_columna_nombra_cual(self):
        sin_barra = resumen_de_prueba().drop(columns=["Barra inyección"])

        with self.assertRaises(nucleo.ErrorEntrada) as caja:
            nucleo.construir_mapa_barra(sin_barra)

        self.assertIn("Barra inyección", str(caja.exception))

    def test_sin_filas_no_hay_umbral(self):
        vacio = resumen_de_prueba().iloc[0:0]

        with self.assertRaises(nucleo.ErrorEntrada):
            nucleo.construir_dic_resumen_factor(vacio)


class FachadaTest(unittest.TestCase):

    def test_la_fachada_exporta_lo_que_dice_exportar(self):
        for nombre in nucleo.__all__:
            self.assertTrue(hasattr(nucleo, nombre), nombre)


if __name__ == "__main__":
    unittest.main()
