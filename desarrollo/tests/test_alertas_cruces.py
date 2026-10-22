"""Avisos de los cruces que la planilla completa con cero sin decirlo.

Complementan tests/test_alertas_mapeos.py: aca estan los cruces que
quedaron sin aviso en la primera pasada (Pmax de RE545, Pmax = 0,
Prorrata SSCC y reservas de Subastas).
"""

import unittest

import pandas as pd

from Script import nucleo


class AlertasCrucesTest(unittest.TestCase):

    def test_re545_avisa_central_sin_pmax(self):
        mensajes = []
        df = pd.DataFrame({
            "clave": ["SAE-A", "SAE-SIN-PMAX"],
            "Mes": [7, 7],
            "Hora Mes": [1, 1],
            "Cuarto de Hora": ["", ""],
            "T": ["V1", "V1"],
            "CMg": [10.0, 10.0],
            "R": [10.0, 10.0],
            "Energia_Positiva": [1.0, 1.0],
            "Energia_Negativa": [0.0, 0.0],
            "S": [1, 1],
            "AU": [0.0, 0.0],
        })
        resumen = pd.DataFrame({
            "clave": ["SAE-A", "SAE-SIN-PMAX"],
            "Ventana": ["V1", "V1"],
            "Edisp_T": [0.0, 0.0],
            "Margen ultima hora": [0.0, 0.0],
            "flag ultima hora": [0, 0],
        })

        try:
            nucleo.calcular_componentes_re545(
                df, resumen, {"sae-a": 10.0}, registrar=mensajes.append
            )
        except Exception:
            # El calculo completo necesita mas columnas; lo que se
            # prueba es que el aviso salga antes de tropezar con eso.
            pass

        self.assertTrue(
            any("sin Pmax (MW)" in m and "SAE-SIN-PMAX" in m
                for m in mensajes),
            mensajes,
        )

    def test_ecostos_avisa_pmax_en_cero(self):
        mensajes = []
        df = pd.DataFrame({
            "clave": ["SAE-CERO", "SAE-OK"],
            "Copia_Ventana": ["V1", "V1"],
            "N": [1.0, 1.0],
            "O": [-1.0, -1.0],
            "W": [1, 1],
        })

        ae, af = nucleo.calcular_ae_af(
            df, {"sae-cero": 0.0, "sae-ok": 10.0},
            registrar=mensajes.append,
        )

        self.assertTrue(pd.isna(ae.iloc[0]) if hasattr(ae, "iloc")
                        else pd.isna(ae[0]))
        self.assertTrue(
            any("Pmax (MW) en 0" in m and "SAE-CERO" in m for m in mensajes),
            mensajes,
        )

    def test_prorrata_avisa_central_sin_ninguna_hora(self):
        mensajes = []
        df = pd.DataFrame({
            "clave": ["SAE-A", "SAE-A", "SAE-B"],
            "Hora Mes": [1, 2, 1],
        })
        dic = {"SAE-A¦1": (0.5, 0.5), "SAE-A¦2": (0.5, 0.5)}

        ag, ah = nucleo.calcular_prorratas(
            df, dic, registrar=mensajes.append
        )

        self.assertEqual(list(ag), [0.5, 0.5, 0.0])
        self.assertTrue(
            any("ninguna hora en la Prorrata SSCC" in m and "SAE-B" in m
                for m in mensajes),
            mensajes,
        )

    def test_prorrata_avisa_solo_filas_cuando_la_central_si_cruza(self):
        mensajes = []
        df = pd.DataFrame({"clave": ["SAE-A", "SAE-A"], "Hora Mes": [1, 2]})

        nucleo.calcular_prorratas(
            df, {"SAE-A¦1": (0.5, 0.5)}, registrar=mensajes.append
        )

        self.assertTrue(any("1 fila(s) sin prorrata SSCC" in m
                            for m in mensajes), mensajes)
        self.assertFalse(any("ninguna hora" in m for m in mensajes))

    def test_re545_avisa_central_ausente_de_subastas(self):
        mensajes = []
        df = pd.DataFrame({
            "clave": ["SAE-A", "SAE-FUERA"],
            "Hora Mes": [1, 1],
        })
        dic_con_sae_a = {("SAE-A", "1", "CPF"): 5.0}
        dics = (dic_con_sae_a, {}, {})

        nucleo.calcular_reservas_re545(df, dics, registrar=mensajes.append)

        self.assertTrue(
            any("ninguna fila de la hoja Subastas" in m and "SAE-FUERA" in m
                for m in mensajes),
            mensajes,
        )
        self.assertFalse(any("SAE-A'" in m for m in mensajes))


if __name__ == "__main__":
    unittest.main()
