# -*- coding: utf-8 -*-
"""Observacion O2026-0052689 (AES Andes): la Componente 2 con una inyeccion chica.

Ejemplo de docs/Obs_O2026-0052689_Componente2_RE545.md: Pmax 1 MW (caja de
250 kWh), Energia Disponible 500 kWh, 8 bloques con CMg 100..35. Hoy, 1 kWh
inyectado en el bloque mas caro descuenta la caja entera; con
C2_SOLO_INYECCION se descuenta solo ese kWh. El programa sigue con la regla
de hoy (False) hasta que el usuario decida.
"""

import unittest
from unittest import mock

import pandas as pd

from Script.nucleo import re545_componentes as rc
from Script.nucleo.utiles import normalizar


def _pago(inyeccion, solo_inyeccion):
    cmg = [100, 95, 90, 85, 50, 45, 40, 35]
    df = pd.DataFrame({
        "clave": ["SAE-X"] * 8, "T": [1] * 8, "Minutos": [0, 15, 30, 45] * 2,
        "S": [1] * 4 + [2] * 4, "R": [92.5] * 4 + [42.5] * 4, "CMg": cmg,
        "Energia_Positiva": inyeccion, "Energia_Negativa": [0.0] * 8,
        "AU": [10.0] * 8, "Hora": [12] * 4 + [13] * 4,
    })
    resumen = pd.DataFrame({
        "Central": ["SAE-X"], "Ventana": [1], "Edisp_T": [500.0],
        "Margen ultima hora": [0.0], "flag ultima hora": [0.0],
    })
    with mock.patch.object(rc, "C2_SOLO_INYECCION", solo_inyeccion):
        r = rc.calcular_componentes_re545(
            df, resumen, {normalizar("SAE-X"): 1.0}, registrar=lambda *a: None
        )
    c1 = float(pd.to_numeric(r["BO"]).fillna(0).sum())
    c2 = float(pd.to_numeric(r["CC"]).fillna(0).sum())
    return round(c1), round(c2), round(c1 - c2)


class TestComponente2(unittest.TestCase):

    def test_por_defecto_sigue_la_regla_de_hoy(self):
        self.assertFalse(rc.C2_SOLO_INYECCION)

    def test_sin_inyeccion_no_cambia(self):
        self.assertEqual(_pago([0.0] * 8, False), (48750, 0, 48750))
        self.assertEqual(_pago([0.0] * 8, True), (48750, 0, 48750))

    def test_un_kwh_descuenta_la_caja_hoy_y_un_kwh_con_la_propuesta(self):
        uno = [1.0] + [0.0] * 7
        self.assertEqual(_pago(uno, False), (48740, 25000, 23740))
        self.assertEqual(_pago(uno, True), (48740, 100, 48640))

    def test_caja_llena_da_lo_mismo(self):
        llena = [250.0] + [0.0] * 7
        self.assertEqual(_pago(llena, False), (46250, 25000, 21250))
        self.assertEqual(_pago(llena, True), (46250, 25000, 21250))


if __name__ == "__main__":
    unittest.main()
