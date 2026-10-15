# -*- coding: utf-8 -*-
"""El SoC del SCADA (hora local, sin zona) en los dias de cambio de hora.

- septiembre: el SCADA salta de las 23:45 del 5-09 a la 01:00 del 6-09;
  las 23:xx del 5-09 tienen que tener su SoC aunque los valores de
  medidas hayan venido en la hora que no existe;
- abril: la hora que se repite (23:xx del sabado) tiene que quedar en
  orden (24 y despues 25, no intercaladas) y, si el SCADA la trae dos
  veces, cada pasada con su SoC; si la trae una vez, las dos con ese.
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb
from Script.nucleo.medidores import construir_medidores
from tests.test_claves_balance import HOMOLOGACION, descarga


def _medidores(df_descarga, mes, timestamps):
    sae, _, _ = cb.construir_por_clave(
        df_descarga, HOMOLOGACION, registrar=lambda *_: None
    )
    soc = pd.DataFrame({
        "central": "A",
        "timestamp": timestamps,
        "soc": [i / 100000 for i in range(len(timestamps))],
        "nombre_scada_original": "A",
    })
    medidores, _ = construir_medidores(sae, soc, mes, registrar=lambda *_: None)
    return medidores[medidores["clave"] == "A"].reset_index(drop=True), soc


class TestSocCambioDeHora(unittest.TestCase):
    def test_septiembre(self):
        timestamps = (
            list(pd.date_range("2026-09-01", "2026-09-05 23:45", freq="15min"))
            + list(pd.date_range("2026-09-06 01:00", "2026-09-30 23:45",
                                 freq="15min"))
        )
        medidores, soc = _medidores(
            descarga("2026-09-01", "2026-10-01", "2026-09-06 00:00",
                     real_con_valor=False),
            9, timestamps,
        )
        self.assertFalse(medidores["SoC"].isna().any())
        ultima = medidores[(medidores["Dia"] == 5) & (medidores["Hora"] == 24)]
        esperado = soc.set_index("timestamp")["soc"]
        self.assertEqual(
            ultima["SoC"].tolist(), esperado[ultima["intervalo"]].tolist()
        )

    def test_abril_con_la_hora_repetida_dos_veces_en_el_scada(self):
        timestamps = (
            list(pd.date_range("2026-04-01", "2026-04-04 23:45", freq="15min"))
            + list(pd.date_range("2026-04-04 23:00", "2026-04-30 23:45",
                                 freq="15min"))
        )
        medidores, _ = _medidores(
            descarga("2026-04-01", "2026-05-01", "2026-04-05 00:00",
                     ofset_antes=-3, ofset_despues=-4),
            4, timestamps,
        )
        dia = medidores[medidores["Dia"] == 4]
        self.assertEqual(dia["Hora"].tolist()[-8:], [24] * 4 + [25] * 4)
        # el SoC avanza en orden: cada pasada con el suyo
        self.assertTrue(dia["SoC"].is_monotonic_increasing)
        self.assertEqual(dia["SoC"].nunique(), len(dia))

    def test_abril_con_la_hora_una_sola_vez_en_el_scada(self):
        timestamps = list(
            pd.date_range("2026-04-01", "2026-04-30 23:45", freq="15min")
        )
        medidores, _ = _medidores(
            descarga("2026-04-01", "2026-05-01", "2026-04-05 00:00",
                     ofset_antes=-3, ofset_despues=-4),
            4, timestamps,
        )
        self.assertFalse(medidores["SoC"].isna().any())
        dia = medidores[medidores["Dia"] == 4]
        h24 = dia[dia["Hora"] == 24]["SoC"].tolist()
        h25 = dia[dia["Hora"] == 25]["SoC"].tolist()
        self.assertEqual(h24, h25)


if __name__ == "__main__":
    unittest.main()
