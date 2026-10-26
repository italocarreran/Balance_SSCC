# -*- coding: utf-8 -*-
"""La API v2 en el cambio de hora de abril: la hora 25 del sabado.

El 4-04-2026 las 23:xx pasan dos veces (-03:00 y despues -04:00). Si
`dateRange` marca las dos pasadas con el mismo desfase, el UTC sacado de
la hora local es igual para las dos y la segunda se sumaba en la hora 24.
Se usa entonces el `intervaloUtc` de la API, que si las distingue.
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb


CAMBIO = pd.Timestamp("2026-04-05 03:00")  # 00:00 -03 = 23:00 -04


def descarga_abril(desfase_segunda=-4, con_segunda=True):
    utc = pd.date_range("2026-04-01 03:00", "2026-05-01 04:00",
                        freq="15min", inclusive="left")
    filas = []
    for u in utc:
        desfase = -3 if u < CAMBIO else -4
        segunda = CAMBIO <= u < CAMBIO + pd.Timedelta(hours=1)
        if segunda and not con_segunda:
            continue
        local = u + pd.Timedelta(hours=desfase)
        marca = desfase_segunda if segunda else desfase
        filas.append({
            "intervalo": f"{local:%Y-%m-%dT%H:%M:%S}.000{marca:+03d}:00",
            "intervaloUtc": f"{u:%Y-%m-%dT%H:%M:%S}.000Z",
        })
    df = pd.DataFrame(filas)
    df["idPuntoMedida"] = "P1"
    df["slugCanal"] = "s"
    df["principal"] = True
    df["canalVal1"] = 1.0
    df["canalVal3"] = 0.0
    return df


HOMOLOGACION = pd.DataFrame({
    "Punto de Medida": ["P1"], "Canal": ["s"], "clave": ["A"], "Flujo": [1],
})


def horas_del_sabado(df):
    log = []
    salida, _, _ = cb.construir_por_clave(df, HOMOLOGACION, registrar=log.append)
    sabado = salida[salida["Dia"] == 4]
    return sabado.groupby("Hora")["Gen_Unidad"].sum(), log


class TestApiV2Abril(unittest.TestCase):
    def test_desfases_correctos(self):
        horas, log = horas_del_sabado(descarga_abril())
        self.assertEqual(list(horas.index), list(range(1, 26)))
        self.assertTrue((horas == 4.0).all())
        self.assertFalse(any("AVISO" in l for l in log), log)

    def test_segunda_pasada_con_el_mismo_desfase(self):
        horas, log = horas_del_sabado(descarga_abril(desfase_segunda=-3))
        self.assertEqual(list(horas.index), list(range(1, 26)))
        self.assertTrue((horas == 4.0).all())
        self.assertTrue(any("choca con otra" in l for l in log), log)

    def test_sin_la_segunda_pasada_avisa(self):
        horas, log = horas_del_sabado(descarga_abril(con_segunda=False))
        self.assertEqual(list(horas.index), list(range(1, 25)))
        self.assertTrue(
            any("no trae la hora repetida (25)" in l for l in log), log
        )


if __name__ == "__main__":
    unittest.main()
