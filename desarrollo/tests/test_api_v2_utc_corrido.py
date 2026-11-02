# -*- coding: utf-8 -*-
"""La API v2 en el cambio de hora de septiembre, como viene de verdad.

Visto con la API real (ARENA_220_JT1_ARE, 202609): `dateRange` trae su
desfase correcto en cada fila (-04:00 hasta el 5-09, -03:00 desde el
6-09 01:00) y ya no trae la hora que no existe; pero `utcRange` viene
corrido una hora en las 20:xx del 5-09 (la 00:xx UTC del 6-09 la pasa a
01:xx), igual al de las 21:xx. Antes el programa las tomaba por la hora
fantasma, descartaba las 21:xx y al 5-09 le faltaba la hora 22 en todas
las centrales.
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb


def descarga_v2_real(puntos=("P1", "P2")):
    utc = pd.date_range("2026-09-01 04:00", "2026-10-01 03:00",
                        freq="15min", inclusive="left")
    cambio = pd.Timestamp("2026-09-06 04:00")  # 00:00 -04 = 01:00 -03
    filas = []
    for u in utc:
        desfase = -4 if u < cambio else -3
        local = u + pd.Timedelta(hours=desfase)
        # el bug: la hora UTC en la franja que en Chile no existe se
        # corre una hora (y se marca con el desfase que tenga)
        u_api = u + pd.Timedelta(hours=1) if (
            pd.Timestamp("2026-09-06 00:00") <= u < pd.Timestamp("2026-09-06 01:00")
        ) else u
        marca_api = -3 if u_api >= pd.Timestamp("2026-09-06 01:00") else -4
        filas.append({
            "intervalo": f"{local:%Y-%m-%dT%H:%M:%S}.000{desfase:+03d}:00",
            "intervaloUtc": f"{u_api:%Y-%m-%dT%H:%M:%S}.000{marca_api:+03d}:00",
        })
    base = pd.DataFrame(filas)
    partes = []
    for punto in puntos:
        for canal in (1, 3):
            parte = base.copy()
            parte["idPuntoMedida"] = punto
            parte["slugCanal"] = "s"
            parte["principal"] = True
            parte["canalVal1"] = 1.0 if canal == 1 else None
            parte["canalVal3"] = 2.0 if canal == 3 else None
            partes.append(parte)
    return pd.concat(partes, ignore_index=True)


HOMOLOGACION = pd.DataFrame({
    "Punto de Medida": ["P1", "P2"], "Canal": ["s", "s"],
    "clave": ["A", "B"], "Flujo": [1, -1],
})


class TestApiV2UtcCorrido(unittest.TestCase):
    def test_la_hora_22_del_5_09_no_se_pierde(self):
        df = descarga_v2_real()
        # el bug esta en los datos: las 20:xx y 21:xx comparten utcRange
        crudo = df[df["idPuntoMedida"] == "P1"].drop_duplicates("intervalo")
        self.assertEqual(crudo["intervaloUtc"].duplicated().sum(), 4)

        log = []
        salida, _, diagnostico = cb.construir_por_clave(
            df, HOMOLOGACION, registrar=log.append
        )
        a = salida[salida["clave"] == "A"]
        self.assertEqual(diagnostico["cuartos_esperados"], 2876)
        self.assertTrue(diagnostico["incompletos"].empty)
        self.assertEqual(sorted(a[a["Dia"] == 5]["Hora"].unique()), list(range(1, 25)))
        self.assertEqual(sorted(a[a["Dia"] == 6]["Hora"].unique()), list(range(2, 25)))
        hora_22 = a[(a["Dia"] == 5) & (a["Hora"] == 22)]
        self.assertEqual(len(hora_22), 4)
        self.assertEqual(hora_22["intervalo"].min(), pd.Timestamp("2026-09-05 21:00"))
        self.assertFalse(any("hora inexistente" in l for l in log), log)
        self.assertTrue(any("intervaloUtc distinto" in l for l in log), log)
        # el Hora Mes sigue correlativo: 5-09 21:00 = 118, 6-09 01:00 = 121
        self.assertEqual(int(hora_22["Hora Mes"].iloc[0]), 118)
        self.assertEqual(
            int(a[(a["Dia"] == 6) & (a["Hora"] == 2)]["Hora Mes"].iloc[0]), 121
        )

    def test_utc_desde_local(self):
        serie = pd.Series(["2026-09-05T20:00:00.000-04:00",
                           "2026-09-06T01:00:00.000-03:00"])
        self.assertEqual(
            list(cb.utc_desde_local(serie)),
            [pd.Timestamp("2026-09-06 00:00"), pd.Timestamp("2026-09-06 04:00")],
        )
        # sin desfase, o con 'Z' (hora local mal marcada): no se usa
        self.assertIsNone(cb.utc_desde_local(pd.Series(["2026-09-05 20:00:00"])))
        self.assertIsNone(cb.utc_desde_local(pd.Series(["2026-09-05T20:00:00.000Z"])))


if __name__ == "__main__":
    unittest.main()
