"""
Claves_Balance con meses de cambio de hora:

  - septiembre: la API entrega los 4 cuartos de hora de la hora local
    que no existe (00:00-00:59 del dia del cambio) con un intervaloUtc
    repetido; se descartan y el mes queda en 2876 cuartos de hora en
    vez de dar todos los puntos de medida por incompletos;
  - abril (2884) y un mes normal (2976) no cambian;
  - las fechas ISO no se leen con dayfirst (2026-09-06 no es 9 de
    junio).
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb


ZONA = "America/Santiago"


def descarga(inicio, fin, hora_fantasma=None):
    utc = pd.date_range(
        pd.Timestamp(inicio, tz=ZONA).tz_convert("UTC"),
        pd.Timestamp(fin, tz=ZONA).tz_convert("UTC"),
        freq="15min", inclusive="left",
    )
    local = utc.tz_convert(ZONA).tz_localize(None)
    base = pd.DataFrame({
        "intervalo": local.strftime("%Y-%m-%d %H:%M:%S"),
        "intervaloUtc": utc.tz_localize(None).strftime(
            "%Y-%m-%dT%H:%M:%S.000Z"
        ),
    })

    if hora_fantasma:
        dia, utc_real = hora_fantasma
        base = pd.concat([base, pd.DataFrame({
            "intervalo": [f"{dia} 00:{m:02d}:00" for m in (0, 15, 30, 45)],
            "intervaloUtc": [
                f"{utc_real}:{m:02d}:00.000Z" for m in (0, 15, 30, 45)
            ],
        })])

    partes = []
    for punto in ("P1", "P2"):
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
    "Punto de Medida": ["P1", "P2"],
    "Canal": ["s", "s"],
    "clave": ["A", "B"],
    "Flujo": [1, -1],
})


class TestCambioDeHora(unittest.TestCase):

    def correr(self, df):
        log = []
        salida, calendario, diagnostico = cb.construir_por_clave(
            df, HOMOLOGACION, registrar=log.append
        )
        return salida, diagnostico, log

    def test_septiembre_descarta_hora_inexistente(self):
        df = descarga(
            "2026-09-01", "2026-10-01",
            hora_fantasma=("2026-09-06", "2026-09-06T04"),
        )
        self.assertEqual(len(df) // 4, 2880)

        salida, diagnostico, log = self.correr(df)

        self.assertEqual(diagnostico["cuartos_esperados"], 2876)
        self.assertTrue(diagnostico["incompletos"].empty)
        self.assertEqual(salida["Cuarto de Hora"].max(), 2876)
        self.assertFalse(
            ((salida["intervalo"] >= "2026-09-06 00:00")
             & (salida["intervalo"] < "2026-09-06 01:00")).any()
        )
        self.assertTrue(any("hora inexistente" in linea for linea in log))

    def test_abril_y_mes_normal_sin_cambios(self):
        for inicio, fin, cuartos in (
            ("2026-04-01", "2026-05-01", 2884),
            ("2026-08-01", "2026-09-01", 2976),
        ):
            salida, diagnostico, log = self.correr(descarga(inicio, fin))
            self.assertEqual(diagnostico["cuartos_esperados"], cuartos)
            self.assertEqual(salida["Cuarto de Hora"].max(), cuartos)
            self.assertFalse(any("hora inexistente" in l for l in log))


class TestParseFecha(unittest.TestCase):

    def test_iso_no_usa_dayfirst(self):
        fechas = cb.parse_fecha_mixta(pd.Series(["2026-09-06 01:00:00"]))
        self.assertEqual(fechas[0], pd.Timestamp("2026-09-06 01:00"))

    def test_dia_primero(self):
        fechas = cb.parse_fecha_mixta(pd.Series(["06-09-2026 01:00"]))
        self.assertEqual(fechas[0], pd.Timestamp("2026-09-06 01:00"))

    def test_ofsets_mezclados(self):
        fechas = cb.parse_fecha_mixta(pd.Series([
            "2026-09-05T23:00:00-04:00", "2026-09-06T01:00:00-03:00",
        ]))
        self.assertEqual(
            fechas.tolist(),
            [pd.Timestamp("2026-09-05 23:00"),
             pd.Timestamp("2026-09-06 01:00")],
        )


if __name__ == "__main__":
    unittest.main()
