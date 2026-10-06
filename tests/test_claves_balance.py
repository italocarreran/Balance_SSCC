"""
Claves_Balance con meses de cambio de hora. El dia y la hora del
cambio NO estan fijos en el codigo: se deducen de la descarga. Por
eso las pruebas arman los datos con ofsets manuales (sin tabla de
zonas horarias) y ponen el cambio en dias y horas arbitrarios.

  - adelanto de hora: la API entrega los 4 cuartos de hora de la hora
    local que no existe con un intervaloUtc repetido; se descartan y
    el mes queda con 4 cuartos de hora menos en vez de dar todos los
    puntos de medida por incompletos;
  - atraso de hora (+4) y un mes normal no cambian;
  - las fechas ISO no se leen con dayfirst (2026-09-06 no es 9 de
    junio).
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb


HORA = pd.Timedelta(hours=1)


def descarga(inicio, fin, cambio=None, ofset_antes=-4, ofset_despues=-3,
             con_fantasma=True):
    """
    Mes [inicio, fin) en hora local. `cambio` es la hora local (con el
    ofset de antes) en que cambia la hora; ofset_despues > ofset_antes
    adelanta la hora (septiembre) y < la atrasa (abril). Con
    con_fantasma, la API entrega ademas la hora local que se salta,
    con el intervaloUtc repetido, que es lo que se ve en la realidad.
    """

    a = pd.Timedelta(hours=ofset_antes)
    d = pd.Timedelta(hours=ofset_despues)
    cambio_utc = pd.Timestamp(cambio) - a if cambio else None

    def ofset(utc):
        return a if cambio_utc is None or utc < cambio_utc else d

    ini_utc = pd.Timestamp(inicio) - a
    fin_utc = pd.Timestamp(fin) - (d if cambio else a)
    utc = pd.date_range(ini_utc, fin_utc, freq="15min", inclusive="left")
    local = [u + ofset(u) for u in utc]

    if cambio and con_fantasma and d > a:
        fantasma = pd.date_range(
            pd.Timestamp(cambio), periods=4, freq="15min"
        )
        utc = list(utc) + [f - a for f in fantasma]
        local = local + list(fantasma)

    base = pd.DataFrame({
        "intervalo": pd.DatetimeIndex(local).strftime("%Y-%m-%d %H:%M:%S"),
        "intervaloUtc": pd.DatetimeIndex(utc).strftime(
            "%Y-%m-%dT%H:%M:%S.000Z"
        ),
    })

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

    def test_adelanto_en_cualquier_dia_y_hora(self):
        for inicio, fin, cambio in (
            ("2026-09-01", "2026-10-01", "2026-09-06 00:00"),
            ("2027-09-01", "2027-10-01", "2027-09-05 00:00"),
            ("2026-09-01", "2026-10-01", "2026-09-13 00:00"),
            ("2026-10-01", "2026-11-01", "2026-10-18 02:00"),
        ):
            with self.subTest(cambio=cambio):
                df = descarga(inicio, fin, cambio)
                dias = (pd.Timestamp(fin) - pd.Timestamp(inicio)).days
                esperados = dias * 96 - 4
                self.assertEqual(len(df) // 4, dias * 96)

                salida, diagnostico, log = self.correr(df)

                self.assertEqual(diagnostico["cuartos_esperados"], esperados)
                self.assertTrue(diagnostico["incompletos"].empty)
                self.assertEqual(salida["Cuarto de Hora"].max(), esperados)
                hueco = pd.Timestamp(cambio)
                self.assertFalse(
                    ((salida["intervalo"] >= hueco)
                     & (salida["intervalo"] < hueco + HORA)).any()
                )
                self.assertTrue(
                    any("hora inexistente" in linea for linea in log)
                )

    def test_atraso_y_mes_normal_sin_cambios(self):
        for inicio, fin, cambio, cuartos in (
            ("2026-04-01", "2026-05-01", "2026-04-05 00:00", 2884),
            ("2026-04-01", "2026-05-01", "2026-04-19 03:00", 2884),
            ("2026-08-01", "2026-09-01", None, 2976),
        ):
            with self.subTest(cambio=cambio):
                df = descarga(
                    inicio, fin, cambio, ofset_antes=-3, ofset_despues=-4
                ) if cambio else descarga(inicio, fin)
                salida, diagnostico, log = self.correr(df)
                self.assertEqual(diagnostico["cuartos_esperados"], cuartos)
                self.assertEqual(salida["Cuarto de Hora"].max(), cuartos)
                self.assertFalse(any("hora inexistente" in l for l in log))

    def test_adelanto_sin_fantasma_no_descarta(self):
        df = descarga(
            "2026-09-01", "2026-10-01", "2026-09-06 00:00",
            con_fantasma=False,
        )
        salida, diagnostico, log = self.correr(df)
        self.assertEqual(diagnostico["cuartos_esperados"], 2876)
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
