"""
Claves_Balance con meses de cambio de hora. El dia y la hora del
cambio NO estan fijos en el codigo: se deducen de la descarga. Por
eso las pruebas arman los datos con ofsets manuales (sin tabla de
zonas horarias) y ponen el cambio en dias y horas arbitrarios.

  - adelanto de hora: la API entrega los 4 cuartos de hora de la hora
    local que no existe (00:00-00:59 si el salto es a medianoche) con
    el intervaloUtc de la hora ANTERIOR al salto (23:00-23:59), como se
    vio con septiembre 2026 real; se descartan los falsos -no los de
    las 23- y el mes queda con 4 cuartos de hora menos en vez de dar
    todos los puntos de medida por incompletos;
  - si una de las dos horas trae valores y la otra no, gana la que
    trae valores;
  - una descarga completa no se reusa: se vuelve a bajar;
  - si los datos llegan solo hasta media mes, se genera igual y el
    hueco queda detallado por punto y canal (detallar_huecos);
  - atraso de hora (+4) y un mes normal no cambian;
  - las fechas ISO no se leen con dayfirst (2026-09-06 no es 9 de
    junio).
"""

import unittest

import pandas as pd

from Script.Medidas import Claves_Balance as cb


HORA = pd.Timedelta(hours=1)


def descarga(inicio, fin, cambio=None, ofset_antes=-4, ofset_despues=-3,
             con_fantasma=True, fantasma_con_valor=True, real_con_valor=True,
             valores_hasta=None):
    """
    Mes [inicio, fin) en hora local. `cambio` es la hora local (con el
    ofset de antes) en que cambia la hora; ofset_despues > ofset_antes
    adelanta la hora (septiembre) y < la atrasa (abril). Con
    con_fantasma, la API entrega ademas la hora local que se salta,
    con el intervaloUtc de la hora anterior al salto (lo que se ve en
    la realidad). valores_hasta: hora local desde la que los valores
    vienen nulos (mes a medio publicar).
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
        utc = list(utc) + [f - d for f in fantasma]
        local = local + list(fantasma)

    base = pd.DataFrame({
        "intervalo": pd.DatetimeIndex(local).strftime("%Y-%m-%d %H:%M:%S"),
        "intervaloUtc": pd.DatetimeIndex(utc).strftime(
            "%Y-%m-%dT%H:%M:%S.000Z"
        ),
    })

    local_ts = pd.to_datetime(base["intervalo"])
    con_valor = pd.Series(True, index=base.index)
    if cambio and con_fantasma and d > a:
        es_fantasma = pd.Series(False, index=base.index)
        es_fantasma.iloc[-4:] = True
        hueco = pd.Timestamp(cambio)
        es_real = (local_ts >= hueco - HORA) & (local_ts < hueco)
        if not fantasma_con_valor:
            con_valor &= ~es_fantasma
        if not real_con_valor:
            con_valor &= ~es_real
    if valores_hasta:
        con_valor &= local_ts < pd.Timestamp(valores_hasta)

    partes = []
    for punto in ("P1", "P2"):
        for canal in (1, 3):
            parte = base.copy()
            parte["idPuntoMedida"] = punto
            parte["slugCanal"] = "s"
            parte["principal"] = True
            parte["canalVal1"] = (1.0 if canal == 1 else None)
            parte["canalVal3"] = (2.0 if canal == 3 else None)
            parte.loc[~con_valor.values, ["canalVal1", "canalVal3"]] = None
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
                # la hora anterior al salto SI existio y se conserva
                self.assertEqual(
                    ((salida["intervalo"] >= hueco - HORA)
                     & (salida["intervalo"] < hueco)).sum(),
                    4 * 2,
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

    def test_gana_la_hora_que_trae_valores(self):
        cambio = "2026-09-06 00:00"
        hueco = pd.Timestamp(cambio)

        # Fantasma sin valores: se descarta el fantasma (00:xx).
        salida, diagnostico, _ = self.correr(descarga(
            "2026-09-01", "2026-10-01", cambio, fantasma_con_valor=False
        ))
        self.assertTrue(diagnostico["incompletos"].empty)
        self.assertFalse(
            ((salida["intervalo"] >= hueco)
             & (salida["intervalo"] < hueco + HORA)).any()
        )

        # Al reves (solo el fantasma trae valores): se queda el fantasma
        # y el punto sigue completo.
        salida, diagnostico, _ = self.correr(descarga(
            "2026-09-01", "2026-10-01", cambio, real_con_valor=False
        ))
        self.assertTrue(diagnostico["incompletos"].empty)
        self.assertEqual(diagnostico["cuartos_esperados"], 2876)

    def test_mes_a_medio_publicar_se_genera_igual(self):
        # Antes: error "no quedo ningun registro principal". Ahora los
        # puntos incompletos entran con lo que traen, y el hueco queda
        # detallado por punto y canal.
        df = descarga(
            "2026-09-01", "2026-10-01", "2026-09-06 00:00",
            valores_hasta="2026-09-13 23:00",
        )
        salida, diagnostico, log = self.correr(df)

        self.assertEqual(len(diagnostico["incompletos"]), 2)
        self.assertEqual(sorted(salida["clave"].unique()), ["A", "B"])
        self.assertEqual(salida["Cuarto de Hora"].max(), 2876)
        self.assertTrue(any("valores solo hasta el 13-09-2026 22:45" in l
                            for l in log))

        huecos = diagnostico["huecos"]
        self.assertEqual(len(huecos), 4)          # 2 puntos x 2 canales
        self.assertEqual(set(huecos["Desde"]),
                         {pd.Timestamp("2026-09-13 23:00")})
        self.assertEqual(set(huecos["Hasta"]),
                         {pd.Timestamp("2026-09-30 23:45")})

    def test_huecos_en_medio_y_canal_que_falta(self):
        df = descarga("2026-08-01", "2026-09-01")
        local = pd.to_datetime(df["intervalo"])
        # P1: canal 1 sin valor el 10-08 de 10:00 a 11:45 (8 cuartos).
        tramo = ((df["idPuntoMedida"] == "P1") & df["canalVal1"].notna()
                 & (local >= "2026-08-10 10:00")
                 & (local < "2026-08-10 12:00"))
        df.loc[tramo, "canalVal1"] = None
        # P2: la API no devolvio el canal 3.
        df = df[~((df["idPuntoMedida"] == "P2") & df["canalVal3"].notna())]

        huecos = cb.detallar_huecos(cb.normalizar_fechas(df))

        self.assertEqual(len(huecos), 2)
        p1 = huecos[huecos["Punto de Medida"] == "P1"].iloc[0]
        self.assertEqual(p1["Canal"], "Canal 1")
        self.assertEqual(p1["Desde"], pd.Timestamp("2026-08-10 10:00"))
        self.assertEqual(p1["Hasta"], pd.Timestamp("2026-08-10 11:45"))
        self.assertEqual(p1["Cuartos de hora"], 8)
        p2 = huecos[huecos["Punto de Medida"] == "P2"].iloc[0]
        self.assertEqual(p2["Canal"], "Canal 3")
        self.assertEqual(p2["Cuartos de hora"], 2976)

    def test_mes_completo_sin_huecos(self):
        _, diagnostico, _ = self.correr(descarga("2026-08-01", "2026-09-01"))
        self.assertTrue(diagnostico["huecos"].empty)

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


def _hay_parquet():
    try:
        import pyarrow  # noqa: F401
        return True
    except ImportError:
        return False


@unittest.skipUnless(_hay_parquet(), "sin pyarrow (los lotes son parquet)")
class TestDescargaNoReusaCompleta(unittest.TestCase):

    def test_descarga_completa_se_vuelve_a_bajar(self):
        import tempfile
        from pathlib import Path
        from unittest import mock

        from Script.Medidas import Descarga_PRMTE as dp

        llamadas = []

        def falsa(sesion, punto, canal, periodo, user_key):
            llamadas.append((punto, canal))
            return pd.DataFrame({"idPuntoMedida": [punto], "v": [len(llamadas)]})

        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(dp, "extraer_datos_api", falsa):
            dp.descargar(["P1", "P2"], "202609", carpeta, user_key="x",
                         registrar=lambda *_: None)
            self.assertEqual(len(llamadas), 4)

            log = []
            df, _ = dp.descargar(["P1", "P2"], "202609", carpeta,
                                 user_key="x", registrar=log.append)
            self.assertEqual(len(llamadas), 8)
            self.assertEqual(len(df), 4)  # sin los lotes viejos
            self.assertTrue(any("vuelve a descargar" in l for l in log))
            self.assertEqual(
                len(list(Path(carpeta).glob("medidas_batch_202609_*"))), 1
            )

    def test_descarga_cortada_se_reanuda(self):
        import tempfile
        from unittest import mock

        from Script.Medidas import Descarga_PRMTE as dp

        llamadas = []

        def falsa(sesion, punto, canal, periodo, user_key):
            llamadas.append(punto)
            return pd.DataFrame({"idPuntoMedida": [punto], "v": [1]})

        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(dp, "extraer_datos_api", falsa):
            dp.descargar(["P1"], "202609", carpeta, user_key="x",
                         registrar=lambda *_: None)
            dp.descargar(["P1", "P2"], "202609", carpeta, user_key="x",
                         registrar=lambda *_: None)
            self.assertEqual(llamadas, ["P1", "P1", "P2", "P2"])
