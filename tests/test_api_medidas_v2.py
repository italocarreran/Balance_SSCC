"""
API de medidas v2 (medidas-v2/measurement), con respuestas sinteticas
que siguen el esquema que le dieron al usuario:

  - traducir_v2() deja las mismas columnas que la v1, tomando SOLO el
    canal pedido (la v2 puede traer channel1..channel12 llenos: sumarlos
    en las dos llamadas duplicaria la energia);
  - descargar() + construir_por_clave() con la v2 dan exactamente lo
    mismo que con la v1;
  - un 4xx no se reintenta y lo que respondio la API sale en el error.
"""

import json
import tempfile
import unittest
from unittest import mock

import pandas as pd
import requests

from Script.Medidas import Claves_Balance as cb
from Script.Medidas import Descarga_PRMTE as dp
from tests.test_claves_balance import HOMOLOGACION, descarga


def _hay_parquet():
    try:
        import pyarrow  # noqa: F401
        return True
    except ImportError:
        return False


def respuesta_v2(df_v1, punto):
    """El registro v2 de un punto a partir de los datos sinteticos v1."""
    del_punto = df_v1[df_v1["idPuntoMedida"] == punto]
    # Una fila por intervalo con los dos canales juntos (la v2 trae
    # todos los canales en cada fila).
    filas = (
        del_punto.groupby(["intervalo", "intervaloUtc"], sort=False)
        [["canalVal1", "canalVal3"]].max().reset_index()
    )
    texto = lambda v: "" if pd.isna(v) else f"{v:.3f}"
    medidas = [
        {
            "dateRange": f.intervalo.replace(" ", "T") + ".000Z",
            "utcRange": f.intervaloUtc,
            "principal": True,
            "channel1": texto(f.canalVal1),
            "channel2": "99.0",            # canales que no se piden
            "channel3": texto(f.canalVal3),
            "channel4": "",
        }
        for f in filas.itertuples()
    ]
    return [{
        "blog": "x",
        "measurement": medidas,
        "measurer": [{"name": "MED-1", "principal": True}],
        "coordinatorId": "COORD",
        "measurePointId": punto,
        "period": 202609,
        "subStation": "SE",
        "lastReadingDate": "2026-10-01T00:00:00.000Z",
        "channel": [
            {"channelId": 1, "slug": "s", "description": "Retiro"},
            {"channelId": 3, "slug": "s", "description": "Inyeccion"},
        ],
    }]


class Respuesta:
    def __init__(self, status, datos=None, texto=""):
        self.status_code = status
        self._datos = datos
        self.text = texto or json.dumps(datos)

    def json(self):
        return self._datos


class TestTraducir(unittest.TestCase):

    def test_columnas_v1_y_solo_el_canal_pedido(self):
        df_v1 = descarga("2026-08-01", "2026-08-02")
        registro = respuesta_v2(df_v1, "P1")[0]

        uno = dp.traducir_v2(registro, 1)
        tres = dp.traducir_v2(registro, 3)

        for columna in ("intervalo", "intervaloUtc", "principal",
                        "canalVal1", "canalVal3", "idPuntoMedida",
                        "slugCanal", "nombreMedidor", "fechaUltimaLectura"):
            self.assertIn(columna, uno.columns)
        self.assertFalse(any(c.startswith("channel") for c in uno.columns))
        self.assertTrue((uno["canalVal1"] == 1.0).all())
        self.assertTrue(uno["canalVal3"].isna().all())
        self.assertTrue((tres["canalVal3"] == 2.0).all())
        self.assertTrue(tres["canalVal1"].isna().all())
        self.assertEqual(tres["descripcionCanal"].iloc[0], "Inyeccion")
        self.assertEqual(uno["idPuntoMedida"].iloc[0], "P1")


@unittest.skipUnless(_hay_parquet(), "sin pyarrow (los lotes son parquet)")
class TestDescargaV2(unittest.TestCase):

    def correr(self, responder):
        llamadas = []

        class Sesion:
            def get(self, url, params=None, timeout=None, headers=None):
                llamadas.append((url, dict(params)))
                return responder(params)

        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(requests, "Session", Sesion), \
                mock.patch.object(dp, "ESPERA_REINTENTO", 0):
            resultado = dp.descargar(
                ["P1", "P2"], "202609", carpeta, user_key="clave",
                registrar=lambda *_: None,
            )
        return resultado, llamadas

    def test_mismo_resultado_que_la_v1(self):
        df_v1 = descarga("2026-09-01", "2026-10-01", "2026-09-06 00:00",
                         valores_hasta="2026-09-20 00:00")
        (df_v2, fallidos), llamadas = self.correr(
            lambda p: Respuesta(200, respuesta_v2(df_v1, p["measurePointId"]))
        )

        self.assertEqual(fallidos, [])
        url, params = llamadas[0]
        self.assertEqual(url, dp.URL_MEDIDAS_V2)
        self.assertEqual(params["period"], "202609010000")
        self.assertEqual(params["channelId"], 1)
        self.assertEqual(params["user_key"], "clave")

        con_v1, _, diag_v1 = cb.construir_por_clave(
            df_v1, HOMOLOGACION, registrar=lambda *_: None)
        con_v2, _, diag_v2 = cb.construir_por_clave(
            df_v2, HOMOLOGACION, registrar=lambda *_: None)

        pd.testing.assert_frame_equal(
            con_v1.reset_index(drop=True), con_v2.reset_index(drop=True),
            check_dtype=False,
        )
        pd.testing.assert_frame_equal(
            diag_v1["huecos"].reset_index(drop=True),
            diag_v2["huecos"].reset_index(drop=True),
        )

    def test_4xx_no_se_reintenta_y_el_error_lo_dice(self):
        with self.assertRaises(dp.ErrorMedidas) as error:
            self.correr(lambda p: Respuesta(
                400, texto='{"message":"period invalido"}'))
        self.assertIn("period invalido", str(error.exception))
        self.assertIn("HTTP 400", str(error.exception))

    def test_4xx_una_llamada_por_punto_y_canal(self):
        contador = []

        def responder(p):
            contador.append(1)
            return Respuesta(403, texto="sin permiso")

        with self.assertRaises(dp.ErrorMedidas):
            self.correr(responder)
        self.assertEqual(len(contador), 4)    # 2 puntos x 2 canales


if __name__ == "__main__":
    unittest.main()
