"""
Descargas de las APIs en paralelo (pedido del usuario 2026-11-03: "lo de
hacerlo en paralelo"): con varios pedidos a la vez el resultado tiene que
ser EXACTAMENTE el de pedirlos uno por uno (mismas filas, mismo orden,
mismos lotes y fallidos), aunque las respuestas lleguen desordenadas.
"""

import random
import tempfile
import threading
import time
import unittest
from unittest import mock

import pandas as pd

from Script.Medidas import Descarga_PRMTE as dp
from Script.Medidas import Generacion_Real as gr


def _falsa_prmte(sesion, punto, canal, periodo, user_key):
    time.sleep(random.uniform(0, 0.01))
    if punto.endswith("7"):
        return pd.DataFrame()
    return pd.DataFrame({
        "idPuntoMedida": [punto] * 3, "canal": canal, "v": [1.5, 2.5, 3.5],
    })


try:
    import pyarrow  # noqa: F401  (los lotes de la descarga son parquet)
    HAY_PARQUET = True
except ImportError:  # pragma: no cover
    HAY_PARQUET = False


@unittest.skipUnless(HAY_PARQUET, "sin pyarrow")
class TestDescargaPrmteParalela(unittest.TestCase):

    def _bajar(self, hilos):
        puntos = [f"P{i}" for i in range(230)]
        log = []
        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(dp, "extraer_datos_api", _falsa_prmte), \
                mock.patch.object(dp, "HILOS_DESCARGA", hilos), \
                mock.patch.object(dp, "TAMANO_LOTE", 40):
            df, fallidos = dp.descargar(
                puntos, "202609", carpeta, user_key="x", registrar=log.append,
            )
        return df, fallidos, log

    def test_igual_que_uno_por_uno(self):
        df1, f1, log1 = self._bajar(1)
        df8, f8, log8 = self._bajar(8)
        pd.testing.assert_frame_equal(df1, df8)
        self.assertEqual(f1, f8)
        self.assertEqual(log1, log8)

    def test_usa_varios_hilos(self):
        vistos = set()

        def falsa(sesion, punto, canal, periodo, user_key):
            vistos.add(threading.get_ident())
            time.sleep(0.005)
            return pd.DataFrame({"idPuntoMedida": [punto], "v": [1]})

        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(dp, "extraer_datos_api", falsa):
            dp.descargar([f"P{i}" for i in range(40)], "202609", carpeta,
                         user_key="x", registrar=lambda *_: None)
        self.assertGreater(len(vistos), 1)


class _Respuesta:
    def __init__(self, dia):
        self.dia = dia
        self.ok = dia != 5
        self.status_code = 200 if self.ok else 503

    def json(self):
        return [{"topologyName": f"C{self.dia % 3}", "topologyTimeHour": h,
                 "measure": self.dia * 10 + h, "idMeasure": h}
                for h in range(3)]


class _Sesion:
    def get(self, url, params=None, timeout=None):
        time.sleep(random.uniform(0, 0.01))
        return _Respuesta(int(params["created"][-2:]))


class TestGeneracionRealParalela(unittest.TestCase):

    def _bajar(self, hilos):
        log = []
        with mock.patch("requests.Session", _Sesion), \
                mock.patch.object(gr, "HILOS_DESCARGA", hilos), \
                mock.patch.object(gr, "ESPERA_REINTENTO", 0), \
                mock.patch.object(gr, "_json_a_dataframe", pd.DataFrame):
            df = gr.descargar_mes(2026, 9, 30, user_key="x",
                                  registrar=log.append)
        return df, log

    def test_igual_que_uno_por_uno(self):
        df1, log1 = self._bajar(1)
        df8, log8 = self._bajar(8)
        pd.testing.assert_frame_equal(df1, df8)
        self.assertEqual(log1, log8)
        self.assertTrue(any("2026-09-05" in l for l in log8))


if __name__ == "__main__":
    unittest.main()
