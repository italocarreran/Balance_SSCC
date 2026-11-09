# -*- coding: utf-8 -*-
"""Herramientas/Buscar_Puntos_PRMTE.py: lista de puntos y busqueda aproximada."""

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "Herramientas"))
import Buscar_Puntos_PRMTE as bp  # noqa: E402

V1 = [
    {"idCoordinado": "X", "idPuntoMedida": "ARENA_220_JT1_ARE", "subEstacion": "ARENA 220", "mediciones": []},
    {"idCoordinado": "X", "idPuntoMedida": "TOCOPILLA_110_J5", "subEstacion": "TOCOPILLA", "mediciones": []},
    {"idCoordinado": "X", "idPuntoMedida": "ARENA_220_JT1_ARE", "subEstacion": "ARENA 220", "mediciones": []},
]
V2 = {"data": [{"measurePointId": "VICTOR_JARA_220_J1", "subStation": "VICTOR JARA"}]}


class TestPuntos(unittest.TestCase):

    def test_lee_respuestas_v1_y_v2_sin_repetidos(self):
        puntos = bp.a_puntos(bp._registros(V1)) + bp.a_puntos(bp._registros(V2))
        self.assertEqual(
            [p["id"] for p in puntos],
            ["ARENA_220_JT1_ARE", "TOCOPILLA_110_J5", "VICTOR_JARA_220_J1"],
        )

    def test_busca_aproximado(self):
        puntos = bp.a_puntos(bp._registros(V1)) + bp.a_puntos(bp._registros(V2))
        puntos += [{"id": f"PUNTO_{i:04d}_220", "detalle": "SE X"} for i in range(1000)]
        for texto, esperado in (
            ("arena 220", "ARENA_220_JT1_ARE"),
            ("tocopila", "TOCOPILLA_110_J5"),
            ("jara victor", "VICTOR_JARA_220_J1"),
            ("punto 7 220", "PUNTO_0007_220"),
        ):
            self.assertEqual(bp.buscar(puntos, texto)[0][0]["id"], esperado, texto)
        self.assertEqual(bp.buscar(puntos, " "), [])

    def test_prueba_las_rutas_hasta_que_una_sirve(self):
        class R:
            def __init__(self, codigo, datos=None):
                self.status_code, self._datos, self.text = codigo, datos, "x"

            def json(self):
                return self._datos

        respuestas = [R(404), R(200, V1)]
        with mock.patch("requests.get", side_effect=lambda *a, **k: respuestas.pop(0)):
            puntos, url = bp.bajar_puntos("t", avisar=lambda *_: None)
        self.assertEqual(url, bp.RUTAS_LISTA[1])
        self.assertEqual(len(puntos), 2)

    def test_token_y_lista_en_config_hacia_arriba(self):
        with tempfile.TemporaryDirectory() as raiz:
            raiz = Path(raiz)
            (raiz / "__config__").mkdir()
            (raiz / "__config__" / "tks.json").write_text(
                json.dumps({"coordinador": {"prmte": "ABC"}}), encoding="utf-8")
            (raiz / "Balance_SSCC" / "Herramientas").mkdir(parents=True)
            with mock.patch.object(bp, "AQUI", raiz / "Balance_SSCC" / "Herramientas"):
                self.assertEqual(bp.carpeta_config(), raiz / "__config__")
                self.assertEqual(bp.token_guardado(), "ABC")
                self.assertEqual(bp.ruta_lista(), raiz / "__config__" / "puntos_prmte.json")


if __name__ == "__main__":
    unittest.main()
