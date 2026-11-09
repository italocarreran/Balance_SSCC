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

    def test_baja_todas_las_paginas(self):
        """La forma de /medidas-v2/measurement-point/all del portal."""

        def pagina(n, total=3):
            return {"page": n, "pageSize": 2, "totalElements": 5, "totalPages": total,
                    "content": [{"subStation": f"SE {n}{k}", "region": "II", "tension": "220",
                                 "coordinatorId": "C1", "measurePointId": f"P{n}{k}"}
                                for k in range(2 if n < 2 else 1)]}

        class R:
            status_code, text = 200, ""

            def __init__(self, datos):
                self._datos = datos

            def json(self):
                return self._datos

        pedidas = []

        def get(url, params=None, **_):
            pedidas.append(params["page"])
            return R(pagina(params["page"]))

        with mock.patch("requests.get", side_effect=get):
            puntos, url = bp.bajar_puntos("t", avisar=lambda *_: None)
        self.assertEqual(url, bp.URL_PUNTOS)
        self.assertEqual(pedidas, [0, 1, 2])
        self.assertEqual([p["id"] for p in puntos], ["P00", "P01", "P10", "P11", "P20"])
        self.assertEqual(puntos[0]["subStation"], "SE 00")
        self.assertEqual(puntos[0]["tension"], "220")
        self.assertEqual(bp.buscar(puntos, "se 11")[0][0]["id"], "P11")

    def test_error_de_la_api_dice_que_respondio(self):
        class R:
            status_code, text = 401, "clave invalida"

        with mock.patch("requests.get", return_value=R()):
            with self.assertRaisesRegex(RuntimeError, "HTTP 401 clave invalida"):
                bp.bajar_puntos("t", avisar=lambda *_: None)

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
