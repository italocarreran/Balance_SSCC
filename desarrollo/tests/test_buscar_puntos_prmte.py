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

class R:
    """Respuesta falsa de requests."""

    def __init__(self, datos, status_code=200, text=""):
        self._datos, self.status_code, self.text = datos, status_code, text

    def json(self):
        return self._datos


class TestListaIncompletaYComprobar(unittest.TestCase):

    def test_sigue_pidiendo_si_la_api_recorta_el_pagesize(self):
        """totalPages=1 pero totalElements=4 y llegan de a 2: pide mas."""

        def get(url, params=None, **_):
            n = params["page"]
            contenido = [{"measurePointId": f"P{n}{k}"} for k in range(2)] if n < 2 else []
            return R({"totalPages": 1, "totalElements": 4, "content": contenido})

        avisos = []
        with mock.patch("requests.get", side_effect=get):
            puntos, _ = bp.bajar_puntos("t", avisar=avisos.append)
        self.assertEqual(len(puntos), 4)
        self.assertNotIn("OJO", avisos[-1])

    def test_avisa_si_faltan_puntos(self):
        def get(url, params=None, **_):
            contenido = [{"measurePointId": "P1"}] if params["page"] == 0 else []
            return R({"totalPages": 3, "totalElements": 10, "content": contenido})

        avisos = []
        with mock.patch("requests.get", side_effect=get):
            bp.bajar_puntos("t", avisar=avisos.append)
        self.assertIn("OJO", avisos[-1])
        self.assertIn("10", avisos[-1])

    def test_comprobar_punto_con_medidas(self):
        pedidos = []

        def get(url, params=None, **_):
            pedidos.append((url, params["channelId"], params["period"]))
            if params["channelId"] == 1:
                return R([{"measurement": []}])
            return R([{"measurement": [{"channel3": "1.5"}], "subStation": "SE X"}])

        with mock.patch("requests.get", side_effect=get):
            punto, mensaje = bp.comprobar_punto("t", " P_NUEVO ", periodo="202610")
        self.assertEqual(punto["id"], "P_NUEVO")
        self.assertEqual(punto["subStation"], "SE X")
        self.assertEqual(punto["origen"], "medidas")
        self.assertIn("existe", mensaje)
        self.assertEqual(pedidos[0], (bp.URL_MEDIDAS, 1, "202610010000"))

    def test_comprobar_punto_sin_medidas(self):
        with mock.patch("requests.get", return_value=R([], 404, "no existe")):
            punto, mensaje = bp.comprobar_punto("t", "NO_EXISTE", periodo="202610")
        self.assertIsNone(punto)
        self.assertIn("HTTP 404", mensaje)

    def test_unir_encontrados_sin_repetir(self):
        lista = [{"id": "A"}, {"id": "B"}]
        unidos = bp.unir_encontrados(lista, [{"id": "B"}, {"id": "C"}])
        self.assertEqual([p["id"] for p in unidos], ["A", "B", "C"])

    def test_mes_anterior(self):
        from datetime import date
        self.assertEqual(bp.mes_anterior(date(2026, 1, 15)), "202512")
        self.assertEqual(bp.mes_anterior(date(2026, 11, 16)), "202610")


if __name__ == "__main__":
    unittest.main()
