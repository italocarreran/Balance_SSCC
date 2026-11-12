# -*- coding: utf-8 -*-
"""Herramientas/Extraer_Datos_PI.py: fechas, tramos, tabla y Excel (sin AF)."""

import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "Herramientas"))
import Extraer_Datos_PI as ed  # noqa: E402

RUTA = "\\\\SRV-SCADA-AF2\\SEN\\Generación\\SEN\\03 Región II\\SAE-ARENA"


class AFFalso:
    """Devuelve un valor por intervalo, con la hora local = UTC - 3."""

    def __init__(self):
        self.pedidos = []

    def atributo(self, ruta, nombre):
        return (ruta, nombre)

    def leer(self, atributo, desde, hasta, intervalo, tipo):
        self.pedidos.append((desde, hasta, intervalo, tipo))
        paso = timedelta(minutes=15 if intervalo == "15m" else 60)
        salida = []
        t = desde
        # Como PI: el instantaneo trae tambien la punta final y el
        # promedio puede traer un intervalo que empieza en 'hasta'.
        while t <= hasta:
            utc = (t + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
            salida.append((utc, t.strftime("%Y-%m-%d %H:%M:%S"),
                           float(t.hour)))
            t += paso
        return salida


PEDIDOS = [
    {"central": "SAE-ARENA", "ruta": RUTA, "atributo": "Potencia Activa",
     "unidad": "MW"},
    {"central": "SAE-ARENA", "ruta": RUTA,
     "atributo": "Porc Carga Disponible", "unidad": "%"},
]


class TestFechas(unittest.TestCase):

    def test_leer_fecha(self):
        self.assertEqual(ed.leer_fecha("2026-10-01 00:00"),
                         datetime(2026, 10, 1))
        self.assertEqual(ed.leer_fecha("2026-10-01"), datetime(2026, 10, 1))
        self.assertEqual(ed.leer_fecha("15-10-2026 13:30"),
                         datetime(2026, 10, 15, 13, 30))
        with self.assertRaises(ValueError):
            ed.leer_fecha("ayer")

    def test_mes_anterior(self):
        self.assertEqual(ed.mes_anterior(datetime(2026, 1, 20)),
                         (datetime(2025, 12, 1), datetime(2026, 1, 1)))

    def test_partir_rango(self):
        tramos = ed.partir_rango(datetime(2026, 1, 1), datetime(2026, 3, 15))
        self.assertEqual(tramos[0], (datetime(2026, 1, 1), datetime(2026, 2, 1)))
        self.assertEqual(tramos[-1][1], datetime(2026, 3, 15))
        for (_, fin), (inicio, _) in zip(tramos, tramos[1:]):
            self.assertEqual(fin, inicio)
        with self.assertRaises(ValueError):
            ed.partir_rango(datetime(2026, 2, 1), datetime(2026, 1, 1))


class TestTabla(unittest.TestCase):

    def test_la_hora_repetida_de_abril_no_se_pisa(self):
        # 23:00 local dos veces (UTC distinto), como el 4 de abril.
        serie = [("2026-04-05 02:00:00", "2026-04-04 23:00:00", 1.0),
                 ("2026-04-05 03:00:00", "2026-04-04 23:00:00", 2.0),
                 ("2026-04-05 04:00:00", "2026-04-05 00:00:00", 3.0)]
        tabla = ed.armar_tabla([("A", serie), ("A", serie[:1])])
        self.assertEqual(len(tabla), 3)
        self.assertEqual(tabla["A"].tolist(), [1.0, 2.0, 3.0])
        self.assertEqual(list(tabla.columns), ["fecha", "fecha UTC", "A", "A (2)"])


class TestExtraer(unittest.TestCase):

    def test_promedio_horario_de_dos_meses(self):
        af = AFFalso()
        desde, hasta = datetime(2026, 9, 1), datetime(2026, 11, 1)
        series = ed.extraer(af, PEDIDOS, desde, hasta, "1h", "promedio",
                            avisar=lambda *_: None)

        horas = (hasta - desde).days * 24
        self.assertEqual([len(v) for _, v in series], [horas, horas])
        self.assertEqual(series[0][0], "SAE-ARENA - Potencia Activa [MW]")
        self.assertEqual(len(af.pedidos), 2 * 2)  # dos tramos por dato
        tabla = ed.armar_tabla(series)
        self.assertEqual(tabla["fecha"].iloc[0], pd.Timestamp("2026-09-01"))
        self.assertEqual(tabla["fecha"].iloc[-1],
                         pd.Timestamp("2026-10-31 23:00"))

    def test_instantaneo_15_minutos_trae_la_punta_final(self):
        af = AFFalso()
        desde, hasta = datetime(2026, 10, 1), datetime(2026, 10, 2)
        series = ed.extraer(af, PEDIDOS[:1], desde, hasta, "15m",
                            "instantaneo", avisar=lambda *_: None)
        self.assertEqual(len(series[0][1]), 97)

    def test_excel(self):
        af = AFFalso()
        desde, hasta = datetime(2026, 10, 1), datetime(2026, 10, 2)
        series = ed.extraer(af, PEDIDOS, desde, hasta, "15m", "promedio",
                            avisar=lambda *_: None)
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "salida.xlsx"
            ed.escribir_excel(ruta, ed.armar_tabla(series), PEDIDOS, desde,
                              hasta, "15 minutos", "Promedio del intervalo")
            datos = pd.read_excel(ruta, sheet_name="Datos")
            pedido = pd.read_excel(ruta, sheet_name="Pedido")
        self.assertEqual(len(datos), 96)
        self.assertEqual(list(datos.columns)[2:],
                         ["SAE-ARENA - Potencia Activa [MW]",
                          "SAE-ARENA - Porc Carga Disponible [%]"])
        self.assertEqual(pedido["ruta"].tolist(), [RUTA, RUTA])
        self.assertEqual(pedido["intervalo"].iloc[0], "15 minutos")


if __name__ == "__main__":
    unittest.main()
