# -*- coding: utf-8 -*-
"""
El SoC desde PI AF: rutas del Diccionario (AF1 -> AF2), busqueda por
Nombre PCP/PID cuando no hay ruta, el archivo nuevo de tres columnas
(central | fecha | soc) y su lectura.

PI no esta en la maquina de pruebas: la conexion es una falsa que
responde como ConexionAF.
"""

import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import Script.nucleo as nucleo  # noqa: E402
from Script.Medidas import Soc_PI  # noqa: E402

RUTA_AF1 = (
    "\\\\SRV-SCADA-AF1\\SEN\\Generación\\SEN\\03 Región II\\"
    "SAE-Tocopilla|Nombre PCP/PID"
)
RUTA_AF2_TOCOPILLA = (
    "\\\\SRV-SCADA-AF2\\SEN\\Generación\\SEN\\03 Región II\\SAE-Tocopilla"
)
RUTA_ARENA = "\\\\SRV-SCADA-AF2\\SEN\\Generación\\SEN\\03 Región II\\SAE-ARENA"


class ElementoFalso:
    def __init__(self, ruta):
        self.ruta = ruta

    def GetPath(self):
        return self.ruta


class ConexionFalsa(Soc_PI.ConexionAF):
    """ConexionAF sin AF: elementos por ruta y por Nombre PCP/PID."""

    def __init__(self, por_ruta, por_nombre, valores):
        self.registrar = lambda *_: None
        self.por_ruta = por_ruta
        self._por_nombre = None
        self._nombres = por_nombre
        self.valores = valores
        self.pedidos = []

    def elemento_por_ruta(self, ruta):
        Soc_PI.partes_de_ruta(ruta)
        return self.por_ruta.get(ruta)

    def elementos_por_nombre_pcp(self):
        return {
            Soc_PI.normalizar(nombre): elemento
            for nombre, elemento in self._nombres.items()
        }

    def leer_soc(self, elemento, inicio, fin):
        self.pedidos.append((elemento.ruta, inicio, fin))
        return self.valores[elemento.ruta]


def escribir_centrales(ruta, filas_diccionario):
    libro = Workbook()
    resumen = libro.active
    resumen.title = "Resumen BESS"
    resumen.append(["Cuadro N° 1: Resumen BESS"])
    resumen.append([])
    resumen.append(["Nombre activo", "Pmax (MW)", "Capacidad (MWh)",
                    "Barra inyección", "Eficiencia"])
    for fila in filas_diccionario:
        resumen.append([fila[0], 10, 40, "BARRA_220", 0.98])

    dic = libro.create_sheet("Diccionario")
    dic.append([])
    dic.append(["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF",
                "SoC_PI"])
    for fila in filas_diccionario:
        dic.append(list(fila))
    libro.save(ruta)


def serie(inicio, cantidad, base):
    fechas = pd.date_range(inicio, periods=cantidad, freq="15min")
    return [
        (f.strftime("%Y-%m-%d %H:%M:%S"), base + i / 1000)
        for i, f in enumerate(fechas)
    ]


class TestRutas(unittest.TestCase):

    def test_af1_se_busca_en_af2_y_se_quita_el_atributo(self):
        self.assertEqual(
            Soc_PI.limpiar_ruta(RUTA_AF1), (RUTA_AF2_TOCOPILLA, True)
        )

    def test_af2_queda_igual(self):
        self.assertEqual(
            Soc_PI.limpiar_ruta(RUTA_ARENA + "|Nombre PCP/PID"),
            (RUTA_ARENA, False),
        )

    def test_vacia(self):
        for vacio in (None, "", "  ", float("nan")):
            self.assertEqual(Soc_PI.limpiar_ruta(vacio), ("", False))

    def test_partes(self):
        servidor, base, elementos = Soc_PI.partes_de_ruta(RUTA_AF2_TOCOPILLA)
        self.assertEqual((servidor, base), ("SRV-SCADA-AF2", "SEN"))
        self.assertEqual(elementos[-1], "SAE-Tocopilla")


class TestUbicar(unittest.TestCase):

    def setUp(self):
        self.toco = ElementoFalso(RUTA_AF2_TOCOPILLA)
        self.arena = ElementoFalso(RUTA_ARENA)
        self.conexion = ConexionFalsa(
            {RUTA_AF2_TOCOPILLA: self.toco}, {"SAE-ARENA": self.arena}, {}
        )

    def test_por_ruta_af1(self):
        elemento, ruta, origen = self.conexion.ubicar(RUTA_AF1, ["SAE-TOCOPILLA"])
        self.assertIs(elemento, self.toco)
        self.assertEqual(origen, Soc_PI.ORIGEN_RUTA_AF1)

    def test_sin_ruta_se_busca_por_nombre(self):
        elemento, ruta, origen = self.conexion.ubicar(None, ["sae-arena"])
        self.assertIs(elemento, self.arena)
        self.assertEqual((ruta, origen), (RUTA_ARENA, Soc_PI.ORIGEN_BUSQUEDA))

    def test_ruta_que_no_existe_cae_a_la_busqueda(self):
        elemento, _, origen = self.conexion.ubicar(
            RUTA_ARENA + "X", ["SAE-ARENA"]
        )
        self.assertIs(elemento, self.arena)
        self.assertEqual(origen, Soc_PI.ORIGEN_BUSQUEDA)

    def test_no_encontrada(self):
        elemento, motivo, origen = self.conexion.ubicar(None, ["SAE-LILE"])
        self.assertIsNone(elemento)
        self.assertIsNone(origen)
        self.assertIn("sin ruta", motivo)


class TestGenerarSoc(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        nucleo.crear_estructura_caso(self.base)
        escribir_centrales(
            self.base / "Auxiliares" / "Centrales.xlsx",
            [
                ("SAE-TOCOPILLA", "SAE TOCOPILLA", "SAE-TOCOPILLA",
                 "BAT_TOCOPILLA", "Tocopilla - BESS", RUTA_AF1),
                ("SAE-ARENA", "SAE ARENA", "SAE-ARENA", None,
                 "Arena - BESS", None),
                ("SAE-LILE", "SAE LILE", "SAE_LILE", "BAT_LILE", None, None),
            ],
        )
        toco = ElementoFalso(RUTA_AF2_TOCOPILLA)
        arena = ElementoFalso(RUTA_ARENA)
        valores_toco = serie("2026-10-01", 4, 0.5)
        valores_toco[1] = (valores_toco[1][0], None)  # mala calidad en PI
        self.conexion = ConexionFalsa(
            {RUTA_AF2_TOCOPILLA: toco},
            {"SAE ARENA": arena},  # su Nombre PCP/PID es el de FD
            {
                RUTA_AF2_TOCOPILLA: valores_toco,
                RUTA_ARENA: serie("2026-10-01", 4, 0.1),
            },
        )

    def tearDown(self):
        self.tmp.cleanup()

    def test_tres_columnas_y_se_lee_como_antes(self):
        log = []
        destino = nucleo.generar_soc(
            self.base, "2610", registrar=log.append, conexion=self.conexion
        )

        self.assertEqual(destino.name, "SOC_2610.xlsx")
        self.assertEqual(
            self.conexion.pedidos[0][1:],
            ("2026-10-01 00:00:00", "2026-11-01 00:00:00"),
        )

        df = pd.read_excel(destino, sheet_name="SoC")
        self.assertEqual(list(df.columns), ["central", "fecha", "soc"])
        self.assertEqual(len(df), 8)
        self.assertEqual(
            list(pd.unique(df["central"])), ["SAE-TOCOPILLA", "SAE-ARENA"]
        )
        self.assertTrue(pd.isna(df.loc[1, "soc"]))

        rutas = pd.read_excel(destino, sheet_name="Rutas PI")
        lile = rutas[rutas["central"] == "SAE-LILE"].iloc[0]
        self.assertEqual(lile["estado"], "No encontrada en AF")
        self.assertTrue(any("SAE-LILE" in linea for linea in log))

        # El programa lo encuentra y lo lee con extraer_soc.
        self.assertEqual(
            nucleo.buscar_soc(self.base / "Medidas", "2610"), destino
        )
        df_soc, incidencias = nucleo.extraer_soc(destino, {})
        self.assertEqual(
            list(df_soc.columns),
            ["central", "timestamp", "soc", "nombre_scada_original"],
        )
        self.assertEqual(len(df_soc), 8)
        self.assertEqual(
            df_soc["timestamp"].iloc[0], pd.Timestamp("2026-10-01 00:00")
        )
        self.assertTrue(any("SAE-TOCOPILLA" in i for i in incidencias))

    def test_la_ruta_no_es_un_sinonimo_de_la_central(self):
        _, diccionario = nucleo.leer_centrales(
            self.base / "Auxiliares" / "Centrales.xlsx"
        )
        mapa = nucleo.construir_homologacion(diccionario)
        self.assertNotIn(nucleo.normalizar(RUTA_AF1), mapa)
        self.assertEqual(mapa[nucleo.normalizar("BAT_TOCOPILLA")],
                         "SAE-TOCOPILLA")

        centrales = nucleo.centrales_para_soc(
            *nucleo.leer_centrales(self.base / "Auxiliares" / "Centrales.xlsx")
        )
        self.assertEqual(centrales[0]["ruta"], RUTA_AF1)
        self.assertNotIn(RUTA_AF1, centrales[0]["nombres"])
        self.assertIn("SAE ARENA", centrales[1]["nombres"])

    def test_sin_ninguna_central_en_pi_es_error(self):
        conexion = ConexionFalsa({}, {}, {})
        with self.assertRaises(nucleo.ErrorEntrada):
            nucleo.generar_soc(
                self.base, "2610", registrar=lambda *_: None,
                conexion=conexion,
            )


class TestFormatoViejoSigueLeyendose(unittest.TestCase):

    def test_soc_real_de_bloques(self):
        ruta = Path(__file__).resolve().parents[2] / "docs" / "SOC_real_2607.xlsx"
        df_soc, _ = nucleo.extraer_soc(ruta, {})
        self.assertGreater(df_soc["central"].nunique(), 5)


if __name__ == "__main__":
    unittest.main()
