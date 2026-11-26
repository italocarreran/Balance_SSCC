# -*- coding: utf-8 -*-
"""
Medidas_CRA.xlsx desde PRMTE (pedido del usuario 2026-11-26): misma idea
que Medidas_SAE, con la forma del Formato_Solicitud_SSAA_SSCC_Hidro_*
para compararlos:

    Unidad, Punto De Medida, Año, Mes, DIA, FECHA_HORA, kWhD, kWhR

La API se reemplaza por la descarga sintetica de test_claves_balance
(canal 1 vale 1, canal 3 vale 2), con slugCanal kWhD / kWhR.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from Script import cra
from Script.cra import medidas_cra
from Script.cra import parametros as p
from desarrollo.tests.test_claves_balance import descarga


def descarga_cra(inicio="2026-08-01", fin="2026-08-03", **kwargs):
    """La descarga sintetica con el slug real de cada canal."""

    df = descarga(inicio, fin, **kwargs)
    # descarga() apila, por punto, primero el canal 1 y despues el 3.
    canal1 = df.groupby("idPuntoMedida").cumcount() < (
        df.groupby("idPuntoMedida")["idPuntoMedida"].transform("size") // 2
    )
    df["slugCanal"] = canal1.map({True: "kWhD", False: "kWhR"})
    return df


HOMOLOGACION = pd.DataFrame({
    "Punto de Medida": ["P1", "P1", "P2"],
    "Canal": ["kWhD", "kWhR", "kWhD"],
    "clave": ["RAPEL-1", "RAPEL-1", "CANUTILLAR-1"],
    "Flujo": [1, -1, 1],
})


class TestConstruir(unittest.TestCase):

    def test_tipo_de_canal(self):
        self.assertEqual(medidas_cra.tipo_de_canal("kWhD"), "kWhD")
        self.assertEqual(medidas_cra.tipo_de_canal(" kwhr "), "kWhR")
        self.assertIsNone(medidas_cra.tipo_de_canal("s"))

    def test_una_fila_por_unidad_punto_y_cuarto_con_d_y_r(self):
        df, calendario, huecos = cra.construir_medidas_cra(
            descarga_cra(), HOMOLOGACION
        )
        self.assertEqual(list(df.columns), p.COLUMNAS_MEDIDAS_CRA)
        self.assertEqual(len(calendario), 2 * 96)
        self.assertTrue(huecos.empty)
        # Dos unidades/puntos, 192 cuartos cada uno.
        self.assertEqual(len(df), 2 * 192)

        rapel = df[df["Unidad"] == "RAPEL-1"]
        self.assertEqual(set(rapel["Punto De Medida"]), {"P1"})
        # Tal como la API (sin Flujo): kWhD del canal 1, kWhR del 3.
        self.assertEqual(set(rapel["kWhD"]), {1.0})
        self.assertEqual(set(rapel["kWhR"]), {2.0})
        canutillar = df[df["Unidad"] == "CANUTILLAR-1"]
        self.assertEqual(set(canutillar["kWhD"]), {1.0})
        self.assertEqual(set(canutillar["kWhR"]), {0.0})   # sin canal R

        primera = rapel.iloc[0]
        self.assertEqual(
            (primera["Año"], primera["Mes"], primera["DIA"]), (2026, 8, 1)
        )
        self.assertEqual(primera["FECHA_HORA"], pd.Timestamp("2026-08-01 00:00"))
        self.assertTrue(rapel["FECHA_HORA"].is_monotonic_increasing)

    def test_canal_desconocido_se_avisa_y_queda_fuera(self):
        homol = pd.concat([HOMOLOGACION, pd.DataFrame({
            "Punto de Medida": ["P2"], "Canal": ["kWhR"],
            "clave": ["CANUTILLAR-1"], "Flujo": [1],
        })], ignore_index=True)
        crudo = descarga_cra()
        crudo.loc[(crudo["idPuntoMedida"] == "P2") & (crudo["slugCanal"] == "kWhR"),
                  "slugCanal"] = "raro"
        homol.loc[3, "Canal"] = "raro"
        log = []
        df, _, _ = cra.construir_medidas_cra(crudo, homol, registrar=log.append)
        self.assertTrue(any("raro" in l and "quedan fuera" in l for l in log))
        self.assertEqual(
            set(df[df["Unidad"] == "CANUTILLAR-1"]["kWhR"]), {0.0}
        )


class TestGenerar(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.mes = Path(self._tmp.name) / "2608"
        cra.crear_carpetas_caso(self.mes)
        self.rutas = cra.resolver_rutas(self.mes, "2608")
        HOMOLOGACION.to_excel(
            self.rutas["auxiliares"] / "Homologacion CRA y PRMTE.xlsx",
            sheet_name="homol", index=False,
        )

    def tearDown(self):
        self._tmp.cleanup()

    def test_genera_en_cra_energia_y_el_arbol_lo_muestra(self):
        filas = {f["id"]: f for f in cra.revisar_estructura(self.mes, "2608")}
        self.assertEqual(filas["homologacion"]["estado"], "ok")
        self.assertEqual(filas["medidas_cra"]["estado"], "pendiente")

        log = []
        with mock.patch.object(
            medidas_cra.Descarga_PRMTE, "descargar",
            lambda *a, **k: (descarga_cra(), []),
        ):
            ruta = cra.generar_medidas_cra(self.mes, "2608", registrar=log.append)

        self.assertEqual(ruta, self.mes / "CRA" / "Energia" / "Medidas_CRA.xlsx")
        leida = pd.read_excel(ruta)
        self.assertEqual(list(leida.columns), p.COLUMNAS_MEDIDAS_CRA)
        self.assertEqual(len(leida), 2 * 192)
        self.assertTrue(self.rutas["puntos_fallidos"].is_file())
        self.assertTrue(pd.read_excel(self.rutas["puntos_fallidos"]).empty)

        filas = {f["id"]: f for f in cra.revisar_estructura(self.mes, "2608")}
        self.assertEqual(filas["medidas_cra"]["estado"], "ok")

    def test_sin_homologacion_se_detiene(self):
        (self.rutas["auxiliares"] / "Homologacion CRA y PRMTE.xlsx").unlink()
        with self.assertRaises(cra.ErrorEntrada):
            cra.generar_medidas_cra(self.mes, "2608")


if __name__ == "__main__":
    unittest.main()
