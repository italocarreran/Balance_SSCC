"""
Medidas_SAE.xlsx con informacion incompleta (pedido del usuario):

  - se genera IGUAL con lo que hay;
  - al lado se escribe SIEMPRE Puntos_fallidos.xlsx: con un renglon por
    punto de medida no encontrado en la API y por tramo sin informacion
    (punto, canal, desde, hasta), o vacio si no falto nada;
  - el arbol de la ventana le cuelga un aviso a la fila de
    Medidas_SAE.xlsx solo si Puntos_fallidos.xlsx tiene renglones.

La API se reemplaza por datos sinteticos (tests.test_claves_balance).
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from Script import nucleo
from Script.nucleo import estructura, medidas_sae
from tests.test_claves_balance import descarga


class TestPuntosFallidos(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.caso = Path(self.tmp.name)
        for sub in nucleo.SUBCARPETAS_CASO:
            (self.caso / sub).mkdir(parents=True, exist_ok=True)
        self.rutas = nucleo.resolver_rutas(self.caso)
        # P3 esta en la homologacion pero la API no lo devuelve.
        pd.DataFrame({
            "clave": ["A", "B", "C"],
            "Punto de Medida": ["P1", "P2", "P3"],
            "Canal": ["s", "s", "s"],
            "Flujo": [1, -1, 1],
        }).to_excel(
            self.rutas["auxiliares_dir"] / "Homologacion ClavesTF y PRMTE.xlsx",
            sheet_name="homol", index=False,
        )

    def tearDown(self):
        self.tmp.cleanup()

    def correr(self, df_api):
        log = []
        with mock.patch.object(
            medidas_sae.Descarga_PRMTE, "descargar",
            lambda *a, **k: (df_api, []),
        ):
            medidas_sae.generar_medidas_sae(
                self.caso, "2609", registrar=log.append
            )
        return pd.read_excel(self.rutas["puntos_fallidos"]), log

    def test_incompleto_se_genera_igual_y_detalla(self):
        df_api = descarga(
            "2026-09-01", "2026-10-01", "2026-09-06 00:00",
            valores_hasta="2026-09-13 23:00",
        )
        fallidos, log = self.correr(df_api)

        sae = pd.read_excel(self.rutas["medidas_sae"])
        self.assertEqual(sorted(sae["clave"].astype(str).unique()), ["A", "B"])

        no_encontrado = fallidos[fallidos["Problema"] == "No encontrado en la API"]
        self.assertEqual(list(no_encontrado["Punto de Medida"]), ["P3"])
        self.assertEqual(list(no_encontrado["Clave"]), ["C"])

        huecos = fallidos[fallidos["Problema"] == "Sin informacion"]
        self.assertEqual(sorted(huecos["Punto de Medida"].unique()), ["P1", "P2"])
        self.assertEqual(set(huecos["Canal"]), {"Canal 1", "Canal 3"})
        self.assertEqual(set(pd.to_datetime(huecos["Desde"])),
                         {pd.Timestamp("2026-09-13 23:00")})
        self.assertEqual(set(pd.to_datetime(huecos["Hasta"])),
                         {pd.Timestamp("2026-09-30 23:45")})
        self.assertTrue(any("informacion incompleta" in l for l in log))

        aviso = estructura.aviso_puntos_fallidos(self.rutas["puntos_fallidos"])
        self.assertIsNotNone(aviso)
        self.assertIn("1 punto(s) de medida no encontrados", aviso["mensaje"])
        self.assertIn("2 punto(s) de medida incompletos", aviso["mensaje"])
        self.assertIn("de 13-09-2026 23:00 a 30-09-2026 23:45", aviso["mensaje"])
        self.assertEqual(aviso["ruta"], str(self.rutas["puntos_fallidos"]))

        _, filas = nucleo.revisar_estructura(self.caso, "2609")
        por_id = {f["id"]: f for f in filas}
        self.assertEqual(por_id["medidas_sae"]["aviso"], aviso)
        self.assertEqual(por_id["puntos_fallidos"]["estado"], "ok")

    def test_sin_problemas_el_archivo_queda_vacio(self):
        pd.read_excel(  # P3 fuera: todos los puntos llegan completos
            self.rutas["auxiliares_dir"] / "Homologacion ClavesTF y PRMTE.xlsx"
        ).iloc[:2].to_excel(
            self.rutas["auxiliares_dir"] / "Homologacion ClavesTF y PRMTE.xlsx",
            sheet_name="homol", index=False,
        )
        fallidos, log = self.correr(descarga("2026-09-01", "2026-10-01",
                                             "2026-09-06 00:00"))

        self.assertTrue(self.rutas["puntos_fallidos"].is_file())
        self.assertTrue(fallidos.empty)
        self.assertEqual(list(fallidos.columns),
                         nucleo.parametros.COLUMNAS_PUNTOS_FALLIDOS)
        self.assertIsNone(
            estructura.aviso_puntos_fallidos(self.rutas["puntos_fallidos"])
        )
        _, filas = nucleo.revisar_estructura(self.caso, "2609")
        self.assertIsNone(
            {f["id"]: f for f in filas}["medidas_sae"]["aviso"]
        )


if __name__ == "__main__":
    unittest.main()
