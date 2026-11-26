# -*- coding: utf-8 -*-
"""
Medidas_CRA.xlsx desde PRMTE (pedido del usuario 2026-11-26): el mismo
proceso y el mismo formato que Medidas_SAE.xlsx, con la homologacion de
CRA/Auxiliares/. Y, para los dos, varios paños en una misma linea de la
homologacion: ALFALFAL_220_J1_AES+J2 = ALFALFAL_220_J1_AES + ALFALFAL_220_J2_AES.

La API se reemplaza por la descarga sintetica de test_claves_balance
(puntos P1 y P2; canal 1 vale 1, canal 3 vale 2).
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from Script import cra
from Script import nucleo
from Script.Medidas import Homologacion
from Script.Medidas.comun import ErrorMedidas
from Script.nucleo import medidas_sae
from desarrollo.tests.test_claves_balance import descarga

J1, J2 = "ALFALFAL_220_J1_AES", "ALFALFAL_220_J2_AES"


def descarga_alfalfal(inicio="2026-08-01", fin="2026-08-03"):
    """La descarga sintetica con los puntos llamados como dos paños."""

    df = descarga(inicio, fin)
    df["idPuntoMedida"] = df["idPuntoMedida"].map({"P1": J1, "P2": J2})
    return df


def homologacion(ruta, puntos):
    pd.DataFrame({
        "Punto de Medida": puntos,
        "Canal": ["s"] * len(puntos),
        "clave": ["ALFALFAL"] * len(puntos),
        "Flujo": [1] * len(puntos),
    }).to_excel(ruta, sheet_name="homol", index=False)


class TestPaños(unittest.TestCase):

    def test_ejemplos_del_usuario(self):
        casos = {
            "ALFALFAL_220_J1_AES+J2": [J1, J2],
            "LSQUILOS_006_G1_CLB+G2+G3": [
                "LSQUILOS_006_G1_CLB", "LSQUILOS_006_G2_CLB", "LSQUILOS_006_G3_CLB"],
            "CPLMAIQN_006_U1_EPP+U2+U3+U4+U5": [
                f"CPLMAIQN_006_U{i}_EPP" for i in range(1, 6)],
            "PULLINQE_066_BT1_EPP+BT2+BT3": [
                "PULLINQE_066_BT1_EPP", "PULLINQE_066_BT2_EPP", "PULLINQE_066_BT3_EPP"],
            "LHIGUERA_220_JT1_HLH+JT2": ["LHIGUERA_220_JT1_HLH", "LHIGUERA_220_JT2_HLH"],
            # El paño de cada '+' va tal como esta escrito.
            "CQLTHUES_012_CG1_AES+G2+G3": [
                "CQLTHUES_012_CG1_AES", "CQLTHUES_012_G2_AES", "CQLTHUES_012_G3_AES"],
            "SAUZAL_013_G1_END": ["SAUZAL_013_G1_END"],
            " SAUZAL_013_G1_END + G2 ": ["SAUZAL_013_G1_END", "SAUZAL_013_G2_END"],
        }
        for texto, esperado in casos.items():
            self.assertEqual(Homologacion.expandir_punto(texto), esperado, texto)

    def test_mal_escritos_se_detienen(self):
        for malo in ("SAUZAL+G2", "SAUZAL_013_G1_END+", "SAUZAL_013_G1_END+G2_X"):
            with self.assertRaises(ErrorMedidas, msg=malo):
                Homologacion.expandir_punto(malo)

    def test_la_hoja_se_expande_con_el_mismo_canal_clave_y_flujo(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "Homologacion.xlsx"
            homologacion(ruta, ["ALFALFAL_220_J1_AES+J2"])
            df = Homologacion.leer_homologacion(ruta)
        self.assertEqual(list(df["Punto de Medida"]), [J1, J2])
        self.assertEqual(set(df["clave"]), {"ALFALFAL"})
        self.assertEqual(
            set(df[Homologacion.COLUMNA_PUNTO_ORIGINAL]), {"ALFALFAL_220_J1_AES+J2"}
        )
        self.assertEqual(Homologacion.puntos_de_medida(df), [J1, J2])


class TestBessConPaños(unittest.TestCase):
    """En Medidas_SAE una linea con '+' baja los dos y los suma en su clave."""

    def test_suma_los_dos_paños(self):
        with tempfile.TemporaryDirectory() as tmp:
            caso = Path(tmp)
            nucleo.crear_estructura_caso(caso, registrar=lambda *a: None)
            rutas = nucleo.resolver_rutas(caso, "2608")
            homologacion(rutas["auxiliares_dir"] / "Homologacion ClavesTF y PRMTE.xlsx",
                         ["ALFALFAL_220_J1_AES+J2"])
            pedidos = {}

            def descargar(puntos, *a, **k):
                pedidos["puntos"] = list(puntos)
                return descarga_alfalfal(), []

            with mock.patch.object(medidas_sae.Descarga_PRMTE, "descargar", descargar):
                nucleo.generar_medidas_sae(caso, "2608", registrar=lambda *a: None)
            sae = pd.read_excel(rutas["medidas_sae"])

        self.assertEqual(pedidos["puntos"], [J1, J2])
        # Cada paño trae 1 (canal 1) + 2 (canal 3) por cuarto: 3 + 3.
        self.assertEqual(set(sae["Gen_Unidad"]), {6.0})
        self.assertEqual(set(sae["clave"]), {"ALFALFAL"})


class TestMedidasCra(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.mes = Path(self._tmp.name) / "2608"
        cra.crear_carpetas_caso(self.mes)
        self.rutas = cra.resolver_rutas(self.mes, "2608")
        homologacion(self.rutas["auxiliares"] / "Homologacion CRA y PRMTE.xlsx",
                     ["ALFALFAL_220_J1_AES+J2"])

    def tearDown(self):
        self._tmp.cleanup()

    def test_mismo_formato_que_medidas_sae_en_cra_energia(self):
        filas = {f["id"]: f for f in cra.revisar_estructura(self.mes, "2608")}
        self.assertEqual(filas["homologacion"]["estado"], "ok")
        self.assertEqual(filas["medidas_cra"]["estado"], "pendiente")

        with mock.patch.object(
            medidas_sae.Descarga_PRMTE, "descargar",
            lambda *a, **k: (descarga_alfalfal(), []),
        ):
            ruta = cra.generar_medidas_cra(self.mes, "2608", registrar=lambda *a: None)

        self.assertEqual(ruta, self.mes / "CRA" / "Energia" / "Medidas_CRA.xlsx")
        libro = pd.ExcelFile(ruta)
        self.assertEqual(libro.sheet_names, [nucleo.HOJA_MEDIDAS_SAE])
        leida = pd.read_excel(libro)
        self.assertEqual(list(leida.columns), nucleo.COLUMNAS_AI)
        self.assertEqual(len(leida), 2 * 96)
        self.assertEqual(set(leida["Gen_Unidad"]), {6.0})
        self.assertTrue(pd.read_excel(self.rutas["puntos_fallidos"]).empty)
        # Los lotes de la descarga quedan en la carpeta del CRA.
        self.assertEqual(self.rutas["trabajo_medidas"].parent, self.mes / "CRA" / "Energia")

        filas = {f["id"]: f for f in cra.revisar_estructura(self.mes, "2608")}
        self.assertEqual(filas["medidas_cra"]["estado"], "ok")

    def test_sin_homologacion_se_detiene(self):
        (self.rutas["auxiliares"] / "Homologacion CRA y PRMTE.xlsx").unlink()
        with self.assertRaises(cra.ErrorEntrada):
            cra.generar_medidas_cra(self.mes, "2608")


if __name__ == "__main__":
    unittest.main()
