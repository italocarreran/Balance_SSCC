"""Los dos tokens de las APIs del Coordinador salen de __config__/tks.json
(o, mientras no esten ahi, de "claves_api" del config.json de antes).

Son distintas entre si, no dependen de quien corra el programa, y el
archivo que las guarda no se versiona: si falta, el error tiene que
decir exactamente que pegar y donde.
"""

import json
import tempfile
import unittest
from pathlib import Path

from Script import config
from Script.Medidas import comun


class ClavesApiTest(unittest.TestCase):

    def setUp(self):
        self.carpeta = tempfile.TemporaryDirectory()
        self.ruta = Path(self.carpeta.name) / "config.json"
        self.original = config.RUTA_CONFIG
        config.RUTA_CONFIG = self.ruta

    def tearDown(self):
        config.RUTA_CONFIG = self.original
        self.carpeta.cleanup()

    def escribir(self, contenido):
        self.ruta.write_text(json.dumps(contenido), encoding="utf-8")

    def test_devuelve_cada_clave_por_separado(self):
        self.escribir({"claves_api": {"prmte": "AAA", "generacion_real": "BBB"}})

        self.assertEqual(config.clave_api(config.CLAVE_PRMTE), "AAA")
        self.assertEqual(config.clave_api(config.CLAVE_GENERACION_REAL), "BBB")

    def test_sin_archivo_el_error_dice_que_pegar_y_donde(self):
        with self.assertRaises(config.ErrorConfig) as caja:
            config.clave_api(config.CLAVE_PRMTE)

        mensaje = str(caja.exception)
        self.assertIn(str(config.RUTA_TOKENS), mensaje)
        self.assertIn('"coordinador"', mensaje)
        self.assertIn('"prmte"', mensaje)
        self.assertIn('"generacion_real"', mensaje)

    def test_el_ejemplo_sin_reemplazar_cuenta_como_vacio(self):
        self.escribir({"claves_api": {"prmte": "PEGAR_AQUI_LA_CLAVE"}})

        with self.assertRaises(config.ErrorConfig):
            config.clave_api(config.CLAVE_PRMTE)

    def test_una_clave_puesta_no_tapa_la_otra_que_falta(self):
        self.escribir({"claves_api": {"prmte": "AAA"}})

        self.assertEqual(config.clave_api(config.CLAVE_PRMTE), "AAA")
        with self.assertRaises(config.ErrorConfig):
            config.clave_api(config.CLAVE_GENERACION_REAL)

    def test_medidas_lo_ve_como_ErrorMedidas(self):
        with self.assertRaises(comun.ErrorMedidas) as caja:
            comun.leer_clave_api(comun.CLAVE_PRMTE)

        self.assertIn("tks.json", str(caja.exception))

    def test_archivo_roto_no_explota(self):
        self.ruta.write_text("{ esto no es json", encoding="utf-8")

        self.assertEqual(config.leer_todo(), {})
        with self.assertRaises(config.ErrorConfig):
            config.clave_api(config.CLAVE_PRMTE)


class SeccionesTest(unittest.TestCase):
    """Guardar lo de un usuario no puede pisar las claves compartidas."""

    def setUp(self):
        self.carpeta = tempfile.TemporaryDirectory()
        self.original = config.RUTA_CONFIG
        config.RUTA_CONFIG = Path(self.carpeta.name) / "config.json"

    def tearDown(self):
        config.RUTA_CONFIG = self.original
        self.carpeta.cleanup()

    def test_actualizar_una_seccion_conserva_las_demas(self):
        config.actualizar_seccion("claves_api", {"prmte": "AAA"})
        config.actualizar_seccion("PC1_italo", {"aamm": "2607"})
        config.actualizar_seccion("PC2_otro", {"aamm": "2608"})
        config.actualizar_seccion("PC1_italo", {"carpeta_base": "C:/casos"})

        todo = config.leer_todo()
        self.assertEqual(todo["claves_api"], {"prmte": "AAA"})
        self.assertEqual(todo["PC1_italo"],
                         {"aamm": "2607", "carpeta_base": "C:/casos"})
        self.assertEqual(todo["PC2_otro"], {"aamm": "2608"})


class ModosYTokensTest(unittest.TestCase):
    """__config__/config.json compartido (modo_bess, modo_cra) y tks.json."""

    def setUp(self):
        self.carpeta = tempfile.TemporaryDirectory()
        c = Path(self.carpeta.name)
        self.originales = (config.RUTA_CONFIG, config.RUTA_TOKENS,
                           config.RUTA_CONFIG_ANTIGUA)
        config.RUTA_CONFIG = c / "__config__" / "config.json"
        config.RUTA_TOKENS = c / "__config__" / "tks.json"
        config.RUTA_CONFIG_ANTIGUA = c / "Balance_SSCC" / "config.json"

    def tearDown(self):
        (config.RUTA_CONFIG, config.RUTA_TOKENS,
         config.RUTA_CONFIG_ANTIGUA) = self.originales
        self.carpeta.cleanup()

    def test_no_borra_lo_de_los_otros_programas(self):
        config.RUTA_CONFIG.parent.mkdir(parents=True)
        config.RUTA_CONFIG.write_text(json.dumps({
            "PC1_italo": {"tab_anio": "2025"},
            "modo_15min": {"PC1_italo": {"ultimo_mes": "2601"}},
        }), encoding="utf-8")
        config.actualizar_seccion_usuario(config.MODO_BESS, "PC1_italo", {"aamm": "2609"})
        config.actualizar_seccion_usuario(config.MODO_CRA, "PC1_italo", {"aamm": "2512"})
        todo = json.loads(config.RUTA_CONFIG.read_text(encoding="utf-8"))
        self.assertEqual(todo["PC1_italo"], {"tab_anio": "2025"})
        self.assertEqual(todo["modo_15min"], {"PC1_italo": {"ultimo_mes": "2601"}})
        self.assertEqual(todo["modo_bess"], {"PC1_italo": {"aamm": "2609"}})
        self.assertEqual(todo["modo_cra"], {"PC1_italo": {"aamm": "2512"}})

    def test_un_archivo_roto_no_se_pisa(self):
        config.RUTA_CONFIG.parent.mkdir(parents=True)
        config.RUTA_CONFIG.write_text("{ roto", encoding="utf-8")
        self.assertFalse(
            config.actualizar_seccion_usuario(config.MODO_BESS, "PC1", {"a": 1})
        )
        self.assertEqual(config.RUTA_CONFIG.read_text(encoding="utf-8"), "{ roto")

    def test_la_primera_vez_copia_lo_de_antes(self):
        config.RUTA_CONFIG_ANTIGUA.parent.mkdir(parents=True)
        config.RUTA_CONFIG_ANTIGUA.write_text(json.dumps({
            "PC1_italo": {"carpeta_base": "T:/caso", "aamm": "2609"},
            "PC1_italo_CRA": {"carpeta_base": "T:/cra"},
            "claves_api": {"prmte": "AAA", "generacion_real": "BBB"},
        }), encoding="utf-8")
        self.assertEqual(
            config.seccion_usuario(config.MODO_BESS, "PC1_italo", "PC1_italo"),
            {"carpeta_base": "T:/caso", "aamm": "2609"},
        )
        self.assertEqual(
            config.seccion_usuario(config.MODO_CRA, "PC1_italo", "PC1_italo_CRA"),
            {"carpeta_base": "T:/cra"},
        )
        self.assertTrue(config.migrar_tokens())
        self.assertEqual(
            json.loads(config.RUTA_TOKENS.read_text(encoding="utf-8")),
            {"coordinador": {"prmte": "AAA", "generacion_real": "BBB"}},
        )
        self.assertFalse(config.migrar_tokens())  # ya estan
        self.assertEqual(config.clave_api(config.CLAVE_PRMTE), "AAA")

    def test_tks_json_manda(self):
        config.RUTA_TOKENS.parent.mkdir(parents=True)
        config.RUTA_TOKENS.write_text(json.dumps(
            {"coordinador": {"prmte": "NUEVO"}}), encoding="utf-8")
        self.assertEqual(config.clave_api(config.CLAVE_PRMTE), "NUEVO")

    def test_carpeta_del_mes_como_abby(self):
        original = config.CARPETA_CONFIG
        config.CARPETA_CONFIG = config.RUTA_CONFIG.parent
        try:
            self.assertEqual(
                config.carpeta_del_mes("2609"),
                config.CARPETA_CONFIG / "2026" / "09 Septiembre" / "Balance BESS",
            )
            (config.CARPETA_CONFIG / "2026" / "9 septiembre").mkdir(parents=True)
            self.assertEqual(
                config.carpeta_del_mes("2609"),
                config.CARPETA_CONFIG / "2026" / "9 septiembre" / "Balance BESS",
            )
        finally:
            config.CARPETA_CONFIG = original


if __name__ == "__main__":
    unittest.main()
