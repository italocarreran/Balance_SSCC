"""Las dos claves de las APIs del Coordinador salen de config.json.

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
        self.assertIn(str(self.ruta), mensaje)
        self.assertIn('"claves_api"', mensaje)
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

        self.assertIn("claves_api", str(caja.exception))

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


if __name__ == "__main__":
    unittest.main()
