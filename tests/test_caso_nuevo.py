# -*- coding: utf-8 -*-
"""
Crear el caso de un periodo nuevo: la carpeta y sus subcarpetas.

Pedido del usuario: "cuando yo elija un mes que no exista, me permita
elegir y crear la carpeta con las carpetas dentro".
"""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import Script.nucleo as nucleo  # noqa: E402


class TestCrearEstructura(unittest.TestCase):

    def test_una_carpeta_que_no_existe_nace_con_todo_adentro(self):
        with tempfile.TemporaryDirectory() as padre:

            base = Path(padre) / "Balance BESS 2608"

            creadas = nucleo.crear_estructura_caso(
                base, registrar=lambda *a: None
            )

            self.assertTrue(base.is_dir())

            for sub in nucleo.SUBCARPETAS_CASO:
                self.assertTrue(
                    (base / sub).is_dir(), f"falta {sub}"
                )

            # la base + una por subcarpeta
            self.assertEqual(
                len(creadas), len(nucleo.SUBCARPETAS_CASO) + 1
            )

    def test_correrlo_dos_veces_no_hace_nada_la_segunda(self):
        with tempfile.TemporaryDirectory() as padre:

            base = Path(padre) / "2608"
            nucleo.crear_estructura_caso(base, registrar=lambda *a: None)

            self.assertEqual(
                nucleo.crear_estructura_caso(base, registrar=lambda *a: None),
                [],
            )

    def test_completa_solo_lo_que_falta_y_no_toca_lo_que_hay(self):
        with tempfile.TemporaryDirectory() as padre:

            base = Path(padre) / "caso a medias"
            (base / "Medidas").mkdir(parents=True)
            archivo = base / "Medidas" / "SOC_2608.xlsx"
            archivo.write_text("no me toques")

            creadas = nucleo.crear_estructura_caso(
                base, registrar=lambda *a: None
            )

            self.assertNotIn(base / "Medidas", creadas)
            self.assertEqual(archivo.read_text(), "no me toques")
            self.assertTrue((base / "Ofertas").is_dir())

    def test_un_caso_viejo_conserva_su_carpeta_de_fd_con_el_nombre_anterior(self):
        """
        Si el caso todavia usa 'SSCC_Desempeño' (el nombre viejo de la
        carpeta), no se le planta una 'FD y FMA' vacia al lado: seria
        una carpeta de mas y el programa igual usa la que ya esta.
        """

        with tempfile.TemporaryDirectory() as padre:

            base = Path(padre) / "caso viejo"
            (base / nucleo.CARPETA_FD_FMA_ANTIGUA).mkdir(parents=True)

            nucleo.crear_estructura_caso(base, registrar=lambda *a: None)

            self.assertFalse((base / nucleo.CARPETA_FD_FMA).exists())
            self.assertTrue(
                (base / nucleo.CARPETA_FD_FMA_ANTIGUA).is_dir()
            )

    def test_si_el_nombre_ya_es_un_archivo_lo_dice(self):
        with tempfile.TemporaryDirectory() as padre:

            ocupado = Path(padre) / "2608"
            ocupado.write_text("soy un archivo")

            with self.assertRaises(nucleo.ErrorEntrada):
                nucleo.crear_estructura_caso(
                    ocupado, registrar=lambda *a: None
                )

    def test_el_caso_creado_es_el_que_espera_el_resto_del_programa(self):
        """
        Las carpetas que se crean son exactamente las que
        revisar_estructura() busca: recien creado, no puede faltar
        ninguna CARPETA (los archivos si faltan, claro).
        """

        with tempfile.TemporaryDirectory() as padre:

            base = Path(padre) / "2608"
            nucleo.crear_estructura_caso(base, registrar=lambda *a: None)

            _, filas = nucleo.revisar_estructura(base, "2608")

            carpetas_en_falta = [
                fila["id"] for fila in filas
                if fila["estado"] == "falta" and fila["id"].endswith("_dir")
            ]

            self.assertEqual(carpetas_en_falta, [])


class TestNombreSugerido(unittest.TestCase):

    def test_cambia_el_aamm_del_nombre_y_conserva_el_resto(self):
        self.assertEqual(
            nucleo.nombre_caso_sugerido("2608", "/casos/Balance BESS 2607"),
            "Balance BESS 2608",
        )

    def test_sin_aamm_en_el_nombre_propone_el_aamm_pelado(self):
        self.assertEqual(
            nucleo.nombre_caso_sugerido("2608", "/casos/caso de julio"),
            "2608",
        )
        self.assertEqual(nucleo.nombre_caso_sugerido("2608"), "2608")

    def test_no_se_come_un_numero_que_no_es_un_aamm(self):
        self.assertEqual(
            nucleo.nombre_caso_sugerido("2608", "/casos/Caso 123456"),
            "2608",
        )

    def test_sin_periodo_valido_avisa(self):
        with self.assertRaises(nucleo.ErrorEntrada):
            nucleo.nombre_caso_sugerido("julio", "/casos/2607")


class TestCarpetaDelPeriodo(unittest.TestCase):
    """
    De que periodo es la carpeta que esta abierta.

    Pedido del usuario: *"cuando cambio el mes en la casilla se
    mantiene en la carpeta; dos meses no pueden tener la misma
    carpeta; si cambio el año y el mes no tiene carpeta debe pedirme
    examinar y seleccionar o crear"*.
    """

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.padre = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def carpeta(self, nombre):
        ruta = self.padre / nombre
        ruta.mkdir()
        return ruta

    def test_el_aamm_en_el_nombre_dice_de_que_periodo_es(self):
        julio = self.carpeta("Balance BESS 2607")

        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(julio, "2607"),
            nucleo.CARPETA_DEL_PERIODO,
        )
        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(julio, "2608"),
            nucleo.CARPETA_DE_OTRO_PERIODO,
        )

    def test_cambiar_de_año_tambien_cambia_de_carpeta(self):
        diciembre = self.carpeta("Caso 2612")

        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(diciembre, "2701"),
            nucleo.CARPETA_DE_OTRO_PERIODO,
        )

    def test_sin_aamm_en_el_nombre_y_sin_registro_no_se_sabe(self):
        """No es "sirve": es "hay que preguntar"."""

        cualquiera = self.carpeta("Caso de julio")

        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(cualquiera, "2607"),
            nucleo.CARPETA_SIN_PERIODO,
        )

    def test_una_carpeta_anotada_para_un_periodo_no_sirve_para_otro(self):
        """La regla del usuario: dos meses no comparten carpeta."""

        cualquiera = self.carpeta("Caso de julio")
        recordadas = {"2607": str(cualquiera)}

        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(
                cualquiera, "2607", recordadas
            ),
            nucleo.CARPETA_DEL_PERIODO,
        )
        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(
                cualquiera, "2608", recordadas
            ),
            nucleo.CARPETA_DE_OTRO_PERIODO,
        )

    def test_lo_anotado_le_gana_al_nombre(self):
        """
        Una carpeta que se llama "...2607" pero que el usuario dijo
        que es la de 2608 es la de 2608: lo que el eligio manda.
        """

        rara = self.carpeta("Balance BESS 2607")

        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(
                rara, "2608", {"2608": str(rara)}
            ),
            nucleo.CARPETA_DEL_PERIODO,
        )

    def test_una_carpeta_que_no_existe_no_es_la_del_periodo(self):
        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo(
                self.padre / "no existe", "2607"
            ),
            nucleo.CARPETA_DE_OTRO_PERIODO,
        )
        self.assertEqual(
            nucleo.carpeta_corresponde_al_periodo("", "2607"),
            nucleo.CARPETA_DE_OTRO_PERIODO,
        )

    def test_los_aamm_del_nombre(self):
        self.assertEqual(
            nucleo.periodos_en_el_nombre("Balance BESS 2607"), ["2607"]
        )
        self.assertEqual(nucleo.periodos_en_el_nombre("Caso 123456"), [])
        self.assertEqual(nucleo.periodos_en_el_nombre("2607 y 2608"),
                         ["2607", "2608"])


if __name__ == "__main__":
    unittest.main()
