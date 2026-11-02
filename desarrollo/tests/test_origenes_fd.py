"""
De donde sale el FD y como se muestra el origen en la ventana:

  - la raiz del arbol de indicadores del DCO es la unidad F: y el FD
    sale de '03 Desempeño para publicar' y de ninguna otra carpeta: ni
    de la de antes ('04 Desempeño para transferencias') ni de otras
    ramas del arbol de la version;
  - entre V1 y V2 se usa SIEMPRE la mas alta que tenga el archivo;
  - el link de la ventana apunta a la carpeta que DE VERDAD se uso
    (V1 o V2), no a la version mas alta que exista;
  - cada fila del diagrama viaja con su ruta (para que el nombre sea
    un link) y, si lo que hay en ella viene de afuera del caso, con el
    id de su origen.
"""

import tempfile
import unittest
from pathlib import Path

from Script import nucleo
from Script.Fd import Indicadores_DCO as dco
from Script.Fd import Indices_FMA as indices_fma


NOMBRE_FD = "SSCC_Disponibilidad_CSF_Agosto_2026_V1.zip"


def arbol_dco(raiz, versiones, subcarpeta="03 Desempeño para publicar"):
    """
    <raiz>/2026/08. Agosto/Indicadores Publicar/<Vn>/<subcarpeta>/, con
    el archivo de FD adentro de las versiones que se pidan.
    """

    publicacion = (
        Path(raiz) / "2026" / "08. Agosto" / "Indicadores Publicar"
    )

    for version, con_archivo in versiones.items():
        carpeta = publicacion / version / subcarpeta
        carpeta.mkdir(parents=True, exist_ok=True)
        if con_archivo:
            (carpeta / NOMBRE_FD).write_bytes(b"")

    return publicacion


class RutaDelFdTest(unittest.TestCase):

    def test_la_raiz_es_la_unidad_f(self):
        self.assertEqual(
            dco.RAIZ_DCO_INDICADORES,
            r"F:\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores",
        )

    def test_la_carpeta_del_fd_es_desempeno_para_publicar(self):
        self.assertEqual(dco.SUBCARPETAS_FD, ("03 Desempeño para publicar",))
        self.assertEqual(
            dco.SUBCARPETAS_FD_ANTIGUA, ("04 Desempeño para transferencias",)
        )

    def test_lo_encuentra_en_la_carpeta_nueva(self):
        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V1": True})
            encontrados = dco.buscar_archivos_fd(publicacion / "V1", 2026)
            self.assertEqual([r.name for r in encontrados], [NOMBRE_FD])
            self.assertEqual(
                encontrados[0].parent.name, "03 Desempeño para publicar"
            )

    def test_no_busca_en_la_carpeta_vieja(self):
        """
        Hasta hace poco el FD vivia en '04 Desempeño para
        transferencias'. Ya no se busca ahi: la regla es esa carpeta y
        ninguna otra.
        """

        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(
                tmp, {"V1": True},
                subcarpeta="04 Desempeño para transferencias",
            )
            self.assertEqual(dco.buscar_archivos_fd(publicacion / "V1", 2026), [])

    def test_no_se_trae_nada_de_otras_ramas_del_arbol(self):
        """
        El problema que reporto el usuario: "cuando corro Traer FD me
        trae muchos que vienen de otras rutas". Pasaba porque, si no
        encontraba, barria recursivamente TODO el arbol de la version.
        """

        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V1": False})
            version = publicacion / "V1"

            # Archivos con el mismo nombre colgando de otras ramas y de
            # una subcarpeta de la propia carpeta del FD.
            for rama in (
                version / "01 Respuesta" / "01 Indices CPF",
                version / "02 Otra cosa",
                version / "03 Desempeño para publicar" / "adjuntos",
            ):
                rama.mkdir(parents=True, exist_ok=True)
                (rama / NOMBRE_FD).write_bytes(b"")

            self.assertEqual(dco.buscar_archivos_fd(version, 2026), [])

    def test_solo_los_archivos_sueltos_de_esa_carpeta(self):
        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V2": True})
            carpeta = publicacion / "V2" / "03 Desempeño para publicar"
            (carpeta / "adjuntos").mkdir()
            (carpeta / "adjuntos" / "SSCC_Desempeño_otro_2026.xlsx").write_bytes(b"")

            encontrados = dco.buscar_archivos_fd(publicacion / "V2", 2026)

            self.assertEqual([r.name for r in encontrados], [NOMBRE_FD])

    def test_si_esta_v2_no_se_usa_v1(self):
        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V1": True, "V2": True})
            carpeta, archivos, _ = dco.buscar_en_versiones(
                publicacion,
                lambda carpeta: dco.buscar_archivos_fd(carpeta, 2026),
                registrar=lambda *_: None,
            )
            self.assertEqual(carpeta.name, "V2")
            self.assertTrue(
                str(archivos[0]).find("V2") > 0, archivos[0]
            )

    def test_si_v2_no_tiene_el_archivo_se_cae_a_v1(self):
        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V1": True, "V2": False})
            carpeta, archivos, revisadas = dco.buscar_en_versiones(
                publicacion,
                lambda carpeta: dco.buscar_archivos_fd(carpeta, 2026),
                registrar=lambda *_: None,
            )
            self.assertEqual(carpeta.name, "V1")
            self.assertEqual(revisadas, ["V2"])
            self.assertEqual(len(archivos), 1)


class RutaDeOrigenTest(unittest.TestCase):
    """
    La ruta que la ventana abre al hacer click en "Origen: DCO". Es la
    del ejemplo que dio el usuario: .../2026/08. Agosto/Indicadores
    Publicar/<Vn>/03 Desempeño para publicar.
    """

    def test_resuelve_la_version_mas_alta_publicada(self):
        with tempfile.TemporaryDirectory() as tmp:
            arbol_dco(tmp, {"V1": True, "V2": True})
            ruta = dco.ruta_origen_fd("2608", raiz=tmp)
            self.assertEqual(ruta.name, "03 Desempeño para publicar")
            self.assertEqual(ruta.parent.name, "V2")

    def test_si_v2_no_tiene_el_archivo_el_link_apunta_a_v1(self):
        """
        Lo que pidio el usuario: el link de la ventana tiene que
        apuntar a la ruta FINAL, la que se ocupo, sea V1 o V2.
        """

        with tempfile.TemporaryDirectory() as tmp:
            arbol_dco(tmp, {"V1": True, "V2": False})
            ruta = dco.ruta_origen_fd("2608", raiz=tmp)
            self.assertEqual(ruta.parent.name, "V1")
            self.assertEqual(ruta.name, "03 Desempeño para publicar")

    def test_el_link_es_la_misma_carpeta_de_la_que_se_copia(self):
        with tempfile.TemporaryDirectory() as tmp:
            publicacion = arbol_dco(tmp, {"V1": True, "V2": True})

            carpeta_version, archivos, _ = dco.buscar_en_versiones(
                publicacion,
                lambda carpeta: dco.buscar_archivos_fd(carpeta, 2026),
                registrar=lambda *_: None,
            )

            self.assertEqual(
                dco.ruta_origen_fd("2608", raiz=tmp), archivos[0].parent
            )
            self.assertEqual(archivos[0].parent.parent, carpeta_version)

    def test_sin_servidor_arma_la_ruta_igual(self):
        """
        Con la unidad desconectada no hay nada que resolver, pero la
        ventana tiene que poder mostrar (y abrir) la ruta del periodo.
        """

        ruta = dco.ruta_origen_fd("2608", raiz=r"F:\no-existe")
        self.assertEqual(
            Path(*ruta.parts[-5:]),
            Path("2026") / "08. Agosto" / "Indicadores Publicar" / "V1"
            / "03 Desempeño para publicar",
        )

    def test_sin_periodo_devuelve_la_raiz(self):
        self.assertEqual(
            str(dco.ruta_origen_fd("", raiz=r"F:\no-existe")),
            r"F:\no-existe",
        )

    def test_los_insumos_del_fma_cuelgan_del_mismo_arbol(self):
        with tempfile.TemporaryDirectory() as tmp:
            arbol_dco(tmp, {"V1": True})
            self.assertEqual(
                indices_fma.ruta_origen_cpf("2608", raiz=tmp).parts[-2:],
                ("01 Respuesta", "01 Indices CPF"),
            )
            self.assertEqual(
                indices_fma.ruta_origen_ctf("2608", raiz=tmp).parts[-2:],
                ("01 Respuesta", "06 Indices CTF"),
            )

        # El CSF no sale del DCO: son los reportes del AGC.
        self.assertEqual(
            str(indices_fma.ruta_origen_csf("2608")),
            indices_fma.RAIZ_AGC_FACE,
        )


class OrigenesDeLaVentanaTest(unittest.TestCase):

    def test_cada_origen_tiene_titulo_etiqueta_y_ruta(self):
        for id_origen in nucleo.ORIGENES:
            with self.subTest(id_origen=id_origen):
                datos = nucleo.origen(id_origen)
                self.assertEqual(datos["id"], id_origen)
                self.assertTrue(datos["etiqueta"])
                self.assertIn(datos["titulo"], ("Origen", "Origen inputs"))
                self.assertTrue(str(nucleo.ruta_origen(id_origen, "2608")))

    def test_el_fd_dice_origen_dco(self):
        self.assertEqual(
            nucleo.origen("fd"), {"id": "fd", "titulo": "Origen", "etiqueta": "DCO"}
        )

    def test_el_fma_dice_origen_inputs(self):
        for tipo in ("cpf", "csf", "ctf"):
            self.assertEqual(
                nucleo.origen(f"fma_{tipo}")["titulo"], "Origen inputs"
            )

    def test_un_id_desconocido_no_revienta(self):
        self.assertIsNone(nucleo.origen("lo-que-sea"))
        self.assertIsNone(nucleo.ruta_origen("lo-que-sea", "2608"))


class FilasConRutaTest(unittest.TestCase):
    """
    Cada archivo y cada carpeta del diagrama viaja con su ruta, que es
    lo que la ventana convierte en link.
    """

    def filas(self, base):
        _, filas = nucleo.revisar_estructura(base, "2608")
        return {fila["id"]: fila for fila in filas}

    def test_las_filas_de_archivos_y_carpetas_traen_su_ruta(self):
        with tempfile.TemporaryDirectory() as tmp:
            filas = self.filas(tmp)

            for id_fila in (
                "base", "medidas_dir", "medidas_sae", "auxiliares_dir",
                "centrales", "ofertas_dir", "cmg_dir", "cmg_csv",
                "cmg_xlsx", "sscc_dir", "sscc", "subastas_dir",
                "db_subastas", "prorrata_dir", "salida", "control",
            ):
                with self.subTest(id_fila=id_fila):
                    self.assertTrue(filas[id_fila]["ruta"], id_fila)

            # Las hojas de la salida no son archivos: no llevan link.
            self.assertEqual(filas["consolidado:medidores"]["ruta"], "")

    def test_las_carpetas_se_marcan_como_carpetas(self):
        with tempfile.TemporaryDirectory() as tmp:
            filas = self.filas(tmp)
            self.assertTrue(filas["medidas_dir"]["es_carpeta"])
            self.assertTrue(filas["db_subastas"]["es_carpeta"])
            self.assertFalse(filas["centrales"]["es_carpeta"])
            self.assertFalse(filas["salida"]["es_carpeta"])

    def test_las_filas_que_traen_algo_dicen_de_donde(self):
        with tempfile.TemporaryDirectory() as tmp:
            filas = self.filas(tmp)

            self.assertEqual(filas["sscc"]["origen"]["etiqueta"], "DCO")
            self.assertEqual(filas["sscc"]["origen"]["titulo"], "Origen")
            self.assertEqual(filas["cmg_csv"]["origen"]["id"], "cmg_csv")
            self.assertEqual(filas["db_subastas"]["origen"]["id"], "subastas")

            for tipo in ("cpf", "csf", "ctf"):
                self.assertEqual(
                    filas[f"fma_{tipo}"]["origen"]["titulo"], "Origen inputs"
                )

            # Lo que se deja a mano en el caso no tiene origen externo.
            self.assertIsNone(filas["centrales"]["origen"])


if __name__ == "__main__":
    unittest.main()
