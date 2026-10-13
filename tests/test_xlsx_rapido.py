# -*- coding: utf-8 -*-
"""
Pegar hojas en un .xlsx a nivel zip (Script/nucleo/xlsx_rapido.py).

Pedido del usuario: "la hoja Centrales se demora mucho en armar" -- con
el libro lleno, actualizar una hoja tardaba minutos porque se abria y se
volvia a guardar el libro entero. Ahora solo se escriben las hojas
nuevas y se pegan en el zip; lo que se controla aca es que el libro que
queda sea el mismo que antes, salvo las hojas pegadas.
"""

import re
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font

from Script import nucleo
from Script.nucleo.xlsx_rapido import _hojas, nombres_de_hojas, pegar_hojas


def libro(ruta, hojas, negrita=()):
    """Un .xlsx con {nombre: DataFrame}; encabezado en negrita en 'negrita'."""

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre, index=False)
            if nombre in negrita:
                for celda in writer.sheets[nombre][1]:
                    celda.font = Font(bold=True)
                writer.sheets[nombre]["B2"].number_format = "#,##0.00"


def partes(ruta):
    with zipfile.ZipFile(ruta) as z:
        return {h["name"]: z.read(h["parte"]) for h in _hojas(z)}


class TestPegarHojas(unittest.TestCase):

    def setUp(self):
        self.carpeta = tempfile.TemporaryDirectory()
        self.dir = Path(self.carpeta.name)
        self.destino = self.dir / "libro.xlsx"
        libro(self.destino, {
            "Uno": pd.DataFrame({"a": [1, 2], "b": ["x", "y"]}),
            "Dos": pd.DataFrame({"c": [3.5, 4.5]}),
            "Tres": pd.DataFrame({"d": ["ñandú", "z"]}),
        }, negrita=("Dos",))

    def tearDown(self):
        self.carpeta.cleanup()

    def test_reemplaza_una_hoja_y_deja_las_otras_identicas(self):
        antes = partes(self.destino)

        nuevas = self.dir / "nuevas.xlsx"
        libro(nuevas, {"Dos": pd.DataFrame({
            "c": [9.0], "texto": ['<a & "b">'],
        })})
        pegar_hojas(self.destino, nuevas)

        despues = partes(self.destino)
        self.assertEqual(antes["Uno"], despues["Uno"])
        self.assertEqual(antes["Tres"], despues["Tres"])
        self.assertNotEqual(antes["Dos"], despues["Dos"])
        self.assertIsNone(zipfile.ZipFile(self.destino).testzip())

        leido = pd.read_excel(self.destino, sheet_name=None)
        self.assertEqual(list(leido), ["Uno", "Dos", "Tres"])
        self.assertEqual(leido["Dos"]["texto"].tolist(), ['<a & "b">'])
        self.assertEqual(leido["Tres"]["d"].tolist(), ["ñandú", "z"])

    def test_agrega_hojas_y_las_ordena(self):
        nuevas = self.dir / "nuevas.xlsx"
        libro(nuevas, {
            "Cero": pd.DataFrame({"e": ["nueva"]}),
            "Cuatro": pd.DataFrame({"f": [4]}),
        })
        pegar_hojas(
            self.destino, nuevas, orden=("Cero", "Uno", "Dos", "Tres", "Cuatro")
        )

        self.assertEqual(
            nombres_de_hojas(self.destino),
            ["Cero", "Uno", "Dos", "Tres", "Cuatro"],
        )
        leido = pd.read_excel(self.destino, sheet_name=None)
        self.assertEqual(leido["Cero"]["e"].tolist(), ["nueva"])
        self.assertEqual(leido["Cuatro"]["f"].tolist(), [4])
        self.assertEqual(leido["Uno"]["b"].tolist(), ["x", "y"])

    def test_los_estilos_de_la_hoja_pegada_se_conservan(self):
        nuevas = self.dir / "nuevas.xlsx"
        libro(nuevas, {"Nueva": pd.DataFrame({"g": [1.5], "h": [2.25]})},
              negrita=("Nueva",))
        pegar_hojas(self.destino, nuevas)

        wb = openpyxl.load_workbook(self.destino)
        self.assertTrue(wb["Nueva"]["A1"].font.b)
        self.assertEqual(wb["Nueva"]["B2"].number_format, "#,##0.00")
        # y los de una hoja que no se toco siguen igual
        self.assertTrue(wb["Dos"]["A1"].font.b)
        self.assertFalse(wb["Uno"]["A1"].font.b)

    def test_un_indice_escrito_como_lo_escribe_excel(self):
        """
        Excel declara el prefijo de relaciones UNA vez en la raiz de
        workbook.xml (openpyxl lo repite en cada <sheet>). Las dos formas
        tienen que servir.
        """

        rel = (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships"
        )
        estilo_excel = self.dir / "excel.xlsx"

        with zipfile.ZipFile(self.destino) as origen, \
                zipfile.ZipFile(estilo_excel, "w", zipfile.ZIP_DEFLATED) as fin:
            for info in origen.infolist():
                datos = origen.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    texto = datos.decode("utf-8")
                    texto = texto.replace(f' xmlns:r="{rel}"', "")
                    texto = re.sub(
                        r"<workbook ", f'<workbook xmlns:r="{rel}" ', texto, 1
                    )
                    datos = texto.encode("utf-8")
                fin.writestr(info, datos)

        nuevas = self.dir / "nuevas.xlsx"
        libro(nuevas, {
            "Dos": pd.DataFrame({"c": [7]}),
            "Cinco": pd.DataFrame({"i": [5]}),
        })
        pegar_hojas(estilo_excel, nuevas)

        leido = pd.read_excel(estilo_excel, sheet_name=None)
        self.assertEqual(list(leido), ["Uno", "Dos", "Tres", "Cinco"])
        self.assertEqual(leido["Dos"]["c"].tolist(), [7])
        openpyxl.load_workbook(estilo_excel)  # que lo pueda abrir entero


class TestEscrituraDeLaPlanilla(unittest.TestCase):

    def test_actualizar_una_hoja_no_reescribe_las_demas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_SALIDA
            nucleo.escribir_salida(
                pd.DataFrame({"Mes": [7, 7]}), ruta, [], [],
                df_cmg=pd.DataFrame({"cmg": [1.0]}),
                hojas_regenerar={nucleo.HOJA_MEDIDORES, nucleo.HOJA_CMG},
                registrar=lambda *a: None,
            )
            antes = partes(ruta)

            nucleo.escribir_salida(
                None, ruta, [], [],
                hojas_regenerar={nucleo.HOJA_CENTRALES},
                centrales=(
                    pd.DataFrame({"Nombre activo": ["SAE UNO"]}),
                    pd.DataFrame([["Balance_BESS"], ["SAE UNO"]]),
                ),
                registrar=lambda *a: None,
            )
            despues = partes(ruta)

            self.assertEqual(
                antes[nucleo.HOJA_MEDIDORES], despues[nucleo.HOJA_MEDIDORES]
            )
            self.assertEqual(antes[nucleo.HOJA_CMG], despues[nucleo.HOJA_CMG])
            self.assertIn(nucleo.HOJA_CENTRALES, despues)
            # sin restos del libro temporal
            self.assertEqual(
                sorted(p.name for p in Path(carpeta).iterdir()),
                [nucleo.ARCHIVO_SALIDA],
            )


if __name__ == "__main__":
    unittest.main()
