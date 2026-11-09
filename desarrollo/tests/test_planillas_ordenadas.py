# -*- coding: utf-8 -*-
"""
Las planillas ordenadas: que hojas salen, en que orden, con que
columnas y con que formato -- y que nada de eso toca un numero.

Pedido del usuario: "sin modificar el calculo, necesito que ordenes
las planillas, quita los auxiliares innecesarios".
"""

import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import Script.nucleo as nucleo  # noqa: E402


class TestMedidoresSinAuxiliares(unittest.TestCase):

    def armar_sae(self, filas=4):
        return pd.DataFrame({
            "Mes": [7] * filas,
            "Dia": [1] * filas,
            "Hora": list(range(1, filas + 1)),
            "Minutos": [0] * filas,
            "Hora Mes": list(range(1, filas + 1)),
            "Cuarto de Hora": list(range(1, filas + 1)),
            "clave": ["SAE UNO"] * filas,
            "intervalo": pd.to_datetime([
                f"2026-07-01 0{h}:00:00" for h in range(1, filas + 1)
            ]),
            "Gen_Unidad": [5.0, -3.0, 2.0, -1.0][:filas],
        })

    def armar_soc(self, filas=4):
        return pd.DataFrame({
            "central": ["SAE UNO"] * filas,
            "nombre_scada_original": ["SAE UNO"] * filas,
            "timestamp": pd.to_datetime([
                f"2026-07-01 0{h}:00:00" for h in range(1, filas + 1)
            ]),
            "soc": [0.5] * filas,
        })

    def test_la_hoja_sale_sin_vacias_sin_clave_auxiliar_y_sin_la_copia(self):
        df, _ = nucleo.construir_medidores(
            self.armar_sae(), self.armar_soc(), 7, registrar=lambda *a: None
        )

        self.assertEqual(
            list(df.columns), nucleo.COLUMNAS_MEDIDORES_SALIDA
        )

        for auxiliar in nucleo.COLUMNAS_AUXILIARES_MEDIDORES:
            self.assertNotIn(auxiliar, df.columns)

    def test_las_columnas_que_quedan_son_las_de_siempre_en_el_mismo_orden(self):
        """
        Quitar auxiliares no puede reordenar lo que queda: el orden
        sigue siendo el de LETRA_A_CAMPO (el de la planilla original).
        """

        quedan = [
            campo for campo in nucleo.LETRA_A_CAMPO.values()
            if campo not in nucleo.COLUMNAS_AUXILIARES_MEDIDORES
        ]

        self.assertEqual(nucleo.COLUMNAS_MEDIDORES_SALIDA, quedan)

    def test_la_copia_de_la_ventana_se_repone_al_leer(self):
        df, _ = nucleo.construir_medidores(
            self.armar_sae(), self.armar_soc(), 7, registrar=lambda *a: None
        )

        repuesto = nucleo.reponer_auxiliares_medidores(df)

        self.assertIn("Copia_Ventana", repuesto.columns)
        self.assertTrue(
            (repuesto["Copia_Ventana"] == repuesto["Ventana"]).all()
        )

    def test_un_libro_viejo_que_todavia_la_trae_se_respeta(self):
        """
        Consolidado_entradas.xlsx de una corrida anterior: la columna
        viene escrita y no hay que pisarla.
        """

        df = pd.DataFrame({
            "Ventana": [1, 1, 2],
            "Copia_Ventana": [9, 9, 9],
        })

        repuesto = nucleo.reponer_auxiliares_medidores(df)

        self.assertEqual(list(repuesto["Copia_Ventana"]), [9, 9, 9])


class TestColumnasDeSalidaDeLosCalculos(unittest.TestCase):

    def test_e_costos_sale_sin_el_ciclo_duplicado(self):
        self.assertNotIn("X", nucleo.COLUMNAS_SALIDA_E_COSTOS)

        # ...pero el dato no se pierde: sigue en la hoja una sola vez,
        # como "Ciclo de Carga del mes" (Copia_Ventana).
        nombres = [
            nucleo.NOMBRES_CALCULO_E_COSTOS[clave]
            for clave in nucleo.COLUMNAS_SALIDA_E_COSTOS
        ]
        self.assertEqual(nombres.count("Ciclo de Carga del mes"), 1)
        self.assertNotIn("Ciclo", nombres)

    def test_re545_sale_sin_la_columna_sin_nombre_ni_la_ventana_repetida(self):
        self.assertNotIn("BL", nucleo.COLUMNAS_SALIDA_RE545)
        self.assertNotIn("BR", nucleo.COLUMNAS_SALIDA_RE545)

        nombres = [
            nucleo.NOMBRES_CALCULO_RE545[clave]
            for clave in nucleo.COLUMNAS_SALIDA_RE545
        ]
        self.assertNotIn("", nombres)
        self.assertEqual(
            sum(n.lower() == "ventana de valorizacion" for n in nombres), 1
        )

    def test_los_intermedios_que_muestran_el_calculo_se_quedan(self):
        """
        El usuario pidio quitar lo trivial, no lo que deja ver de
        donde sale cada peso: las curvas monotonas, el ranking de CMg
        y las energias con FD siguen en la hoja.
        """

        for clave in ("W", "Y", "AB", "AC", "AD", "R", "AS", "AT"):
            self.assertIn(clave, nucleo.COLUMNAS_SALIDA_E_COSTOS)

        # Salen (pedido 2026-11-03): las CTF (siempre 0) y las del metodo
        # anterior que no llegan al monto.
        for clave in ("AI", "AL", "AO", "AR", "M", "N", "O", "AE", "AF"):
            self.assertNotIn(clave, nucleo.COLUMNAS_SALIDA_E_COSTOS)

        for clave in ("S", "U", "V", "BM", "BN", "BS", "BT"):
            self.assertIn(clave, nucleo.COLUMNAS_SALIDA_RE545)

        # Salen (pedido 2026-11-03): CTF, las que no llegan al monto.
        for clave in ("AE", "AH", "AK", "AN", "AQ", "AT", "M", "N", "O",
                      "BK", "BY"):
            self.assertNotIn(clave, nucleo.COLUMNAS_SALIDA_RE545)

    def test_los_grupos_siguen_cayendo_sobre_columnas_de_la_hoja(self):
        """
        Los encabezados de grupo se ubican con .index() sobre el orden
        de salida: si un grupo nombra una columna que ya no se
        escribe, escribir el libro revienta.
        """

        for grupos, columnas in (
            (nucleo.GRUPOS_CALCULO_E_COSTOS, nucleo.COLUMNAS_SALIDA_E_COSTOS),
            (nucleo.GRUPOS_CALCULO_RE545, nucleo.COLUMNAS_SALIDA_RE545),
        ):
            for etiqueta, claves in grupos:
                posiciones = [columnas.index(clave) for clave in claves]
                self.assertEqual(
                    posiciones,
                    list(range(min(posiciones), max(posiciones) + 1)),
                    f"el grupo '{etiqueta}' quedo partido en dos",
                )


class TestLibroDePagosOrdenado(unittest.TestCase):

    def armar_calculo(self):
        filas = 3
        base = {
            clave: list(range(1, filas + 1))
            for clave in nucleo.COLUMNAS_SALIDA_E_COSTOS
        }
        base["clave"] = ["SAE UNO"] * filas
        base["Barra"] = ["BARRA A"] * filas
        base["Energia_Positiva"] = [1.5, 2.5, 3.5]
        df = pd.DataFrame(base)
        return (
            df[nucleo.COLUMNAS_SALIDA_E_COSTOS]
            .rename(columns=nucleo.NOMBRES_CALCULO_E_COSTOS)
        )

    def escribir(self, carpeta):
        ruta = Path(carpeta) / nucleo.nombre_salida("2607")
        nucleo.escribir_pagos_bess(
            ruta,
            df_ecostos=self.armar_calculo(),
            df_resumen=pd.DataFrame({
                "NOMBRE": ["EMPRESA UNO"],
                "RECIBE": [1000.5],
                "PAGA": [0.0],
                "NETO": [1000.5],
            }),
            registrar=lambda *a: None,
        )
        return ruta

    def test_puesto_sin_fila_en_la_curva_monotona_lleva_guion(self):
        df = self.armar_calculo()
        df.loc[1:, "Curva monotona CMg Descarga"] = pd.NA
        df.loc[2, "Curva monotona CMg Carga"] = pd.NA
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            nucleo.escribir_pagos_bess(
                ruta, df_ecostos=df, registrar=lambda *a: None
            )
            # data_only: la fila 4 lleva la formula de prueba, con el
            # valor del programa guardado.
            ws = openpyxl.load_workbook(ruta, data_only=True)[
                nucleo.HOJA_CALCULO_ECOSTOS
            ]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        descarga = nombres["Curva monotona CMg Descarga [CLP/kWh]"]
        carga = nombres["Curva monotona CMg Carga [CLP/kWh]"]
        self.assertEqual(
            [ws[f"{descarga}{f}"].value for f in (4, 5, 6)], [1, "–", "–"]
        )
        self.assertEqual([ws[f"{carga}{f}"].value for f in (4, 5, 6)], [1, 2, "–"])

    def _escribir_re545(self, carpeta, ctf_en_cero=True):
        filas = 3
        base = {c: [1.0, 2.0, 3.0] for c in nucleo.COLUMNAS_SALIDA_RE545}
        base["clave"] = ["SAE UNO"] * filas
        base["Barra"] = ["BARRA A"] * filas
        df = (
            pd.DataFrame(base)[nucleo.COLUMNAS_SALIDA_RE545]
            .rename(columns=nucleo.NOMBRES_CALCULO_RE545)
        )
        df.attrs["ctf_en_cero"] = ctf_en_cero
        resumen = pd.DataFrame({
            nombre: [1.0] for nombre in nucleo.NOMBRES_RESUMEN_RE545.values()
        })
        ruta = Path(carpeta) / nucleo.nombre_salida("2607")
        nucleo.escribir_pagos_bess(
            ruta, df_re545=df, df_resumen_re545=resumen,
            registrar=lambda *a: None,
        )
        return openpyxl.load_workbook(ruta)[nucleo.HOJA_CALCULO_RE545]

    def test_re545_en_cinco_cuadros_con_formulas_y_totales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ws = self._escribir_re545(carpeta)
        titulos = [c.value for c in ws[2] if c.value]
        self.assertEqual(len(titulos), 5, titulos)
        self.assertTrue(titulos[4].endswith("Resumen por central y ventana"))
        self.assertEqual(ws.freeze_panes, "I4")
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        self.assertIn("Central", nombres)  # antes sin nombre
        self.assertIn("Reserva CPF(-) [MW]", nombres)
        self.assertNotIn("CTF(-)", nombres)
        reserva = nombres["Reserva con FD y FMA [kWh]"]
        self.assertTrue(str(ws[f"{reserva}4"].value).startswith("=("))
        c1 = nombres["Total C1 [CLP]"]
        self.assertEqual(
            ws[f"{c1}4"].value,
            f'=IF(COUNT({nombres["Energía asignada [kWh]"]}4,'
            f'{nombres["Curva Cmg Decendente [CLP/kWh]"]}4)<2,"",'
            f'{nombres["Energía asignada [kWh]"]}4*'
            f'{nombres["Curva Cmg Decendente [CLP/kWh]"]}4)',
        )
        monto = nombres["Monto a compensar [CLP]"]
        self.assertEqual(ws[f"{monto}1"].value, f"=SUM({monto}4:{monto}6)")
        self.assertEqual(ws[f"{monto}3"].fill.start_color.rgb[-6:], "375623")
        self.assertTrue(all(ws.cell(3, c).comment is None for c in range(2, 9)))
        self.assertIn("Hoja origen: CMg", ws[f'{nombres["CMg Promedio [CLP/kWh]"]}3'].comment.text)
        self.assertIn("Edisp_T", ws[f'{nombres["Energía asignada [kWh]"]}3'].comment.text)
        self.assertIn(
            "Fórmula de Excel equivalente en la fila 4",
            ws[f'{nombres["Monto a compensar [CLP]"]}3'].comment.text,
        )
        # La formula de prueba va en la fila 4; abajo, valores.
        self.assertTrue(str(ws[f"{monto}4"].value).startswith("=IF(SUMIFS("))
        self.assertEqual([ws[f"{monto}{f}"].value for f in (5, 6)], [2, 3])

    def test_re545_con_aporte_ctf_deja_la_reserva_como_valor(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ws = self._escribir_re545(carpeta, ctf_en_cero=False)
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        reserva = nombres["Reserva con FD y FMA [kWh]"]
        self.assertEqual(ws[f"{reserva}4"].value, 1)

    def test_el_resumen_lleva_sus_totales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = self.escribir(carpeta)
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_RESUMEN]
            resumen = nucleo.leer_hoja_con_fila_vacia(ruta, nucleo.HOJA_RESUMEN)
        # B NOMBRE, C RECIBE, D PAGA, E NETO; una empresa (fila 4).
        self.assertEqual(ws["C1"].value, "=SUM(C4:C4)")
        self.assertEqual(ws["D1"].value, "=SUM(D4:D4)")
        self.assertEqual(ws["C1"].font.color.rgb[-6:], "244062")
        self.assertEqual(ws["E1"].value, "=ROUND(SUM(E4:E4),0)")
        # Verde si da 0, rojo si no (en el formato de numero).
        self.assertEqual(ws["E1"].font.color.rgb[-6:], "375623")
        self.assertIn("[Red]", ws["E1"].number_format)
        # La lectura sigue encontrando los nombres.
        self.assertEqual(list(resumen["NOMBRE"]), ["EMPRESA UNO"])

    def test_la_columna_del_total_alcanza_para_el_numero(self):
        """
        Pedido del usuario: el total de la fila 1 se ve entero al abrir
        la planilla, sin "#####" ni ajustar el ancho a mano.
        """

        df = self.armar_calculo()
        df["Monto a compensar"] = [4.0e9, 3.0e9, 2.5e9]
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            nucleo.escribir_pagos_bess(
                ruta, df_ecostos=df,
                df_resumen=pd.DataFrame({
                    "NOMBRE": ["EMPRESA UNO", "EMPRESA DOS"],
                    "RECIBE": [9.5e9, 2.0e9],
                    "PAGA": [1.0, 2.0],
                    "NETO": [9.5e9, 2.0e9],
                }),
                registrar=lambda *a: None,
            )
            wb = openpyxl.load_workbook(ruta)

        def ancho_minimo(total):
            return len(f"Total: {round(total):,}")

        resumen = wb[nucleo.HOJA_RESUMEN]
        for letra, total in (("C", 11.5e9), ("E", 11.5e9)):
            self.assertGreater(
                resumen.column_dimensions[letra].width, ancho_minimo(total),
                letra,
            )
        ws = wb[nucleo.HOJA_CALCULO_ECOSTOS]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        monto = nombres["Monto a compensar [CLP]"]
        self.assertGreater(
            ws.column_dimensions[monto].width, ancho_minimo(9.5e9)
        )

    def test_los_comentarios_traen_la_formula_para_probar(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = self.escribir(carpeta)
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_CALCULO_ECOSTOS]
        nombres = {c.value: c for c in ws[3] if c.value}
        monto = nombres["Monto a compensar [CLP]"].comment.text
        self.assertIn("Fórmula de Excel equivalente en la fila 4", monto)
        # En español y con ";" (para pegar en un Excel en español).
        self.assertIn("=MAX(0;(SUMAR.SI.CONJUNTO(", monto)
        self.assertNotIn("SUMIFS", monto)
        bloque = nombres["Bloque ordenado"].comment.text
        self.assertIn("en la fila 5", bloque)  # mira la fila de arriba
        curva = nombres["Curva monotona CMg Descarga [CLP/kWh]"].comment.text
        self.assertIn("matricial", curva)

    def test_la_formula_de_prueba_va_en_la_primera_fila(self):
        """
        Pedido del usuario: la formula equivalente en la primera fila de
        datos (no solo en el comentario) y el resto de la columna como
        valores del programa, para extenderla y comparar. La celda
        guarda tambien el valor del programa (lo que leen la
        compensacion y el CSV sin abrir Excel).
        """

        from openpyxl.worksheet.formula import ArrayFormula

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = self.escribir(carpeta)
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_CALCULO_ECOSTOS]
            valores = openpyxl.load_workbook(ruta, data_only=True)[
                nucleo.HOJA_CALCULO_ECOSTOS
            ]
            # Como la vuelve a leer el programa (pandas, valores).
            leida = pd.read_excel(
                ruta, sheet_name=nucleo.HOJA_CALCULO_ECOSTOS, header=2
            )
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        monto = nombres["Monto a compensar [CLP]"]
        self.assertEqual(
            ws[f"{monto}4"].value,
            f"=MAX(0,(SUMIFS($AM$4:$AM$6,$H$4:$H$6,H4,$M$4:$M$6,M4)"
            f"-SUMIFS($AQ$4:$AQ$6,$H$4:$H$6,H4,$M$4:$M$6,M4))"
            f"/COUNTIFS($H$4:$H$6,H4,$M$4:$M$6,M4))",
        )
        self.assertEqual(
            [valores[f"{monto}{f}"].value for f in (4, 5, 6)], [1, 2, 3]
        )
        self.assertEqual([ws[f"{monto}{f}"].value for f in (5, 6)], [2, 3])
        # Mira la fila de arriba: va en la 5, y la 4 queda como valor.
        bloque = nombres["Bloque ordenado"]
        self.assertTrue(str(ws[f"{bloque}5"].value).startswith("=IF("))
        self.assertNotIsInstance(ws[f"{bloque}4"].value, str)
        # Matricial: como formula de matriz.
        curva = nombres["Curva monotona CMg Descarga [CLP/kWh]"]
        self.assertIsInstance(ws[f"{curva}4"].value, ArrayFormula)
        self.assertEqual(valores[f"{curva}4"].value, 1)
        self.assertEqual(
            leida["Monto a compensar [CLP]"].tolist(), [1, 2, 3]
        )

    def test_el_resumen_queda_primero_y_el_calculo_detras(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = self.escribir(carpeta)
            hojas = pd.ExcelFile(ruta).sheet_names

            self.assertEqual(hojas[0], nucleo.HOJA_RESUMEN)
            self.assertLess(
                hojas.index(nucleo.HOJA_RESUMEN),
                hojas.index(nucleo.HOJA_CALCULO_ECOSTOS),
            )

    def test_la_hoja_queda_legible_sin_tocar_los_datos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = self.escribir(carpeta)
            wb = openpyxl.load_workbook(ruta)
            ws = wb[nucleo.HOJA_CALCULO_ECOSTOS]

            # Tres cuadros: fila 1 vacia, "Cuadro N°" en la 2, nombres
            # en la 3 (en negrita), margen en A; panel fijo en C:H.
            # Fila 1: solo los totales, sobre Total C1, Total C2 y Monto.
            nombres = {c.value: c.column_letter for c in ws[3] if c.value}
            totales = {c.column_letter: c.value for c in ws[1] if c.value}
            self.assertEqual(set(totales), {
                nombres["Total C1 [CLP]"], nombres["Total C2 [CLP]"],
                nombres["Monto a compensar [CLP]"],
            })
            monto = nombres["Monto a compensar [CLP]"]
            self.assertEqual(totales[monto], f"=SUM({monto}4:{monto}6)")
            self.assertEqual(ws[f"{monto}1"].number_format, '"Total: "#,##0')
            # El Monto en su propio cuadro, en verde.
            self.assertEqual(ws[f"{monto}3"].fill.start_color.rgb[-6:], "375623")
            self.assertTrue(str(ws[f"{monto}2"].value).startswith("Cuadro N° 4"))
            # Lo que se calcula en la misma fila va como formula.
            ajustada = nombres["Descarga ajustada FD-prorrata [kWh]"]
            self.assertTrue(str(ws[f"{ajustada}4"].value).startswith(
                f"={nombres['Descarga [kWh]']}4*("
            ))
            self.assertEqual(
                ws[f"{nombres['Total C1 [CLP]']}5"].value,
                f"={nombres['Valorización Descarga [CLP]']}5"
                f"+{nombres['Valorización Carga [CLP]']}5",
            )
            # Curva monotona sin fila en ese puesto: guion, no vacio.
            # (la fila 4 lleva la formula de prueba, ver
            # test_la_formula_de_prueba_va_en_la_primera_fila)
            curva = nombres["Curva monotona CMg Descarga [CLP/kWh]"]
            self.assertEqual(
                [ws[f"{curva}{f}"].value for f in (5, 6)], [2, 3]
            )
            # B:H sin comentario; CMg con su origen.
            self.assertTrue(all(ws.cell(3, c).comment is None for c in range(2, 9)))
            titulos = [c.value for c in ws[2] if c.value]
            self.assertEqual(len(titulos), 4, titulos)
            self.assertTrue(titulos[0].startswith("Cuadro N° 1"))
            self.assertEqual(ws["B3"].value, "Mes")
            self.assertTrue(ws["B3"].font.bold)
            self.assertEqual(ws["H3"].value, "Configuracion")
            self.assertEqual(ws.freeze_panes, "I4")
            # El origen va como comentario en el encabezado.
            self.assertIn("Hoja origen: CMg", next(
                c for c in ws[3] if str(c.value).startswith("CMg")
            ).comment.text)

            # ancho de columna puesto (pandas no pone ninguno)
            self.assertIsNotNone(ws.column_dimensions["A"].width)

            # y los datos son los mismos
            # La hoja muestra la unidad entre corchetes...
            self.assertIn("Descarga [kWh]", [c.value for c in ws[3]])
            # ...y al leerla vuelve el nombre de siempre.
            from Script.nucleo.unidades import sin_unidades
            df = sin_unidades(nucleo.sin_margen_izquierdo(pd.read_excel(
                ruta, sheet_name=nucleo.HOJA_CALCULO_ECOSTOS, header=2
            )), nucleo.HOJA_CALCULO_ECOSTOS)
            self.assertEqual(list(df["Descarga kWh"]), [1.5, 2.5, 3.5])
            self.assertEqual(list(df["Configuracion"]), ["SAE UNO"] * 3)


class TestUnaSolaPlanilla(unittest.TestCase):
    """
    Pedido del usuario: "combinar el consolidado entradas con pagos
    bess pero ordenados de fin a inicio, el fin es el resumen y el
    inicio las entradas", y las hojas de control en otra planilla.

    Las dos mitades (entradas y calculo) escriben el MISMO archivo por
    separado: la que escribe tiene que dejar intacta la otra mitad.
    """

    def entradas(self, ruta, control=None):
        nucleo.escribir_salida(
            pd.DataFrame({"Mes": [7, 7]}),
            ruta,
            ["un aviso de la corrida"],
            [],
            df_cmg=pd.DataFrame({"cmg": [1.0, 2.0]}),
            df_subastas=pd.DataFrame({"subasta": [1, 2]}),
            hojas_regenerar={
                nucleo.HOJA_MEDIDORES, nucleo.HOJA_CMG, nucleo.HOJA_SUBASTAS,
                nucleo.HOJA_CENTRALES,
            },
            ruta_existente=ruta,
            ruta_control=control,
            registrar=lambda *a: None,
            centrales=(
                pd.DataFrame({"Nombre activo": ["SAE UNO"]}),
                pd.DataFrame([["Balance_BESS", "FD"], ["SAE UNO", "SAE 1"]]),
            ),
        )

    def calculo(self, ruta):
        nucleo.escribir_pagos_bess(
            ruta,
            df_resumen=pd.DataFrame({
                "NOMBRE": ["EMPRESA UNO"], "RECIBE": [1.0],
                "PAGA": [0.0], "NETO": [1.0],
            }),
            hojas_regenerar={nucleo.HOJA_RESUMEN},
            ruta_existente=ruta,
            registrar=lambda *a: None,
        )

    def test_las_hojas_quedan_del_resumen_a_las_entradas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            self.entradas(ruta)
            self.calculo(ruta)

            hojas = pd.ExcelFile(ruta).sheet_names

            self.assertEqual(hojas, list(nucleo.ORDEN_HOJAS_SALIDA))
            self.assertEqual(hojas[0], nucleo.HOJA_RESUMEN)
            # La ultima es la foto de Centrales.xlsx: el maestro del
            # que sale todo lo demas.
            self.assertEqual(hojas[-1], nucleo.HOJA_CENTRALES)

    def test_cada_mitad_preserva_la_otra(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")

            self.entradas(ruta)
            self.calculo(ruta)

            # El calculo no se llevo puestas las entradas...
            self.assertEqual(
                len(nucleo.leer_hoja_con_fila_vacia(
                    ruta, nucleo.HOJA_MEDIDORES
                )), 2
            )

            # ...ni las entradas se llevan puesto el calculo.
            self.entradas(ruta)
            resumen = nucleo.leer_hoja_con_fila_vacia(ruta, nucleo.HOJA_RESUMEN)
            self.assertEqual(list(resumen["NOMBRE"]), ["EMPRESA UNO"])

    def test_el_control_va_en_su_propio_archivo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            control = Path(carpeta) / nucleo.CARPETA_LOG / nucleo.ARCHIVO_CONTROL

            self.entradas(ruta, control)

            self.assertNotIn(nucleo.HOJA_LOG, pd.ExcelFile(ruta).sheet_names)

            log = nucleo.leer_hoja_con_fila_vacia(control, nucleo.HOJA_LOG)
            # Mismo formato que la planilla: fila 1 vacia, cuadro en B2.
            hoja_log = openpyxl.load_workbook(control)[nucleo.HOJA_LOG]
            self.assertIsNone(hoja_log["A1"].value)
            self.assertEqual(
                hoja_log["B2"].value, "Cuadro N° 1 — Avisos e incidencias"
            )
            self.assertIn("un aviso de la corrida", list(log["detalle"]))

    def test_el_control_conserva_las_hojas_que_no_reescribe(self):
        with tempfile.TemporaryDirectory() as carpeta:
            control = Path(carpeta) / nucleo.CARPETA_LOG / nucleo.ARCHIVO_CONTROL

            registro = nucleo.Registro()
            nucleo.escribir_control(
                control, registro=registro, periodo="07",
                hojas_regeneradas=["Calculo RE545"],
            )
            nucleo.escribir_control(
                control,
                df_log=pd.DataFrame(
                    [("ok", "Sin observaciones.")],
                    columns=["tipo", "detalle"],
                ),
            )

            hojas = pd.ExcelFile(control).sheet_names

            # La matriz de centrales solo sale cuando el calculo lee las
            # entradas; las demas siempre, en su orden.
            self.assertEqual(
                hojas, [h for h in nucleo.ORDEN_HOJAS_CONTROL if h in hojas]
            )
            for hoja in (nucleo.HOJA_EJECUCION, nucleo.HOJA_ALERTAS, nucleo.HOJA_LOG):
                self.assertIn(hoja, hojas)


class TestFormatoNoTocaLosDatos(unittest.TestCase):

    def test_formatear_hoja_deja_los_valores_como_estaban(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "x.xlsx"
            df = pd.DataFrame({
                "texto": ["a", "b"],
                "entero": [1, 2],
                "decimal": [1.25, 2.5],
                "vacia": [None, None],
            })
            df.to_excel(ruta, index=False)

            wb = openpyxl.load_workbook(ruta)
            ws = wb.active
            antes = [[c.value for c in fila] for fila in ws.iter_rows()]

            nucleo.formatear_hoja(ws)

            self.assertEqual(
                [[c.value for c in fila] for fila in ws.iter_rows()], antes
            )
            self.assertEqual(ws.cell(2, 2).number_format, "#,##0")
            self.assertEqual(ws.cell(2, 3).number_format, "#,##0.00")

    def test_una_hoja_vacia_no_rompe(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "v.xlsx"
            pd.DataFrame().to_excel(ruta, index=False)

            wb = openpyxl.load_workbook(ruta)
            nucleo.formatear_libro(wb)  # no debe levantar


if __name__ == "__main__":
    unittest.main()


class TestNoSeAbrenPlanillasDeMas(unittest.TestCase):
    """
    Pedido del usuario: "que no se abran planillas innecesarias" y
    "si alguna informacion esta en el consolidado que se saque de ahi".
    """

    def test_el_cmg_de_los_pagos_sale_del_consolidado(self):
        """
        La hoja 'CMg' del consolidado y cmg.xlsx son el mismo dato: se
        lee el consolidado (la foto de las entradas), no el archivo
        original.
        """

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            cmg = pd.DataFrame({
                "A": ["x"], "B": ["y"], "C": ["z"],
                "D_Barra": ["BARRA A"], "E": [""], "F_CMg": [50.0],
                "G": [""], "H_Cuarto": [1], "I_CMgProm": [45.0],
            })
            with pd.ExcelWriter(ruta) as w:
                cmg.to_excel(w, sheet_name="CMg", index=False)

            leido = nucleo.leer_cmg_consolidado(ruta, registrar=lambda *a: None)

            self.assertEqual(len(leido.columns), 9)
            self.assertEqual(
                nucleo.construir_dic_cmg(leido),
                nucleo.construir_dic_cmg(cmg),
            )

    def test_sin_la_hoja_cmg_el_error_dice_que_hay_que_generarla(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.nombre_salida("2607")
            pd.DataFrame({"a": [1]}).to_excel(
                ruta, sheet_name="Medidores", index=False
            )

            with self.assertRaises(nucleo.ErrorEntrada) as fallo:
                nucleo.leer_cmg_consolidado(ruta, registrar=lambda *a: None)

            self.assertIn("CMg", str(fallo.exception))
