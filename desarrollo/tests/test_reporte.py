# -*- coding: utf-8 -*-
"""
Los reportes en palabras (Script/nucleo/reporte.py).

Pedido del usuario (2026-11-10):
  - en el control de corrida, un boton "Generar reporte": si una central
    tiene cero ingreso, que diga por que ("segun FD no tuvo participacion,
    o FD=0");
  - en Herramientas, comparar dos planillas (mes o version anterior) y
    explicar los cambios ("tal central aumento su ingreso por RE545 dado
    que las ofertas aumentaron tanto"), con el resultado en Log/.
"""

import math
import tempfile
import unittest
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo import reporte as r


def fila(**valores):
    """Una central con todo "normal" (recibe por los dos lados)."""

    base = {
        r.C_CENTRAL: "SAE-UNO", r.C_EMPRESA: "EMPRESA",
        r.C_ECOSTOS: 100.0, r.C_RE545: 200.0, r.C_TOTAL: 300.0,
        r.C_DESCARGA: 500.0, r.C_CARGA: 600.0, r.C_CMG: 50.0,
        r.C_HORAS_ADJ: 300.0, r.C_ADJ_CPF: 10.0, r.C_ADJ_CSF: 5.0,
        r.C_RESERVA: 1000.0, r.C_FD_CPF: 0.9, r.C_FD_CSF: 0.8,
        r.C_OFERTA_DIAS: 20.0, r.C_DIAS: 31.0, r.C_HORAS_SUB: 300.0,
    }
    base.update(valores)
    return base


class TestPorQueNoRecibe(unittest.TestCase):

    def test_recibe_por_los_dos_lados_no_dice_nada(self):
        self.assertEqual(r.explicar_central(fila()), "")

    def test_fd_cero(self):
        texto = r.explicar_central(fila(
            **{r.C_ECOSTOS: 0.0, r.C_RE545: 0.0, r.C_TOTAL: 0.0,
               r.C_FD_CPF: 0.0, r.C_ADJ_CSF: 0.0}
        ))
        self.assertEqual(
            texto, "Según FD no tuvo participación (FD CPF = 0 en todo el mes)."
        )

    def test_sin_adjudicacion(self):
        texto = r.explicar_central(fila(
            **{r.C_TOTAL: 0.0, r.C_ADJ_CPF: 0.0, r.C_ADJ_CSF: 0.0}
        ))
        self.assertEqual(texto, "No se adjudicó nada en Subastas en el mes.")

    def test_sin_energia_manda(self):
        texto = r.explicar_central(fila(
            **{r.C_TOTAL: 0.0, r.C_DESCARGA: 0.0, r.C_CARGA: 0.0,
               r.C_FD_CPF: 0.0}
        ))
        self.assertTrue(texto.startswith("Sin energía medida"), texto)

    def test_adjudicada_sin_cruce(self):
        texto = r.explicar_central(fila(**{r.C_TOTAL: 0.0, r.C_HORAS_ADJ: 0.0}))
        self.assertIn("Diccionario", texto)

    def test_sin_motivo_en_las_entradas(self):
        texto = r.explicar_central(fila(**{r.C_TOTAL: 0.0}))
        self.assertIn("Total C1 ≤ Total C2", texto)

    def test_solo_e_costos(self):
        texto = r.explicar_central(fila(
            **{r.C_RE545: 0.0, r.C_TOTAL: 100.0, r.C_FD_CPF: 0.0}
        ))
        self.assertTrue(texto.startswith("Solo recibe por E Costos: según FD"), texto)

    def test_resumen_de_la_corrida(self):
        resumen = pd.DataFrame([
            fila(), fila(**{r.C_CENTRAL: "SAE-DOS", r.C_ECOSTOS: 0.0,
                            r.C_RE545: 0.0, r.C_TOTAL: 0.0}),
        ])
        lineas, tabla = r.reporte_corrida(resumen)
        self.assertEqual(lineas[0], "Total del mes: $300 entre 2 centrales.")
        self.assertIn("SAE-DOS", lineas[-1])
        self.assertEqual(tabla["Observación"].iloc[0], "")


class TestComparacion(unittest.TestCase):

    def test_que_paso_y_lo_que_mas_cambio(self):
        # Cambios de menos de $1.000 no se comentan: montos realistas.
        anterior = pd.DataFrame([
            fila(**{r.C_RE545: 200000.0, r.C_TOTAL: 200100.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        actual = pd.DataFrame([
            fila(**{r.C_RE545: 400000.0, r.C_TOTAL: 400100.0, r.C_OFERTA_DIAS: 28.0,
                    r.C_ADJ_CPF: 15.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        lineas, tabla = r.comparar_resumenes(actual, anterior)
        uno = tabla.set_index(r.C_CENTRAL).loc["SAE-UNO"]
        self.assertEqual(uno["Qué pasó"], "RE545 sube $200.000 (+100 %).")
        # Los responsables en pesos van en sus columnas (test_responsables).
        self.assertIn("Principal responsable", tabla.columns)
        self.assertIn("Desglose en $", tabla.columns)
        self.assertNotIn("Lo que más cambió", tabla.columns)
        # Lo que no cambio va en una sola linea, no en su fila.
        self.assertEqual(tabla.set_index(r.C_CENTRAL).loc["SAE-DOS", "Qué pasó"], "")
        self.assertIn("Sin cambios relevantes: SAE-DOS.", lineas)
        # Lo que mas cambio, arriba.
        self.assertEqual(tabla[r.C_CENTRAL].iloc[0], "SAE-UNO")

    def test_caso_gabriela_con_c1_c2_y_ciclos(self):
        """
        Caso real (SAE-CRCA-PFV-GABRIELA, sept -> oct): E Costos +2.209 %.
        C1 casi igual, C2 se cae (el CMg baja) y los ciclos con monto pasan
        de 4 a 30: eso es lo que tiene que decir.
        """

        anterior = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 9845069.2, r.C_RE545: 0.0, r.C_TOTAL: 9845069.2,
            r.C_EC_C1: 518e6, r.C_EC_C2: 653e6, r.C_EC_CICLOS: 4.0,
            r.C_EC_CICLOS_MES: 31.0, r.C_ADJ_CPF: 17274.0,
        })])
        actual = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 227344681.2, r.C_RE545: 0.0, r.C_TOTAL: 227344681.2,
            r.C_EC_C1: 471e6, r.C_EC_C2: 244e6, r.C_EC_CICLOS: 30.0,
            r.C_EC_CICLOS_MES: 32.0, r.C_ADJ_CPF: 0.0,
        })])
        _, tabla = r.comparar_resumenes(actual, anterior)
        self.assertEqual(
            tabla["Qué pasó"].iloc[0],
            "E Costos sube $217 MM (+2.209 %): ciclos con monto 4 → 30 de 32; "
            "C1 $518 MM → $471 MM y C2 $653 MM → $244 MM.",
        )

    def test_nueva_sin_ingreso_dice_por_que(self):
        """Caso real (SAE-LILE): adjudicada, pero FD = 0 todo el mes."""

        anterior = pd.DataFrame([fila(**{r.C_CENTRAL: "OTRA"})])
        actual = pd.DataFrame([
            fila(**{r.C_CENTRAL: "OTRA"}),
            fila(**{r.C_CENTRAL: "SAE-LILE", r.C_ECOSTOS: 0.0, r.C_RE545: 0.0,
                    r.C_TOTAL: 0.0, r.C_FD_CPF: 0.0, r.C_FD_CSF: 0.0}),
        ])
        _, tabla = r.comparar_resumenes(actual, anterior)
        self.assertEqual(
            tabla.set_index(r.C_CENTRAL).loc["SAE-LILE", "Qué pasó"],
            "Nueva: no estaba en la planilla anterior. Recibe $0. Según FD "
            "no tuvo participación (FD CPF y CSF = 0 en todo el mes).",
        )

    def test_queda_y_sigue_en_cero_dicen_por_que(self):
        anterior = pd.DataFrame([fila(**{r.C_RE545: 2e6, r.C_TOTAL: 2.1e6})])
        actual = pd.DataFrame([fila(**{r.C_ECOSTOS: 0.0, r.C_RE545: 0.0,
                                       r.C_TOTAL: 0.0, r.C_FD_CPF: 0.0})])
        _, tabla = r.comparar_resumenes(actual, anterior)
        self.assertIn("Queda en $0: Según FD no tuvo participación",
                      tabla["Qué pasó"].iloc[0])
        cero = fila(**{r.C_ECOSTOS: 0.0, r.C_RE545: 0.0, r.C_TOTAL: 0.0,
                       r.C_FD_CSF: 0.0})
        _, tabla = r.comparar_resumenes(pd.DataFrame([cero]), pd.DataFrame([cero]))
        self.assertTrue(tabla["Qué pasó"].iloc[0].startswith("Sigue en $0. Según FD"))

    def test_central_nueva_y_que_ya_no_esta(self):
        anterior = pd.DataFrame([fila(**{r.C_CENTRAL: "VIEJA"})])
        actual = pd.DataFrame([fila(**{r.C_CENTRAL: "NUEVA"})])
        _, tabla = r.comparar_resumenes(actual, anterior)
        textos = dict(zip(tabla[r.C_CENTRAL], tabla["Qué pasó"]))
        self.assertTrue(textos["NUEVA"].startswith("Nueva"))
        self.assertTrue(textos["VIEJA"].startswith("Ya no está"))

    def test_plata_a_la_chilena(self):
        self.assertEqual(r.plata(1234567.6), "$1.234.568")
        self.assertEqual(r.plata(-5), "-$5")
        self.assertEqual(r.plata_corta(217499612.0), "$217 MM")
        self.assertEqual(r.plata_corta(436684.0), "$436.684")


class TestFormulasVacias(unittest.TestCase):
    """
    Las planillas de sept y oct del usuario se escribieron antes de que
    las formulas llevaran su valor: Total C1, Total C2... se leian vacios
    (y el reporte decia 0). Se calculan con la misma plantilla de la hoja.
    """

    def test_e_costos(self):
        from Script.nucleo.ecostos import FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS

        df = pd.DataFrame({
            "Descarga kWh": [10.0, 0.0], "Carga kWh": [0.0, -4.0],
            "CMg": [2.0, 3.0], "Adj SSCC": [1, 0],
            "Valorización Descarga [$]": [50.0, 0.0],
            "Valorización Carga [$]": [0.0, -8.0],
            "Total [CLP]": [None, None], "Total": [None, None],
        })
        r.completar_formulas(df, FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS)
        self.assertEqual(df["Total [CLP]"].tolist(), [50.0, -8.0])     # C1 = AY + AZ
        self.assertEqual(df["Valorizacion Descarga"].tolist(), [20.0, 0.0])
        self.assertEqual(df["Total"].tolist(), [20.0, 0.0])            # C2 = L*(S+T)

    def test_no_pisa_lo_que_trae_valor(self):
        from Script.nucleo.ecostos import FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS

        df = pd.DataFrame({
            "Valorización Descarga [$]": [50.0], "Valorización Carga [$]": [1.0],
            "Total [CLP]": [99.0],
        })
        r.completar_formulas(df, FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS)
        self.assertEqual(df["Total [CLP]"].tolist(), [99.0])

    def test_re545_con_if_count(self):
        from Script.nucleo.re545 import FORMULAS_RE545, NOMBRES_CALCULO_RE545

        df = pd.DataFrame({
            "Energía asignada": [2.0, None], "Curva Cmg Decendente": [5.0, 5.0],
            "Total C1": [None, None],
        })
        r.completar_formulas(df, FORMULAS_RE545, NOMBRES_CALCULO_RE545)
        self.assertEqual(df["Total C1"].iloc[0], 10.0)
        self.assertTrue(math.isnan(df["Total C1"].iloc[1]))   # el "" de la formula


class TestCambiosPorCentral(unittest.TestCase):
    """
    La sub tabla (pedido del usuario: "por cada cosa que diga el cambio y
    su variacion y si cambio o no").
    """

    def test_un_renglon_por_concepto_con_su_cambio(self):
        anterior = pd.DataFrame([fila(**{r.C_RE545: 200000.0, r.C_TOTAL: 200100.0})])
        actual = pd.DataFrame([fila(**{r.C_RE545: 400000.0, r.C_TOTAL: 400100.0,
                                       r.C_ADJ_CPF: 15.0})])
        cambios = r.cambios_por_central(actual, anterior).set_index("Concepto")
        self.assertEqual(list(r.cambios_por_central(actual, anterior).columns),
                         list(r.COLUMNAS_CAMBIOS))
        self.assertEqual(cambios.loc["Recibe RE545", "Cambio"], r.SUBE)
        self.assertEqual(cambios.loc["Recibe RE545", "Variación [%]"], 100.0)
        self.assertEqual(cambios.loc["Adjudicado CPF", "Anterior"], 10.0)
        self.assertEqual(cambios.loc["Adjudicado CPF", "Actual"], 15.0)
        self.assertEqual(cambios.loc["FD CPF promedio", "Cambio"], r.IGUAL)
        # La plata, primero.
        self.assertEqual(
            list(r.cambios_por_central(actual, anterior)["Concepto"])[:3],
            ["Recibe total", "Recibe E Costos", "Recibe RE545"],
        )

    def test_horas_con_adj_sscc_solo_si_no_coinciden(self):
        """
        El usuario: "horas adjudicadas y horas con adjudicacion es lo
        mismo". Lo son cuando el nombre cruza bien; si no coinciden (la
        central no cruza con Medidores) si va, como alerta.
        """

        iguales = fila(**{r.C_HORAS_SUB: 300.0, r.C_HORAS_ADJ: 300.0})
        cambios = r.cambios_por_central(
            pd.DataFrame([iguales]), pd.DataFrame([fila(**{r.C_HORAS_ADJ: 300.0})])
        )
        self.assertFalse(cambios["Concepto"].str.startswith("Horas con Adj SSCC").any())
        no_cruza = fila(**{r.C_HORAS_SUB: 300.0, r.C_HORAS_ADJ: 0.0})
        cambios = r.cambios_por_central(pd.DataFrame([no_cruza]), pd.DataFrame([iguales]))
        self.assertTrue(cambios["Concepto"].str.startswith("Horas con Adj SSCC").any())

    def test_sin_ruido_y_sin_dato(self):
        anterior = pd.DataFrame([fila(**{r.C_ADJ_CSF: 0.0})])
        actual = pd.DataFrame([fila(**{r.C_ADJ_CSF: 0.0, r.C_FD_CSF: math.nan})])
        cambios = r.cambios_por_central(actual, anterior).set_index("Concepto")
        self.assertNotIn("Adjudicado CSF", cambios.index)       # 0 en los dos
        self.assertEqual(cambios.loc["FD CSF promedio", "Cambio"], r.SIN_DATO)

    def test_nueva(self):
        cambios = r.cambios_por_central(
            pd.DataFrame([fila()]), pd.DataFrame(columns=list(r.COLUMNAS_RESUMEN))
        )
        self.assertTrue((cambios["Cambio"] == r.NUEVO).all())


def _planilla(carpeta, nombre, multiplicar=1.0):
    from desarrollo.tests import test_relleno_rapido as t

    d = t._datos(60)
    # Como la hoja Subastas real (columna J): la hora del mes.
    d["sub"] = d["sub"].assign(Hora_mes=np.arange(len(d["sub"])) % 24 + 1)
    for clave in ("ec", "re545"):
        d[clave] = d[clave].copy()
        d[clave]["Monto a compensar"] = (
            np.abs(d[clave]["Monto a compensar"]) * 1e5 * multiplicar
        )
    Path(carpeta).mkdir(parents=True, exist_ok=True)
    ruta = Path(t._escribir(carpeta, True, d))
    destino = ruta.with_name(nombre)
    ruta.rename(destino)
    return destino


class TestArchivos(unittest.TestCase):

    def test_la_comparacion_queda_en_el_log_de_la_actual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            anterior = _planilla(Path(carpeta) / "sep", "11_PAGOS_BESS_2609.xlsx")
            actual = _planilla(Path(carpeta) / "oct", "11_PAGOS_BESS_2610.xlsx", 1.5)
            destino, lineas, tabla = nucleo.comparar_planillas(
                actual, anterior, registrar=lambda *a: None
            )
            self.assertEqual(destino, actual.parent / "Log" / "Comparacion_2610_vs_2609.xlsx")
            libro = openpyxl.load_workbook(destino)
            self.assertEqual(
                libro.sheetnames, ["Comparacion", "Responsables", "Cambios por central"]
            )
            responsables = libro["Responsables"]
            self.assertEqual(responsables["B3"].value, "Responsable")
            nombres = [c.value for c in libro["Comparacion"][3] if c.value]
            self.assertIn("Principal responsable", nombres)
            self.assertIn("Desglose en $", nombres)
            sub = libro["Cambios por central"]
            self.assertEqual(
                [c.value for c in sub[3] if c.value], list(r.COLUMNAS_CAMBIOS)
            )
            self.assertIsNotNone(sub.auto_filter.ref)
            ws = libro["Comparacion"]
            self.assertTrue(str(ws["B2"].value).startswith("Cuadro N° 1"))
            self.assertEqual(ws["B3"].value, "Resumen")
            self.assertTrue(lineas[0].startswith("Total del mes:"))
            self.assertTrue(any("RE545 sube" in e for e in tabla["Qué pasó"]))
            # Las columnas con formula (Total C1, Total C2...) llevan el
            # valor del programa: se leen sin abrir el libro en Excel.
            self.assertTrue((resumen_tmp := nucleo.resumen_por_central(
                actual, registrar=lambda *a: None))[r.C_EC_C1].abs().sum() > 0)
            self.assertTrue(resumen_tmp[r.C_RESERVA].abs().sum() > 0)
            con_formula = openpyxl.load_workbook(actual)["Calculo E Costos"]
            valores = openpyxl.load_workbook(actual, data_only=True)["Calculo E Costos"]
            letra = next(c.column_letter for c in con_formula[3]
                         if str(c.value).startswith("Total C1"))
            self.assertTrue(str(con_formula[f"{letra}10"].value).startswith("="))
            self.assertIsInstance(valores[f"{letra}10"].value, (int, float))
            # Los dias y las horas adjudicadas salen de la planilla.
            resumen = nucleo.resumen_por_central(actual, registrar=lambda *a: None)
            self.assertTrue((resumen[r.C_DIAS].dropna() > 0).any())
            self.assertTrue((resumen[r.C_HORAS_SUB].dropna() > 0).any())

    def test_la_hoja_reporte_del_control_y_el_resto_se_preserva(self):
        from Script.nucleo.rutas import resolver_rutas

        with tempfile.TemporaryDirectory() as carpeta:
            rutas = resolver_rutas(carpeta, "2610")
            _planilla(carpeta, rutas["salida"].name)
            # Un control que ya tiene otra hoja: no se pierde.
            nucleo.escribir_control(
                rutas["control"],
                df_log=pd.DataFrame({"mensaje": ["algo"]}),
            )
            nucleo.generar_reporte_corrida(carpeta, "2610", registrar=lambda *a: None)
            libro = openpyxl.load_workbook(rutas["control"])
        self.assertIn(nucleo.HOJA_REPORTE, libro.sheetnames)
        self.assertIn(nucleo.HOJA_LOG, libro.sheetnames)
        ws = libro[nucleo.HOJA_REPORTE]
        nombres = [c.value for c in ws[3] if c.value]
        self.assertEqual(nombres[0], "Resumen")
        self.assertIn("Observación", nombres)
        self.assertTrue(str(ws["B4"].value).startswith("Total del mes:"))

    def test_sin_planilla_dice_que_hay_que_calcular(self):
        with tempfile.TemporaryDirectory() as carpeta:
            with self.assertRaises(nucleo.ErrorEntrada):
                nucleo.generar_reporte_corrida(carpeta, "2610", registrar=lambda *a: None)


class TestTablasEnteras(unittest.TestCase):
    """
    El control y los reportes se guardan sin pasar por el pegado de las
    filas: sus tablas quedaban cortadas en 200 filas (la sub tabla real de
    sept/oct tiene 355 renglones; el Log de una corrida, cientos).
    """

    def test_control_con_un_log_largo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_CONTROL
            nucleo.escribir_control(
                ruta, df_log=pd.DataFrame({"mensaje": [f"linea {i}" for i in range(500)]})
            )
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_LOG]
        mensajes = [c.value for c in ws["B"][3:] if c.value]
        self.assertEqual(len(mensajes), 500)
        self.assertEqual(mensajes[-1], "linea 499")

    def test_reporte_con_una_sub_tabla_larga(self):
        tabla = pd.DataFrame({
            r.C_CENTRAL: [f"C{i // 20}" for i in range(355)],
            "Concepto": ["Recibe total"] * 355, "Unidad": ["CLP"] * 355,
            "Anterior": [1.0] * 355, "Actual": [2.0] * 355, "Diferencia": [1.0] * 355,
            "Variación [%]": [100.0] * 355, "Cambio": [r.SUBE] * 355,
        })
        with tempfile.TemporaryDirectory() as carpeta:
            destino = r.escribir_reporte(
                Path(carpeta) / "reporte.xlsx", [("Cambios por central", [], tabla, "x")]
            )
            ws = openpyxl.load_workbook(destino)["Cambios por central"]
        self.assertEqual(ws.max_row, 3 + 355)


class TestNoParticipo(unittest.TestCase):
    """
    El usuario: "el valor de los FD puede ser 1 en las columnas L:M y
    AA:AB, pero en las columnas H y V son 'No participó', ese 'No
    participó' explica que sea cero el monto final".
    """

    def _planilla_fd(self, carpeta, respuesta_csf):
        csf = pd.DataFrame({
            k: [1.0, 1.0] for k in list(nucleo.NOMBRES_FD_CSF.values())[1:-1]
        })
        csf["Hora Mes"] = [1, 2]
        respuesta = next(c for c in csf.columns if c.startswith("Respuesta CSF"))
        csf[respuesta] = respuesta_csf
        csf["Unidad"] = "U"
        csf["Central"] = "SAE-UNO"
        cpf = pd.DataFrame({
            k: [1.0] for k in list(nucleo.NOMBRES_FD_CPF.values())[1:-1]
        })
        cpf["Unidad"] = "U"
        cpf["Central"] = "SAE-UNO"
        subastas = pd.DataFrame({
            "Concepto": ["CSF(+)", "CSF(-)"], "Fecha": ["2026-10-01"] * 2,
            "Año": [2026] * 2, "Mes": [10] * 2, "Dia": [1] * 2,
            "Hora_dia": [1, 2], "Hora_mes": [1, 2],
            "Configuración": ["SAE-UNO"] * 2, "Energía SSCC": [5.0, 5.0],
        })
        resumen = pd.DataFrame({"Nombre activo": ["SAE-UNO"], "Propietario": ["E1"]})
        diccionario = pd.DataFrame([["Balance_BESS", "FD"], ["SAE-UNO", "U"]])
        ruta = Path(carpeta) / "11_PAGOS_BESS_2610.xlsx"
        nucleo.escribir_salida(
            None, ruta, [], [], df_fd_csf=csf, df_fd_cpf=cpf,
            df_subastas=subastas, centrales=(resumen, diccionario),
            hojas_regenerar={nucleo.HOJA_FD, nucleo.HOJA_SUBASTAS,
                             nucleo.HOJA_CENTRALES},
            registrar=lambda *a: None,
        )
        return nucleo.resumen_por_central(ruta, registrar=lambda *a: None)

    def test_cuenta_las_horas_aunque_el_fd_diga_1(self):
        with tempfile.TemporaryDirectory() as carpeta:
            resumen = self._planilla_fd(carpeta, ["No Participó", "No participo"])
        uno = resumen.set_index(r.C_CENTRAL).loc["SAE-UNO"]
        self.assertEqual(uno[r.C_FD_CSF], 1.0)        # el factor no lo ve
        self.assertEqual(uno[r.C_NP_CSF], 2.0)
        self.assertEqual(uno[r.C_FDH_CSF], 2.0)
        fila_cero = dict(uno)
        fila_cero.update({r.C_ECOSTOS: 0.0, r.C_RE545: 0.0, r.C_TOTAL: 0.0})
        self.assertEqual(
            r.explicar_central(fila_cero),
            "Según FD no tuvo participación en CSF ('No participó' en todas "
            "sus 2 horas, aunque el FD diga 1).",
        )

    def test_participo_no_cuenta(self):
        with tempfile.TemporaryDirectory() as carpeta:
            resumen = self._planilla_fd(carpeta, [0.95, "No participó"])
        self.assertEqual(resumen.set_index(r.C_CENTRAL).loc["SAE-UNO", r.C_NP_CSF], 1.0)

    def test_la_comparacion_lo_usa(self):
        anterior = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 9.8e6, r.C_RE545: 0.0, r.C_TOTAL: 9.8e6,
            r.C_NP_CSF: 700.0, r.C_FDH_CSF: 720.0,
        })])
        actual = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 227.3e6, r.C_RE545: 0.0, r.C_TOTAL: 227.3e6,
            r.C_NP_CSF: 0.0, r.C_FDH_CSF: 744.0,
        })])
        cambios = r.cambios_por_central(actual, anterior).set_index("Concepto")
        self.assertEqual(cambios.loc["FD CSF: horas 'No participó'", "Anterior"], 700.0)
        self.assertEqual(cambios.loc["FD CSF: horas 'No participó'", "Cambio"], r.BAJA)


class TestHerramienta(unittest.TestCase):

    def test_sugiere_la_planilla_del_mes_anterior(self):
        import importlib.util

        ruta = Path(__file__).resolve().parents[2] / "Herramientas" / "Comparar_Meses.py"
        spec = importlib.util.spec_from_file_location("comparar_meses", ruta)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)

        self.assertEqual(modulo.aamm_anterior("2601"), "2512")
        self.assertEqual(modulo.aamm_anterior("2610"), "2609")
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "2609" / "SSCC").mkdir(parents=True)
            (base / "2610" / "SSCC").mkdir(parents=True)
            anterior = base / "2609" / "SSCC" / "11_PAGOS_BESS_2609.xlsx"
            anterior.write_bytes(b"")
            actual = base / "2610" / "SSCC" / "11_PAGOS_BESS_2610.xlsx"
            self.assertEqual(modulo.sugerir_anterior(actual), anterior)


if __name__ == "__main__":
    unittest.main()
