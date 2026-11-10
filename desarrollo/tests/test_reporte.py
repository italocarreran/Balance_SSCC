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
        r.C_DESCARGA: 500.0, r.C_CARGA: -600.0, r.C_CMG: 50.0,
        r.C_HORAS_ADJ: 100.0, r.C_ADJ_CPF: 10.0, r.C_ADJ_CSF: 5.0,
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

    def test_explica_con_lo_que_se_movio_en_el_mismo_sentido(self):
        # Cambios de menos de $1.000 no se comentan: montos realistas.
        anterior = pd.DataFrame([
            fila(**{r.C_RE545: 200000.0, r.C_TOTAL: 200100.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        actual = pd.DataFrame([
            fila(**{r.C_RE545: 400000.0, r.C_TOTAL: 400100.0, r.C_OFERTA_DIAS: 28.0,
                    r.C_ADJ_CPF: 15.0, r.C_CMG: 10.0}),
            fila(**{r.C_CENTRAL: "SAE-DOS"}),
        ])
        lineas, tabla = r.comparar_resumenes(actual, anterior)
        uno = tabla.set_index(r.C_CENTRAL).loc["SAE-UNO", "Explicación"]
        self.assertTrue(uno.startswith("RE545 sube $200.000 (+100 %):"), uno)
        self.assertIn("lo adjudicado en CPF subió (10 → 15 MW)", uno)
        self.assertIn("los días con oferta completa subieron (20 → 28 días)", uno)
        # El CMg bajo, pero el RE545 sube: no lo explica, no se nombra.
        self.assertNotIn("CMg", uno)
        # Lo que no cambio va en una sola linea, no en su fila.
        self.assertEqual(tabla.set_index(r.C_CENTRAL).loc["SAE-DOS", "Explicación"], "")
        self.assertIn("Sin cambios relevantes: SAE-DOS.", lineas)
        # Lo que mas cambio, arriba.
        self.assertEqual(tabla[r.C_CENTRAL].iloc[0], "SAE-UNO")

    def test_central_que_estuvo_parte_del_mes_anterior(self):
        """
        Caso real (SAE-CRCA-PFV-GABRIELA): E Costos +2.209 % y el reporte
        decia "sin un cambio grande en sus entradas". El mes anterior
        estuvo solo una parte del mes: eso tiene que salir, y la columna
        Datos muestra lo adjudicado "tanto vs tanto".
        """

        anterior = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 9.8e6, r.C_RE545: 0.0, r.C_TOTAL: 9.8e6,
            r.C_DIAS: 9.0, r.C_HORAS_SUB: 150.0, r.C_ADJ_CPF: 900.0,
        })])
        actual = pd.DataFrame([fila(**{
            r.C_ECOSTOS: 227.3e6, r.C_RE545: 0.0, r.C_TOTAL: 227.3e6,
            r.C_ADJ_CPF: 4300.0, r.C_HORAS_SUB: 700.0,
        })])
        _, tabla = r.comparar_resumenes(actual, anterior)
        texto = tabla["Explicación"].iloc[0]
        self.assertIn("los días en el mes subieron (9 → 31 días)", texto)
        self.assertIn("las horas adjudicadas subieron (150 → 700 h)", texto)
        datos = tabla["Datos (anterior → actual)"].iloc[0]
        self.assertIn("Adjudicado CPF: 900 → 4.300 MW", datos)
        # Sin ruido: lo que no cambio no va, salvo el FD, que va siempre
        # ("yo creo que tiene que ver con los FD"): dice que fue igual.
        self.assertIn("FD CPF: 0,90 (igual)", datos)
        self.assertNotIn("Días con oferta completa", datos)

    def test_nueva_sin_ingreso_dice_por_que(self):
        """Caso real (SAE-LILE): adjudicada, pero FD = 0 todo el mes."""

        anterior = pd.DataFrame([fila(**{r.C_CENTRAL: "OTRA"})])
        actual = pd.DataFrame([
            fila(**{r.C_CENTRAL: "OTRA"}),
            fila(**{r.C_CENTRAL: "SAE-LILE", r.C_ECOSTOS: 0.0, r.C_RE545: 0.0,
                    r.C_TOTAL: 0.0, r.C_FD_CPF: 0.0, r.C_FD_CSF: 0.0}),
        ])
        _, tabla = r.comparar_resumenes(actual, anterior)
        lile = tabla.set_index(r.C_CENTRAL).loc["SAE-LILE"]
        self.assertEqual(
            lile["Explicación"],
            "Nueva: no estaba en la planilla anterior. Recibe $0. Según FD "
            "no tuvo participación (FD CPF y CSF = 0 en todo el mes).",
        )
        self.assertIn("Adjudicado CPF: 10 MW", lile["Datos (anterior → actual)"])
        self.assertNotIn("sin dato", lile["Datos (anterior → actual)"])

    def test_queda_en_cero_dice_por_que(self):
        anterior = pd.DataFrame([fila(**{r.C_RE545: 2e6, r.C_TOTAL: 2.1e6})])
        actual = pd.DataFrame([fila(**{r.C_ECOSTOS: 0.0, r.C_RE545: 0.0,
                                       r.C_TOTAL: 0.0, r.C_FD_CPF: 0.0})])
        _, tabla = r.comparar_resumenes(actual, anterior)
        self.assertIn("Queda en $0: Según FD no tuvo participación",
                      tabla["Explicación"].iloc[0])

    def test_central_nueva_y_que_ya_no_esta(self):
        anterior = pd.DataFrame([fila(**{r.C_CENTRAL: "VIEJA"})])
        actual = pd.DataFrame([fila(**{r.C_CENTRAL: "NUEVA"})])
        _, tabla = r.comparar_resumenes(actual, anterior)
        textos = dict(zip(tabla[r.C_CENTRAL], tabla["Explicación"]))
        self.assertTrue(textos["NUEVA"].startswith("Nueva"))
        self.assertTrue(textos["VIEJA"].startswith("Ya no está"))

    def test_plata_a_la_chilena(self):
        self.assertEqual(r.plata(1234567.6), "$1.234.568")
        self.assertEqual(r.plata(-5), "-$5")


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
            self.assertEqual(libro.sheetnames, ["Comparacion", "Detalle"])
            ws = libro["Comparacion"]
            self.assertTrue(str(ws["B2"].value).startswith("Cuadro N° 1"))
            self.assertEqual(ws["B3"].value, "Resumen")
            self.assertTrue(lineas[0].startswith("Total del mes:"))
            self.assertTrue(any("RE545 sube" in e for e in tabla["Explicación"]))
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
        _, tabla = r.comparar_resumenes(actual, anterior)
        self.assertIn(
            "las horas con 'No participó' en FD CSF bajaron (700 → 0 h)",
            tabla["Explicación"].iloc[0],
        )
        self.assertIn(
            "Horas FD CSF 'No participó': 700 de 720 → 0 de 744",
            tabla["Datos (anterior → actual)"].iloc[0],
        )


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
