# -*- coding: utf-8 -*-
"""
Los responsables en pesos del cambio de lo que recibe una central
(Script/nucleo/responsables.py).

Pedido del usuario (2026-11-10): "necesito que se entienda claramente el
principal responsable del aumento o bajada de ingresos por central [...]
me preguntan 'porque en sept eran 1000 palitos y ahora son 1800?'".
La suma de los responsables tiene que dar EXACTO la diferencia.
"""

import math
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script.nucleo import reporte as r
from Script.nucleo import responsables as rs


def fila_ec(grupos, e_tot=1000.0, e=1000.0, l=800.0, q=600.0, nombre="SAE-UNO", **extra):
    """
    Una central solo con descarga en E Costos. grupos: [(C1, C2)] por
    ciclo; C1 y C2 totales salen de ahi.
    """

    c1 = sum(g[0] for g in grupos)
    c2 = sum(g[1] for g in grupos)
    ingreso = sum(max(0.0, a - b) for a, b in grupos)
    fila = {
        r.C_CENTRAL: nombre, r.C_TOTAL: ingreso, r.C_ECOSTOS: ingreso, r.C_RE545: 0.0,
        rs.X_EC_E_TOT[0]: e_tot, rs.X_EC_E[0]: e, rs.X_EC_L[0]: l, rs.X_EC_Q[0]: q,
        rs.X_EC_C1[0]: c1, rs.X_EC_C2[0]: c2,
        rs.X_EC_E_TOT[1]: 0.0, rs.X_EC_E[1]: 0.0, rs.X_EC_L[1]: 0.0, rs.X_EC_Q[1]: 0.0,
        rs.X_EC_C1[1]: 0.0, rs.X_EC_C2[1]: 0.0,
        rs.X_EC_I: ingreso, rs.X_RE_I: 0.0,
        rs.X_EC_GRUPOS: [(a, 0.0, b, 0.0) for a, b in grupos],
    }
    fila.update(extra)
    return fila


class TestShapley(unittest.TestCase):

    def test_la_suma_da_exacto_y_reparte_parejo(self):
        efectos = rs.shapley_producto([2.0, 3.0, 5.0], [4.0, 3.0, 10.0])
        self.assertAlmostEqual(sum(efectos), 4 * 3 * 10 - 2 * 3 * 5)
        self.assertEqual(efectos[1], 0.0)                 # no cambio: nada
        # Dos factores que se duplican aportan lo mismo.
        a, b = rs.shapley_producto([1.0, 1.0], [2.0, 2.0])
        self.assertAlmostEqual(a, b)


class TestDescomponer(unittest.TestCase):

    def test_solo_baja_el_cmg_de_c2(self):
        """Sept: C1 100 < C2 120 (monto 0). Oct: el CMg de C2 a la mitad."""
        anterior = fila_ec([(100.0, 120.0)])
        actual = fila_ec([(100.0, 60.0)])
        efectos = rs.descomponer(actual, anterior)
        self.assertEqual(list(efectos), ["E Costos"])
        # C2 baja 60 (Componente 2 = +60) y el max por ciclo recorta 20:
        # antes C1 - C2 era -20 y no se pagaba.
        self.assertAlmostEqual(efectos["E Costos"][rs.CMG], 60.0)
        self.assertAlmostEqual(efectos["E Costos"][rs.REPARTO], -20.0)
        self.assertAlmostEqual(sum(efectos["E Costos"].values()), 40.0)

    def test_sube_el_factor_fd_prorrata(self):
        """Caso GABRIELA: el factor FD x prorrata 0,5 -> 1 (C1 se duplica)."""
        anterior = fila_ec([(100.0, 80.0)], q=400.0)
        actual = fila_ec([(200.0, 80.0)], q=800.0)
        efectos = rs.descomponer(actual, anterior)["E Costos"]
        self.assertAlmostEqual(efectos[rs.FD_PRORRATA], 100.0)
        self.assertNotIn(rs.CMG, efectos)

    def test_el_max_se_aplica_por_ciclo(self):
        """
        Un ciclo positivo y uno negativo: el negativo no "resta". Lo que
        cambia el margen de cada ciclo se reparte, y la suma cierra.
        """
        anterior = fila_ec([(100.0, 50.0), (10.0, 90.0)])
        actual = fila_ec([(100.0, 25.0), (10.0, 45.0)])
        efectos = rs.descomponer(actual, anterior)["E Costos"]
        self.assertAlmostEqual(sum(efectos.values()), 75.0 - 50.0)
        # C2 baja de 140 a 70; el ciclo negativo no paga: el reparto lo recorta.
        self.assertAlmostEqual(efectos[rs.CMG], 70.0)
        self.assertAlmostEqual(efectos[rs.REPARTO], 25.0 - 70.0)

    def test_factor_que_aparece(self):
        """Caso ARENA: sept con C1 = 0 (FD CPF 0); oct con C1."""
        anterior = fila_ec([(0.0, 90.0)], q=0.0)
        actual = fila_ec([(150.0, 90.0)], q=600.0)
        efectos = rs.descomponer(actual, anterior)["E Costos"]
        # C1 aparece entero por el FD; el max por ciclo recorta lo que C2 se come.
        self.assertAlmostEqual(efectos[rs.FD_PRORRATA], 150.0)
        self.assertNotIn(rs.CMG, efectos)
        self.assertAlmostEqual(sum(efectos.values()), 60.0)

    def test_central_nueva_y_que_ya_no_esta(self):
        actual = fila_ec([(100.0, 20.0)])
        self.assertEqual(rs.descomponer(actual, {}), {"E Costos": {rs.NUEVA: 80.0}})
        self.assertEqual(rs.descomponer({}, actual), {"E Costos": {rs.YA_NO: -80.0}})


class TestPerdidaPorFD(unittest.TestCase):
    """Un solo mes: cuanto MAS habria recibido con FD = 1."""

    def test_factor_fd_a_la_mitad(self):
        # Energia ajustada con FD 400; solo con prorrata 800 (FD = 1).
        fila = fila_ec([(100.0, 80.0)], q=400.0, **{rs.X_EC_QP[0]: 800.0})
        fila[rs.X_RE_R0] = 0.0
        perdida = rs.perdida_por_fd(fila)
        self.assertAlmostEqual(perdida["E Costos"], 200.0 - 80.0 - 20.0)
        self.assertEqual(perdida["RE545"], 0.0)          # sin reserva: no aplica

    def test_fd_perfecto_no_pierde(self):
        fila = fila_ec([(100.0, 80.0)], q=800.0, **{rs.X_EC_QP[0]: 800.0})
        fila[rs.X_RE_R0] = 0.0
        self.assertAlmostEqual(rs.perdida_por_fd(fila)["E Costos"], 0.0)

    def test_factor_casi_cero_no_se_extrapola(self):
        """SAE-LILE oct: factor 0,013 -> escalarlo a 1 seria multiplicar por 77."""
        fila = fila_ec([(1.0, 80.0)], q=10.0, **{rs.X_EC_QP[0]: 800.0})
        fila[rs.X_RE_R0] = 0.0
        self.assertTrue(math.isnan(rs.perdida_por_fd(fila)["E Costos"]))

    def test_fd_cero_no_se_puede(self):
        fila = fila_ec([(0.0, 80.0)], q=0.0, **{rs.X_EC_QP[0]: 800.0})
        fila[rs.X_RE_R0] = 0.0
        self.assertTrue(math.isnan(rs.perdida_por_fd(fila)["E Costos"]))


class TestTablas(unittest.TestCase):

    def test_el_mes_suma_la_diferencia_y_hay_principal(self):
        anterior = pd.DataFrame([
            fila_ec([(100.0, 120.0)], nombre="A"), fila_ec([(100.0, 80.0)], q=400.0, nombre="B"),
        ])
        actual = pd.DataFrame([
            fila_ec([(100.0, 60.0)], nombre="A"), fila_ec([(200.0, 80.0)], q=800.0, nombre="B"),
        ])
        mes, por_central, textos = rs.tablas_responsables(actual, anterior)
        total = mes.set_index("Responsable")["Aporte [CLP]"]
        self.assertAlmostEqual(total["Total (= diferencia del mes)"], (40.0 + 120.0) - (0.0 + 20.0))
        self.assertAlmostEqual(total.drop("Total (= diferencia del mes)").sum(), 140.0)
        self.assertEqual(set(por_central[r.C_CENTRAL]), {"A", "B"})


class TestConPlanillas(unittest.TestCase):
    """Con dos planillas escritas por el programa: cierra central por central."""

    def test_cierra_exacto(self):
        from desarrollo.tests.test_reporte import _planilla

        with tempfile.TemporaryDirectory() as carpeta:
            anterior = r.resumen_por_central(
                _planilla(Path(carpeta) / "sep", "11_PAGOS_BESS_2609.xlsx"),
                registrar=lambda *a: None,
            )
            actual = r.resumen_por_central(
                _planilla(Path(carpeta) / "oct", "11_PAGOS_BESS_2610.xlsx", 1.7),
                registrar=lambda *a: None,
            )
        self.assertTrue(actual[rs.X_EC_GRUPOS].map(lambda g: isinstance(g, list)).any())
        for central in actual[r.C_CENTRAL]:
            a = actual.set_index(r.C_CENTRAL).loc[central].to_dict()
            b = anterior.set_index(r.C_CENTRAL).loc[central].to_dict()
            efectos = rs.descomponer(a, b)
            suma = sum(sum(v.values()) for v in efectos.values())
            delta = (a[r.C_TOTAL] or 0.0) - (b[r.C_TOTAL] or 0.0)
            self.assertLess(abs(suma - delta), 1.0, central)


if __name__ == "__main__":
    unittest.main()


class TestTextoSoloCambiaC2(unittest.TestCase):
    """
    Obs. O2026-0052689 (2026-11-18): con el cambio de la C2, la comparacion
    decia "energía asignada / reserva 0,58 → 0,58" y "precio de la curva
    22,7 → 22,7" (los factores de C1, que no se movieron). Ahora muestra
    los de C2.
    """

    def test_muestra_el_factor_de_c2(self):
        from unittest import mock

        from Script.nucleo import responsables as resp

        thetas = {
            "antes": [1.0, 1.0, 1.0, 0.58, 0.30, 26.2, 18.0],
            "ahora": [1.0, 1.0, 1.0, 0.58, 0.15, 26.2, 21.5],
        }
        with mock.patch.object(resp, "_theta_re", lambda fila: thetas[fila["cual"]]):
            uso = resp._texto_movimiento(resp.USO, "RE545", {"cual": "ahora"}, {"cual": "antes"})
            cmg = resp._texto_movimiento(resp.CMG, "RE545", {"cual": "ahora"}, {"cual": "antes"})
        self.assertEqual(uso, "energía contada en C2 / reserva con FD y FMA 0,30 → 0,15")
        self.assertEqual(cmg, "precio de lo contado en C2 18,0 → 21,5 CLP/kWh")


class TestComponentesYReemplazos(unittest.TestCase):
    """
    Pedido del usuario (2026-11-20): los reemplazos del 7_Reasignaciones
    como responsable, y cada aporte con su lado y componente (para filtrar
    E Costos > Componente 1 > FD, CMg). Los totales por responsable no
    cambian: dentro de cada responsable se reparte con Owen.
    """

    def test_el_cmg_de_c2_va_al_componente_2(self):
        anterior = fila_ec([(100.0, 120.0)])
        actual = fila_ec([(100.0, 60.0)])
        detalle = rs.descomponer_detalle(actual, anterior)["E Costos"]
        self.assertEqual(set(detalle), {(rs.CMG, rs.C2), (rs.REPARTO, rs.COMP_REPARTO)})
        self.assertAlmostEqual(detalle[(rs.CMG, rs.C2)], 60.0)

    def test_componente_1_es_el_cambio_de_c1(self):
        """2026-11-20 (ANDES3): Componente 1 tiene que ser el cambio de Total C1."""
        anterior = fila_ec([(100.0, 80.0), (50.0, 10.0)])
        actual = fila_ec([(90.0, 85.0), (70.0, 20.0)], q=640.0)
        detalle = rs.descomponer_detalle(actual, anterior)["E Costos"]
        c1 = sum(v for (_, comp), v in detalle.items() if comp == rs.C1)
        c2 = sum(v for (_, comp), v in detalle.items() if comp == rs.C2)
        self.assertAlmostEqual(c1, 160.0 - 150.0)
        self.assertAlmostEqual(c2, -(105.0 - 90.0))
        self.assertAlmostEqual(sum(detalle.values()), 55.0 - 60.0)

    def test_con_el_cmg_igual_el_precio_no_sale_como_cmg(self):
        """ANDES3: el CMg igual; el precio promedio cambia porque entran otros bloques."""
        anterior = fila_ec([(100.0, 80.0)], **{r.C_CMG: 10.0})
        actual = fila_ec([(120.0, 80.0)], **{r.C_CMG: 10.0})
        detalle = rs.descomponer_detalle(actual, anterior)["E Costos"]
        self.assertEqual(set(detalle), {(rs.FD_PRORRATA, rs.C1)})
        self.assertAlmostEqual(detalle[(rs.FD_PRORRATA, rs.C1)], 20.0)

    def test_reemplazos_exacto_con_la_planilla_sin_reemplazos(self):
        anterior = pd.DataFrame([fila_ec([(100.0, 80.0)], nombre="A")])        # 20
        actual = pd.DataFrame([fila_ec([(130.0, 80.0)], nombre="A")])          # 50
        actual_sin = pd.DataFrame([fila_ec([(110.0, 80.0)], nombre="A")])      # 30
        mes, por_central, _ = rs.tablas_responsables(actual, anterior, actual_sin=actual_sin)
        total = mes.set_index("Responsable")["Aporte [CLP]"]
        self.assertAlmostEqual(total[rs.REEMPLAZOS], 20.0)
        self.assertAlmostEqual(total["Total (= diferencia del mes)"], 30.0)
        fila = por_central[por_central["Responsable"] == rs.REEMPLAZOS].iloc[0]
        # C1 sube 20 por los reemplazos (130 con, 110 sin), C2 igual.
        self.assertEqual((fila["Lado"], fila["Componente"]), ("E Costos", rs.C1))
        self.assertAlmostEqual(fila["Aporte [CLP]"], 20.0)
        arbol = rs.arbol_mes(por_central)
        self.assertEqual(arbol[rs.COLUMNA_ARBOL_MES].iloc[-1], rs.TOTAL_MES)
        self.assertAlmostEqual(arbol["Aporte [CLP]"].iloc[-1], 30.0)
        self.assertIn(rs.REEMPLAZOS, list(arbol[rs.COLUMNA_ARBOL_MES]))

    def test_la_hoja_responsables_es_un_arbol(self):
        """2026-11-20: "que no sea desplegable, que sea jerarquico"."""
        import openpyxl

        anterior = pd.DataFrame([fila_ec([(100.0, 120.0)], nombre="A")])
        actual = pd.DataFrame([fila_ec([(100.0, 60.0)], nombre="A")])
        _, por_central, _ = rs.tablas_responsables(actual, anterior)
        mes = rs.arbol_mes(por_central)
        self.assertEqual(list(mes[rs.COLUMNA_ARBOL_MES]),
                         ["E Costos", rs.C2, rs.CMG, rs.COMP_REPARTO, rs.REPARTO, rs.TOTAL_MES])
        self.assertEqual(mes.attrs["niveles"], [0, 1, 2, 1, 2, -1])
        self.assertAlmostEqual(mes["Aporte [CLP]"].iloc[0], 40.0)
        self.assertAlmostEqual(mes["Aporte [CLP]"].iloc[-1], 40.0)
        central = rs.arbol_por_central(por_central)
        self.assertEqual(list(central[rs.COLUMNA_ARBOL_CENTRAL])[:4],
                         ["A", "E Costos", rs.C2, rs.CMG])
        self.assertEqual(central.attrs["niveles"][:4], [0, 1, 2, 3])
        with tempfile.TemporaryDirectory() as carpeta:
            destino = Path(carpeta) / "c.xlsx"
            r.escribir_reporte(destino, [("Responsables", mes, central, "Mes")])
            ws = openpyxl.load_workbook(destino)["Responsables"]
            self.assertEqual(len(ws.tables), 0)
            self.assertEqual(ws["B4"].value, "E Costos")
            self.assertTrue(ws["B4"].font.bold)
            self.assertEqual(ws["B5"].alignment.indent, 2)
            self.assertEqual(ws["B6"].alignment.indent, 4)
            self.assertFalse(ws["B6"].font.bold)


class TestPesoSobreLoMovido(unittest.TestCase):
    """
    2026-11-20: "cuando la variacion es muy baja pone numeros enormes". El %
    era sobre la diferencia neta: con +44 y -42 que se compensan (neto +2)
    salia 2.200 %. Ahora es sobre lo movido (la suma sin signo).
    """

    def _por_central(self):
        return pd.DataFrame([
            {"Central": "G", "Lado": "E Costos", "Componente": rs.C1, "Responsable": rs.CMG,
             "Aporte [CLP]": 44e6, "Qué se movió": ""},
            {"Central": "G", "Lado": "E Costos", "Componente": rs.AMBAS, "Responsable": rs.ENERGIA,
             "Aporte [CLP]": -42e6, "Qué se movió": ""},
        ])

    def test_los_pesos_quedan_entre_menos_100_y_100(self):
        arbol = rs.arbol_por_central(self._por_central())
        pesos = arbol[rs.PESO].tolist()
        self.assertTrue(all(abs(p) <= 100.0 + 1e-9 for p in pesos))
        fila = arbol.set_index(rs.COLUMNA_ARBOL_CENTRAL)
        self.assertAlmostEqual(fila.loc[rs.CMG, rs.PESO], 44 / 86 * 100)
        self.assertAlmostEqual(fila.loc["G", rs.PESO], 2 / 86 * 100)

    def test_el_texto_avisa_que_se_compensan(self):
        lista = [("E Costos", rs.CMG, 44e6), ("E Costos", rs.ENERGIA, -42e6)]
        principal, desglose = rs._textos_central(lista, 2e6, {}, {})
        self.assertIn("(51 %)", principal)
        self.assertIn("Se compensan", desglose)


class TestTotalesAlLadoDelNombre(unittest.TestCase):
    """2026-11-20: "que se vea el total al lado del nombre, no como columna"."""

    def test_la_central_y_el_mes_traen_lo_que_reciben(self):
        anterior = pd.DataFrame([fila_ec([(100.0, 120.0)], nombre="A")])
        actual = pd.DataFrame([fila_ec([(100.0, 60.0)], nombre="A")])
        _, por_central, _ = rs.tablas_responsables(actual, anterior)
        central = rs.arbol_por_central(por_central, {"A": (0.0, 40.0)})
        self.assertEqual(central[rs.COLUMNA_ARBOL_CENTRAL].iloc[0],
                         f"A  ·  recibe {r.plata_corta(0.0)} → {r.plata_corta(40.0)} (+{r.plata_corta(40.0)})")
        mes = rs.arbol_mes(por_central, {"total": (0.0, 40.0), "E Costos": (0.0, 40.0)})
        self.assertTrue(mes[rs.COLUMNA_ARBOL_MES].iloc[0].startswith("E Costos  ·  recibe"))
        self.assertTrue(mes[rs.COLUMNA_ARBOL_MES].iloc[-1].startswith(rs.TOTAL_MES + "  ·  recibe"))
        # Sin totales, el nombre solo (como antes).
        self.assertEqual(rs.arbol_por_central(por_central)[rs.COLUMNA_ARBOL_CENTRAL].iloc[0], "A")


class TestTextoReparto(unittest.TestCase):
    """
    2026-11-20 (ANDES3): C1 y C2 cambian pero lo que recibe no. El texto del
    reparto dice cuanto de C1 - C2 cambio en ventanas que pagan y cuanto en
    las que no.
    """

    def test_lo_que_cambia_en_ventanas_que_no_pagan(self):
        anterior = {rs.X_RE_MARGEN: {1: (100.0, 100.0), 2: (-50.0, 0.0)}}
        actual = {rs.X_RE_MARGEN: {1: (100.0, 100.0), 2: (-20.0, 0.0)}}
        texto = rs._texto_reparto("RE545", actual, anterior)
        self.assertIn("C1 − C2 cambia +", texto)
        self.assertIn("en ventanas que no pagan", texto)
        self.assertNotIn("que pagan;", texto)
