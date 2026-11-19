# -*- coding: utf-8 -*-
"""
Reemplazos de SSCC (7_Reasignaciones_SSCC_AAMM) en Subastas.

Pedido del usuario (2026-11-19, obs. de SAE-CRCA-PFV-ELENA en 2610): las
horas en que una BESS es REEMPLAZANTE de otra unidad (asignacion directa
del Coordinador) no estaban en Subastas y quedaban con Adj SSCC = 0; y
las horas en que una BESS adjudicada fue REEMPLAZADA seguian contando.
El 7_ va en Subastas/: el Cuadro N° 1 lleva la bandera "Reemplazado" y
el Cuadro N° 2 los reemplazos con BESS. El calculo usa lo adjudicado sin
lo reemplazado mas lo que la BESS presto como reemplazante.
"""

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

from Script.nucleo import reasignaciones as R
from Script.nucleo.columnas_compartidas import calcular_l
from Script.nucleo.ecostos_prorratas import construir_prorrata_sscc
from Script.nucleo.hojas_entrada import reponer_auxiliares_subastas

DICCIONARIO = pd.DataFrame([
    [None] * 4,
    ["Balance_BESS", "FD", "Subastas", "Reasignaciones"],
    ["SAE-CRCA-PFV-ELENA", "BESS ELENA", "SAE-CRCA-PFV-ELENA", None],
    ["SAE-DEL-DESIERTO", "BESS DD", "SAE-DEL-DESIERTO", None],
    ["SAE-OTRA", "BESS OTRA", "SAE-OTRA", "COORD-OTRA"],
])

ENCABEZADO_7 = [
    "Fecha", "Inicio Periodo", "Fin Periodo", "Instrucción - SSCC Requerida",
    "ID Configuración", "Central / Subestación (PRS)", "BARRA CT",
    "Central-Unidad", "Configuración / Paño (PRS-EV)", "Disponibilidad",
    "BAJA", "SUBE", "(kV, MVAr, MW)", "Motivo", "Comentario", "Estado",
    "Fecha de Edición", None, "Reemplazado 1", "Reemplazado 2",
]


def _fila_7(dia, ini, fin, servicio, configuracion, mw, *reemplazados):
    d = datetime(2026, 10, dia)
    h = lambda t: d.replace(hour=int(t[:2]), minute=int(t[3:]))  # noqa: E731
    fila = [d, h(ini), h(fin), servicio, 1, configuracion, None,
            configuracion, configuracion, mw, None, None, "MW",
            "Asignación directa", "Reemplaza a ...", "CERRADA", d, None]
    return fila + list(reemplazados) + [None] * (2 - len(reemplazados))


def _escribir_7(carpeta, filas, nombre="7_Reasignaciones_SSCC_2610_pre.xlsm"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = R.HOJA_7
    ws.append(ENCABEZADO_7)
    for fila in filas:
        ws.append(fila)
    ruta = Path(carpeta) / nombre
    wb.save(ruta)
    return ruta


def _subastas(*filas):
    """Cuadro N° 1 minimo: (Concepto, Dia, Hora_dia, Configuración, MW)."""
    datos = []
    for concepto, dia, hora, conf, mw in filas:
        datos.append({
            "Concepto": concepto, "Fecha": pd.Timestamp(2026, 10, dia),
            "Año": 2026, "Mes": 10, "Dia": dia, "Hora_dia": hora,
            "Hora_mes": (dia - 1) * 24 + hora, "Configuración": conf,
            "Energía SSCC": mw,
        })
    return pd.DataFrame(datos)


class TestReasignaciones(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        # ELENA (como "PFV-ELENA", el nombre del Coordinador) reemplaza a
        # un parque eolico de 09:07 a 10:00 y a DEL-DESIERTO (BESS) de
        # 06:45 a 07:00; PEHUENCHE reemplaza a ELENA de 23:00 a 23:59; una
        # unidad que se reemplaza a si misma no cuenta.
        self.ruta = _escribir_7(self.tmp.name, [
            _fila_7(17, "09:07", "10:00", "CSF(-)", "PFV-ELENA", 75, "PE-SANJUAN"),
            _fila_7(15, "06:45", "07:00", "CSF(+)", "PFV-ELENA", 80, "SAE-DEL-DESIERTO"),
            _fila_7(17, "23:00", "23:59", "CSF(-)", "PEHUENCHE", 76, "SAE-CRCA-PFV-ELENA"),
            _fila_7(29, "17:00", "18:00", "CSF(+)", "PFV-ELENA", 10, "PFV-ELENA"),
            _fila_7(1, "01:00", "02:00", "CTF(+)", "TOCOPILLA U16", 16, "RUCATAYO"),
        ])
        self.sub = _subastas(
            ("CSF(+)", 15, 7, "SAE-DEL-DESIERTO", 50.0),
            ("CSF(-)", 17, 24, "SAE-CRCA-PFV-ELENA", 75.5),
            ("CSF(+)", 17, 24, "SAE-CRCA-PFV-ELENA", 20.0),
            ("CSF(-)", 17, 10, "SAE-CRCA-PFV-ELENA", 5.0),
        )
        self.avisos = []
        self.c2 = R.reemplazos_desde_archivo(
            self.ruta, "2610", DICCIONARIO, self.sub,
            registrar=self.avisos.append,
        )

    def test_busca_el_archivo_en_subastas(self):
        self.assertEqual(R.buscar_archivo_reasignaciones(self.tmp.name, "2610"), self.ruta)
        self.assertIsNone(R.buscar_archivo_reasignaciones(self.tmp.name, "2609"))

    def test_cuadro_2_solo_con_bess_y_por_hora(self):
        c2 = self.c2.set_index(["Configuración", "Hora_mes", "Concepto"])
        # 09:07-10:00 es la hora 10 (la que termina), 53 minutos,
        # MW ponderado 75*53/60.
        fila = c2.loc[("SAE-CRCA-PFV-ELENA", 16 * 24 + 10, "CSF(-)")]
        self.assertEqual(fila[R.COLUMNA_MINUTOS], 53)
        self.assertAlmostEqual(fila["Energía SSCC"], 75 * 53 / 60)
        self.assertEqual(fila[R.COLUMNA_ROL_BESS], R.ROL_REEMPLAZANTE)
        # BESS que reemplaza a BESS.
        self.assertEqual(
            c2.loc[("SAE-CRCA-PFV-ELENA", 14 * 24 + 7, "CSF(+)"), R.COLUMNA_ROL_BESS],
            R.ROL_AMBAS,
        )
        # 23:00-23:59 es la hora 24 completa; la reemplazante no es BESS.
        fila = c2.loc[("PEHUENCHE", 16 * 24 + 24, "CSF(-)")]
        self.assertEqual(fila[R.COLUMNA_MINUTOS], 60)
        self.assertEqual(fila[R.COLUMNA_REEMPLAZA_A], "SAE-CRCA-PFV-ELENA")
        self.assertEqual(fila[R.COLUMNA_ROL_BESS], R.ROL_REEMPLAZADA)
        # Sin BESS (TOCOPILLA U16 por RUCATAYO) y la que se reemplaza a
        # si misma no entran.
        self.assertEqual(len(self.c2), 3)
        self.assertTrue(any("a si misma" in a for a in self.avisos))
        self.assertTrue(any("PFV-ELENA" in a and "SAE-CRCA-PFV-ELENA" in a for a in self.avisos))

    def test_columna_reasignaciones_del_diccionario(self):
        nombres = R.NombresBess(DICCIONARIO, [], registrar=lambda *_: None)
        self.assertEqual(nombres.resolver("COORD-OTRA"), "SAE-OTRA")
        self.assertTrue(nombres.es_bess("SAE-OTRA"))
        self.assertEqual(nombres.resolver("PE-SANJUAN"), "PE-SANJUAN")
        self.assertFalse(nombres.es_bess("PE-SANJUAN"))

    def test_bandera_y_subastas_para_calculo(self):
        marcada = R.marcar_reemplazadas(self.sub, self.c2, registrar=lambda *_: None)
        self.assertEqual(marcada[R.COLUMNA_REEMPLAZADO].tolist(), [1, 1, 0, 0])

        ef = R.subastas_para_calculo(marcada, self.c2, registrar=lambda *_: None)
        claves = set(zip(ef["Configuración"], ef["Hora_mes"], ef["Concepto"]))
        # Fuera: ELENA CSF(-) h24 del 17 (la reemplazo PEHUENCHE) y
        # DEL-DESIERTO el 15 h7 (la reemplazo ELENA).
        self.assertNotIn(("SAE-DEL-DESIERTO", 14 * 24 + 7, "CSF(+)"), claves)
        self.assertNotIn(("SAE-CRCA-PFV-ELENA", 16 * 24 + 24, "CSF(-)"), claves)
        # Dentro: ELENA como reemplazante el 15 h7 (nueva fila).
        self.assertIn(("SAE-CRCA-PFV-ELENA", 14 * 24 + 7, "CSF(+)"), claves)
        # La del 17 h10 ya estaba adjudicada: se suma el MW, no se repite.
        h10 = ef[(ef["Configuración"] == "SAE-CRCA-PFV-ELENA") & (ef["Hora_mes"] == 16 * 24 + 10)]
        self.assertEqual(len(h10), 1)
        self.assertAlmostEqual(h10["Energía SSCC"].iloc[0], 5 + 75 * 53 / 60)
        # No se cuenta una PEHUENCHE (no es BESS).
        self.assertNotIn("PEHUENCHE", set(ef["Configuración"]))

        # L de E Costos y la prorrata ven la hora nueva.
        ef = reponer_auxiliares_subastas(ef)
        ecostos = pd.DataFrame({
            "clave": ["SAE-CRCA-PFV-ELENA"] * 2, "Mes": [10, 10],
            "Dia": [15, 17], "Hora": [7, 24],
        })
        self.assertEqual(calcular_l(ecostos, ef).tolist(), [1, 1])
        prorrata = construir_prorrata_sscc(ef).set_index(["Configuración", "Hora_mes"])
        self.assertEqual(prorrata.loc[("SAE-CRCA-PFV-ELENA", 16 * 24 + 24), "CSF(+)"], 1.0)

    def test_sin_reemplazos_queda_igual(self):
        self.assertIs(R.subastas_para_calculo(self.sub, None), self.sub)

    def test_leer_los_dos_cuadros_de_la_hoja(self):
        marcada = R.marcar_reemplazadas(self.sub, self.c2, registrar=lambda *_: None)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Subastas"
        ancho = len(marcada.columns)
        ws.cell(2, 2, "Cuadro N° 1 — " + R.TITULO_CUADRO_1)
        ws.cell(2, 3 + ancho, "Cuadro N° 2 — " + R.TITULO_CUADRO_2)
        for j, nombre in enumerate(marcada.columns):
            ws.cell(3, 2 + j, "Energía SSCC [MW]" if nombre == "Energía SSCC" else nombre)
        for j, nombre in enumerate(self.c2.columns):
            ws.cell(3, 3 + ancho + j, "Energía SSCC [MW]" if nombre == "Energía SSCC" else nombre)
        for i, fila in enumerate(marcada.itertuples(index=False)):
            for j, valor in enumerate(fila):
                ws.cell(4 + i, 2 + j, valor.to_pydatetime() if isinstance(valor, pd.Timestamp) else valor)
        for i, fila in enumerate(self.c2.itertuples(index=False)):
            for j, valor in enumerate(fila):
                ws.cell(4 + i, 3 + ancho + j, valor.to_pydatetime() if isinstance(valor, pd.Timestamp) else valor)
        ruta = Path(self.tmp.name) / "planilla.xlsx"
        wb.save(ruta)

        c1, c2 = R.leer_subastas_consolidado(ruta, "Subastas")
        self.assertEqual(list(c1.columns), list(marcada.columns))
        self.assertEqual(len(c1), 4)
        self.assertEqual(c1[R.COLUMNA_REEMPLAZADO].tolist(), [1, 1, 0, 0])
        self.assertEqual(list(c2.columns), list(R.COLUMNAS_REEMPLAZOS))
        self.assertEqual(len(c2), 3)
        self.assertEqual(str(c2["Hora_mes"].dtype), "Int64")

    def test_hoja_de_antes_una_sola_tabla(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Subastas"
        ws.append(list(self.sub.columns))
        for fila in self.sub.itertuples(index=False):
            ws.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in fila])
        ruta = Path(self.tmp.name) / "vieja.xlsx"
        wb.save(ruta)
        c1, c2 = R.leer_subastas_consolidado(ruta, "Subastas")
        self.assertIsNone(c2)
        self.assertEqual(len(c1), 4)
        self.assertIs(R.subastas_para_calculo(c1, c2), c1)


if __name__ == "__main__":
    unittest.main()
