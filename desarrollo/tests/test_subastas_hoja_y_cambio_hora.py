# -*- coding: utf-8 -*-
"""Hoja Subastas sin auxiliares y Hora_mes en el cambio de hora.

- la hoja no trae Control, Sub_Baj, Clave horaria ni Ciclo (antes B, C,
  L y M): se arman de Concepto y de Configuración + Dia + Hora_dia al
  leerla, y quedan en el orden de siempre;
- HORA de los Access es el periodo 1..24: en septiembre el 6-09 no trae
  la 1 (desde ahi -1) y en abril el sabado trae la 25 (desde el domingo
  +1). Hora_mes queda correlativo, igual que el de Medidores.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo.cambio_hora import hora_mes_correlativa
from Script.nucleo.hojas_entrada import (
    NOMBRES_SUBASTAS, reponer_auxiliares_subastas, subastas_para_hoja,
)


def _subastas():
    return pd.DataFrame({
        "Concepto": ["CSF(+)", "CPF(-)"],
        "Control": ["CSF", "CPF"],
        "Sub_Baj": ["SUBIDA", "BAJADA"],
        "Fecha": pd.to_datetime(["2026-07-01", "2026-07-01"]),
        "Año": [2026, 2026], "Mes": [7, 7], "Dia": [1, 1],
        "Hora_dia": [3, 4], "Hora_mes": [3, 4],
        "Configuración": ["BESS UNO", "SAE DOS"],
        "Propietario": ["COLBUN", "ENGIE"],
        "Clave horaria": ["BESS UNO13", "SAE DOS14"],
        "Ciclo": [pd.NA, pd.NA],
        "Energía SSCC": [1.5, 2.5],
    })


class TestHojaSubastas(unittest.TestCase):
    def test_se_escribe_sin_auxiliares_y_se_reponen_al_leer(self):
        original = _subastas()
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "Consolidado.xlsx"
            nucleo.escribir_salida(
                pd.DataFrame({"Mes": [7]}), ruta, [], [],
                df_subastas=original,
            )
            encabezado = [
                c.value for c in openpyxl.load_workbook(ruta)["Subastas"][1]
            ]
            leida = pd.read_excel(ruta, sheet_name="Subastas")

        for auxiliar in ("Control", "Sub_Baj", "Clave horaria", "Ciclo"):
            self.assertNotIn(auxiliar, encabezado)
        self.assertEqual(encabezado[0], "Concepto")

        repuesta = reponer_auxiliares_subastas(leida)
        self.assertEqual(list(repuesta.columns), list(NOMBRES_SUBASTAS.values())[:14])
        for columna in ("Control", "Sub_Baj", "Clave horaria"):
            self.assertEqual(
                repuesta[columna].tolist(), original[columna].tolist(), columna
            )
        self.assertTrue(repuesta["Ciclo"].isna().all())

    def test_una_hoja_de_antes_queda_igual(self):
        original = _subastas()
        repuesta = reponer_auxiliares_subastas(original)
        pd.testing.assert_frame_equal(repuesta, original)

    def test_para_hoja(self):
        self.assertEqual(
            list(subastas_para_hoja(_subastas()).columns),
            ["Concepto", "Fecha", "Año", "Mes", "Dia", "Hora_dia",
             "Hora_mes", "Configuración", "Propietario", "Energía SSCC"],
        )


class TestHoraMesSubastas(unittest.TestCase):
    def _hora_mes(self, filas):
        fechas = pd.Series([pd.Timestamp(f) for f, _ in filas])
        horas = pd.Series([h for _, h in filas])
        return [None if pd.isna(v) else int(v)
                for v in hora_mes_correlativa(fechas, horas)]

    def test_septiembre_sin_la_hora_1(self):
        self.assertEqual(
            self._hora_mes([("2026-09-05", 24), ("2026-09-06", 1),
                            ("2026-09-06", 2), ("2026-09-07", 1)]),
            [120, None, 121, 144],
        )

    def test_abril_con_la_hora_25(self):
        self.assertEqual(
            self._hora_mes([("2026-04-04", 24), ("2026-04-04", 25),
                            ("2026-04-05", 1), ("2026-04-30", 24)]),
            [96, 97, 98, 30 * 24 + 1],
        )

    def test_mes_normal(self):
        self.assertEqual(
            self._hora_mes([("2026-07-01", 1), ("2026-07-31", 24)]),
            [1, 744],
        )


if __name__ == "__main__":
    unittest.main()


class TestSubastasConsolidadas(unittest.TestCase):
    """Los Access se consolidan afuera (Subastas/subastas_AAMM.xlsx, boton
    'Consolidar') y la hoja 'Subastas' se lleva de ese archivo."""

    def _crudo(self):
        return pd.DataFrame({
            "CONFIGURACIÓN": ["BESS UNO", "BESS UNO", "TERMICA X"],
            "SERVICIO": ["CSF(+)", "CPF(-)", "CSF(+)"],
            "AÑO": [2026, 2026, 2026], "MES": [9, 9, 9],
            "DIA": [6, 6, 6], "HORA": [2, 3, 2], "BANDA": [1, 1, 1],
            "CANTIDAD MW": [5.0, 3.0, 9.0],
            "PRECIO USD/MW": [1.0, 1.0, 1.0],
            "CANTIDAD PONDERADA MW": [5.0, 3.0, 9.0],
            "Hora_PID": [0, 0, 0],
        })

    def test_consolidar_escribe_el_archivo_y_la_hoja_sale_de_ahi(self):
        from unittest import mock
        from Script.nucleo import traer
        from Script.Subastas import Ofertas_Adjudicadas as ofertas_adj

        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for sub in nucleo.SUBCARPETAS_CASO:
                (base / sub).mkdir(parents=True, exist_ok=True)
            resumen = {"archivos_leidos": 2, "dias_sin_archivo": []}
            with mock.patch.object(
                ofertas_adj, "construir_crudo",
                lambda *a, **k: (self._crudo(), resumen),
            ):
                destino = nucleo.consolidar_subastas(
                    base, "2609", registrar=lambda *_: None
                )
            self.assertEqual(destino.name, "subastas_2609.xlsx")
            self.assertEqual(destino.parent, base / "Subastas")
            self.assertEqual(len(pd.read_excel(destino)), 3)

            hoja = nucleo.construir_subastas_desde_consolidado(
                base / "Subastas", "2609", registrar=lambda *_: None
            )
            directo = nucleo.construir_subastas_desde_crudo(
                self._crudo(), "2609", registrar=lambda *_: None
            )
            pd.testing.assert_frame_equal(
                hoja.reset_index(drop=True), directo.reset_index(drop=True),
                check_dtype=False,
            )
            self.assertEqual(hoja["Configuración"].tolist(), ["BESS UNO"] * 2)
            # 6-09 hora 2 = Hora_mes 121 (cambio de hora)
            self.assertEqual(int(hoja["Hora_mes"].iloc[0]), 121)

            # La fila del arbol queda "ok" con el archivo
            _, filas = nucleo.revisar_estructura(base, "2609")
            fila = next(f for f in filas if f["id"] == "subastas_xlsx")
            self.assertEqual(fila["estado"], "ok")
            self.assertEqual(fila["etiqueta"], "subastas_2609.xlsx")

    def test_sin_consolidado_el_error_dice_que_hay_que_consolidar(self):
        with tempfile.TemporaryDirectory() as carpeta:
            with self.assertRaises(nucleo.ErrorEntrada) as caso:
                nucleo.construir_subastas_desde_consolidado(
                    Path(carpeta), "2609", registrar=lambda *_: None
                )
        self.assertIn("Consolidar", str(caso.exception))
