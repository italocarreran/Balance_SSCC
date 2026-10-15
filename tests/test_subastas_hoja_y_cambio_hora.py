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
