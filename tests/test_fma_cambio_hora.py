# -*- coding: utf-8 -*-
"""FMA CPF y CTF en el cambio de hora.

Comprobado con los reportes reales del DCO de 2026 (huella de las
columnas de frecuencia T-/T+): el tabla_resumen sin sufijo va en el
desfase vigente al empezar el dia, y el `_UTC-4` del 4-04 es el mismo dia
corrido una hora, con la segunda pasada de las 23:00 en su ultima fila.
Aca se arman reportes sinteticos con esa forma.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script.Fd import Indices_FMA as fma
from Script.Subastas import Fma as fma_subastas


def _reporte(ruta, marca):
    """tabla_resumen con una unidad: (a1) de la fila k = marca + k."""

    columnas = fma.NOMBRES_CPF[:29]
    filas = []
    for k in range(24):
        fila = {c: 0 for c in columnas}
        fila["Hora Local"] = k
        fila["(a1)"] = marca + k
        filas.append(fila)
    with pd.ExcelWriter(ruta) as writer:
        pd.DataFrame({"Unidad": ["X"]}).to_excel(writer, sheet_name="Resumen")
        pd.DataFrame(filas, columns=columnas).to_excel(
            writer, sheet_name="BESS UNO", index=False, startrow=4
        )


class TestFmaCpfCambioDeHora(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.carpeta = Path(self.tmp.name)
        fma.limpiar_indice_cpf()

    def tearDown(self):
        fma.limpiar_indice_cpf()
        self.tmp.cleanup()

    def _cpf(self, aamm, archivos):
        for nombre, marca in archivos:
            sub = self.carpeta / f"Reporte diario {nombre}"
            sub.mkdir(exist_ok=True)
            _reporte(sub / f"{nombre}.xlsx", marca)
        df = fma.construir_fma_cpf(self.carpeta, aamm, registrar=lambda *_: None)
        return df.set_index("Hora")["Hace CPF [hrs]"]

    def test_abril_hora_25_sale_del_utc_4(self):
        hace = self._cpf("2604", [
            ("tabla_resumen_4_4_2026", 0),
            ("tabla_resumen_4_4_2026_UTC-4", 100),
        ])
        self.assertEqual(list(hace.index), list(range(1, 26)))
        # 1..24 del sin sufijo (fila 0..23); 25 de la fila 23 del UTC-4
        self.assertEqual(hace.loc[1], 0)
        self.assertEqual(hace.loc[24], 23)
        self.assertEqual(hace.loc[25], 123)

    def test_septiembre_sin_la_hora_1(self):
        hace = self._cpf("2609", [("tabla_resumen_6_9_2026", 0)])
        # UTC-4 al empezar el dia: fila 0 = 01:00 local = hora 2; la
        # fila 23 ya es el 7-09 y se descarta.
        self.assertEqual(list(hace.index), list(range(2, 25)))
        self.assertEqual(hace.loc[2], 0)
        self.assertEqual(hace.loc[24], 22)

    def test_mes_normal_igual_que_antes(self):
        hace = self._cpf("2607", [("tabla_resumen_15_7_2026", 0)])
        self.assertEqual(list(hace.index), list(range(1, 25)))
        self.assertEqual(hace.loc[24], 23)


class TestFmaCtfCambioDeHora(unittest.TestCase):
    def test_la_segunda_pasada_de_abril_va_a_la_25(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "CTF_202604.csv"
            pd.DataFrame({
                "t0": ["2026-04-04 23:10:00-03:00", "2026-04-04 23:10:00-04:00",
                       "2026-09-06 01:10:00-03:00"],
                "tfin": ["2026-04-04 23:40:00-03:00", "2026-04-04 23:40:00-04:00",
                         "2026-09-06 01:40:00-03:00"],
                "unidad": ["U1", "U1", "U1"],
                "Configuracion": ["C1", "C1", "C1"],
                "variacion": ["1,0", "1,0", "-1,0"],
            }).to_csv(ruta, sep=";", index=False)
            df = fma.construir_fma_ctf(ruta, registrar=lambda *_: None)
            salida = Path(carpeta) / "fma_cft_2604.xlsx"
            df.to_excel(salida, index=False)
            por_unidad, _ = fma_subastas.tablas_ctf(salida, registrar=lambda *_: None)

        self.assertEqual(df[fma.COLUMNA_HORA_CTF].tolist(), [24, 25, 2])
        # sin la zona, las dos quedaban en la 24 y se sumaban
        abril = por_unidad[por_unidad["Dia"] == 4].set_index("Hora")
        self.assertEqual(sorted(abril.index), [24, 25])
        self.assertAlmostEqual(abril.loc[24, "FMA_CTF_mas"], 0.5)
        self.assertAlmostEqual(abril.loc[25, "FMA_CTF_mas"], 0.5)


if __name__ == "__main__":
    unittest.main()
