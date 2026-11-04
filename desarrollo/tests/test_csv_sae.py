# -*- coding: utf-8 -*-
"""
SAE_AAMM.csv (pedido del usuario 2026-11-04: "un boton final que genere
el siguiente documento"), igual que la macro C_Exporta_pagos.Generar_CSV_SAE
de 11_PAGOS_BESS, y los nombres nuevos del caso: 11_PAGOS_BESS_AAMM.xlsx y
la carpeta Log/ con el control, el registro y los sellos.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo
from Script.nucleo.csv_sae import construir_csv_sae
from Script.nucleo.sellos import leer_sellos, ruta_sellos
from desarrollo.tests.test_pasos_sueltos import armar_caso, correr_pagos, dejar_prorrata


def hoja(filas):
    return pd.DataFrame(filas, columns=[
        "Mes", "Dia", "Hora", "Minuto", "Bloque horario", "Configuracion",
        "Monto a compensar",
    ])


class TestConstruirCsv(unittest.TestCase):

    def test_como_la_macro(self):
        ecostos = hoja([
            (9, 1, 1, 0, 1, "SAE-B", 10.0),
            (9, 1, 1, 15, 2, "SAE-B", 0.0),          # en cero: no va
            (9, 1, 1, 0, 1, "SAE-A", 1.005),
            (9, 5, 24, 45, 480, "SAE-A", 2.0),
            # 6-09: no existe la hora 0; el bloque 481 es la 01:00
            (9, 6, 2, 0, 481, "SAE-A", 3.0),
        ])
        re545 = hoja([
            (9, 1, 1, 0, 1, "SAE-B", 5.25),          # se suma al de E Costos
            (9, 1, 1, 15, 2, "SAE-A", -4.0),
        ])
        texto, totales, verificaciones = construir_csv_sae(ecostos, re545, "2609")
        self.assertTrue(texto.endswith("\r\n"))
        self.assertEqual(texto.split("\r\n")[:-1], [
            "Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min",
            "01-09-2026 00:00,SAE-A,,1.01,BESS,1",
            "01-09-2026 00:00,SAE-B,,15.25,BESS,1",
            "01-09-2026 00:15,SAE-A,,-4.00,BESS,2",
            "05-09-2026 23:45,SAE-A,,2.00,BESS,480",
            "06-09-2026 01:00,SAE-A,,3.00,BESS,481",
        ])
        self.assertAlmostEqual(totales["SAE-B"], 15.25)
        self.assertTrue(any("Cambio hora +1" in v for v in verificaciones))

    def test_todo_en_cero_no_hay_csv(self):
        texto, totales, _ = construir_csv_sae(
            hoja([(9, 1, 1, 0, 1, "SAE-A", 0.0)]), hoja([]), "2609"
        )
        self.assertIsNone(texto)
        self.assertEqual(totales, {})


def _libro_con_hojas(ruta, ecostos, re545):
    """Las dos hojas como las deja el programa: fila 1 vacia, titulo del
    cuadro en la 2, encabezado (con unidades) en la 3 y columna A de margen."""

    with pd.ExcelWriter(ruta) as writer:
        for nombre, df in ((nucleo.HOJA_CALCULO_ECOSTOS, ecostos),
                           (nucleo.HOJA_CALCULO_RE545, re545)):
            df = df.rename(columns={"Monto a compensar": "Monto a compensar [CLP]"})
            pd.DataFrame([["Cuadro N° 1 — Informacion"]]).to_excel(
                writer, sheet_name=nombre, startrow=1, startcol=1,
                header=False, index=False,
            )
            df.to_excel(writer, sheet_name=nombre, startrow=2, startcol=1,
                        index=False)


class TestCasoCompleto(unittest.TestCase):

    def test_boton_final_lee_la_planilla(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            _libro_con_hojas(
                base / nucleo.nombre_salida("2609"),
                hoja([(9, 1, 1, 0, 1, "SAE-A", 100.0),
                      (9, 1, 1, 15, 2, "SAE-A", 0.0)]),
                hoja([(9, 1, 1, 0, 1, "SAE-A", 23.456)]),
            )
            log = []
            ruta = nucleo.generar_csv_sae(base, "2609", registrar=log.append)

            self.assertEqual(ruta, base / "SAE_2609.csv")
            crudo = ruta.read_bytes()
            self.assertTrue(crudo.startswith(b"\xef\xbb\xbf"))
            self.assertEqual(crudo.decode("utf-8-sig"), (
                "Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min\r\n"
                "01-09-2026 00:00,SAE-A,,123.46,BESS,1\r\n"
            ))
            self.assertTrue(any("TOTAL GENERAL: 123.46" in l for l in log))

            _, filas = nucleo.revisar_estructura(base, "2609")
            ids = {f["id"]: f for f in filas}
            self.assertEqual(ids["sae_csv"]["etiqueta"], "SAE_2609.csv")
            self.assertEqual(ids["sae_csv"]["estado"], "ok")
            self.assertEqual(ids["salida"]["etiqueta"], "11_PAGOS_BESS_2609.xlsx")

    def test_corrida_deja_lo_suyo_en_log(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)
            dejar_prorrata(base)
            correr_pagos(base, {"ecostos", "re545"}, [])
            self.assertTrue((base / "Log" / nucleo.ARCHIVO_CONTROL).is_file())
            # Los sellos (json) van a __config__/2026/07 Julio/Balance BESS/
            sellos = ruta_sellos(base / nucleo.nombre_salida("2607"))
            self.assertTrue(sellos.is_file())
            self.assertEqual(sellos.parent.parent.parent.name, "07 Julio")
            self.assertFalse(list(base.rglob("*.json")))
            self.assertFalse((base / nucleo.ARCHIVO_CONTROL).exists())
            # Este caso chico no tiene pagos: no se escribe un CSV vacio.
            with self.assertRaisesRegex(nucleo.ErrorEntrada, "en cero"):
                nucleo.generar_csv_sae(base, "2607", registrar=lambda *_: None)
            self.assertFalse((base / "SAE_2607.csv").exists())


class TestCasoDeAntes(unittest.TestCase):

    def test_se_renombra_y_se_ordena_en_log(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "Balance_BESS.xlsx").write_bytes(b"libro")
            (base / ".sellos_Balance_BESS.json").write_text("{}", encoding="utf-8")
            (base / "Control_corrida.xlsx").write_bytes(b"control")
            (base / "Registro_ultima_corrida.txt").write_text("r", encoding="utf-8")

            rutas = nucleo.resolver_rutas(base, "2609")

            self.assertEqual(rutas["salida"], base / "11_PAGOS_BESS_2609.xlsx")
            self.assertEqual(rutas["salida"].read_bytes(), b"libro")
            self.assertEqual(
                sorted(p.name for p in base.iterdir()),
                ["11_PAGOS_BESS_2609.xlsx", "Log"],
            )
            self.assertEqual(sorted(p.name for p in (base / "Log").iterdir()), [
                "Control_corrida.xlsx", "Registro_ultima_corrida.txt",
            ])
            self.assertTrue(ruta_sellos(rutas["salida"]).is_file())

    def test_los_sellos_que_quedaron_en_log_pasan_a_config(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "11_PAGOS_BESS_2609.xlsx").write_bytes(b"x")
            (base / "Log").mkdir()
            (base / "Log" / "sellos_11_PAGOS_BESS_2609.json").write_text(
                '{"Resumen": 1.0}', encoding="utf-8")
            rutas = nucleo.resolver_rutas(base, "2609")
            self.assertFalse((base / "Log").exists())  # quedo vacia
            self.assertEqual(leer_sellos(rutas["salida"]), {"Resumen": 1.0})

    def test_sin_periodo_usa_la_planilla_que_hay(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "11_PAGOS_BESS_2609.xlsx").write_bytes(b"x")
            self.assertEqual(
                nucleo.resolver_rutas(base)["salida"].name,
                "11_PAGOS_BESS_2609.xlsx",
            )


if __name__ == "__main__":
    unittest.main()
