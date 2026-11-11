# -*- coding: utf-8 -*-
"""Los Access del periodo salen de UNA lectura de 'DB subastas/'.

Plan de la ventana (desarrollo/PLAN_OPTIMIZACION_VENTANAS_BESS.md, Fase
1): antes se preguntaba is_file() por cada uno de los 24 x dias nombres
posibles del mes en cada repintado; sobre una unidad de red eso eran
segundos de ventana congelada. El resultado tiene que ser el mismo.
"""

import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from Script.Subastas import Ofertas_Adjudicadas as ofertas_adj


class TestAccdbPresentes(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.subastas = Path(self.tmp.name) / "Subastas"
        self.db = ofertas_adj.asegurar_carpeta_db(self.subastas)
        nombres = {
            (dia, hora): nombre
            for dia, hora, nombre in ofertas_adj.nombres_del_periodo("2609")
        }
        # Desordenados a proposito: el resultado sale en orden (dia, hora).
        self.esperados = [(3, 14), (1, 0), (1, 5), (30, 23)]
        for i, clave in enumerate(self.esperados):
            ruta = self.db / nombres[clave]
            ruta.write_bytes(b"x")
            os.utime(ruta, (1_000_000 + i, 1_000_000 + i))
        self.esperados.sort()
        # Lo que no cuenta: otro mes, otro tipo y una carpeta con nombre
        # de Access del periodo.
        (self.db / ofertas_adj.nombre_accdb("2610", 1, 0)).write_bytes(b"x")
        (self.db / "notas.txt").write_text("x")
        (self.db / nombres[(2, 0)]).mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def test_devuelve_solo_los_del_periodo_en_orden(self):
        presentes = ofertas_adj.accdb_presentes(self.subastas, "2609")
        self.assertEqual([(d, h) for d, h, _ in presentes], self.esperados)
        for dia, hora, ruta in presentes:
            self.assertEqual(ruta, self.db / ofertas_adj.nombre_accdb("2609", dia, hora))
            self.assertTrue(ruta.is_file())

    def test_carpeta_que_no_existe(self):
        otra = Path(self.tmp.name) / "no existe"
        self.assertEqual(ofertas_adj.accdb_presentes(otra, "2609"), [])
        self.assertEqual(
            ofertas_adj.accdb_presentes_con_fecha(otra, "2609"), ([], None)
        )

    def test_una_sola_lectura_y_ningun_stat_por_nombre(self):
        scandir = mock.Mock(side_effect=os.scandir)
        stat = mock.Mock(side_effect=os.stat)
        with mock.patch.object(os, "scandir", scandir), \
                mock.patch.object(os, "stat", stat):
            presentes = ofertas_adj.accdb_presentes(self.subastas, "2609")
        self.assertEqual(len(presentes), len(self.esperados))
        self.assertEqual(scandir.call_count, 1)
        self.assertEqual(stat.call_count, 0)

    def test_fecha_del_mas_nuevo(self):
        presentes, fecha = ofertas_adj.accdb_presentes_con_fecha(
            self.subastas, "2609"
        )
        self.assertEqual(fecha, max(os.stat(r).st_mtime for _, _, r in presentes))

    def test_revisar_estructura_deja_la_fecha_para_vigencia(self):
        from Script import nucleo

        base = Path(self.tmp.name)
        for sub in nucleo.SUBCARPETAS_CASO:
            (base / sub).mkdir(parents=True, exist_ok=True)
        rutas, filas = nucleo.revisar_estructura(str(base), "2609")
        self.assertEqual(len(rutas["accdb_subastas"]), len(self.esperados))
        self.assertEqual(
            rutas["accdb_subastas_mtime"],
            max(os.stat(r).st_mtime for r in rutas["accdb_subastas"]),
        )
        fila = next(f for f in filas if f["id"] == "db_subastas")
        self.assertEqual(fila["estado"], "ok")


if __name__ == "__main__":
    unittest.main()
