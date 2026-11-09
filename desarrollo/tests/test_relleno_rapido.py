# -*- coding: utf-8 -*-
"""
La escritura rapida (relleno_rapido) da el MISMO libro que escribir todo
con openpyxl.

Pedido del usuario (2026-11-03): "mucho más rápido [...] que no cague la
planilla, mismos resultados". Se escribe el mismo libro de las dos
formas (escritura.ESCRITURA_RAPIDA) con tablas mas largas que la
muestra, y se compara celda por celda: valor, formato de numero, fuente,
relleno, borde, alineacion y comentario; ademas celdas combinadas,
panel, anchos y altos.
"""

import datetime as dt
import os
import tempfile
import unittest

import numpy as np
import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo import escritura
from Script.nucleo.ecostos import COLUMNAS_SALIDA_E_COSTOS as CE, NOMBRES_CALCULO_E_COSTOS as NE
from Script.nucleo.re545 import COLUMNAS_SALIDA_RE545 as CR, NOMBRES_CALCULO_RE545 as NR
from Script.nucleo.re545_resumen import NOMBRES_RESUMEN_RE545 as NRR
from Script.nucleo.hojas_entrada import NOMBRES_FD_CSF, NOMBRES_FD_CPF


def _datos(n):
    rng = np.random.default_rng(7)
    def marco(cols, nombres):
        df = pd.DataFrame({c: rng.uniform(-50, 50, n).round(3) for c in cols})
        for c in ("Mes","Dia","Hora","Hora Mes","Minutos","Cuarto de Hora","Copia_Ventana","T","L"):
            if c in df: df[c] = rng.integers(0, 30, n)
        df["clave"] = rng.choice(["SAE UNO", " SAE DOS ", "A&B <x>"], n); df["Barra"] = "B"
        for c in ("AB","AD","BM","BN"):
            if c in df: df.loc[rng.random(n) < .3, c] = np.nan
        return df[cols].rename(columns=nombres)
    ec = marco(CE, NE); re545 = marco(CR, NR); re545.attrs["ctf_en_cero"] = True
    res = pd.DataFrame({v: rng.uniform(0, 9, 250) for v in NRR.values()}); res["Central"] = "SAE"
    pr = pd.DataFrame({"Cuarto de Hora": rng.integers(1, 2900, 3*n), "Suministrador": rng.choice(["E1","E2"], 3*n), "Prorrata": rng.random(3*n), "Monto a pagar [$]": rng.random(3*n)*100})
    cu = pd.DataFrame({"Cuarto de Hora": range(1, 301), "Monto a compensar [$]": rng.random(300)})
    pg = pd.DataFrame({"Suministrador": ["E1","E2"], "Total a pagar [$]": [1.0, 2.0]})
    rs = pd.DataFrame({"NOMBRE": [f"E{i}" for i in range(250)], "RECIBE": rng.random(250), "PAGA": rng.random(250), "NETO": rng.random(250)})
    med = pd.DataFrame({"Mes": 7, "Dia": rng.integers(1,30,n), "intervalo": [dt.datetime(2026,7,1)+dt.timedelta(minutes=15*i) for i in range(n)], "clave": "SAE", "Gen_Unidad": rng.normal(size=n), "SoC": rng.random(n), "Ventana_No_Completa": rng.integers(0,2,n), "flag": rng.random(n) < .5})
    med.loc[5, "Gen_Unidad"] = np.nan
    cmg = pd.DataFrame({"A": None, "B": None, "C": None, "D_Barra": "B", "E": None, "F_CMg": rng.random(n)*50, "G": None, "H_Cuarto": range(n), "I_CMgProm": rng.random(n)})
    csf = pd.DataFrame({k: rng.random(n) for k in list(NOMBRES_FD_CSF.values())[1:-1]}); csf["Unidad"]="U"; csf["Central"]="SAE"
    c5 = csf.columns[5]; csf[c5] = csf[c5].astype(object); csf.loc[3, c5] = "Participó"
    cpf = pd.DataFrame({k: rng.random(n//2) for k in list(NOMBRES_FD_CPF.values())[1:-1]}); cpf["Unidad"]="U"; cpf["Central"]="SAE"
    sub = pd.DataFrame({"Concepto": rng.choice(["CPF(+)","CSF(-)"], n), "Fecha": "2026-07-01", "Año": 2026, "Configuración": "SAE", "Energía SSCC": rng.random(n)})
    fma = {"cpf": pd.DataFrame({"Central": "SAE", "Año": 2026, "Hora": range(n), "FMA_CPF_mas": rng.random(n)}), "csf": pd.DataFrame({"Año": [2026]*50, "FMA_CSF_mas_base": rng.random(50)}), "ctf_unidad": None, "ctf_configuracion": None}
    return locals()


def _escribir(carpeta, rapida, d):
    escritura.ESCRITURA_RAPIDA = rapida
    try:
        ruta = os.path.join(carpeta, "Balance_BESS.xlsx")
        nucleo.escribir_salida(
            d["med"], ruta, [], [], df_cmg=d["cmg"], df_fd_csf=d["csf"],
            df_fd_cpf=d["cpf"], df_subastas=d["sub"], tablas_fma=d["fma"],
            hojas_regenerar={"Medidores", "CMg", "FD", "Subastas", "FMA"},
            registrar=lambda *a: None,
        )
        nucleo.escribir_pagos_bess(
            ruta, df_ecostos=d["ec"], df_re545=d["re545"],
            df_resumen_re545=d["res"], df_prorrata_retiros=d["pr"],
            df_compensacion_cuarto=d["cu"], df_pagos_suministrador=d["pg"],
            df_resumen=d["rs"], registrar=lambda *a: None,
        )
    finally:
        escritura.ESCRITURA_RAPIDA = True
    return ruta


def _valor(celda):
    from openpyxl.worksheet.formula import ArrayFormula

    if isinstance(celda.value, ArrayFormula):
        return ("matricial", celda.value.ref, celda.value.text)
    return celda.value


def _atributos(celda):
    return (
        _valor(celda), celda.number_format, repr(celda.font), repr(celda.fill),
        repr(celda.border), repr(celda.alignment),
        celda.comment.text if celda.comment else None,
    )


class TestEscrituraRapida(unittest.TestCase):
    def test_el_libro_sale_igual_celda_por_celda(self):
        d = _datos(450)
        with tempfile.TemporaryDirectory() as uno,                 tempfile.TemporaryDirectory() as dos:
            ruta_normal = _escribir(uno, False, d)
            ruta_rapida = _escribir(dos, True, d)
            normal = openpyxl.load_workbook(ruta_normal)
            rapida = openpyxl.load_workbook(ruta_rapida)
            # Los valores guardados (la formula de prueba de la primera
            # fila lleva el valor del programa) tambien salen iguales.
            for nombre in ("Calculo E Costos", "Calculo RE545"):
                a = openpyxl.load_workbook(ruta_normal, data_only=True)[nombre]
                b = openpyxl.load_workbook(ruta_rapida, data_only=True)[nombre]
                filas_a = [[c.value for c in f] for f in a.iter_rows(max_row=10)]
                filas_b = [[c.value for c in f] for f in b.iter_rows(max_row=10)]
                self.assertEqual(filas_a, filas_b, nombre)
                self.assertTrue(
                    any(v is not None for v in filas_a[3]), nombre
                )
            self.assertEqual(normal.sheetnames, rapida.sheetnames)
            for nombre in normal.sheetnames:
                a, b = normal[nombre], rapida[nombre]
                self.assertEqual(
                    (a.max_row, a.max_column), (b.max_row, b.max_column), nombre
                )
                for fila_a, fila_b in zip(a.iter_rows(), b.iter_rows()):
                    for ca, cb in zip(fila_a, fila_b):
                        self.assertEqual(
                            _atributos(ca), _atributos(cb),
                            f"{nombre}!{ca.coordinate}",
                        )
                self.assertEqual(
                    sorted(map(str, a.merged_cells.ranges)),
                    sorted(map(str, b.merged_cells.ranges)), nombre,
                )
                self.assertEqual(a.freeze_panes, b.freeze_panes, nombre)
                self.assertEqual(
                    {k: v.width for k, v in a.column_dimensions.items()},
                    {k: v.width for k, v in b.column_dimensions.items()},
                    nombre,
                )
                self.assertEqual(
                    {k: v.height for k, v in a.row_dimensions.items() if v.height},
                    {k: v.height for k, v in b.row_dimensions.items() if v.height},
                    nombre,
                )


if __name__ == "__main__":
    unittest.main()
