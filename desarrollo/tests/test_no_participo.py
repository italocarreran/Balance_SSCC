# -*- coding: utf-8 -*-
"""
Las horas "No participó": el FD se toma TAL CUAL del archivo de desempeno.

El 2026-11-11 se puso en 0 el FD de la direccion "No participó"; el
2026-11-18 el usuario lo devolvio (ver no_participo.py). Estas pruebas
fijan que E Costos y RE545 ya no lo tocan, que el Reporte las sigue
contando por direccion y que la comparacion avisa si una planilla se
calculo con la regla devuelta.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo
from Script.nucleo import no_participo as npar
from Script.nucleo import reporte as r
from Script.nucleo.factores_subastas import tablas_fd_desde_hoja
from Script.nucleo.fma import calcular_fd_subastas


def bloque_cpf(resp_mas, resp_menos, fd=1.0):
    """Un bloque CPF de la hoja FD, dos horas, una unidad 'U'."""

    # Como la devuelve leer_fd_consolidado: "id" primero, la hora del mes
    # en la segunda columna.
    df = pd.DataFrame({k: [fd, fd] for k in list(nucleo.NOMBRES_FD_CPF.values())[:-1]})
    df["id"] = ["a", "b"]
    df["Hora Mes"] = [1, 2]
    columnas = list(df.columns)
    mas = next(c for c in columnas if c.startswith("Respuesta CPF+"))
    menos = next(c for c in columnas if c.startswith("Respuesta CPF-"))
    df[mas], df[menos] = resp_mas, resp_menos
    df["Unidad"] = "U"
    df["Central"] = "SAE-UNO"
    return df


def bloque_csf(resp, fd=1.0):
    df = pd.DataFrame({k: [fd, fd] for k in list(nucleo.NOMBRES_FD_CSF.values())[:-1]})
    df["id"] = ["a", "b"]
    df["Hora Mes"] = [1, 2]
    columna = next(c for c in df.columns if c.startswith("Respuesta CSF"))
    df[columna] = resp
    df["Unidad"] = "U"
    df["Central"] = "SAE-UNO"
    return df


class TestMascarasPorDireccion(unittest.TestCase):

    def test_cpf_cada_direccion_con_su_respuesta(self):
        m = npar.mascaras_no_participo(
            bloque_cpf(["No Participó", 0.9], [0.8, "No participo"]), "CPF"
        )
        self.assertEqual(m["+"].tolist(), [True, False])
        self.assertEqual(m["-"].tolist(), [False, True])

    def test_csf_una_respuesta_para_las_dos(self):
        m = npar.mascaras_no_participo(bloque_csf(["No participo", 0.95]), "CSF")
        self.assertEqual(m["+"].tolist(), [True, False])
        self.assertEqual(m["-"].tolist(), [True, False])


class TestRE545PorFilaDeSubastas(unittest.TestCase):
    """El FD de cada fila de Subastas es el del archivo, aunque diga "No participó"."""

    def test_fd_de_cada_fila(self):
        tablas = tablas_fd_desde_hoja(
            None, bloque_cpf(["No Participó", 1], [1, "No Participó"]),
            registrar=lambda *a: None,
        )
        subastas = pd.DataFrame({
            "Concepto": ["CPF(+)", "CPF(-)", "CPF(+)", "CPF(-)"],
            "Control": ["CPF"] * 4,
            "Hora_mes": [1, 1, 2, 2],
            "Configuración": ["SAE-UNO"] * 4,
        })
        fd = calcular_fd_subastas(
            subastas, tablas, dic_unidad_fd={"sae-uno": "U"}, registrar=lambda *a: None
        )
        self.assertEqual([float(v) for v in fd], [1.0, 1.0, 1.0, 1.0])

    def test_csf_no_participo_conserva_el_fd_del_archivo(self):
        # EMOAC (O2026-0052605/0052606): FD CSF 1 en el archivo, 0 en la hoja.
        tablas = tablas_fd_desde_hoja(
            bloque_csf(["No Participó", 0.95]), None, registrar=lambda *a: None,
        )
        subastas = pd.DataFrame({
            "Concepto": ["CSF(+)", "CSF(-)"],
            "Control": ["CSF"] * 2,
            "Hora_mes": [1, 1],
            "Configuración": ["SAE-UNO"] * 2,
        })
        fd = calcular_fd_subastas(
            subastas, tablas, dic_unidad_fd={"sae-uno": "U"}, registrar=lambda *a: None
        )
        self.assertEqual([float(v) for v in fd], [1.0, 1.0])


class TestReportePorDireccion(unittest.TestCase):
    """
    El reporte cuenta "No participó" y el FD solo en la direccion
    adjudicada (antes contaba si lo decia cualquiera de las dos: inflaba).
    """

    def test_subida_no_participo_adjudicada_en_bajada(self):
        cpf = bloque_cpf(["No Participó", "No Participó"], [1, 1], fd=1.0)
        csf = bloque_csf([0.9, 0.9])
        subastas = pd.DataFrame({
            "Concepto": ["CPF(-)", "CPF(-)"], "Fecha": ["2026-10-01"] * 2,
            "Año": [2026] * 2, "Mes": [10] * 2, "Dia": [1] * 2,
            "Hora_dia": [1, 2], "Hora_mes": [1, 2],
            "Configuración": ["SAE-UNO"] * 2, "Energía SSCC": [5.0, 5.0],
        })
        resumen = pd.DataFrame({"Nombre activo": ["SAE-UNO"], "Propietario": ["E1"]})
        diccionario = pd.DataFrame([["Balance_BESS", "FD"], ["SAE-UNO", "U"]])
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "11_PAGOS_BESS_2610.xlsx"
            nucleo.escribir_salida(
                None, ruta, [], [], df_fd_csf=csf, df_fd_cpf=cpf, df_subastas=subastas,
                centrales=(resumen, diccionario),
                hojas_regenerar={nucleo.HOJA_FD, nucleo.HOJA_SUBASTAS, nucleo.HOJA_CENTRALES},
                registrar=lambda *a: None,
            )
            uno = nucleo.resumen_por_central(ruta, registrar=lambda *a: None).set_index(
                r.C_CENTRAL).loc["SAE-UNO"]
        self.assertEqual(uno[r.C_NPA_CPF], 0.0)       # la bajada si participo
        self.assertEqual(uno[r.C_HA_CPF], 2.0)
        self.assertEqual(uno[r.C_FD_CPF], 1.0)        # FD de la bajada


class TestAvisoPlanillaDeAntes(unittest.TestCase):
    """
    Comparar una planilla calculada con la regla devuelta con una de ahora
    mostraria la regla como un cambio de FD: la comparacion lo avisa.
    """

    def test_avisa(self):
        from unittest import mock

        from Script.nucleo import reporte

        def resumen(ruta, registrar=None):
            df = pd.DataFrame([{r.C_CENTRAL: "SAE-UNO", r.C_TOTAL: 1.0e6,
                                r.C_ECOSTOS: 1.0e6, r.C_RE545: 0.0}],
                              columns=[*r.COLUMNAS_RESUMEN])
            df.attrs["no_participo_en_cero"] = 340 if "2609" in str(ruta) else 0
            return df

        with tempfile.TemporaryDirectory() as carpeta, \
                mock.patch.object(reporte, "resumen_por_central", resumen):
            _, lineas, _ = reporte.comparar_planillas(
                Path(carpeta) / "11_PAGOS_BESS_2610.xlsx",
                Path(carpeta) / "11_PAGOS_BESS_2609.xlsx", registrar=lambda *a: None,
            )
        self.assertTrue(lineas[0].startswith("ATENCIÓN: 11_PAGOS_BESS_2609.xlsx se calculó con la regla"))
        self.assertIn("340 filas", lineas[0])
        self.assertFalse(any("2610.xlsx se calculó con" in l for l in lineas))


if __name__ == "__main__":
    unittest.main()
