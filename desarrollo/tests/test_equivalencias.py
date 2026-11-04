# -*- coding: utf-8 -*-
"""
Cada cruce busca la central con TODOS sus nombres del Diccionario.

Pedido del usuario: "para subastas, ofertas y todo, que se busque con
todos los nombres del diccionario [...] y que se indique tipo 'central
conejo no coincide con nombre de subastas del diccionario, pero si se
encuentra por el nombre de FD'". Primero se prueba el nombre que le
corresponde al cruce; si no esta, los demas de la fila, con aviso
DIC-002.
"""

import unittest

import pandas as pd

from Script import nucleo
from Script.nucleo.equivalencias import Equivalencias
from Script.nucleo.lectura import ROL_FD, ROL_FMA_CPF, ROL_OFERTAS


def diccionario(*filas):
    """Diccionario en el formato nuevo (fila vacia y encabezados)."""

    return pd.DataFrame(
        [[None] * 5, ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"]]
        + [list(fila) for fila in filas]
    )


# La FD y la FMA_CPF del Diccionario estan mal escritas: no cruzan.
DICCIONARIO = diccionario(
    ["SAE CONEJO", "BESS CONEJO MAL", "BAT_CONEJO", "OFERTA CONEJO",
     "Conejo mal - BESS"],
)


class TestEquivalencias(unittest.TestCase):
    def test_primero_el_nombre_del_cruce_sin_aviso(self):
        eq = Equivalencias(DICCIONARIO)
        encontrado = eq.buscar(
            "BAT_CONEJO", ROL_FD, {"bess conejo mal", "sae conejo"}, "FD"
        )
        self.assertEqual(encontrado, "BESS CONEJO MAL")
        self.assertEqual(eq.avisos, [])

    def test_si_no_esta_prueba_los_otros_nombres_y_avisa(self):
        eq = Equivalencias(DICCIONARIO)
        encontrado = eq.buscar(
            "BAT_CONEJO", ROL_FD, {"sae conejo"}, "las unidades de la hoja FD"
        )
        self.assertEqual(encontrado, "SAE CONEJO")
        self.assertEqual(len(eq.avisos), 1)
        aviso = eq.avisos[0]
        self.assertIn("SAE CONEJO", aviso)
        self.assertIn("nombre FD", aviso)
        self.assertIn("BESS CONEJO MAL", aviso)
        self.assertIn("por su nombre Balance_BESS", aviso)

        lineas = []
        eq.informar(lineas.append)
        self.assertTrue(lineas[0].startswith("  [AVISO] DIC-002:"))

    def test_un_aviso_por_central_y_cruce(self):
        eq = Equivalencias(DICCIONARIO)
        for _ in range(3):
            eq.buscar("BAT_CONEJO", ROL_FD, {"sae conejo"}, "FD")
        self.assertEqual(len(eq.avisos), 1)

    def test_si_no_esta_en_ninguna_parte_queda_none(self):
        eq = Equivalencias(DICCIONARIO)
        self.assertIsNone(eq.buscar("BAT_CONEJO", ROL_FD, {"otra"}, "FD"))
        self.assertIsNone(eq.buscar("NADIE", ROL_FD, {"otra"}, "FD"))
        # Fuera del Diccionario, el nombre tal cual (como antes).
        self.assertEqual(eq.buscar("OTRA", ROL_FD, {"otra"}, "FD"), "OTRA")

    def test_central_de_un_nombre_de_otra_columna(self):
        eq = Equivalencias(DICCIONARIO)
        # Una unidad de FD que en el Diccionario esta como Subastas.
        self.assertEqual(
            eq.central_de("BAT_CONEJO", ROL_FD, "la unidad de FD"),
            "SAE CONEJO",
        )
        self.assertIn("por su nombre Subastas", eq.avisos[0])
        self.assertIsNone(eq.central_de("NADIE", ROL_OFERTAS, "Ofertas"))

    def test_formato_viejo_no_cambia_nada(self):
        viejo = pd.DataFrame([["FD", None], ["SAE CONEJO", "BESS CONEJO"]])
        eq = Equivalencias(viejo)
        self.assertFalse(eq)
        self.assertIsNone(eq.buscar("SAE CONEJO", ROL_FD, {"x"}, "FD"))


class TestRe545BuscaConTodosLosNombres(unittest.TestCase):
    def test_fd_y_fma_cpf_se_encuentran_por_otro_nombre(self):
        subastas = pd.DataFrame({
            "Concepto": ["CPF(+)", "CSF(-)"],
            "Control": ["CPF", "CSF"],
            "Sub_Baj": ["SUBIDA", "BAJADA"],
            "Fecha": ["2026-07-01"] * 2,
            "Año": [2026, 2026], "Mes": [7, 7], "Dia": [1, 1],
            "Hora_dia": [1, 2], "Hora_mes": [1, 2],
            "Configuración": ["BAT_CONEJO"] * 2,
            "Propietario": ["E1"] * 2,
            "Clave horaria": ["x", "y"], "Ciclo": [1, 1],
            "Energía SSCC": [10.0, 20.0],
        })
        # La hoja FD y el fma_cpf traen la central con su nombre
        # Balance_BESS, no con el de la columna FD / FMA_CPF.
        csf = pd.DataFrame(
            [["2SAE CONEJO", 2, 1, "2026-07-01", 1, "SAE CONEJO",
              "Participó", 1.0, 1.0, 0.8, 0.8, 0.8, 2]]
        ).set_axis(list(nucleo.NOMBRES_FD_CSF.values()), axis=1)
        cpf = pd.DataFrame(
            [["1SAE CONEJO", 1, 1, "2026-07-01", 0, "SAE CONEJO",
              1.0, 1.0, 1.0, 1.0, 0.7, "si", 0.7, 0.7, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CPF.values()), axis=1)
        tablas_fma = {
            "cpf": pd.DataFrame({
                "Central": ["SAE CONEJO"], "Año": [2026], "Mes": [7],
                "Dia": [1], "Hora": [1],
                "FMA_CPF_mas": [0.5], "FMA_CPF_menos": [0.4],
            }),
            "csf": pd.DataFrame({
                "Año": [2026], "Mes": [7], "Dia": [1], "Hora": [2],
                "FMA_CSF_mas_base": [0.3], "FMA_CSF_menos_base": [0.2],
            }),
            "ctf_unidad": None, "ctf_configuracion": None,
        }

        lineas = []
        con_factores = nucleo.subastas_con_fd_y_fma(
            subastas, csf, cpf, tablas_fma, DICCIONARIO,
            registrar=lineas.append,
        )

        self.assertEqual(con_factores["FD"].tolist(), [0.7, 0.8])
        self.assertEqual(con_factores["FMA"].tolist(), [0.5, 0.2])
        avisos = [l for l in lineas if "DIC-002" in l]
        self.assertEqual(len(avisos), 2, lineas)  # FD y FMA CPF
        self.assertTrue(any("nombre FD" in a for a in avisos))
        self.assertTrue(any("nombre FMA_CPF" in a for a in avisos))


class TestOfertasYECostos(unittest.TestCase):
    def test_ofertas_cruza_por_el_nombre_balance_bess(self):
        from Script.nucleo.ofertas_sscc import agregar_central_ofertas

        df_wxy = pd.DataFrame({
            "Nombre": ["SAE CONEJO", "OTRA"], "Dia": [1, 1],
            "Oferta completa": [1, 1],
        })
        lineas = []
        con_central = agregar_central_ofertas(
            df_wxy, DICCIONARIO, registrar=lineas.append
        )
        self.assertEqual(
            con_central[nucleo.COLUMNA_CENTRAL].tolist(), ["SAE CONEJO", ""]
        )
        self.assertTrue(any("nombre Ofertas" in l for l in lineas), lineas)

    def test_e_costos_busca_la_unidad_por_otro_nombre(self):
        from Script.nucleo.ecostos import _completar_mapeo_con_otros_nombres

        fd = pd.DataFrame({"Unidad": ["BAT_CONEJO"]})
        dic_mapeo = {"sae conejo": "BESS CONEJO MAL"}
        lineas = []
        _completar_mapeo_con_otros_nombres(
            pd.DataFrame({"clave": ["SAE CONEJO"]}), dic_mapeo, DICCIONARIO,
            fd, fd, lineas.append,
        )
        self.assertEqual(dic_mapeo["sae conejo"], "BAT_CONEJO")
        self.assertTrue(any("por su nombre Subastas" in l for l in lineas))


if __name__ == "__main__":
    unittest.main()
