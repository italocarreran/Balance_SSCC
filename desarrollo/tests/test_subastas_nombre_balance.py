# -*- coding: utf-8 -*-
"""
Una central que en Subastas se llama distinto que en Balance_BESS.

Pedido del usuario (2026-11-10): "hay una central que en la columna
Balance_BESS se llama diferente a como se llama en subasta y la esta
buscando en subastas con el nombre de Balance BESS". Hasta ese mes se
llamaban siempre igual. Los cruces del calculo contra Subastas (L Adj
SSCC, la prorrata SSCC de E Costos, las reservas de RE545) traducen el
nombre de Subastas al de Balance_BESS por el Diccionario.
"""

import unittest

import pandas as pd

from Script.nucleo.columnas_compartidas import calcular_l
from Script.nucleo.ecostos_prorratas import (
    construir_dic_prorrata, construir_prorrata_sscc,
)
from Script.nucleo.hojas_entrada import subastas_con_nombre_balance_bess
from Script.nucleo.re545_reservas import construir_dic_reservas_subastas

DICCIONARIO = pd.DataFrame([
    [None] * 5,
    ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
    ["SAE-UNO", "U1", "BAT_UNO_SUB", "UNO OF", "Uno"],
    ["SAE-DOS", "U2", "SAE-DOS", "DOS OF", "Dos"],
])


def subastas():
    return pd.DataFrame({
        "Concepto": ["CPF(+)", "CSF(-)", "CPF(+)"],
        "Sub_Baj": ["SUBIDA", "BAJADA", "SUBIDA"],
        "Mes": [10, 10, 10], "Dia": [1, 1, 1], "Hora_dia": [3, 3, 3],
        "Hora_mes": [3, 3, 3],
        "Configuración": ["BAT_UNO_SUB", "BAT_UNO_SUB", "SAE-DOS"],
        "Energía SSCC": [5.0, 2.0, 1.0],
        "FD": [1.0, 1.0, 1.0], "FMA": [1.0, 1.0, 1.0],
    })


class TestNombreBalanceEnSubastas(unittest.TestCase):

    def test_traduce_por_el_diccionario(self):
        avisos = []
        df = subastas_con_nombre_balance_bess(
            subastas(), DICCIONARIO, registrar=avisos.append
        )
        self.assertEqual(
            df["Configuración"].tolist(), ["SAE-UNO", "SAE-UNO", "SAE-DOS"]
        )
        self.assertTrue(any("BAT_UNO_SUB" in a for a in avisos), avisos)
        # La original no se toca (la hoja Subastas sigue igual).
        self.assertEqual(subastas()["Configuración"].iloc[0], "BAT_UNO_SUB")

    def test_sin_diccionario_o_sin_diferencias_queda_igual(self):
        original = subastas()
        self.assertIs(
            subastas_con_nombre_balance_bess(original, None), original
        )
        iguales = original.assign(**{"Configuración": "SAE-DOS"})
        self.assertIs(
            subastas_con_nombre_balance_bess(
                iguales, DICCIONARIO, registrar=lambda *a: None
            ),
            iguales,
        )

    def test_l_y_prorrata_y_reservas_cruzan_con_el_nombre_balance(self):
        calculo = pd.DataFrame({
            "clave": ["SAE-UNO", "SAE-DOS"], "Mes": [10, 10],
            "Dia": [1, 1], "Hora": [3, 3],
        })
        crudo = subastas()
        # Con el nombre de Subastas tal cual, SAE-UNO no cruzaba.
        self.assertEqual(calcular_l(calculo, crudo).tolist(), [0, 1])

        traducido = subastas_con_nombre_balance_bess(
            crudo, DICCIONARIO, registrar=lambda *a: None
        )
        self.assertEqual(calcular_l(calculo, traducido).tolist(), [1, 1])

        dic = construir_dic_prorrata(
            construir_prorrata_sscc(traducido), registrar=lambda *a: None
        )
        self.assertIn("SAE-UNO¦3", dic)
        self.assertNotIn("BAT_UNO_SUB¦3", dic)

        energia, _fd, _fma = construir_dic_reservas_subastas(traducido)
        self.assertEqual(energia[("SAE-UNO", "3", "CPF(+)")], 5.0)


class _Parar(Exception):
    pass


class TestElCalculoRecibeElNombreBalance(unittest.TestCase):
    """Que E Costos y RE545 de verdad crucen con la Subastas traducida."""

    def test_e_costos(self):
        from unittest import mock

        from Script.nucleo import ecostos

        vistos = []

        def espia(df, df_subastas):
            vistos.append(df_subastas["Configuración"].tolist())
            raise _Parar

        with mock.patch.object(ecostos, "calcular_l", espia):
            with self.assertRaises(_Parar):
                ecostos.completar_calculo_e_costos_grupos(
                    pd.DataFrame({"clave": ["SAE-UNO"]}), subastas(), {},
                    DICCIONARIO, None, None, registrar=lambda *a: None,
                )
        self.assertEqual(vistos, [["SAE-UNO", "SAE-UNO", "SAE-DOS"]])

    def test_re545(self):
        from unittest import mock

        from Script.nucleo import proceso

        vistos = {}

        def fd_fma(df_subastas, *a, **k):
            # El FD y el FMA traducen desde el nombre de Subastas.
            vistos["fd_fma"] = df_subastas["Configuración"].tolist()
            return df_subastas

        def completar(df_re545, df_subastas, *a, **k):
            vistos["calculo"] = df_subastas["Configuración"].tolist()
            raise _Parar

        with mock.patch.object(proceso, "construir_calculo_re545",
                               lambda *a, **k: pd.DataFrame()), \
                mock.patch.object(proceso, "subastas_con_fd_y_fma", fd_fma), \
                mock.patch.object(proceso, "completar_calculo_re545", completar):
            with self.assertRaises(_Parar):
                proceso.calcular_hoja_re545(
                    None, None, None, subastas(), None, None, None,
                    DICCIONARIO, None, None, None, None,
                    registrar=lambda *a: None,
                )
        self.assertEqual(vistos["fd_fma"][0], "BAT_UNO_SUB")
        self.assertEqual(vistos["calculo"], ["SAE-UNO", "SAE-UNO", "SAE-DOS"])


if __name__ == "__main__":
    unittest.main()
