# -*- coding: utf-8 -*-
"""Calculo E Costos con el metodo nuevo del FD (macro SSCC_Calcular_Integrado).

El FD se aplica a la energia medida de cada cuarto de hora, en su propia
hora, antes de acumular y repartir en la curva; ya no hay descuento FD.
Caso armado a mano: una central (Pmax 0,4 MW = 100 kWh por cuarto), un
ciclo, la hora 1 descargando y la hora 2 cargando.
"""

import unittest

import pandas as pd

from Script import nucleo
from Script.nucleo.parametros import COLUMNA_CENTRAL


def subastas():
    filas = [
        # hora 1: CPF(-) y CSF(-) -> 50% / 50%
        ("CPF(-)", "BAJADA", 1, 1),
        ("CSF(-)", "BAJADA", 1, 1),
        # hora 2: dos CPF(+) y un CSF(+) -> 2/3 / 1/3
        ("CPF(+)", "SUBIDA", 2, 2),
        ("CPF(+)", "SUBIDA", 2, 2),
        ("CSF(+)", "SUBIDA", 2, 2),
    ]
    return pd.DataFrame([
        {
            "Concepto": concepto, "Sub_Baj": sub_baj,
            "Configuración": "SAE A", "Mes": 9, "Dia": 1,
            "Hora_dia": hora_dia, "Hora_mes": hora_mes,
        }
        for concepto, sub_baj, hora_dia, hora_mes in filas
    ])


def bloque_fd(mas, menos, valores):
    return pd.DataFrame([
        {"id": f"{hora}BESS A", "Hora Mes": hora, "Unidad": "BESS A",
         COLUMNA_CENTRAL: "SAE A", mas: v_mas, menos: v_menos}
        for hora, (v_mas, v_menos) in valores.items()
    ])


FD_CPF = bloque_fd("CPF(+)", "CPF(-)", {1: (0.7, 0.9), 2: (0.6, 0.95)})
FD_CSF = bloque_fd("CSF(+)", "CSF(-)", {1: (0.5, 0.8), 2: (0.75, 0.4)})


def calculo_base():
    cmg = [10, 20, 30, 40, 5, 6, 7, 8]
    return pd.DataFrame({
        "Mes": 9, "Dia": 1,
        "Hora": [1] * 4 + [2] * 4,
        "Hora Mes": [1] * 4 + [2] * 4,
        "Minutos": [0, 15, 30, 45] * 2,
        "Cuarto de Hora": range(1, 9),
        "clave": "SAE A", "Barra": "B",
        "Energia_Positiva": [50.0] * 4 + [0.0] * 4,
        "Energia_Negativa": [0.0] * 4 + [-40.0] * 4,
        "SoC": 0.5, "Copia_Ventana": 1,
        "CMg": [float(v) for v in cmg],
    })


def calcular(fd_cpf=FD_CPF, fd_csf=FD_CSF):
    mensajes = []
    df = nucleo.completar_calculo_e_costos_grupos(
        calculo_base(), subastas(), {"sae a": 0.4}, 0.06, None,
        fd_csf, fd_cpf, registrar=mensajes.append,
    )
    return df, mensajes


class TestMetodoNuevoFD(unittest.TestCase):

    def test_prorrata_con_direccion_por_hora(self):
        df, _ = calcular()
        hora1, hora2 = df.iloc[0], df.iloc[4]
        self.assertEqual((hora1["AG"], hora1["AH"], hora1["AJ"], hora1["AK"]),
                         (0.5, 0.5, 0.0, 0.0))
        self.assertAlmostEqual(hora2["AJ"], 2 / 3)
        self.assertAlmostEqual(hora2["AK"], 1 / 3)
        self.assertEqual((hora2["AG"], hora2["AH"]), (0.0, 0.0))

    def test_fd_de_la_hora_de_la_fila_solo_con_prorrata(self):
        df, _ = calcular()
        hora1, hora2 = df.iloc[0], df.iloc[4]
        # hora 1: solo (-) tiene prorrata; los (+) no se buscan
        self.assertEqual((hora1["AM"], hora1["AN"], hora1["AP"], hora1["AQ"]),
                         (0.9, 0.8, 0.0, 0.0))
        self.assertEqual((hora2["AM"], hora2["AN"], hora2["AP"], hora2["AQ"]),
                         (0.0, 0.0, 0.6, 0.75))

    def test_energia_ajustada(self):
        df, _ = calcular()
        # 50 x (0,5x0,9 + 0,5x0,8) = 42,5 ; -40 x (2/3x0,6 + 1/3x0,75) = -26
        self.assertEqual(list(df["AS"].round(9)), [42.5] * 4 + [0.0] * 4)
        self.assertEqual(list(df["AT"].round(9)), [0.0] * 4 + [-26.0] * 4)

    def test_acumulados_y_asignacion(self):
        df, _ = calcular()
        self.assertEqual(list(df["AU"].round(9)),
                         [170, 127.5, 85, 42.5, 0, 0, 0, 0])
        self.assertEqual(list(df["AV"].round(9)),
                         [104, 104, 104, 104, 104, 78, 52, 26])
        # 170 kWh en bloques de 100: 100 + 70 ; 104: 100 + 4
        self.assertEqual(list(df["AW"].round(9)), [100, 70, 0, 0, 0, 0, 0, 0])
        self.assertEqual(list(df["AX"].round(9)), [-100, -4, 0, 0, 0, 0, 0, 0])

    def test_valorizacion_sin_descuento_fd_y_monto(self):
        df, _ = calcular()
        # curva descarga 40,30,20,10: promedio de los 2 puestos asignados 35
        # curva carga 5,6,7,8: promedio 5,5
        self.assertEqual(list(df["AY"].round(9)[:2]), [3500, 2450])
        self.assertEqual(list(df["AZ"].round(9)[:2]), [-550, -22])
        self.assertAlmostEqual(df["BA"].sum(), 5378)
        # U = 50x(10+20+30+40) - 40x(5+6+7+8) = 3960
        self.assertAlmostEqual(df["U"].sum(), 3960)
        self.assertEqual(set(df["BC"].round(9)), {177.25})

    def test_sin_adj_sscc_todo_en_cero(self):
        base = calculo_base()
        sin_subastas = subastas().iloc[0:0]
        df = nucleo.completar_calculo_e_costos_grupos(
            base, sin_subastas, {"sae a": 0.4}, 0.06, None,
            FD_CSF, FD_CPF, registrar=lambda *_: None,
        )
        for columna in ("AG", "AM", "AS", "AT", "AW", "BA", "BC"):
            self.assertTrue((df[columna] == 0).all(), columna)

    def test_fd_vacio_se_toma_como_cero_y_se_avisa(self):
        fd_cpf = FD_CPF.copy()
        fd_cpf.loc[0, "CPF(-)"] = None
        df, mensajes = calcular(fd_cpf=fd_cpf)
        self.assertEqual(df.iloc[0]["AM"], 0.0)
        self.assertAlmostEqual(df.iloc[0]["AS"], 50 * 0.5 * 0.8)
        self.assertTrue(any("FD-008" in m for m in mensajes), mensajes)

    def test_la_hoja_sale_con_los_nombres_nuevos(self):
        df, _ = calcular()
        hoja = nucleo.renombrar_calculo_e_costos(df)
        for nombre in ("Descarga ajustada FD-prorrata kWh",
                       "Asignación Descarga kWh", "Valorización Carga [$]",
                       "Total [CLP]", "Monto a compensar"):
            self.assertIn(nombre, hoja.columns)
        self.assertNotIn("Descuento FD", hoja.columns)
        self.assertEqual(hoja.columns[-1], "Monto a compensar")


if __name__ == "__main__":
    unittest.main()
