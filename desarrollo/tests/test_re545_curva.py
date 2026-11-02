# -*- coding: utf-8 -*-
"""
Curva descendente de RE545 (BM): el k-esimo CMg de cada puesto se busca
dentro de la misma central + ventana + orden de hora (observacion 4 de la
reliquidacion de enero 2026).
"""

import unittest

import pandas as pd

from Script.nucleo.re545_componentes import (
    calcular_bi_bj_re545, calcular_bk_bl_bm_bs_re545,
)

# Hora 1 barata, hora 2 cara; mismos CMg para las dos centrales (misma barra)
CMG = {1: [22.0, 20.0, 18.0, 21.0], 2: [120.0, 135.0, 128.0, 115.0]}
RANKING = {1: 2, 2: 1}  # S: puesto de la hora por CMg promedio
# Inyeccion distinta en cada bloque de la hora cara
INYECCION = {1: [0.0] * 4, 2: [100.0, 200.0, 300.0, 400.0]}


def _filas(central):
    filas = []
    for hora in (1, 2):
        prom = sum(CMG[hora]) / 4
        for q, cmg in enumerate(CMG[hora]):
            filas.append({
                "clave": central, "T": 1, "Minutos": q * 15, "Hora": hora,
                "CMg": cmg, "R": prom, "S": RANKING[hora],
                "Energia_Positiva": INYECCION[hora][q],
            })
    return filas


def _calcular(*centrales):
    # BI se arma por central, como en un mes real (la ventana cambia de
    # 31 a 1 al pasar de una central a la siguiente).
    partes = []
    for central in centrales:
        df = pd.DataFrame(_filas(central))
        df["BI"], df["BJ"] = calcular_bi_bj_re545(df)
        partes.append(df)
    _, _, bm, bs = calcular_bk_bl_bm_bs_re545(pd.concat(partes, ignore_index=True))
    return list(bm), list(bs)


def _curva(*centrales):
    return _calcular(*centrales)[0]


class CurvaRe545(unittest.TestCase):

    ESPERADA = [135.0, 128.0, 120.0, 115.0, 22.0, 21.0, 20.0, 18.0]

    def test_una_central(self):
        self.assertEqual(_curva("SAE-A"), self.ESPERADA)

    def test_dos_centrales_en_la_misma_barra_no_se_mezclan(self):
        bm = _curva("SAE-A", "SAE-B")
        self.assertEqual(bm[:8], self.ESPERADA)
        self.assertEqual(bm[8:], self.ESPERADA)


    def test_la_inyeccion_sale_del_mismo_bloque_que_el_cmg(self):
        # Hora cara: CMg 135 (:15), 128 (:30), 120 (:00), 115 (:45)
        _, bs = _calcular("SAE-A")
        self.assertEqual(bs[:4], [200.0, 300.0, 100.0, 400.0])


if __name__ == "__main__":
    unittest.main()
