"""SOC-019 (carga/descarga contra el SoC) y CMG-007 (barra sin CMg).

Observacion a 11_PAGOS_BESS_2610 sobre SAE-CRCA-PFV-ELENA.
"""

import unittest

import pandas as pd

from Script import nucleo


def _medidores(central, energias, socs, inicio=1):
    n = len(energias)
    intervalo = pd.date_range("2026-10-01", periods=n, freq="15min")
    return pd.DataFrame({
        "clave": [central] * n,
        "intervalo": intervalo,
        "Cuarto de Hora": list(range(inicio, inicio + n)),
        "Gen_Unidad": energias,
        "SoC": socs,
    })


def _ciclo(invertido=False, bloques=40):
    """Carga y luego descarga, con el SoC siguiendo (o al reves)."""
    mitad = bloques // 2
    energia = [-100.0] * mitad + [100.0] * mitad
    soc = [0.1]
    for e in energia[:-1]:
        paso = -e / 1000.0
        soc.append(soc[-1] + (-paso if invertido else paso))
    return energia, soc


class SocVsEnergiaTest(unittest.TestCase):

    def test_bess_sana_no_alerta(self):
        energia, soc = _ciclo()
        df = _medidores("SAE-SANA", energia, soc)
        mensajes = []
        alertas = nucleo.verificar_soc_vs_energia(df, mensajes.append)
        self.assertEqual(alertas, [])

    def test_canales_invertidos_alerta_alta(self):
        energia, soc = _ciclo(invertido=True)
        df = _medidores("SAE-CRCA-PFV-ELENA", energia, soc)
        registro = nucleo.Registro(salida=None)
        alertas = nucleo.verificar_soc_vs_energia(df, registro)
        self.assertEqual(len(alertas), 1)
        self.assertEqual(alertas[0].id_alerta, "SOC-019")
        self.assertEqual(alertas[0].severidad, nucleo.ALTA)
        self.assertEqual(alertas[0].central, "SAE-CRCA-PFV-ELENA")
        self.assertIn("CRCA", alertas[0].mensaje)
        self.assertEqual(registro.estado(), nucleo.NO_APROBADA)

    def test_pocos_bloques_avisa_que_no_se_pudo(self):
        energia, soc = _ciclo(invertido=True, bloques=10)
        df = _medidores("SAE-X", energia, soc)
        alertas = nucleo.verificar_soc_vs_energia(df, lambda *_: None)
        self.assertEqual([a.id_alerta for a in alertas], ["SOC-020"])
        self.assertEqual(alertas[0].severidad, nucleo.MEDIA)

    def test_sin_soc_avisa_que_no_se_pudo(self):
        energia, _ = _ciclo()
        df = _medidores("SAE-SIN-SOC", energia, [None] * len(energia))
        alertas = nucleo.verificar_soc_vs_energia(df, lambda *_: None)
        self.assertEqual([a.id_alerta for a in alertas], ["SOC-020"])
        self.assertIn("sin SoC", alertas[0].mensaje)


def _pfv_dia(signo, dias=3):
    """Un PFV sin SoC: inyecta de 8 a 18 y retira un poco de noche."""
    intervalo = pd.date_range("2026-10-01", periods=96 * dias, freq="15min")
    energia = [
        signo * (500.0 if 8 <= t.hour < 18 else -20.0) for t in intervalo
    ]
    return pd.DataFrame({
        "clave": "SAE-CRCA-PFV-X", "intervalo": intervalo,
        "Cuarto de Hora": range(1, len(intervalo) + 1),
        "Gen_Unidad": energia, "SoC": [None] * len(intervalo),
    })


class MediodiaPfvTest(unittest.TestCase):

    def test_pfv_bien_no_alerta(self):
        alertas = nucleo.verificar_soc_vs_energia(_pfv_dia(+1), print)
        self.assertEqual(alertas, [])

    def test_pfv_invertido_alerta_soc021(self):
        alertas = nucleo.verificar_soc_vs_energia(
            _pfv_dia(-1), lambda *_: None
        )
        self.assertEqual([a.id_alerta for a in alertas], ["SOC-021"])
        self.assertEqual(alertas[0].severidad, nucleo.ALTA)

    def test_signo_por_central(self):
        energia, soc = _ciclo(invertido=True)
        df = pd.concat([
            _pfv_dia(+1), _medidores("SAE-B", energia, soc),
        ])
        dato = nucleo.signo_por_central(df)
        self.assertEqual(dato["SAE-B"]["metodo"], "SoC")
        self.assertLess(dato["SAE-B"]["coincidencia"], 0.1)
        self.assertEqual(dato["SAE-CRCA-PFV-X"]["metodo"], "mediodia PFV")
        self.assertEqual(dato["SAE-CRCA-PFV-X"]["coincidencia"], 1.0)


class MatrizSocVsMedidaTest(unittest.TestCase):

    def test_columna_en_la_matriz(self):
        from Script.nucleo.matriz_centrales import (
            COL_SOC_VS_MEDIDA, FALTA, OK, construir_matriz_centrales,
        )
        sana, soc_sana = _ciclo()
        inv, soc_inv = _ciclo(invertido=True)
        df = pd.concat([
            _medidores("SAE-SANA", sana, soc_sana),
            _medidores("SAE-INV", inv, soc_inv),
        ])
        matriz = construir_matriz_centrales(df_medidores=df)
        fila = dict(zip(matriz["Central"], matriz[COL_SOC_VS_MEDIDA]))
        self.assertTrue(fila["SAE-SANA"].startswith("100 % coincide"))
        self.assertIn("signo invertido", fila["SAE-INV"])
        estados = matriz.attrs["estados"]
        self.assertEqual(estados[("SAE-SANA", COL_SOC_VS_MEDIDA)], OK)
        self.assertEqual(estados[("SAE-INV", COL_SOC_VS_MEDIDA)], FALTA)


class HomologacionTest(unittest.TestCase):

    def _homol(self, filas):
        return pd.DataFrame(
            filas, columns=["Punto de Medida", "Canal", "clave", "Flujo"]
        )

    def test_filas_repetidas_hom001(self):
        from Script.Medidas.Homologacion import revisar_homologacion
        df = self._homol([
            ["TENGIS_J1", "kWhD", "SAE-LILE", -1],
            ["TENGIS_J1", "kWhD", "SAE-LILE", -1],
            ["OTRO", "kWhD", "SAE-B", 1],
            ["OTRO", "kWhR", "SAE-B", -1],
        ])
        ids = [p["id"] for p in revisar_homologacion(df)]
        self.assertEqual(ids, ["HOM-001"])

    def test_dos_canales_mismo_signo_hom002(self):
        from Script.Medidas.Homologacion import revisar_homologacion
        df = self._homol([
            ["P1", "kWhD", "SAE-B", 1],
            ["P1", "kWhR", "SAE-B", 1],
        ])
        ids = [p["id"] for p in revisar_homologacion(df)]
        self.assertEqual(ids, ["HOM-002"])

    def test_convenciones_distintas_no_son_error(self):
        from Script.Medidas.Homologacion import revisar_homologacion
        df = self._homol([
            ["P1", "kWhD", "SAE-A", 1], ["P1", "kWhR", "SAE-A", -1],
            ["P2", "kWhD", "SAE-B", -1], ["P2", "kWhR", "SAE-B", 1],
        ])
        self.assertEqual(revisar_homologacion(df), [])

    def test_desorden_de_filas_no_importa(self):
        energia, soc = _ciclo()
        df = _medidores("SAE-SANA", energia, soc).sample(frac=1, random_state=1)
        self.assertEqual(nucleo.verificar_soc_vs_energia(df, print), [])


class BarrasConCmgTest(unittest.TestCase):

    def _dic_cmg(self):
        return {f"CRUCERO_______220|{c}": (50.0, 50.0) for c in range(1, 5)}

    def test_barra_ausente_de_cmg(self):
        df = pd.concat([
            _medidores("SAE-CRCA-PFV-ELENA", [1.0] * 4, [0.5] * 4),
            _medidores("SAE-OK", [1.0] * 4, [0.5] * 4),
        ])
        mapa = {"sae-crca-pfv-elena": "KIMAL_________220",
                "sae-ok": "CRUCERO_______220"}
        registro = nucleo.Registro(salida=None)
        alertas = nucleo.verificar_barras_con_cmg(
            df, mapa, self._dic_cmg(), registro
        )
        self.assertEqual(len(alertas), 1)
        a = alertas[0]
        self.assertEqual(a.id_alerta, "CMG-007")
        self.assertEqual(a.central, "SAE-CRCA-PFV-ELENA")
        self.assertIn("KIMAL_________220", a.mensaje)
        self.assertIn("queda en 0", a.mensaje)
        self.assertEqual(registro.estado(), nucleo.NO_APROBADA)

    def test_barra_presente_pero_ningun_cuarto_cruza(self):
        df = _medidores("SAE-OK", [1.0] * 4, [0.5] * 4, inicio=100)
        alertas = nucleo.verificar_barras_con_cmg(
            df, {"sae-ok": "CRUCERO_______220"}, self._dic_cmg(), print
        )
        self.assertEqual(len(alertas), 1)
        self.assertIn("ninguno", alertas[0].mensaje)

    def test_sin_barra_no_duplica_mae001(self):
        df = _medidores("SAE-SIN", [1.0] * 4, [0.5] * 4)
        self.assertEqual(
            nucleo.verificar_barras_con_cmg(df, {}, self._dic_cmg(), print), []
        )

    def test_verificar_medidores_junta_las_dos(self):
        energia, soc = _ciclo(invertido=True)
        df = _medidores("SAE-CRCA-PFV-ELENA", energia, soc)
        alertas = nucleo.verificar_medidores(
            df, {"sae-crca-pfv-elena": "KIMAL_________220"},
            self._dic_cmg(), registrar=lambda *_: None,
        )
        self.assertEqual(
            sorted(a.id_alerta for a in alertas), ["CMG-007", "SOC-019"]
        )


if __name__ == "__main__":
    unittest.main()
