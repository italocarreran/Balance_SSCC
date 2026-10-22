"""
Las dos lecturas nuevas del consolidado y los dos formatos de la hoja
Diccionario.

Cubre lo que se cambio en esta sesion:
  - "Calculo E Costos" toma el FD de la hoja 'FD' del consolidado en
    vez de releer el SSCC_Desempeño_*;
  - la hoja 'Medidores' ya no trae R, S ni T: salen de la hoja
    'Ofertas SSCC';
  - la hoja Diccionario acepta el formato nuevo de una sola tabla
    (Balance_BESS | FD | Subastas | Ofertas | FMA_CPF) sin dejar de
    entender el viejo de bloques lado a lado.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo


# El Diccionario real anterior a la propuesta del usuario: bloques
# separados por columnas vacias, "FD" en A:B y "Subastas"/"ofertas" en
# E:F:G, sin ninguna columna de FMA CPF.
DICCIONARIO_VIEJO = pd.DataFrame([
    ["FD", None, None, None, "Subastas", None, "ofertas"],
    [None, None, None, None, None, None, None],
    ["SAE-TOCOPILLA", "SAE TOCOPILLA", None, None,
     "SAE-TOCOPILLA", "BAT_TOCOPILLA", None],
    ["SAE-CRCA-PFV-ANDES3", "BESS PFV ANDES SOLAR III", None, None,
     "SAE-CRCA-PFV-ANDES3", "SAE-CRCA-PFV-ANDES3", "BAT_ANDES_3_FV"],
])

# La propuesta del usuario: una sola tabla con encabezados.
DICCIONARIO_NUEVO = pd.DataFrame([
    [None, None, None, None, None],
    ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
    ["SAE-TOCOPILLA", "SAE TOCOPILLA", "BAT_TOCOPILLA", None,
     "Tocopilla - BESS"],
    ["SAE-CRCA-PFV-ANDES3", "BESS PFV ANDES SOLAR III",
     "SAE-CRCA-PFV-ANDES3", "BAT_ANDES_3_FV", "Andes Solar 3 - PFV"],
])


class DiccionarioFormatoNuevoTest(unittest.TestCase):

    def test_el_formato_viejo_se_sigue_reconociendo_como_viejo(self):
        fila, columnas = nucleo.encabezado_diccionario(DICCIONARIO_VIEJO)

        self.assertIsNone(fila)
        self.assertEqual(columnas, {})
        self.assertEqual(nucleo.filas_diccionario(DICCIONARIO_VIEJO), [])

    def test_el_formato_nuevo_se_detecta_con_sus_cinco_columnas(self):
        fila, columnas = nucleo.encabezado_diccionario(DICCIONARIO_NUEVO)

        self.assertEqual(fila, 1)
        self.assertEqual(
            columnas,
            {"balance_bess": 0, "fd": 1, "subastas": 2, "ofertas": 3,
             "fma_cpf": 4},
        )

    def test_la_nomenclatura_de_fma_cpf_solo_existe_en_el_formato_nuevo(self):
        """Es la causa de las diferencias de FMA CPF que reporto el usuario."""

        self.assertEqual(
            nucleo.construir_dic_bloque_diccionario(
                DICCIONARIO_VIEJO, nucleo.TITULO_BLOQUE_FMA_CPF
            ),
            {},
        )

        dic = nucleo.construir_dic_bloque_diccionario(
            DICCIONARIO_NUEVO, nucleo.TITULO_BLOQUE_FMA_CPF
        )
        self.assertEqual(dic["bat_tocopilla"], "Tocopilla - BESS")

    def test_el_fd_se_homologa_por_el_nombre_de_subastas(self):
        """
        BAT_TOCOPILLA es como se llama Tocopilla en Subastas. Con el
        formato viejo la clave era el nombre canonico y esa fila no
        cruzaba; con el nuevo cruza por los dos nombres.
        """

        dic = nucleo.construir_dic_bloque_diccionario(
            DICCIONARIO_NUEVO, nucleo.TITULO_BLOQUE_FD
        )

        self.assertEqual(dic["bat_tocopilla"], "SAE TOCOPILLA")
        self.assertEqual(dic["sae-tocopilla"], "SAE TOCOPILLA")

    def test_los_dos_formatos_dan_la_misma_homologacion(self):
        viejo = nucleo.construir_homologacion(DICCIONARIO_VIEJO)
        nuevo = nucleo.construir_homologacion(DICCIONARIO_NUEVO)

        for mapa in (viejo, nuevo):
            self.assertEqual(mapa["bat_tocopilla"], "SAE-TOCOPILLA")
            self.assertEqual(mapa["sae tocopilla"], "SAE-TOCOPILLA")
            self.assertEqual(
                mapa["bat_andes_3_fv"], "SAE-CRCA-PFV-ANDES3"
            )

        # Y la fila de encabezados del formato nuevo NO entra como si
        # fuera una central mas.
        self.assertNotIn("balance_bess", nuevo)
        self.assertNotIn("fma_cpf", nuevo)

    def test_el_mapeo_a_b_sigue_dando_lo_mismo_en_los_dos(self):
        for diccionario in (DICCIONARIO_VIEJO, DICCIONARIO_NUEVO):
            dic = nucleo.construir_dic_mapeo_diccionario(diccionario)
            self.assertEqual(dic["sae-tocopilla"], "SAE TOCOPILLA")


class LecturaDelConsolidadoTest(unittest.TestCase):
    """Escribir la hoja y volver a leerla tiene que dar lo mismo."""

    def setUp(self):
        self.carpeta = tempfile.TemporaryDirectory()
        self.ruta = Path(self.carpeta.name) / "Consolidado_entradas.xlsx"

    def tearDown(self):
        self.carpeta.cleanup()

    def test_el_fd_vuelve_con_sus_dos_bloques_y_su_largo(self):
        csf = pd.DataFrame(
            [[f"1U{i}", 1, 1, "2026-07-01", 1, f"U{i}",
              0.1, 0.2, 0.3, 0.9, 0.9, 0.9, 1] for i in range(3)]
        ).set_axis(list(nucleo.NOMBRES_FD_CSF.values()), axis=1)
        cpf = pd.DataFrame(
            [[f"1U{i}", 1, 1, "2026-07-01", 1, f"U{i}",
              0.1, 0.2, 0.3, 0.4, 0.8, "si", 0.8, 0.8, 1]
             for i in range(5)]
        ).set_axis(list(nucleo.NOMBRES_FD_CPF.values()), axis=1)

        nucleo.escribir_salida(
            pd.DataFrame({"Mes": [7]}), self.ruta, [], [],
            df_fd_csf=csf, df_fd_cpf=cpf,
        )

        leido_csf, leido_cpf = nucleo.leer_fd_consolidado(
            self.ruta, registrar=lambda _: None
        )

        # El bloque corto no se lleva las filas de relleno del largo.
        self.assertEqual(len(leido_csf), 3)
        self.assertEqual(len(leido_cpf), 5)
        self.assertEqual(
            list(leido_csf.columns), list(nucleo.NOMBRES_FD_CSF.values())
        )
        self.assertEqual(
            list(leido_cpf.columns), list(nucleo.NOMBRES_FD_CPF.values())
        )

    def test_sin_hoja_fd_el_error_dice_que_hay_que_generarla(self):
        pd.DataFrame({"Mes": [7]}).to_excel(self.ruta, index=False)

        with self.assertRaises(nucleo.ErrorEntrada) as caso:
            nucleo.leer_fd_consolidado(self.ruta, registrar=lambda _: None)

        self.assertIn("FD", str(caso.exception))

    def test_ofertas_sscc_vuelve_con_sus_dos_tablas(self):
        wxy = pd.DataFrame({
            "Nombre": ["SAE UNO"] * 3,
            "Dia": [1, 2, 3],
            "Oferta completa": [1, 0, 1],
        })
        resumen = pd.DataFrame({
            "Central": ["SAE UNO", "SAE UNO"],
            "Ventana T": [0, 1],
            "Oferta": [36, 96],
            "Completa": [1, 0],
        })

        nucleo.escribir_salida(
            pd.DataFrame({"Mes": [7]}), self.ruta, [], [],
            df_wxy=wxy, df_resumen_ventana=resumen,
        )

        leido_wxy, leido_resumen = nucleo.leer_ofertas_sscc_consolidado(
            self.ruta, registrar=lambda _: None
        )

        pd.testing.assert_frame_equal(
            leido_wxy.reset_index(drop=True), wxy, check_dtype=False
        )
        pd.testing.assert_frame_equal(
            leido_resumen.reset_index(drop=True), resumen, check_dtype=False
        )


class ColumnasDeOfertasEnMedidoresTest(unittest.TestCase):
    """R, S y T se reconstruyen igual que cuando vivian en Medidores."""

    def test_r_s_y_t_salen_de_las_dos_tablas(self):
        medidores = pd.DataFrame({
            "Dia": [1, 1, 2, 2],
            "clave": ["SAE UNO"] * 4,
            "Ventana": [0, 0, 1, 1],
        })
        wxy = pd.DataFrame({
            "Nombre": ["SAE UNO", "SAE UNO"],
            "Dia": [1, 2],
            "Oferta completa": [1, 0],
        })
        resumen = pd.DataFrame({
            "Central": ["SAE UNO", "SAE UNO"],
            "Ventana T": [0, 1],
            "Oferta": [2, 0],
            "Completa": [1, 0],
        })

        diccionario = pd.DataFrame([
            ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
            ["SAE UNO", "SAE UNO", "SAE UNO", "SAE UNO", "SAE UNO"],
        ])

        completado, avisos = nucleo.completar_ofertas_en_medidores(
            medidores, wxy, resumen, diccionario, registrar=lambda _: None,
        )

        self.assertEqual(avisos, [])
        # R: el dia 1 tiene oferta completa, el 2 no.
        self.assertEqual(
            list(completado["Oferta_Completa_Dia"]), [1, 1, 0, 0]
        )
        # S: 1 si R==1 al cambiar la ventana, 2 si no.
        self.assertEqual(
            list(completado["Indicador_Ventana_Oferta"]), [1, 1, 2, 2]
        )
        # T = 1 - Completa.
        self.assertEqual(list(completado["Ventana_No_Completa"]), [0, 0, 1, 1])

        # Y Medidores en si no las trae.
        self.assertNotIn("Ventana_No_Completa", medidores.columns)


class ConciliacionConNombresRealesTest(unittest.TestCase):

    def test_una_hoja_ya_renombrada_no_pasa_silenciosamente(self):
        """
        El error que reporto el usuario: la hoja llegaba con los nombres
        del Excel ("Descarga kWh") y con nombres de columna repetidos,
        asi que pandas tiraba un KeyError pelado. Ahora el mensaje dice
        que se esperaban los nombres internos.
        """

        renombrada = pd.DataFrame({
            "Descarga kWh": [1.0],
            "Carga kWh": [-1.0],
        })

        with self.assertRaises(KeyError) as caso:
            nucleo.conciliar_energia(
                pd.DataFrame({
                    "Gen_Unidad": [0.0], "Ventana_No_Completa": [1],
                }),
                renombrada, None, registrar=lambda _: None,
            )

        self.assertIn("Energia_Positiva", str(caso.exception))
        self.assertIn("INTERNOS", str(caso.exception))


if __name__ == "__main__":
    unittest.main()
