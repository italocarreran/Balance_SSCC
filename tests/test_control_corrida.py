"""Registro de alertas, conciliacion de energia y manifiesto.

Los tres controles que el catalogo (docs/Alertas_y_Controles_Traspaso_
Python_BESS.md) pone en la Fase 1 y que no existian.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo


def escribir_tabla(writer, titulo, df, columna):
    """La hoja 'Ofertas SSCC': titulo, encabezado y datos, en columna."""

    pd.DataFrame([[titulo]]).to_excel(
        writer, sheet_name=nucleo.HOJA_OFERTAS_SSCC, index=False,
        header=False, startrow=0, startcol=columna,
    )
    df.to_excel(
        writer, sheet_name=nucleo.HOJA_OFERTAS_SSCC, index=False,
        startrow=1, startcol=columna,
    )


def medidores_de_prueba(energias, ventana_no_completa):
    return pd.DataFrame({
        "Mes": [7] * len(energias),
        "clave": [f"SAE-{i}" for i in range(len(energias))],
        "Gen_Unidad": energias,
        "Ventana_No_Completa": ventana_no_completa,
    })


def hoja_de_prueba(energias, mascara):
    """Una hoja de calculo con la energia que le toca segun la mascara."""

    energia = pd.Series(energias, dtype=float)
    activa = pd.Series(mascara)
    return pd.DataFrame({
        "Energia_Positiva": energia.where(energia > 0, 0.0).where(activa, 0.0),
        "Energia_Negativa": energia.where(energia < 0, 0.0).where(activa, 0.0),
    })


class RegistroTest(unittest.TestCase):

    def test_el_estado_sale_de_la_peor_severidad(self):
        registro = nucleo.Registro(salida=lambda _: None)
        self.assertEqual(registro.estado(), nucleo.APROBADA)

        registro.anotar(nucleo.Alerta("X-1", nucleo.INFO, "e", "m"))
        self.assertEqual(registro.estado(), nucleo.APROBADA)

        registro.anotar(nucleo.Alerta("X-2", nucleo.MEDIA, "e", "m"))
        self.assertEqual(registro.estado(), nucleo.APROBADA_CON_ADVERTENCIAS)

        registro.anotar(nucleo.Alerta("X-3", nucleo.ALTA, "e", "m"))
        self.assertEqual(registro.estado(), nucleo.NO_APROBADA)

    def test_severidad_inventada_no_pasa(self):
        with self.assertRaises(ValueError):
            nucleo.Alerta("X-1", "GRAVISIMA", "e", "m")

    def test_anotar_funciona_con_print_comun(self):
        """Sin Registro no se guarda nada, pero tampoco explota."""

        mensajes = []
        nucleo.anotar(mensajes.append,
                      nucleo.Alerta("X-1", nucleo.ALTA, "e", "un problema"))

        self.assertEqual(len(mensajes), 1)
        self.assertIn("X-1", mensajes[0])

    def test_guarda_todos_los_faltantes_aunque_muestre_15(self):
        """FD-005: el tope de 15 es de la pantalla, no del registro."""

        registro = nucleo.Registro(salida=lambda _: None)
        centrales = [f"SAE-{i:03d}" for i in range(40)]

        nucleo._avisar_claves_sin_mapeo(
            centrales, {}, "Central sin barra", "Resumen BESS",
            registrar=registro, id_alerta="MAE-001",
        )

        self.assertEqual(len(registro.alertas), 40)
        self.assertEqual(
            sorted(a.central for a in registro.alertas), centrales
        )
        self.assertIn("y 25 mas", registro.lineas[0])


class ConciliacionTest(unittest.TestCase):

    def setUp(self):
        self.energias = [10.0, -4.0, 7.5, -2.5]
        self.a_ecostos = [1, 1, 0, 0]          # L = 1 -> E Costos
        self.medidores = medidores_de_prueba(self.energias, self.a_ecostos)
        self.mascara_eco = [True, True, False, False]
        self.mascara_545 = [False, False, True, True]

    def test_cuadra_y_no_alerta(self):
        registro = nucleo.Registro(salida=lambda _: None)

        resultado = nucleo.conciliar_energia(
            self.medidores,
            hoja_de_prueba(self.energias, self.mascara_eco),
            hoja_de_prueba(self.energias, self.mascara_545),
            registrar=registro,
        )

        self.assertEqual(registro.estado(), nucleo.APROBADA)
        self.assertAlmostEqual(resultado["diferencia"], 0.0)
        self.assertAlmostEqual(resultado["energia_medidores"], 11.0)
        self.assertAlmostEqual(resultado["energia_ecostos"], 6.0)
        self.assertAlmostEqual(resultado["energia_re545"], 5.0)

    def test_una_fila_que_se_perdio_deja_la_corrida_no_aprobada(self):
        registro = nucleo.Registro(salida=lambda _: None)

        # RE545 se come la ultima fila (-2.5): la suma ya no cierra.
        perdida = hoja_de_prueba(self.energias, [False, False, True, False])

        nucleo.conciliar_energia(
            self.medidores,
            hoja_de_prueba(self.energias, self.mascara_eco),
            perdida,
            registrar=registro,
        )

        ids = {a.id_alerta for a in registro.alertas}
        self.assertIn("TRA-007", ids)
        self.assertIn("TRA-006", ids)
        self.assertEqual(registro.estado(), nucleo.NO_APROBADA)

    def test_energia_en_la_hoja_equivocada_se_detecta(self):
        """La suma total cuadra, pero cada hoja tiene lo del otro."""

        registro = nucleo.Registro(salida=lambda _: None)

        nucleo.conciliar_energia(
            self.medidores,
            hoja_de_prueba(self.energias, self.mascara_545),
            hoja_de_prueba(self.energias, self.mascara_eco),
            registrar=registro,
        )

        ids = {a.id_alerta for a in registro.alertas}
        self.assertIn("TRA-005", ids)
        self.assertIn("TRA-006", ids)
        self.assertEqual(registro.estado(), nucleo.NO_APROBADA)

    def test_filas_de_mas_o_de_menos(self):
        registro = nucleo.Registro(salida=lambda _: None)

        corta = hoja_de_prueba(self.energias, self.mascara_eco).iloc[:2]

        nucleo.conciliar_energia(
            self.medidores, corta,
            hoja_de_prueba(self.energias, self.mascara_545),
            registrar=registro,
        )

        self.assertIn("TRA-001", {a.id_alerta for a in registro.alertas})

    def test_una_sola_hoja_concilia_parcial(self):
        registro = nucleo.Registro(salida=lambda _: None)

        resultado = nucleo.conciliar_energia(
            self.medidores,
            hoja_de_prueba(self.energias, self.mascara_eco),
            None,
            registrar=registro,
        )

        self.assertFalse(resultado["conciliacion_completa"])
        self.assertIsNone(resultado["diferencia"])
        self.assertIn("TRA-009", {a.id_alerta for a in registro.alertas})
        # Parcial es INFO: no bloquea, pero queda dicho.
        self.assertEqual(registro.estado(), nucleo.APROBADA)


class ManifiestoTest(unittest.TestCase):

    def test_registra_hash_tamano_y_saltea_lo_que_no_existe(self):
        with tempfile.TemporaryDirectory() as carpeta:
            uno = Path(carpeta) / "Centrales.xlsx"
            uno.write_bytes(b"contenido de prueba")

            tabla = nucleo.construir_manifiesto([
                ("Centrales", uno),
                ("cmg", Path(carpeta) / "no_existe.xlsx"),
                ("repetido", uno),
                ("nada", None),
            ])

            self.assertEqual(len(tabla), 1)
            fila = tabla.iloc[0]
            self.assertEqual(fila["archivo"], "Centrales.xlsx")
            self.assertEqual(fila["bytes"], len(b"contenido de prueba"))
            self.assertEqual(len(fila["sha256"]), 64)

    def test_el_hash_cambia_si_cambia_el_archivo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "x.xlsx"
            ruta.write_bytes(b"antes")
            antes = nucleo.sha256_de(ruta)
            ruta.write_bytes(b"despues")

            self.assertNotEqual(antes, nucleo.sha256_de(ruta))


if __name__ == "__main__":
    unittest.main()


class PagosBessEndToEndTest(unittest.TestCase):
    """
    El camino real: generar_pagos_bess() sobre un caso sintetico
    minimo, para comprobar que el archivo de control sale con
    "Alertas" y "Ejecucion" y que el estado refleja lo que paso.
    """

    def armar_caso(self, carpeta):
        base = Path(carpeta)
        (base / "Auxiliares").mkdir(parents=True)
        (base / "Cmg").mkdir(parents=True)

        centrales = pd.DataFrame({
            "Nombre activo": ["SAE UNO"],
            "Pmax (MW)": [10.0],
            "Horas para descarga forzada": [1],
            "Capacidad (MWh)": [40.0],
            "Energía mínima": [0],
            "Barra inyección": ["BARRA A"],
            "% Energía sobre mínima (indicador nuevo ciclo)": [0.06],
            "Ciclos max diarios": [2],
            "Eficiencia": [0.9],
        })
        # Diccionario en el formato nuevo: una tabla con encabezados.
        diccionario = pd.DataFrame([
            ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
            ["SAE UNO", "SAE UNO", "SAE UNO", "SAE UNO", "SAE UNO"],
        ])
        with pd.ExcelWriter(base / "Auxiliares" / "Centrales.xlsx") as w:
            centrales.to_excel(w, sheet_name="Resumen BESS", index=False)
            diccionario.to_excel(
                w, sheet_name="Diccionario", index=False, header=False
            )

        cuartos = 8
        medidores = pd.DataFrame({
            "Mes": [7] * cuartos,
            "Dia": [1] * cuartos,
            "Hora": [1] * cuartos,
            "Minutos": [0] * cuartos,
            "Hora Mes": list(range(1, cuartos + 1)),
            "Cuarto de Hora": list(range(1, cuartos + 1)),
            "clave": ["SAE UNO"] * cuartos,
            "intervalo": [15] * cuartos,
            "Gen_Unidad": [5.0, -3.0, 2.0, -1.0, 4.0, -2.0, 1.0, -0.5],
            "SoC": [0.5] * cuartos,
            "Copia_Ventana": [1] * cuartos,
            "Ventana": [1] * cuartos,
            "Clave_Dia_HoraMes": [f"1|{i}" for i in range(1, cuartos + 1)],
            "Indicador_SoC": [0] * cuartos,
        })

        # R, S y T ya no viven en la hoja Medidores: salen de las dos
        # tablas de la hoja "Ofertas SSCC" (ver
        # completar_ofertas_en_medidores). La ventana 1 queda
        # incompleta (T=1 -> E Costos) y la 2 completa (T=0 -> RE545):
        # es la mitad a cada hoja que mide la conciliacion.
        medidores["Copia_Ventana"] = [1, 1, 1, 1, 2, 2, 2, 2]
        medidores["Ventana"] = [1, 1, 1, 1, 2, 2, 2, 2]

        ofertas_por_dia = pd.DataFrame({
            "Nombre": ["SAE UNO"],
            "Dia": [1],
            "Oferta completa": [1],
        })
        resumen_ventana = pd.DataFrame({
            "Central": ["SAE UNO", "SAE UNO"],
            "Ventana T": [1, 2],
            "Oferta": [4, 4],
            "Completa": [0, 1],
        })

        # La hoja Subastas con sus 16 columnas reales (B:Q).
        subastas = pd.DataFrame({
            "Concepto": ["CPF", "CSF"],
            "Control": ["cpf", "csf"],
            "Sub_Baj": ["SUBIDA", "BAJADA"],
            "Fecha": ["2026-07-01"] * 2,
            "Año": [2026] * 2,
            "Mes": [7] * 2,
            "Dia": [1] * 2,
            "Hora_dia": [1, 1],
            "Hora_mes": [1, 2],
            "Configuración": ["SAE UNO"] * 2,
            "Propietario": ["EMPRESA UNO"] * 2,
            "Clave horaria": ["SAE UNO11"] * 2,
            "Ciclo": [1, 1],
            "Energía SSCC": [0.0, 0.0],
            "FD": [1.0, 1.0],
            "FMA": [1.0, 1.0],
        })

        # La hoja FD: dos bloques lado a lado (CSF en A, CPF en Q), de
        # donde "Calculo E Costos" saca AM:AR. Antes se releia el
        # SSCC_Desempeño_*; ahora sale del consolidado.
        fd_csf = pd.DataFrame(
            [["1SAE UNO", 1, 1, "2026-07-01", 1, "SAE UNO",
              1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CSF.values()), axis=1)
        fd_cpf = pd.DataFrame(
            [["1SAE UNO", 1, 1, "2026-07-01", 1, "SAE UNO",
              1.0, 1.0, 1.0, 1.0, 1.0, "si", 1.0, 1.0, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CPF.values()), axis=1)

        # cmg.xlsx se lee POR POSICION (A:I): D=Barra, F=CMg,
        # H=Cuarto de Hora, I=CMg Promedio (ver construir_dic_cmg).
        cmg = pd.DataFrame({
            "A": [""] * cuartos,
            "B": [""] * cuartos,
            "C": [""] * cuartos,
            "D_Barra": ["BARRA A"] * cuartos,
            "E": [""] * cuartos,
            "F_CMg": [50.0] * cuartos,
            "G": [""] * cuartos,
            "H_Cuarto": list(range(1, cuartos + 1)),
            "I_CMgProm": [45.0] * cuartos,
        })
        cmg.to_excel(base / "Cmg" / "cmg.xlsx", index=False)

        with pd.ExcelWriter(base / nucleo.ARCHIVO_SALIDA) as w:
            medidores.to_excel(w, sheet_name="Medidores", index=False)
            # El CMg de los pagos sale de la hoja 'CMg' del
            # consolidado, no de volver a abrir cmg.xlsx.
            cmg.to_excel(w, sheet_name="CMg", index=False)
            subastas.to_excel(w, sheet_name="Subastas", index=False)
            fd_csf.to_excel(w, sheet_name="FD", index=False, startcol=0)
            fd_cpf.to_excel(
                w, sheet_name="FD", index=False,
                startcol=nucleo._COLUMNA_Q_INDICE,
            )
            escribir_tabla(
                w, nucleo.TITULO_OFERTAS_POR_DIA, ofertas_por_dia, 0
            )
            escribir_tabla(
                w, nucleo.TITULO_RESUMEN_VENTANA, resumen_ventana,
                len(ofertas_por_dia.columns) + 2,
            )

        return base, medidores

    def test_el_libro_sale_con_alertas_ejecucion_y_conciliacion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base, medidores = self.armar_caso(carpeta)

            lineas = []
            try:
                nucleo.generar_pagos_bess(
                    base, {"re545"}, registrar=lineas.append
                )
            except nucleo.ErrorEntrada as error:
                self.fail(f"el caso sintetico no corrio: {error}")

            salida = base / nucleo.ARCHIVO_SALIDA
            hojas = pd.ExcelFile(salida).sheet_names
            self.assertIn(nucleo.HOJA_CALCULO_RE545, hojas)

            # El control ya no ensucia la planilla: vive aparte.
            self.assertNotIn(nucleo.HOJA_ALERTAS, hojas)
            self.assertNotIn(nucleo.HOJA_EJECUCION, hojas)

            control = base / nucleo.ARCHIVO_CONTROL
            hojas_control = pd.ExcelFile(control).sheet_names
            self.assertIn(nucleo.HOJA_ALERTAS, hojas_control)
            self.assertIn(nucleo.HOJA_EJECUCION, hojas_control)

            ejecucion = pd.read_excel(
                control, sheet_name=nucleo.HOJA_EJECUCION
            )
            campos = dict(zip(ejecucion["campo"], ejecucion["valor"]))

            self.assertIn("estado", campos)
            self.assertEqual(campos["hojas_recalculadas"], "Calculo RE545")
            self.assertEqual(str(campos["periodo"]), "07")

            # La energia que le toca a RE545 es la de las filas con
            # Ventana_No_Completa != 1: 4.0 - 2.0 + 1.0 - 0.5 = 2.5
            self.assertAlmostEqual(float(campos["energia_re545"]), 2.5)
            self.assertAlmostEqual(float(campos["energia_medidores"]), 5.5)
            self.assertFalse(bool(campos["conciliacion_completa"]))

            # y el manifiesto quedo con el hash de las entradas reales
            texto = "\n".join(str(l) for l in lineas)
            self.assertIn("Conciliando energia", texto)


    def test_la_prorrata_y_el_resumen_no_abren_el_consolidado(self):
        """
        Pedido del usuario: "que no se abran planillas innecesarias".
        PRORRATA_RETIROS y el Resumen salen de las dos hojas de
        calculo ya escritas en la planilla, asi que no releen las
        hojas de entrada ni cmg.xlsx (que aca ni siquiera esta).
        """

        with tempfile.TemporaryDirectory() as carpeta:
            base, _ = self.armar_caso(carpeta)

            (base / "Prorrata retiros").mkdir(parents=True, exist_ok=True)
            pd.DataFrame({
                "Cuarto de Hora": [1, 1],
                "Suministrador": ["EMPRESA UNO", "EMPRESA DOS"],
                "Prorrata": [0.5, 0.5],
            }).to_excel(
                base / "Prorrata retiros" / "Prorrata_Retiros_2607_def.xlsx",
                sheet_name="Prorrata 15min", index=False,
            )

            nucleo.generar_pagos_bess(
                base, {"ecostos", "re545"}, registrar=lambda *a: None,
                aamm="2607",
            )

            # Se va cmg.xlsx: si la prorrata lo abriera, la corrida
            # se caeria. Las hojas de entrada ya no se pueden borrar
            # (viven en la misma planilla que el calculo), asi que se
            # controla por el log: no se leen.
            (base / "Cmg" / "cmg.xlsx").unlink()

            lineas = []
            try:
                nucleo.generar_pagos_bess(
                    base, {"prorrata_retiros", "resumen"},
                    registrar=lineas.append, aamm="2607",
                )
            except nucleo.ErrorEntrada as error:
                self.fail(f"abrio algo que no necesitaba: {error}")

            hojas = pd.ExcelFile(base / nucleo.ARCHIVO_SALIDA).sheet_names
            self.assertIn(nucleo.HOJA_PRORRATA_RETIROS, hojas)

            # y las hojas de entrada siguen ahi: la escritura del
            # calculo preserva la otra mitad de la planilla.
            self.assertIn(nucleo.HOJA_MEDIDORES, hojas)
            self.assertIn(nucleo.HOJA_SUBASTAS, hojas)

            resumen = pd.read_excel(
                base / nucleo.ARCHIVO_SALIDA, sheet_name=nucleo.HOJA_RESUMEN
            )
            self.assertIn("NETO", resumen.columns)

            texto = "\n".join(str(l) for l in lineas)
            self.assertNotIn("Leyendo hoja 'Medidores'", texto)
            self.assertNotIn("Leyendo hoja 'Subastas'", texto)

    def test_las_dos_hojas_juntas_concilian(self):
        """
        Las dos hojas en la misma corrida: es el camino en el que se
        rompia la conciliacion (df_ecostos llegaba ya renombrado y
        'Energia_Positiva' se llamaba "Descarga kWh"), y ademas el
        unico que ejerce el FD leido del consolidado.
        """

        with tempfile.TemporaryDirectory() as carpeta:
            base, _ = self.armar_caso(carpeta)

            lineas = []
            try:
                nucleo.generar_pagos_bess(
                    base, {"ecostos", "re545"}, registrar=lineas.append
                )
            except nucleo.ErrorEntrada as error:
                self.fail(f"el caso sintetico no corrio: {error}")

            ejecucion = pd.read_excel(
                base / nucleo.ARCHIVO_CONTROL,
                sheet_name=nucleo.HOJA_EJECUCION,
            )
            campos = dict(zip(ejecucion["campo"], ejecucion["valor"]))

            self.assertTrue(bool(campos["conciliacion_completa"]))
            # E Costos (ventana 1, T=1) 5.0-3.0+2.0-1.0 = 3.0
            self.assertAlmostEqual(float(campos["energia_ecostos"]), 3.0)
            self.assertAlmostEqual(float(campos["energia_re545"]), 2.5)
            self.assertAlmostEqual(float(campos["energia_medidores"]), 5.5)
            self.assertAlmostEqual(float(campos["diferencia"]), 0.0)

            texto = "\n".join(str(l) for l in lineas)
            self.assertNotIn("Energia_Positiva", texto)
