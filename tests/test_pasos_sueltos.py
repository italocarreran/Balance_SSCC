# -*- coding: utf-8 -*-
"""
Cada paso se corre solo, en el dia y el orden en que llega su entrada.

Pedido del usuario: "la informacion del proceso me llega de a poco
[...] corro los calculos de E Costos un dia, al dia siguiente el RE545
y despues la compensacion u otras cosas, por esto necesito que no de
error o problemas si corro cosas por separado", y "que no se abran
cosas innecesarias en cada paso".

Lo que se controla aca:
  - la hoja "Centrales" (la foto de Centrales.xlsx dentro de la
    planilla): se copia una vez y despues ningun paso abre
    Centrales.xlsx;
  - los pasos de calculo corridos de a uno dan lo mismo que leyendo
    Centrales.xlsx;
  - 'Ofertas SSCC' sola no abre Medidas_SAE ni el SoC;
  - la columna 'Central' (el nombre Balance_BESS) que se agrega al
    importar FD y Ofertas SSCC, y que el calculo usa sin homologar.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from Script import nucleo
from Script.nucleo import proceso
from Script.nucleo.parametros import HOJA_COMPENSACION_CENTRAL
from tests.test_control_corrida import PagosBessEndToEndTest


def armar_caso(carpeta):
    """El caso sintetico de test_control_corrida (un libro de antes de
    la hoja 'Centrales', con Centrales.xlsx en Auxiliares/)."""

    base, _ = PagosBessEndToEndTest.armar_caso(None, carpeta)
    return base


def correr_pagos(base, secciones, lineas):
    nucleo.generar_pagos_bess(
        base, set(secciones), registrar=lineas.append, aamm="2607"
    )


def hoja(base, nombre, header=1):
    return pd.read_excel(
        base / nucleo.ARCHIVO_SALIDA, sheet_name=nombre, header=header
    )


def dejar_prorrata(base):
    (base / "Prorrata retiros").mkdir(parents=True, exist_ok=True)
    pd.DataFrame({
        "Cuarto de Hora": [1, 1],
        "Suministrador": ["EMPRESA UNO", "EMPRESA DOS"],
        "Prorrata": [0.5, 0.5],
    }).to_excel(
        base / "Prorrata retiros" / "Prorrata_Retiros_2607_def.xlsx",
        sheet_name="Prorrata 15min", index=False,
    )


class TestHojaCentrales(unittest.TestCase):

    def test_se_copia_una_vez_y_despues_no_se_abre_centrales_xlsx(self):
        """
        El camino del usuario: E Costos un dia, RE545 al otro, despues
        la compensacion, la prorrata y el Resumen, cada uno con su
        boton. Con la hoja 'Centrales' ya copiada, Centrales.xlsx
        puede no estar y nada se rompe.
        """

        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)

            nucleo.generar_consolidado(
                base, "2607", {"centrales"}, registrar=lambda *a: None
            )
            self.assertIn(
                nucleo.HOJA_CENTRALES,
                pd.ExcelFile(base / nucleo.ARCHIVO_SALIDA).sheet_names,
            )

            (base / "Auxiliares" / nucleo.ARCHIVO_CENTRALES).unlink()
            dejar_prorrata(base)

            lineas = []
            for secciones in (
                {"ecostos"}, {"re545"}, {"compensacion_central"},
                {"prorrata_retiros"}, {"resumen"},
            ):
                try:
                    correr_pagos(base, secciones, lineas)
                except nucleo.ErrorEntrada as error:
                    self.fail(f"{secciones} solo fallo: {error}")

            texto = "\n".join(str(l) for l in lineas)
            self.assertNotIn(f"Leyendo {nucleo.ARCHIVO_CENTRALES}", texto)
            # La conciliacion solo corre cuando se recalcula una hoja de
            # calculo: la compensacion sola no la dispara (el TRA-010
            # falso que vio el usuario).
            self.assertNotIn("TRA-010", texto)

            hojas = pd.ExcelFile(base / nucleo.ARCHIVO_SALIDA).sheet_names
            for nombre in (
                nucleo.HOJA_CALCULO_ECOSTOS, nucleo.HOJA_CALCULO_RE545,
                HOJA_COMPENSACION_CENTRAL,
                nucleo.HOJA_PRORRATA_RETIROS, nucleo.HOJA_RESUMEN,
                nucleo.HOJA_CENTRALES,
            ):
                self.assertIn(nombre, hojas)

            # COMPENSACION_CENTRAL: detalle y total por central de cada
            # hoja de calculo, y al final el total por empresa.
            import openpyxl
            libro = openpyxl.load_workbook(
                base / nucleo.ARCHIVO_SALIDA, read_only=True
            )
            fila_titulos = next(libro[HOJA_COMPENSACION_CENTRAL].iter_rows(
                min_row=2, max_row=2, values_only=True
            ))
            libro.close()
            titulos = [t for t in fila_titulos if t]
            self.assertEqual(len(titulos), 5, titulos)
            self.assertIn("E Costos: total por central", titulos[1])
            self.assertIn("RE545: total por central", titulos[3])

    def test_da_lo_mismo_que_leyendo_centrales_xlsx(self):
        with tempfile.TemporaryDirectory() as uno, \
                tempfile.TemporaryDirectory() as dos:

            con_archivo = armar_caso(uno)
            correr_pagos(con_archivo, {"ecostos"}, [])
            correr_pagos(con_archivo, {"re545"}, [])

            con_hoja = armar_caso(dos)
            nucleo.generar_consolidado(
                con_hoja, "2607", {"centrales"},
                registrar=lambda *a: None,
            )
            (con_hoja / "Auxiliares" / nucleo.ARCHIVO_CENTRALES).unlink()
            correr_pagos(con_hoja, {"ecostos"}, [])
            correr_pagos(con_hoja, {"re545"}, [])

            for nombre in (
                nucleo.HOJA_CALCULO_ECOSTOS, nucleo.HOJA_CALCULO_RE545,
            ):
                pd.testing.assert_frame_equal(
                    hoja(con_archivo, nombre), hoja(con_hoja, nombre)
                )

    def test_un_libro_sin_la_hoja_usa_el_archivo_y_avisa(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)

            lineas = []
            correr_pagos(base, {"re545"}, lineas)

            texto = "\n".join(str(l) for l in lineas)
            self.assertIn(f"Leyendo {nucleo.ARCHIVO_CENTRALES}", texto)
            self.assertIn(
                f"todavia no tiene la hoja '{nucleo.HOJA_CENTRALES}'", texto
            )

    def test_sin_hoja_ni_archivo_el_error_dice_que_falta(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)
            (base / "Auxiliares" / nucleo.ARCHIVO_CENTRALES).unlink()

            with self.assertRaises(nucleo.ErrorEntrada) as error:
                correr_pagos(base, {"re545"}, [])

            self.assertIn(nucleo.HOJA_CENTRALES, str(error.exception))

    def test_los_dos_formatos_del_diccionario_van_y_vuelven_iguales(self):
        resumen = pd.DataFrame({
            "Nombre activo": ["SAE UNO", "SAE DOS"],
            "Propietario": ["E1", "E2"],
            "Pmax (MW)": [10.0, 5.5],
        })
        nuevo = pd.DataFrame([
            [None, None, None],
            ["Balance_BESS", "FD", "Subastas"],
            ["SAE UNO", "SAE 1", "BAT_UNO"],
        ])
        viejo = pd.DataFrame([
            ["FD", None, None, None, "Subastas", None, "ofertas"],
            ["SAE UNO", "SAE 1", None, None, "SAE UNO", "BAT_UNO", None],
            ["SAE DOS", "SAE 2", None, None, "SAE DOS", None, "Of dos"],
        ])

        for diccionario in (nuevo, viejo):
            with tempfile.TemporaryDirectory() as carpeta:
                ruta = Path(carpeta) / nucleo.ARCHIVO_SALIDA
                nucleo.escribir_salida(
                    None, ruta, [], [],
                    hojas_regenerar={nucleo.HOJA_CENTRALES},
                    centrales=(resumen, diccionario),
                    registrar=lambda *a: None,
                )

                leido_resumen, leido_dic = nucleo.leer_centrales_planilla(ruta)

                pd.testing.assert_frame_equal(resumen, leido_resumen)
                self.assertEqual(
                    nucleo.construir_homologacion(diccionario),
                    nucleo.construir_homologacion(leido_dic),
                )
                self.assertEqual(
                    nucleo.construir_dic_mapeo_diccionario(diccionario),
                    nucleo.construir_dic_mapeo_diccionario(leido_dic),
                )
                self.assertEqual(
                    nucleo._mapas_homologacion_fge(diccionario),
                    nucleo._mapas_homologacion_fge(leido_dic),
                )


class TestOfertasSola(unittest.TestCase):

    def test_ofertas_sola_lee_medidores_de_la_planilla(self):
        """
        Antes 'Ofertas SSCC' rehacia Medidores en memoria y para eso
        abria Medidas_SAE.xlsx y el SoC. Ahora toma la hoja ya escrita:
        en este caso ninguno de los dos archivos existe.
        """

        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)
            (base / "Ofertas").mkdir()
            (base / "Ofertas" / "Ofertas_OfertasSSCC_2607.xlsx").touch()

            recibido = {}

            def ofertas_falsas(df_medidores, archivo, diccionario, anio, mes,
                               registrar=print):
                recibido["filas"] = len(df_medidores)
                recibido["periodo"] = (anio, mes)
                return (
                    pd.DataFrame({
                        "Nombre": ["SAE UNO"], "Central": ["SAE UNO"],
                        "Dia": [1], "Oferta completa": [1],
                    }),
                    pd.DataFrame({
                        "Central": ["SAE UNO"], "Ventana T": [1],
                        "Oferta": [4], "Completa": [1],
                    }),
                    [],
                )

            lineas = []
            with mock.patch.object(
                proceso, "construir_ofertas_sscc", ofertas_falsas
            ):
                nucleo.generar_consolidado(
                    base, "2607", {"ofertas_sscc"},
                    registrar=lineas.append,
                )

            self.assertEqual(recibido["filas"], 8)
            self.assertEqual(recibido["periodo"], (2026, 7))

            texto = "\n".join(str(l) for l in lineas)
            self.assertNotIn(nucleo.ARCHIVO_MEDIDAS_SAE, texto)

            # El libro no tenia la hoja 'Centrales': se aprovecho la
            # lectura de Centrales.xlsx para dejarla escrita.
            self.assertIn(
                nucleo.HOJA_CENTRALES,
                pd.ExcelFile(base / nucleo.ARCHIVO_SALIDA).sheet_names,
            )


class TestColumnaCentral(unittest.TestCase):

    DICCIONARIO_NUEVO = pd.DataFrame([
        ["Balance_BESS", "FD", "Subastas", "Ofertas"],
        ["SAE UNO", "BESS UNO FD", "BAT_UNO", "Uno ofertas"],
    ])

    def test_la_unidad_del_fd_se_traduce_al_nombre_balance_bess(self):
        viejo = pd.DataFrame([
            ["FD", None],
            ["SAE UNO", "BESS UNO FD"],
        ])

        for diccionario in (self.DICCIONARIO_NUEVO, viejo):
            mapa = nucleo.mapa_unidad_fd_a_central(diccionario)
            self.assertEqual(mapa, {"bess uno fd": "SAE UNO"})

        fd = pd.DataFrame({"Unidad": ["BESS UNO FD", "OTRA"]})
        con_central = nucleo.agregar_central_fd(
            fd, nucleo.mapa_unidad_fd_a_central(self.DICCIONARIO_NUEVO)
        )
        self.assertEqual(
            con_central[nucleo.COLUMNA_CENTRAL].tolist(), ["SAE UNO", ""]
        )

    def test_la_hoja_fd_vuelve_con_su_columna_central(self):
        csf = pd.DataFrame(
            [["1BESS UNO FD", 1, 1, "2026-07-01", 1, "BESS UNO FD",
              1.0, 1.0, 1.0, 0.9, 0.9, 0.8, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CSF.values()), axis=1)
        cpf = pd.DataFrame(
            [["1BESS UNO FD", 1, 1, "2026-07-01", 1, "BESS UNO FD",
              1.0, 1.0, 1.0, 1.0, 0.7, "si", 0.7, 0.6, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CPF.values()), axis=1)

        mapa = nucleo.mapa_unidad_fd_a_central(self.DICCIONARIO_NUEVO)

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_SALIDA
            nucleo.escribir_salida(
                None, ruta, [], [],
                df_fd_csf=nucleo.agregar_central_fd(csf, mapa),
                df_fd_cpf=nucleo.agregar_central_fd(cpf, mapa),
                hojas_regenerar={nucleo.HOJA_FD},
                registrar=lambda *a: None,
            )

            leido_csf, leido_cpf = nucleo.leer_fd_consolidado(
                ruta, registrar=lambda *a: None
            )

        for leido, original in ((leido_csf, csf), (leido_cpf, cpf)):
            self.assertEqual(
                list(leido.columns),
                list(original.columns) + [nucleo.COLUMNA_CENTRAL],
            )
            self.assertEqual(leido[nucleo.COLUMNA_CENTRAL].tolist(), ["SAE UNO"])

        self.assertEqual(
            nucleo.dic_mapeo_desde_fd(leido_csf, leido_cpf),
            {"sae uno": "BESS UNO FD"},
        )

    def test_e_costos_usa_la_columna_central_sin_homologar(self):
        """
        El Diccionario de este caso NO tiene la columna FD: sin la
        columna 'Central' de la hoja FD, AM:AR quedarian vacias. Con
        ella, E Costos da lo mismo que con el Diccionario completo.
        """

        with tempfile.TemporaryDirectory() as uno, \
                tempfile.TemporaryDirectory() as dos:

            referencia = armar_caso(uno)
            correr_pagos(referencia, {"ecostos"}, [])

            base = armar_caso(dos)
            ruta = base / nucleo.ARCHIVO_SALIDA
            csf, cpf = nucleo.leer_fd_consolidado(
                ruta, registrar=lambda *a: None
            )
            resumen, _ = nucleo.leer_centrales(
                base / "Auxiliares" / nucleo.ARCHIVO_CENTRALES
            )
            sin_fd = pd.DataFrame([
                ["Balance_BESS", "Subastas", "Ofertas"],
                ["SAE UNO", "SAE UNO", "SAE UNO"],
            ])
            mapa = {"sae uno": "SAE UNO"}
            nucleo.escribir_salida(
                None, ruta, [], [],
                df_fd_csf=nucleo.agregar_central_fd(csf, mapa),
                df_fd_cpf=nucleo.agregar_central_fd(cpf, mapa),
                hojas_regenerar={nucleo.HOJA_FD, nucleo.HOJA_CENTRALES},
                centrales=(resumen, sin_fd),
                ruta_existente=ruta,
                registrar=lambda *a: None,
            )

            lineas = []
            correr_pagos(base, {"ecostos"}, lineas)

            texto = "\n".join(str(l) for l in lineas)
            self.assertNotIn("DIC-001", texto)
            pd.testing.assert_frame_equal(
                hoja(referencia, nucleo.HOJA_CALCULO_ECOSTOS),
                hoja(base, nucleo.HOJA_CALCULO_ECOSTOS),
            )

    def test_ofertas_usa_la_columna_central_sin_diccionario(self):
        df_wxy = pd.DataFrame({
            "Nombre": ["Uno ofertas", "Uno ofertas"],
            "Dia": [1, 2],
            "Oferta completa": [1, 0],
        })

        con_central = nucleo.agregar_central_ofertas(
            df_wxy, self.DICCIONARIO_NUEVO
        )
        self.assertEqual(
            list(con_central.columns),
            ["Nombre", nucleo.COLUMNA_CENTRAL, "Dia", "Oferta completa"],
        )
        self.assertEqual(
            con_central[nucleo.COLUMNA_CENTRAL].tolist(), ["SAE UNO"] * 2
        )

        medidores = pd.DataFrame({
            "Dia": [1, 2], "clave": ["SAE UNO", "SAE UNO"],
        })

        # Sin Diccionario: sale directo de la columna.
        r, _ = nucleo.calcular_r(
            medidores, con_central, None, registrar=lambda *a: None
        )
        # Y es lo mismo que homologando con el Diccionario (libro viejo).
        r_viejo, _ = nucleo.calcular_r(
            medidores, df_wxy, self.DICCIONARIO_NUEVO,
            registrar=lambda *a: None,
        )

        self.assertEqual(r.tolist(), [1, 0])
        self.assertEqual(r.tolist(), r_viejo.tolist())


if __name__ == "__main__":
    unittest.main()


def escribir_sscc_desempeno(ruta):
    """
    Un SSCC_Desempeño_* sintetico con el layout real: encabezados en la
    fila 11 y datos desde la 12, CPF Horario en B:J y CSF Horario en
    B:H. Trae una unidad BESS, una SAE y una termica (que la hoja FD
    filtra), y horas en que la BESS "No Participó" en CSF.
    """

    filas_cpf, filas_csf = [], []

    for dia in (1, 2):
        for hora in range(24):
            fecha = pd.Timestamp(2026, 7, dia)
            for unidad, base in (
                ("BESS PFV UNO", 0.9), ("SAE DOS", 0.8), ("TERMICA TG", 0.5),
            ):
                fd = round(base - hora / 1000 - dia / 100, 6)
                filas_cpf.append(
                    [None, fecha, hora, unidad, 1, 1, 1, 1, fd, "si"]
                )
                respuesta = (
                    "No Participó" if unidad == "BESS PFV UNO" and hora in (3, 4)
                    else "Participó"
                )
                filas_csf.append(
                    [None, fecha, hora, unidad, respuesta, 1, 1, fd - 0.05]
                )

    vacias = [[None] * 10 for _ in range(11)]

    with pd.ExcelWriter(ruta) as writer:
        pd.DataFrame(vacias + filas_cpf).to_excel(
            writer, sheet_name="CPF Horario", index=False, header=False
        )
        pd.DataFrame(
            [fila[:8] for fila in vacias] + filas_csf
        ).to_excel(
            writer, sheet_name="CSF Horario", index=False, header=False
        )


class TestFdYFmaSeAplicanEnRe545(unittest.TestCase):
    """
    Pedido del usuario: el FD y el FMA salen de la hoja Subastas y se
    calculan y aplican en el calculo que los usa (RE545), para que no
    haya un descuento doble. El valor tiene que ser EXACTAMENTE el que
    antes quedaba escrito en Subastas!P/Q: solo cambia de donde sale.
    """

    def test_el_fd_de_la_hoja_fd_es_el_mismo_que_el_del_archivo(self):
        from Script.nucleo.externos import desempeno_fd

        with tempfile.TemporaryDirectory() as carpeta:
            archivo = Path(carpeta) / "SSCC_Desempeño_Julio_2026_V2.xlsx"
            escribir_sscc_desempeno(archivo)

            # Antes: Subastas leia el archivo entero.
            del_archivo = desempeno_fd.construir_tablas_fd(
                archivo, registrar=lambda *a: None
            )

            # Ahora: la hoja FD de la planilla (ida y vuelta por Excel).
            csf, cpf = nucleo.construir_fd(archivo, registrar=lambda *a: None)
            ruta = Path(carpeta) / nucleo.ARCHIVO_SALIDA
            nucleo.escribir_salida(
                None, ruta, [], [], df_fd_csf=csf, df_fd_cpf=cpf,
                hojas_regenerar={nucleo.HOJA_FD}, registrar=lambda *a: None,
            )
            csf, cpf = nucleo.leer_fd_consolidado(
                ruta, registrar=lambda *a: None
            )
            de_la_hoja = nucleo.tablas_fd_desde_hoja(
                csf, cpf, registrar=lambda *a: None
            )

        # La hoja FD solo guarda BESS/SAE: la termica queda afuera, y
        # es lo unico que cambia (las subastas tambien son solo BESS/SAE).
        def solo_bess(dic):
            return {
                clave: valor for clave, valor in dic.items()
                if "bess" in clave[0] or "sae" in clave[0]
            }

        for tabla in ("cpf", "csf", "participacion_csf"):
            self.assertEqual(
                solo_bess(del_archivo[tabla]), de_la_hoja[tabla], tabla
            )
            self.assertTrue(de_la_hoja[tabla], tabla)

        # Las dos horas en que la BESS no participo quedan en 0.
        self.assertEqual(
            de_la_hoja["participacion_csf"][("bess pfv uno", 4)], 0
        )

    def test_subastas_vieja_con_fd_y_fma_escritos_no_se_usa(self):
        """
        Un libro de antes trae Subastas!FD y FMA ya calculados. RE545
        los reemplaza por los que calcula: un solo origen, nunca dos.
        """

        subastas = pd.DataFrame({
            "Concepto": ["CPF(+)", "CSF(-)"],
            "Control": ["CPF", "CSF"],
            "Sub_Baj": ["SUBIDA", "BAJADA"],
            "Fecha": ["2026-07-01"] * 2,
            "Año": [2026, 2026], "Mes": [7, 7], "Dia": [1, 1],
            "Hora_dia": [1, 2], "Hora_mes": [1, 2],
            "Configuración": ["SAE UNO"] * 2,
            "Propietario": ["E1"] * 2,
            "Clave horaria": ["x", "y"], "Ciclo": [1, 1],
            "Energía SSCC": [10.0, 20.0],
            "FD": [99.0, 99.0], "FMA": [99.0, 99.0],
        })
        csf = pd.DataFrame(
            [["2SAE UNO", 2, 1, "2026-07-01", 1, "SAE UNO",
              "Participó", 1.0, 1.0, 0.8, 0.8, 0.8, 2]]
        ).set_axis(list(nucleo.NOMBRES_FD_CSF.values()), axis=1)
        cpf = pd.DataFrame(
            [["1SAE UNO", 1, 1, "2026-07-01", 0, "SAE UNO",
              1.0, 1.0, 1.0, 1.0, 0.7, "si", 0.7, 0.7, 1]]
        ).set_axis(list(nucleo.NOMBRES_FD_CPF.values()), axis=1)
        tablas_fma = {
            "cpf": pd.DataFrame({
                "Central": ["SAE UNO"], "Año": [2026], "Mes": [7],
                "Dia": [1], "Hora": [1],
                "FMA_CPF_mas": [0.5], "FMA_CPF_menos": [0.4],
            }),
            "csf": pd.DataFrame({
                "Año": [2026], "Mes": [7], "Dia": [1], "Hora": [2],
                "FMA_CSF_mas_base": [0.3], "FMA_CSF_menos_base": [0.2],
            }),
            "ctf_unidad": None, "ctf_configuracion": None,
        }

        con_factores = nucleo.subastas_con_fd_y_fma(
            subastas, csf, cpf, tablas_fma, None, registrar=lambda *a: None
        )

        self.assertEqual(con_factores["FD"].tolist(), [0.7, 0.8])
        self.assertEqual(con_factores["FMA"].tolist(), [0.5, 0.2])

        energia, fd, fma = nucleo.construir_dic_reservas_subastas(con_factores)
        self.assertEqual(sorted(fd.values()), [0.7, 0.8])
        self.assertEqual(sorted(energia.values()), [10.0, 20.0])

    def test_la_hoja_fma_va_y_vuelve(self):
        tablas = {
            "cpf": pd.DataFrame({
                "Central": ["A", "B"], "Año": [2026, 2026], "Mes": [7, 7],
                "Dia": [1, 1], "Hora": [1, 2],
                "FMA_CPF_mas": [0.5, 0.6], "FMA_CPF_menos": [0.4, 0.3],
            }),
            "csf": pd.DataFrame({
                "Año": [2026], "Mes": [7], "Dia": [1], "Hora": [2],
                "FMA_CSF_mas_base": [0.3], "FMA_CSF_menos_base": [0.2],
            }),
            "ctf_unidad": None,
            "ctf_configuracion": None,
        }

        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_SALIDA
            nucleo.escribir_salida(
                None, ruta, [], [], hojas_regenerar={nucleo.HOJA_FMA},
                tablas_fma=tablas, registrar=lambda *a: None,
            )
            leidas = nucleo.leer_fma_consolidado(
                ruta, registrar=lambda *a: None
            )

        for clave in ("cpf", "csf"):
            pd.testing.assert_frame_equal(
                tablas[clave], leidas[clave], check_dtype=False
            )
        self.assertIsNone(leidas["ctf_unidad"])

    def test_re545_sin_hoja_fma_dice_que_hay_que_generarla(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = armar_caso(carpeta)
            ruta = base / nucleo.ARCHIVO_SALIDA
            hojas = pd.read_excel(ruta, sheet_name=None, header=None)
            del hojas[nucleo.HOJA_FMA]
            with pd.ExcelWriter(ruta) as writer:
                for nombre, df in hojas.items():
                    df.to_excel(
                        writer, sheet_name=nombre, index=False, header=False
                    )

            with self.assertRaises(nucleo.ErrorEntrada) as error:
                correr_pagos(base, {"re545"}, [])

            self.assertIn(nucleo.HOJA_FMA, str(error.exception))
