# -*- coding: utf-8 -*-
"""
El plan de "Ejecutar todo": que se hace, que no, en que orden y que
bloquea el boton.

No corre ningun calculo real: el grafo se ejercita con filas de estado
armadas a mano (las que devuelve revisar_estructura) y con grupos
falsos que solo anotan cuando corrieron.
"""

import sys
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import Script.nucleo as nucleo  # noqa: E402
from Script.nucleo import orquestador  # noqa: E402


# Ids de fila de revisar_estructura() que son ENTRADAS de la persona:
# con estas cuatro en "ok" el caso arranca.
ENTRADAS_OK = {
    "centrales": "ok",
    "centrales:Resumen BESS": "ok",
    "centrales:Diccionario": "ok",
    "homologacion": "ok",
    "homologacion:homol": "ok",
    "soc": "ok",
    "ofertas": "ok",
    "prorrata_archivo": "ok",
}


def filas_de(estados):
    return [
        {"id": id_fila, "etiqueta": id_fila, "nivel": 1,
         "estado": estado, "detalle": ""}
        for id_fila, estado in estados.items()
    ]


def caso_vacio(**cambios):
    """Entradas puestas, nada generado todavia."""

    estados = dict(ENTRADAS_OK)
    for tarea in orquestador.TAREAS:
        estados.setdefault(tarea.fila_estado, "falta")
    estados.update(cambios)
    return filas_de(estados)


def caso_completo(**cambios):
    """Todo generado y al dia."""

    estados = dict(ENTRADAS_OK)
    for tarea in orquestador.TAREAS:
        estados[tarea.fila_estado] = "ok"
    estados.update(cambios)
    return filas_de(estados)


class TestGrafoCompleto(unittest.TestCase):
    """
    Que ninguna hoja se quede afuera del boton "Ejecutar todo".

    Paso de verdad al juntar dos ramas: COMPENSACION_CENTRAL se agrego
    en una y el grafo en la otra, asi que la hoja tenia su boton de
    fila pero "Ejecutar todo" no la generaba nunca. Esta prueba es para
    que la proxima hoja que se agregue no pueda repetirlo.
    """

    def test_cada_hoja_de_las_dos_salidas_tiene_su_tarea(self):
        for grupo, secciones in (
            ("consolidado", nucleo.SECCIONES_CONSOLIDADO),
            ("pagos", nucleo.SECCIONES_PAGOS),
        ):
            con_tarea = {
                tarea.seccion for tarea in orquestador.TAREAS
                if tarea.grupo == grupo
            }
            with self.subTest(grupo=grupo):
                self.assertEqual(
                    {seccion[0] for seccion in secciones} - con_tarea, set()
                )

    def test_la_compensacion_va_despues_de_las_dos_hojas_de_calculo(self):
        tarea = orquestador.TAREA_POR_ID["pagos:compensacion_central"]

        self.assertEqual(
            set(tarea.depende), {"pagos:ecostos", "pagos:re545"}
        )

    def test_el_plan_de_cero_incluye_la_compensacion(self):
        plan = nucleo.planificar("/caso", "2607", filas=caso_vacio())
        seleccionadas = {t["id"] for t in plan["tareas"] if t["seleccionada"]}

        self.assertIn("pagos:compensacion_central", seleccionadas)


class TestPlan(unittest.TestCase):

    def test_de_cero_se_hace_todo_y_el_boton_se_puede_apretar(self):
        plan = nucleo.planificar("/caso", "2607", filas=caso_vacio())

        self.assertTrue(plan["puede_ejecutar"], plan["bloqueos"])
        self.assertEqual(
            {t["id"] for t in plan["tareas"] if t["seleccionada"]},
            {t.id for t in orquestador.TAREAS},
        )

    def test_con_todo_al_dia_no_hay_nada_que_hacer(self):
        plan = nucleo.planificar("/caso", "2607", filas=caso_completo())

        self.assertEqual(
            [t["id"] for t in plan["tareas"] if t["seleccionada"]], []
        )
        self.assertFalse(plan["puede_ejecutar"])
        self.assertIn("No hay nada tildado", " ".join(plan["bloqueos"]))
        self.assertTrue(
            all(t["estado"] == "al dia" for t in plan["tareas"])
        )

    def test_solo_se_hace_lo_que_falta(self):
        """
        El caso normal del boton: ya esta casi todo, falta el Resumen.
        """

        filas = caso_completo(**{"pagos:resumen": "pendiente"})
        plan = nucleo.planificar("/caso", "2607", filas=filas)

        self.assertEqual(
            [t["id"] for t in plan["tareas"] if t["seleccionada"]],
            ["pagos:resumen"],
        )
        self.assertEqual(plan["orden"], ["pagos"])

    def test_sin_centrales_el_boton_se_bloquea_y_dice_que_falta(self):
        filas = caso_vacio(centrales="falta")
        filas = [f for f in filas if f["id"] != "centrales:Resumen BESS"]

        plan = nucleo.planificar("/caso", "2607", filas=filas)

        self.assertFalse(plan["puede_ejecutar"])
        self.assertTrue(
            any("Centrales.xlsx" in b for b in plan["bloqueos"]),
            plan["bloqueos"],
        )

    def test_sin_el_excel_de_prorrata_se_hace_todo_lo_demas(self):
        """
        Ese archivo llega despues: no puede bloquear la corrida
        entera, solo quedan afuera sus dos hojas (a la vista y con el
        motivo).
        """

        filas = caso_vacio(prorrata_archivo="falta")
        plan = nucleo.planificar("/caso", "2607", filas=filas)

        self.assertTrue(plan["puede_ejecutar"], plan["bloqueos"])

        elegidas = {t["id"] for t in plan["tareas"] if t["seleccionada"]}
        self.assertNotIn("pagos:prorrata_retiros", elegidas)
        self.assertNotIn("pagos:resumen", elegidas)
        self.assertIn("pagos:ecostos", elegidas)

        prorrata = next(
            t for t in plan["tareas"] if t["id"] == "pagos:prorrata_retiros"
        )
        self.assertEqual(prorrata["estado"], "sin tildar")
        self.assertTrue(prorrata["faltan_requisitos"])

    def test_sin_ofertas_el_boton_se_bloquea(self):
        plan = nucleo.planificar(
            "/caso", "2607", filas=caso_vacio(ofertas="falta")
        )

        self.assertFalse(plan["puede_ejecutar"])
        self.assertTrue(
            any("OfertasSSCC" in b for b in plan["bloqueos"]),
            plan["bloqueos"],
        )

    def test_sin_periodo_no_se_puede_ejecutar(self):
        plan = nucleo.planificar("/caso", "", filas=caso_vacio())

        self.assertFalse(plan["puede_ejecutar"])

    def test_una_tarea_sin_su_dependencia_queda_bloqueada(self):
        """
        Tildar solo 'Calculo E Costos' con el consolidado sin generar
        no puede correr: lo dice, no lo calcula con datos viejos.
        """

        plan = nucleo.planificar(
            "/caso", "2607", seleccion={"pagos:ecostos"},
            filas=caso_vacio(),
        )

        ecostos = next(
            t for t in plan["tareas"] if t["id"] == "pagos:ecostos"
        )
        self.assertEqual(ecostos["estado"], "bloqueada")
        self.assertTrue(ecostos["faltan_dependencias"])
        self.assertFalse(plan["puede_ejecutar"])

    def test_rehacer_algo_del_medio_avisa_de_lo_que_queda_viejo(self):
        plan = nucleo.planificar(
            "/caso", "2607", seleccion={"consolidado:medidores"},
            filas=caso_completo(),
        )

        self.assertTrue(plan["puede_ejecutar"], plan["bloqueos"])
        self.assertTrue(
            any("Calculo E Costos" in a for a in plan["advertencias"]),
            plan["advertencias"],
        )

    def test_propagar_seleccion_arrastra_lo_que_sale_de_ahi(self):
        completa = nucleo.propagar_seleccion({"consolidado:medidores"})

        self.assertIn("pagos:ecostos", completa)
        self.assertIn("pagos:re545", completa)
        self.assertIn("pagos:resumen", completa)
        # ...pero no lo que no depende de Medidores
        self.assertNotIn("cmg_csv", completa)
        self.assertNotIn("consolidado:fd", completa)

    def test_el_fma_no_bloquea_pero_ordena(self):
        """
        Falta el FMA: la hoja 'Subastas' se genera igual (esa columna
        queda en 0, como la formula original), pero si el FMA se
        genera en la misma corrida va antes.
        """

        filas = caso_vacio()
        plan = nucleo.planificar(
            "/caso", "2607",
            seleccion={"subastas_db", "fd_sscc", "consolidado:subastas"},
            filas=filas,
        )

        subastas = next(
            t for t in plan["tareas"] if t["id"] == "consolidado:subastas"
        )
        self.assertEqual(subastas["estado"], "se genera")
        self.assertTrue(plan["puede_ejecutar"], plan["bloqueos"])

        con_fma = nucleo.planificar(
            "/caso", "2607",
            seleccion={
                "subastas_db", "fd_sscc", "fma_cpf", "consolidado:subastas",
            },
            filas=filas,
        )
        self.assertLess(
            con_fma["orden"].index("fma"),
            con_fma["orden"].index("consolidado"),
        )

    def test_el_orden_respeta_las_dependencias(self):
        plan = nucleo.planificar("/caso", "2607", filas=caso_vacio())
        orden = plan["orden"]

        for antes, despues in (
            ("cmg_csv", "cmg_xlsx"),
            ("cmg_xlsx", "consolidado"),
            ("medidas_sae", "consolidado"),
            ("fd_sscc", "consolidado"),
            ("subastas_db", "consolidado"),
            ("consolidado", "pagos"),
        ):
            self.assertLess(
                orden.index(antes), orden.index(despues),
                f"{antes} tiene que ir antes que {despues}: {orden}",
            )


class TestCorrida(unittest.TestCase):
    """
    La ejecucion, con los grupos reemplazados por funciones falsas: se
    mira el orden real, el paralelismo y que un error corte solo la
    rama que sale de el.
    """

    def correr(self, filas, seleccion=None, falla=(), hilos=3):

        corridos = []
        candado = threading.Lock()
        simultaneos = {"max": 0, "ahora": 0}

        def fabricar(id_grupo):
            def _ejecutar(carpeta, aamm, secciones, registrar, progreso):
                with candado:
                    simultaneos["ahora"] += 1
                    simultaneos["max"] = max(
                        simultaneos["max"], simultaneos["ahora"]
                    )
                time.sleep(0.02)
                with candado:
                    corridos.append((id_grupo, tuple(sorted(secciones))))
                    simultaneos["ahora"] -= 1
                if id_grupo in falla:
                    raise RuntimeError(f"{id_grupo} exploto")
            return _ejecutar

        grupos_falsos = tuple(
            orquestador.Grupo(
                g.id, g.etiqueta, g.recurso, fabricar(g.id)
            )
            for g in orquestador.GRUPOS
        )

        original_grupos = orquestador.GRUPOS
        original_por_id = orquestador.GRUPO_POR_ID
        orquestador.GRUPOS = grupos_falsos
        orquestador.GRUPO_POR_ID = {g.id: g for g in grupos_falsos}

        try:
            plan = nucleo.planificar(
                "/caso", "2607", seleccion=seleccion, filas=filas
            )
            error = None
            try:
                orquestador.ejecutar_plan(
                    "/caso", "2607", registrar=lambda *a: None,
                    hilos=hilos, plan=plan,
                )
            except nucleo.ErrorEntrada as fallo:
                error = str(fallo)
        finally:
            orquestador.GRUPOS = original_grupos
            orquestador.GRUPO_POR_ID = original_por_id

        return corridos, simultaneos["max"], error

    def test_corre_todo_en_orden_y_en_paralelo_lo_que_puede(self):
        corridos, simultaneos, error = self.correr(caso_vacio())

        self.assertIsNone(error)

        ids = [g for g, _ in corridos]
        self.assertEqual(len(ids), len(set(ids)), "algo corrio dos veces")
        self.assertLess(ids.index("cmg_csv"), ids.index("cmg_xlsx"))
        self.assertLess(ids.index("consolidado"), ids.index("pagos"))
        self.assertGreater(
            simultaneos, 1, "las bajadas tendrian que ir en paralelo"
        )

    def test_las_hojas_van_juntas_en_una_sola_escritura(self):
        corridos, _, _ = self.correr(caso_vacio())

        consolidado = [s for g, s in corridos if g == "consolidado"]
        pagos = [s for g, s in corridos if g == "pagos"]

        self.assertEqual(len(consolidado), 1)
        self.assertEqual(
            consolidado[0],
            ("cmg", "fd", "medidores", "ofertas_sscc", "subastas"),
        )
        self.assertEqual(len(pagos), 1)
        self.assertEqual(
            pagos[0],
            (
                "compensacion_central", "ecostos", "prorrata_retiros",
                "re545", "resumen",
            ),
        )

    def test_dos_pasos_del_mismo_recurso_no_corren_a_la_vez(self):
        """
        'Traer FD' y 'Generar FMA' escriben en la misma carpeta:
        comparten recurso y se turnan aunque el grafo los deje sueltos.
        """

        corridos, simultaneos, _ = self.correr(
            caso_vacio(),
            seleccion={"fd_sscc", "fma_cpf", "fma_csf", "fma_ctf"},
        )

        self.assertEqual(simultaneos, 1)
        self.assertEqual({g for g, _ in corridos}, {"fd_sscc", "fma"})

    def test_si_algo_falla_no_se_corre_lo_que_depende_de_eso(self):
        corridos, _, error = self.correr(
            caso_vacio(), falla={"cmg_csv"}
        )

        ids = [g for g, _ in corridos]

        self.assertIn("cmg_csv", ids)
        self.assertNotIn("cmg_xlsx", ids)      # sale del CSV que fallo
        self.assertNotIn("consolidado", ids)   # y de ahi para abajo
        self.assertNotIn("pagos", ids)
        # lo que no depende del CMg si corrio
        self.assertIn("medidas_sae", ids)
        self.assertIn("subastas_db", ids)
        self.assertIsNotNone(error)
        self.assertIn("cmg", error)

    def test_con_un_solo_hilo_hace_lo_mismo_en_fila(self):
        corridos, simultaneos, error = self.correr(caso_vacio(), hilos=1)

        self.assertIsNone(error)
        self.assertEqual(simultaneos, 1)
        ids = [g for g, _ in corridos]
        self.assertEqual(len(ids), 8)
        self.assertLess(ids.index("consolidado"), ids.index("pagos"))

    def test_no_se_ejecuta_un_plan_bloqueado(self):
        with self.assertRaises(nucleo.ErrorEntrada):
            orquestador.ejecutar_plan(
                "/caso", "2607", seleccion={"pagos:ecostos"},
                registrar=lambda *a: None,
                plan=nucleo.planificar(
                    "/caso", "2607", seleccion={"pagos:ecostos"},
                    filas=caso_vacio(),
                ),
            )


if __name__ == "__main__":
    unittest.main()
