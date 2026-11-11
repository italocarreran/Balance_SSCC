"""El Registro de la ventana: hora en cada linea, colores y archivo.

Pedido del usuario: control de tiempo en el log (cuanto se demora cada
cosa) y colores: verde lo que esta ok, amarillo los avisos y rojo los
errores. Al terminar, el registro queda en Registro_ultima_corrida.txt
en la carpeta del caso.

Necesita tkinter Y un display (en Linux: `xvfb-run -a python -m
unittest discover`); sin ellos se saltea.
"""

import re
import unittest
from unittest import mock

from desarrollo.tests import test_ventana_actualizar as base
from desarrollo.tests.test_ventana_rueda import _hay_ventana


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestRegistro(base.TestVentanaActualizar):

    test_actualizar_muestra_un_archivo_pegado_a_mano = None
    test_medidas_arranca_sin_advertencia = None
    test_aviso_de_medidas_incompletas = None
    test_actualizar_sin_cambios_no_rehace_el_arbol = None
    test_cambia_el_estado_sin_rehacer_las_filas = None
    test_despues_de_un_proceso_los_botones_quedan_habilitados = None

    def test_nivel_del_mensaje(self):
        nivel = self.app.nivel_del_mensaje
        self.assertEqual(nivel("[AVISO] sin Access para el dia 3"), "aviso")
        self.assertEqual(nivel("[cmg_xlsx] Listo (3,2 s)."), "ok")
        self.assertEqual(nivel("Listo: 5 fila(s)"), "ok")
        self.assertEqual(nivel("[fma] ERROR (1,0 s): no hay"), "error")
        self.assertEqual(nivel("  cmg.xlsx: omitido: no corrio cmg_csv"), "aviso")
        self.assertIsNone(nivel("Leyendo cmg.xlsx..."))

    def test_hora_colores_y_archivo(self):
        def proceso_falso(registrar=None, progreso=None, **_):
            registrar("Leyendo algo...")
            registrar("[AVISO] faltan 2 dias")
            registrar("Listo: 10 fila(s)")

        class HiloInmediato:
            def __init__(self, target, daemon=None, **_):
                self.target = target

            def start(self):
                self.target()

        txt = next(
            w for w in self._todos(self.root) if w.winfo_class() == "Text"
            and w.master.cget("text") == "Registro"
        )
        with mock.patch.object(self.nucleo, "traer_subastas", proceso_falso), \
                mock.patch.object(self.app.threading, "Thread", HiloInmediato), \
                mock.patch.object(self.app.messagebox, "showinfo"), \
                mock.patch.object(self.app, "mostrar_error"):
            self._boton("Traer", self._fila("DB subastas")).invoke()
            for _ in range(5):
                self.root.update()

        texto = txt.get("1.0", "end")
        # cada linea con hora y tiempo transcurrido
        self.assertRegex(texto, r"\[\d\d:\d\d:\d\d \+00:\d\d\] Leyendo algo")
        # Cada boton como encabezado destacado y su cierre con el tiempo.
        self.assertRegex(texto, r"■ .+: listo en ")

        def lineas_con(tag):
            rangos = txt.tag_ranges(tag)
            return " ".join(
                txt.get(rangos[i], rangos[i + 1]) for i in range(0, len(rangos), 2)
            )

        self.assertIn("[AVISO] faltan 2 dias", lineas_con("aviso"))
        self.assertIn("Listo: 10 fila(s)", lineas_con("ok"))
        self.assertNotIn("Leyendo algo", lineas_con("ok") + lineas_con("aviso"))
        self.assertEqual(txt.tag_cget("ok", "foreground"), self.app.COLOR_REGISTRO["ok"])
        self.assertIn("▶ ", lineas_con("encabezado"))
        self.assertIn("listo en", lineas_con("encabezado_fin"))

        guardado = self.caso / "Log" / self.app.ARCHIVO_REGISTRO
        self.assertTrue(guardado.is_file())
        self.assertIn("[AVISO] faltan 2 dias", guardado.read_text(encoding="utf-8"))


    def test_rafaga_grande_queda_entera_y_en_orden(self):
        # El registro escribe por tandas con tope (plan de la ventana,
        # Fase 4.3); al terminar tiene que vaciar la cola entera antes
        # de guardar el archivo.
        total = 5000

        def proceso_falso(registrar=None, progreso=None, **_):
            for i in range(total):
                registrar(f"mensaje numero {i:05d}")

        class HiloInmediato:
            def __init__(self, target, daemon=None, **_):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(self.nucleo, "traer_subastas", proceso_falso), \
                mock.patch.object(self.app.threading, "Thread", HiloInmediato), \
                mock.patch.object(self.app.messagebox, "showinfo"), \
                mock.patch.object(self.app, "mostrar_error"):
            self._boton("Traer", self._fila("DB subastas")).invoke()
            for _ in range(5):
                self.root.update()

        guardado = self.caso / "Log" / self.app.ARCHIVO_REGISTRO
        numeros = re.findall(
            r"mensaje numero (\d{5})", guardado.read_text(encoding="utf-8")
        )
        self.assertEqual(numeros, [f"{i:05d}" for i in range(total)])


if __name__ == "__main__":
    unittest.main()
