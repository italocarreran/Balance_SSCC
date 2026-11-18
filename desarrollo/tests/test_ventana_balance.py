# -*- coding: utf-8 -*-
"""El balance del caso en la ventana: año _ mes _ proceso (2026-11-18).

Pedido del usuario: en vez del AAMM, tres casillas que se leen
"2026_Marzo_Definitivo"; "Ir a este balance" azul si el balance tiene
carpeta y bloqueado si no; "Examinar" destacado cuando no la tiene; el
"+" de la lista de procesos agrega uno (R01P, R01D, ...).

Las funciones necesitan tkinter (Balance_BESS lo importa); la ventana,
ademas, un display (en Linux: xvfb-run).
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from desarrollo.tests.test_ventana_rueda import _hay_ventana, tk


@unittest.skipUnless(tk is not None, "sin tkinter")
class TestFuncionesDelBalance(unittest.TestCase):

    def setUp(self):
        import Balance_BESS as app

        self.app = app
        self.cfg = {}
        self.parches = [
            mock.patch.object(app, "leer_config", lambda: dict(self.cfg)),
            mock.patch.object(app, "guardar_config", self.cfg.update),
        ]
        for parche in self.parches:
            parche.start()

    def tearDown(self):
        for parche in self.parches:
            parche.stop()

    def test_aamm_y_nombre(self):
        app = self.app
        self.assertEqual(app.aamm_de("2026", "Marzo"), "2603")
        self.assertEqual(app.aamm_de("2025", "Diciembre"), "2512")
        self.assertEqual(app.aamm_de("2026", "Marzito"), "")
        self.assertEqual(app.ano_mes_de("2610"), ("2026", "Octubre"))
        self.assertEqual(app.ano_mes_de("2613"), ("", ""))
        self.assertEqual(
            app.nombre_balance("2026", "Marzo", "Definitivo"), "2026_Marzo_Definitivo"
        )
        self.assertEqual(app.nombre_de_clave("2602_R01P"), "2026_Febrero_R01P")

    def test_anos_desde_2025(self):
        self.assertEqual(self.app.anos_disponibles(2027), ["2025", "2026", "2027"])
        self.assertEqual(self.app.anos_disponibles(2025), ["2025"])

    def test_procesos(self):
        app = self.app
        self.assertEqual(app.procesos_guardados(), ["Preliminar", "Definitivo"])
        self.cfg[app.CLAVE_PROCESOS] = ["R01P", "r01p", "preliminar", "R 01/D"]
        self.assertEqual(
            app.procesos_guardados(), ["Preliminar", "Definitivo", "R01P", "R01D"]
        )
        self.assertEqual(app.proceso_en_el_nombre("D:/BESS 2610 Preliminar"), "Preliminar")
        self.assertEqual(app.proceso_en_el_nombre("D:/BESS 2610"), "")

    def test_carpeta_del_balance(self):
        app = self.app
        with tempfile.TemporaryDirectory() as carpeta:
            pre = Path(carpeta) / "BESS 2610 Preliminar"
            defi = Path(carpeta) / "BESS 2610 Def"
            pre.mkdir()
            defi.mkdir()
            self.cfg[app.CLAVE_CARPETAS_PERIODO] = {
                "2610": str(pre), "2610_Definitivo": str(defi),
            }
            self.assertEqual(app.carpeta_del_balance("2610", "Definitivo"), str(defi))
            # Anotada solo con el AAMM (de antes): vale si el nombre dice el proceso.
            self.assertEqual(app.carpeta_del_balance("2610", "Preliminar"), str(pre))
            self.assertEqual(app.carpeta_del_balance("2610", "R01P"), "")
            self.assertEqual(app.balance_de_la_carpeta(str(defi)), "2610_Definitivo")

    def test_el_cra_lee_las_carpetas_del_bess_por_aamm(self):
        import Balance_CRA as cra

        datos = {cra.get_usuario(): {cra.CLAVE_CARPETAS_PERIODO: {
            "2603_Definitivo": "D:/BESS 2603", "2602": "D:/BESS 2602",
        }}}
        with mock.patch.object(cra.config, "seccion", lambda nombre: datos):
            self.assertEqual(
                cra.carpetas_del_bess(), {"2603": "D:/BESS 2603", "2602": "D:/BESS 2602"}
            )


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentanaBalance(unittest.TestCase):

    def setUp(self):
        import Balance_BESS as app
        from Script import nucleo

        self.app = app
        self.tmp = tempfile.TemporaryDirectory()
        self.caso = Path(self.tmp.name) / "Caso 2610 Preliminar"
        for sub in nucleo.SUBCARPETAS_CASO:
            (self.caso / sub).mkdir(parents=True, exist_ok=True)
        self.cfg = {"carpeta_base": str(self.caso), "aamm": "2610"}
        self.parches = [
            mock.patch.object(app, "leer_config", lambda: dict(self.cfg)),
            mock.patch.object(app, "guardar_config", self.cfg.update),
        ]
        for parche in self.parches:
            parche.start()

        ventana = {}
        original = tk.Tk.mainloop
        tk.Tk.mainloop = lambda self_: ventana.setdefault("root", self_)
        try:
            app.main()
        finally:
            tk.Tk.mainloop = original
        self.root = ventana["root"]
        self.root.update()

    def tearDown(self):
        self.root.destroy()
        for parche in self.parches:
            parche.stop()
        self.tmp.cleanup()

    def _todos(self, widget=None):
        for hijo in (widget or self.root).winfo_children():
            yield hijo
            yield from self._todos(hijo)

    def _boton(self, texto):
        for w in self._todos():
            if w.winfo_class() == "Button" and w.cget("text") == texto:
                return w
        self.fail(f"no hay boton '{texto}'")

    def _casillas(self):
        return [w for w in self._todos() if w.winfo_class() == "TCombobox"]

    def _elegir(self, posicion, valor):
        caja = self._casillas()[posicion]
        self.root.setvar(str(caja.cget("textvariable")), valor)
        caja.event_generate("<<ComboboxSelected>>")
        self.root.update()

    def test_arranca_con_el_caso_de_antes(self):
        ano, mes, proceso = (c.get() for c in self._casillas())
        self.assertEqual((ano, mes, proceso), ("2026", "Octubre", "Preliminar"))
        # Quedo anotada como la del balance (el proceso salio del nombre).
        self.assertEqual(
            self.cfg[self.app.CLAVE_CARPETAS_PERIODO], {"2610_Preliminar": str(self.caso)}
        )
        self.assertEqual(str(self._boton("Ir a este balance").cget("state")), "normal")

    def test_sin_carpeta_ir_bloqueado(self):
        self._elegir(1, "Noviembre")
        self.assertEqual(str(self._boton("Ir a este balance").cget("state")), "disabled")
        self._elegir(1, "Octubre")
        self.assertEqual(str(self._boton("Ir a este balance").cget("state")), "normal")

    def test_examinar_anota_el_balance(self):
        otra = Path(self.tmp.name) / "Caso 2611"
        otra.mkdir()
        self._elegir(1, "Noviembre")
        self._elegir(2, "Definitivo")
        with mock.patch.object(self.app.filedialog, "askdirectory", return_value=str(otra)), \
                mock.patch.object(self.app.messagebox, "askyesno", return_value=False):
            self._boton("Examinar").invoke()
        self.root.update()
        self.assertEqual(self.cfg["aamm"], "2611")
        self.assertEqual(self.cfg["proceso"], "Definitivo")
        self.assertEqual(self.cfg["carpeta_base"], str(otra))
        self.assertEqual(
            self.cfg[self.app.CLAVE_CARPETAS_PERIODO]["2611_Definitivo"], str(otra)
        )
        self.assertEqual(str(self._boton("Ir a este balance").cget("state")), "normal")

    def test_el_mas_agrega_un_proceso(self):
        from tkinter import simpledialog

        with mock.patch.object(simpledialog, "askstring", return_value=" R01P "):
            self._elegir(2, self.app.OPCION_AGREGAR_PROCESO)
        self.assertEqual(self._casillas()[2].get(), "R01P")
        self.assertEqual(self.cfg[self.app.CLAVE_PROCESOS], ["R01P"])
        self.assertIn("R01P", self._casillas()[2].cget("values"))

    def test_cancelar_el_mas_vuelve_al_anterior(self):
        from tkinter import simpledialog

        with mock.patch.object(simpledialog, "askstring", return_value=None):
            self._elegir(2, self.app.OPCION_AGREGAR_PROCESO)
        self.assertEqual(self._casillas()[2].get(), "Preliminar")

    def test_botones_que_ya_no_estan(self):
        textos = {w.cget("text") for w in self._todos() if w.winfo_class() == "Button"}
        for quitado in ("Crear carpeta del caso", "Abrir carpeta del caso",
                        "Ir a este mes"):
            self.assertNotIn(quitado, textos)
        for queda in ("Ejecutar todo", "Generar antecedentes SAE",
                      "Actualizar estado", "Salir"):
            self.assertIn(queda, textos)


if __name__ == "__main__":
    unittest.main()
