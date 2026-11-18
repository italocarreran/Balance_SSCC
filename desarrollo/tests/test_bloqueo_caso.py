# -*- coding: utf-8 -*-
"""Bloquear/Desbloquear caso (2026-11-18).

Pedido del usuario: "un boton de Bloquear/Desbloquear caso, para no correr
un caso anterior por error", arriba y que bloquee todos los botones del
caso. El bloqueo es un archivo en la carpeta del BESS (nucleo/bloqueo.py).
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from Script import nucleo
from desarrollo.tests.test_ventana_rueda import _hay_ventana, tk


class TestArchivoDeBloqueo(unittest.TestCase):

    def test_bloquear_y_desbloquear(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / nucleo.CARPETA_BESS).mkdir()
            self.assertIsNone(nucleo.info_bloqueo(base))
            ruta = nucleo.bloquear(base, quien="PC_usuario", motivo="ya publicado")
            # En <mes>/BESS/, donde vive lo propio del BESS.
            self.assertEqual(ruta.parent, base / nucleo.CARPETA_BESS)
            info = nucleo.info_bloqueo(base)
            self.assertEqual(info["quien"], "PC_usuario")
            self.assertEqual(info["motivo"], "ya publicado")
            self.assertTrue(info["cuando"])
            self.assertTrue(nucleo.esta_bloqueado(base))
            self.assertTrue(nucleo.desbloquear(base))
            self.assertFalse(nucleo.esta_bloqueado(base))
            self.assertFalse(nucleo.desbloquear(base))

    def test_sin_carpeta_no_esta_bloqueado(self):
        self.assertIsNone(nucleo.info_bloqueo(""))
        self.assertIsNone(nucleo.info_bloqueo("/no/existe/esta/carpeta"))


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentanaBloqueo(unittest.TestCase):

    def setUp(self):
        import Balance_BESS as app

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

    def _botones_del_arbol(self):
        """Los botones de accion del diagrama (los que estan en sus filas)."""
        quedan = {"Ir a este balance", "Examinar", "Bloquear caso",
                  "Desbloquear caso", "Ejecutar todo", "Generar antecedentes SAE",
                  "Actualizar estado", "Salir"}
        return [w for w in self._todos() if w.winfo_class() == "Button"
                and w.cget("text") not in quedan]

    def _estado(self, boton):
        return str(boton.cget("state"))

    def test_bloquear_apaga_todos_los_botones_del_caso(self):
        arbol = self._botones_del_arbol()
        self.assertTrue(arbol)
        self.assertTrue(all(self._estado(b) == "normal" for b in arbol))

        with mock.patch.object(self.app.messagebox, "askyesno", return_value=True):
            self._boton("Bloquear caso").invoke()
        self.root.update()

        self.assertTrue(nucleo.esta_bloqueado(self.caso))
        self.assertTrue(all(self._estado(b) == "disabled" for b in self._botones_del_arbol()))
        self.assertEqual(self._estado(self._boton("Ejecutar todo")), "disabled")
        self.assertEqual(self._estado(self._boton("Generar antecedentes SAE")), "disabled")
        # Lo que no corre procesos sigue andando.
        for texto in ("Desbloquear caso", "Examinar", "Actualizar estado", "Salir"):
            self.assertEqual(self._estado(self._boton(texto)), "normal")

        # Despues de Actualizar estado (el arbol se repinta) siguen apagados.
        self._boton("Actualizar estado").invoke()
        self.root.update()
        self.assertTrue(all(self._estado(b) == "disabled" for b in self._botones_del_arbol()))

        with mock.patch.object(self.app.messagebox, "askyesno", return_value=True):
            self._boton("Desbloquear caso").invoke()
        self.root.update()
        self.assertFalse(nucleo.esta_bloqueado(self.caso))
        self.assertTrue(all(self._estado(b) == "normal" for b in self._botones_del_arbol()))
        self.assertEqual(self._estado(self._boton("Ejecutar todo")), "normal")

    def test_si_se_cancela_no_se_bloquea(self):
        with mock.patch.object(self.app.messagebox, "askyesno", return_value=False):
            self._boton("Bloquear caso").invoke()
        self.assertFalse(nucleo.esta_bloqueado(self.caso))

    def test_un_caso_bloqueado_afuera_no_corre_nada(self):
        # Alguien lo bloqueo desde otro PC: el boton del arbol (si se
        # alcanza a apretar antes de repintar) avisa y no corre.
        nucleo.bloquear(self.caso, quien="otro PC")
        boton = self._botones_del_arbol()[0]
        with mock.patch.object(self.app.messagebox, "showwarning") as aviso, \
                mock.patch.object(self.app.threading, "Thread") as hilo:
            boton.invoke()
        self.root.update()
        aviso.assert_called()
        self.assertEqual(aviso.call_args[0][0], "Caso bloqueado")
        hilo.assert_not_called()
        self.assertEqual(self._estado(self._boton("Ejecutar todo")), "disabled")


if __name__ == "__main__":
    unittest.main()
