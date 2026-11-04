"""La ventana de "Ejecutar todo" con el mismo orden que el arbol.

Pedido del usuario: que salgan todas las cosas que tienen boton en la
ventana principal, y las hojas del Excel adentro de Balance_BESS.xlsx,
con su propio cuadro selector. Y la fila nueva de las subastas
consolidadas (subastas_AAMM.xlsx) con su boton "Consolidar".

Necesita tkinter Y un display (en Linux: `xvfb-run -a python -m
unittest discover`); sin ellos se saltea.
"""

import unittest

from Script.nucleo import orquestador
from desarrollo.tests import test_ventana_actualizar as base
from desarrollo.tests.test_ventana_rueda import _hay_ventana


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentanaEjecutarTodo(base.TestVentanaActualizar):

    # Solo las pruebas de esta clase, no las heredadas.
    test_actualizar_muestra_un_archivo_pegado_a_mano = None
    test_medidas_arranca_sin_advertencia = None
    test_aviso_de_medidas_incompletas = None

    def _dialogo(self):
        self._boton("Ejecutar todo", self.root).invoke()
        self.root.update()
        dialogos = [
            w for w in self.root.winfo_children()
            if w.winfo_class() == "Toplevel" and w.title() == "Ejecutar todo"
        ]
        self.assertTrue(dialogos, "no se abrio la ventana 'Ejecutar todo'")
        return dialogos[-1]

    def _textos(self, widget):
        return [
            str(w.cget("text")) for w in self._todos(widget)
            if w.winfo_class() == "Label"
        ]

    def test_la_fila_de_subastas_consolidadas_tiene_su_boton(self):
        fila = self._fila("subastas_2609.xlsx")
        self._boton("Consolidar", fila)

    def test_la_lista_sigue_el_arbol_y_las_hojas_van_dentro_del_libro(self):
        dialogo = self._dialogo()
        try:
            textos = self._textos(dialogo)
            # Todas las tareas aparecen
            for tarea in self.nucleo.TAREAS:
                self.assertIn(tarea.etiqueta, textos, tarea.id)
            # carpetas como titulo, en el orden del arbol
            medidas = self.nucleo.TAREA_POR_ID["medidas_sae"].etiqueta
            for antes, despues in (
                ("Medidas/", medidas),
                ("Subastas/", "Traer subastas (Access del periodo)"),
                ("Traer subastas (Access del periodo)",
                 "Consolidar subastas_AAMM.xlsx"),
                ("Consolidar subastas_AAMM.xlsx", "Balance_BESS.xlsx"),
                ("Balance_BESS.xlsx", "hoja 'Resumen'"),
                ("hoja 'Resumen'", "hoja 'Centrales'"),
            ):
                indice = {t: i for i, t in reversed(list(enumerate(textos)))}
                self.assertLess(
                    indice[antes], indice[despues], f"{antes} / {despues}"
                )
            hojas = [
                t for t in self.nucleo.TAREAS
                if t.id.split(":", 1)[0] in ("consolidado", "pagos")
            ]

            def contador():
                return next(
                    t for t in self._textos(dialogo) if "hoja(s) tildada" in t
                )

            # El cuadro del libro tilda y destilda todas las hojas
            libro = next(
                w for w in self._todos(dialogo)
                if w.winfo_class() == "Label"
                and w.cget("text") == "Balance_BESS.xlsx"
            ).master
            cuadro = next(
                w for w in libro.winfo_children()
                if w.winfo_class() == "Checkbutton"
            )
            cuadro.invoke()
            self.root.update()
            # Las que esperan el Excel de prorrata (que este caso no
            # tiene) quedan fuera con su aviso, como siempre.
            esperan = [
                t for t in hojas
                if set(t.requisitos) & set(orquestador.ENTRADAS_TARDIAS)
            ]
            self.assertIn(
                f"{len(hojas) - len(esperan)} de {len(hojas)}", contador()
            )
            libro = next(
                w for w in self._todos(dialogo)
                if w.winfo_class() == "Label"
                and w.cget("text") == "Balance_BESS.xlsx"
            ).master
            next(
                w for w in libro.winfo_children()
                if w.winfo_class() == "Checkbutton"
            ).invoke()
            self.root.update()
            self.assertIn(f"0 de {len(hojas)}", contador())
        finally:
            dialogo.destroy()


    def _principal(self):
        """El Canvas de la ventana principal (el que mueve la rueda)."""
        return next(
            w for w in self.root.winfo_children()
            if w.winfo_class() == "Canvas"
        )

    def test_la_rueda_sobre_el_dialogo_no_mueve_la_ventana_principal(self):
        """Pedido del usuario: "al mover la ruedita del mouse se mueve la
        ventana principal"."""

        principal = self._principal()
        self.root.geometry("1000x300")  # que la principal tenga que bajar
        self.root.update()
        principal.yview_moveto(0.0)
        dialogo = self._dialogo()
        try:
            dialogo.geometry("900x300")
            dialogo.update()
            lienzo = next(
                w for w in self._todos(dialogo) if w.winfo_class() == "Canvas"
            )
            for widget in (lienzo, next(
                w for w in self._todos(dialogo) if w.winfo_class() == "Button"
            )):
                x = widget.winfo_rootx() + widget.winfo_width() // 2
                y = widget.winfo_rooty() + widget.winfo_height() // 2
                for _ in range(3):
                    widget.event_generate(
                        "<MouseWheel>", delta=-120, rootx=x, rooty=y,
                        x=widget.winfo_width() // 2,
                        y=widget.winfo_height() // 2,
                    )
                self.root.update()
            self.assertEqual(principal.yview()[0], 0.0)
            # ...y sobre la lista, se mueve la lista.
            self.assertGreater(lienzo.yview()[0], 0.0)
        finally:
            dialogo.destroy()

    def test_tildar_no_vuelve_a_armar_la_lista(self):
        """Pedido del usuario: "cuando selecciono una casilla hace un
        parpadeo". La lista se arma una vez: tildar solo cambia valores."""

        dialogo = self._dialogo()
        try:
            antes = [str(w) for w in self._todos(dialogo)]
            cuadro = next(
                w for w in self._todos(dialogo)
                if w.winfo_class() == "Checkbutton"
                and w.master.winfo_children()[1].cget("text")
                != "Balance_BESS.xlsx"
            )
            cuadro.invoke()
            self.root.update()
            despues = [str(w) for w in self._todos(dialogo)]
            self.assertEqual(antes, despues)
        finally:
            dialogo.destroy()


if __name__ == "__main__":
    unittest.main()
