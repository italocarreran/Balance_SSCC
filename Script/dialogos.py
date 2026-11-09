# -*- coding: utf-8 -*-
"""
Cuadro de error con boton "Copiar".

Pedido del usuario (2026-11-09): "los mensajes de error que tengan un
boton de copiar para mandarlo escrito y no con foto". Reemplaza a
messagebox.showerror: el texto va en un recuadro que se puede
seleccionar, y "Copiar" lo deja en el portapapeles (titulo + mensaje).

Herramientas/ tiene su propia copia (esas ventanas no importan nada del
programa, funcionan solas en cualquier carpeta).
"""


def mostrar_error(titulo, mensaje, parent=None, **_opciones):
    """Como messagebox.showerror(titulo, mensaje), con boton Copiar."""

    import tkinter as tk

    raiz = parent or tk._default_root
    ventana = tk.Toplevel(raiz)
    ventana.title(titulo)
    ventana.resizable(True, True)
    if raiz is not None:
        ventana.transient(raiz)

    texto_completo = f"{titulo}\n\n{mensaje}"
    lineas = max(4, min(20, str(mensaje).count("\n") + 2 + len(str(mensaje)) // 80))

    cuerpo = tk.Frame(ventana, padx=14, pady=12)
    cuerpo.pack(fill="both", expand=True)
    tk.Label(
        cuerpo, text="✖", fg="#c0392b", font=("Segoe UI", 20, "bold"),
    ).pack(side="left", anchor="n", padx=(0, 12))
    caja = tk.Text(
        cuerpo, wrap="word", width=78, height=lineas, font=("Segoe UI", 9),
        relief="flat", background=ventana.cget("background"),
    )
    caja.insert("1.0", str(mensaje))
    caja.config(state="disabled")  # se puede seleccionar, no editar
    caja.pack(side="left", fill="both", expand=True)

    botones = tk.Frame(ventana, padx=14, pady=10)
    botones.pack(fill="x")
    aviso = tk.Label(botones, text="", fg="#2d7a2d", font=("Segoe UI", 8))

    def copiar():
        ventana.clipboard_clear()
        ventana.clipboard_append(texto_completo)
        aviso.config(text="Copiado: pégalo donde quieras.")

    tk.Button(botones, text="Aceptar", width=12, command=ventana.destroy).pack(
        side="right"
    )
    tk.Button(
        botones, text="Copiar", width=12, bg="#1f5fa8", fg="white",
        font=("Segoe UI", 9, "bold"), command=copiar,
    ).pack(side="right", padx=8)
    aviso.pack(side="left")

    ventana.bind("<Escape>", lambda e: ventana.destroy())
    ventana.bind("<Return>", lambda e: ventana.destroy())
    ventana.grab_set()
    ventana.focus_set()
    ventana.wait_window()
    return "ok"
