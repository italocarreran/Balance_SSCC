# -*- coding: utf-8 -*-
"""
Buscar barras en el CSV de CMg 15-minutal.

Pedido del usuario (2026-11-05): "una ventana, selecciono el archivo csv
y indico el nombre de la barra aproximado". Sirve para encontrar el
nombre EXACTO de una barra (el que va en la columna 'Barra inyección' de
Centrales.xlsx), por ejemplo 'arena 220' -> 'ARENA_________220'.

Como se busca:
  - se comparan los nombres "limpios": mayusculas, sin tildes, y los
    guiones bajos y simbolos como espacios ('TOCOPILLA_____110' ->
    'TOCOPILLA 110');
  - una barra que trae todas las palabras escritas (en cualquier orden)
    queda arriba; despues las que se parecen (letras cambiadas o de mas),
    con su % de parecido.

Doble clic en una barra (o el boton 'Copiar') copia su nombre exacto.

Uso:  python Herramientas/Buscar_Barras_CMg.py (funciona en cualquier carpeta)
"""

import difflib
import re
import threading
import unicodedata
from pathlib import Path

# El CSV del Coordinador: ';' y latin1 (ver Script/Cmg/Extrae_CMG_barras.py).
SEPARADOR_CSV = ";"
CODIFICACION_CSV = "latin1"
COLUMNA_BARRA = "BARRA"
MAXIMO_RESULTADOS = 50


# ------------------------------------------------------------
# La busqueda (sin ventana, para poder probarla sola)
# ------------------------------------------------------------

def limpiar(texto):
    """'Tocopilla_____110 ' -> 'TOCOPILLA 110' (sin tildes ni simbolos)."""

    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^0-9A-Za-z]+", " ", texto.upper())
    return " ".join(texto.split())


def buscar_barras(barras, texto, limite=MAXIMO_RESULTADOS):
    """
    [(barra, parecido 0..100)] de las barras que mas se parecen a texto,
    de mas a menos parecida. Arriba las que traen todas las palabras
    escritas; despues las parecidas aunque esten mal escritas.
    """

    buscado = limpiar(texto)
    if not buscado:
        return []
    palabras = buscado.split()
    resultados = []
    for barra in barras:
        nombre = limpiar(barra)
        sus_palabras = nombre.split()
        # Cada palabra escrita: aparece tal cual, o como parte de una
        # palabra de la barra ('quill' en 'QUILLAGUA'), o se le parece.
        puntos = []
        for palabra in palabras:
            if palabra in sus_palabras or (
                palabra.isdigit()
                and any(suya.isdigit() and int(suya) == int(palabra)
                        for suya in sus_palabras)
            ):
                # '12' es '0012'; '33' es '033'.
                puntos.append(1.0)
            elif any(palabra in suya for suya in sus_palabras):
                puntos.append(0.9)
            else:
                puntos.append(max(
                    (difflib.SequenceMatcher(None, palabra, suya).ratio()
                     for suya in sus_palabras),
                    default=0.0,
                ))
        por_palabras = sum(puntos) / len(puntos)
        todo = difflib.SequenceMatcher(None, buscado, nombre).ratio()
        parecido = max(por_palabras * 0.95 + todo * 0.05, todo)
        if parecido >= 0.5:
            resultados.append((barra, round(100 * parecido)))
    resultados.sort(key=lambda r: (-r[1], r[0]))
    return resultados[:limite]


def leer_barras(ruta):
    """{barra: filas} del CSV (solo lee la columna BARRA: es rapido)."""

    import pandas as pd

    ruta = Path(ruta)
    try:
        try:
            # El lector de pyarrow (si esta) es varias veces mas rapido
            # con los millones de filas del CSV.
            serie = pd.read_csv(
                ruta, sep=SEPARADOR_CSV, encoding=CODIFICACION_CSV,
                usecols=[COLUMNA_BARRA], dtype=str, engine="pyarrow",
            )[COLUMNA_BARRA]
        except ImportError:
            serie = pd.read_csv(
                ruta, sep=SEPARADOR_CSV, encoding=CODIFICACION_CSV,
                usecols=[COLUMNA_BARRA], dtype=str,
            )[COLUMNA_BARRA]
    except ValueError:
        # Sin columna BARRA con ese separador: se intenta adivinarlo.
        serie = pd.read_csv(
            ruta, sep=None, engine="python", encoding=CODIFICACION_CSV,
            dtype=str,
        )
        columna = next(
            (c for c in serie.columns if limpiar(c) == COLUMNA_BARRA), None
        )
        if columna is None:
            raise ValueError(
                f"{ruta.name} no tiene una columna '{COLUMNA_BARRA}'. "
                f"Columnas: {list(serie.columns)}"
            )
        serie = serie[columna]
    serie = serie.dropna().str.strip()
    return serie[serie != ""].value_counts().to_dict()


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Buscar barras en el CSV de CMg")
    root.geometry("760x560")

    estado = {"barras": {}}
    var_ruta = tk.StringVar(value="")
    var_texto = tk.StringVar(value="")
    var_info = tk.StringVar(value="Selecciona el CSV 15-minutal de CMg.")

    # --- archivo
    marco_archivo = tk.LabelFrame(root, text="Archivo CSV de CMg", padx=10, pady=6)
    marco_archivo.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(
        marco_archivo, textvariable=var_ruta, anchor="w", justify="left",
        wraplength=600, font=("Segoe UI", 9),
    ).pack(side="left", fill="x", expand=True)
    boton_examinar = tk.Button(marco_archivo, text="Examinar...")
    boton_examinar.pack(side="right")

    # --- busqueda
    marco_buscar = tk.LabelFrame(
        root, text="Nombre aproximado de la barra", padx=10, pady=6
    )
    marco_buscar.pack(fill="x", padx=14, pady=4)
    entrada = tk.Entry(
        marco_buscar, textvariable=var_texto, font=("Consolas", 11, "bold")
    )
    entrada.pack(side="left", fill="x", expand=True)
    tk.Label(
        marco_buscar, text="ej. arena 220, tocopilla, quillagua",
        fg="#444444", font=("Segoe UI", 8),
    ).pack(side="left", padx=8)

    # --- resultados
    marco_lista = tk.Frame(root)
    marco_lista.pack(fill="both", expand=True, padx=14, pady=4)
    tabla = ttk.Treeview(
        marco_lista, columns=("barra", "parecido", "filas"), show="headings",
        selectmode="browse",
    )
    tabla.heading("barra", text="Barra (nombre exacto)")
    tabla.heading("parecido", text="Parecido")
    tabla.heading("filas", text="Filas en el CSV")
    tabla.column("barra", width=420, anchor="w")
    tabla.column("parecido", width=90, anchor="center")
    tabla.column("filas", width=120, anchor="e")
    barra_scroll = ttk.Scrollbar(marco_lista, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=barra_scroll.set)
    barra_scroll.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)

    # --- abajo
    abajo = tk.Frame(root)
    abajo.pack(fill="x", padx=14, pady=(4, 10))
    boton_copiar = tk.Button(
        abajo, text="Copiar nombre", bg="#1f5fa8", fg="white",
        font=("Segoe UI", 9, "bold"),
    )
    boton_copiar.pack(side="left")
    tk.Label(abajo, textvariable=var_info, anchor="w", fg="#444444").pack(
        side="left", padx=12
    )

    def mostrar(*_):
        tabla.delete(*tabla.get_children())
        barras = estado["barras"]
        if not barras:
            return
        texto = var_texto.get()
        if not limpiar(texto):
            var_info.set(f"{len(barras):,} barras en el CSV. Escribe un nombre.")
            return
        resultados = buscar_barras(barras, texto)
        for barra, parecido in resultados:
            tabla.insert("", "end", values=(barra, f"{parecido} %", f"{barras[barra]:,}"))
        if resultados:
            primero = tabla.get_children()[0]
            tabla.selection_set(primero)
            var_info.set(
                f"{len(resultados)} parecida(s) de {len(barras):,} barras. "
                f"Doble clic o 'Copiar nombre' para copiar."
            )
        else:
            var_info.set("Ninguna barra se parece. Prueba con menos palabras.")

    def copiar(*_):
        seleccion = tabla.selection()
        if not seleccion:
            return
        barra = tabla.item(seleccion[0], "values")[0]
        root.clipboard_clear()
        root.clipboard_append(barra)
        var_info.set(f"Copiado: {barra}")

    def cargar(ruta):
        var_ruta.set(str(ruta))
        var_info.set("Leyendo el CSV...")
        boton_examinar.config(state="disabled")

        def trabajo():
            try:
                barras, error = leer_barras(ruta), None
            except Exception as exc:  # noqa: BLE001 - se muestra al usuario
                barras, error = {}, exc
            root.after(0, lambda: terminado(barras, error))

        def terminado(barras, error):
            boton_examinar.config(state="normal")
            if error is not None:
                var_info.set("No se pudo leer el archivo.")
                messagebox.showerror("No se pudo leer el CSV", str(error))
                return
            estado["barras"] = barras
            mostrar()
            entrada.focus_set()

        threading.Thread(target=trabajo, daemon=True).start()

    def examinar():
        ruta = filedialog.askopenfilename(
            title="CSV 15-minutal de CMg",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        )
        if ruta:
            cargar(ruta)

    boton_examinar.config(command=examinar)
    boton_copiar.config(command=copiar)
    tabla.bind("<Double-1>", copiar)
    var_texto.trace_add("write", mostrar)
    entrada.bind("<Return>", copiar)

    root.mainloop()


if __name__ == "__main__":
    main()
