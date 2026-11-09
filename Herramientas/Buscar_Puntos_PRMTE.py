# -*- coding: utf-8 -*-
"""
Buscar puntos de medida PRMTE por nombre aproximado.

Funciona guardado en cualquier carpeta: no usa nada del resto del
programa. Solo necesita Python con 'requests'.

  1. Baja de la API de medidas del Coordinador la lista de puntos de
     medida y la guarda (en __config__ si la encuentra hacia arriba de
     esta carpeta; si no, al lado de este archivo). Despues busca sin
     conectarse, hasta que se aprieta 'Actualizar lista'.
  2. Se escribe el nombre aproximado ('arena 220', 'tocopila') y la
     lista se filtra mientras se escribe. Doble clic copia el id exacto.

El token se toma de __config__/tks.json ("coordinador" -> "prmte") si
esta; si no, se pega en la ventana.

OJO: la ruta de la API que entrega la LISTA de puntos no esta
confirmada (la que usa Balance BESS baja las medidas de UN punto). Se
prueban las rutas de RUTAS_LISTA en orden; si ninguna sirve, la ventana
dice que respondio cada una. Cuando se sepa la correcta, se deja sola.
"""

import difflib
import json
import re
import threading
import unicodedata
from datetime import date
from pathlib import Path

AQUI = Path(__file__).resolve().parent

# Rutas que se prueban para bajar la lista de puntos (SIN CONFIRMAR).
# {periodo} = mes anterior como AAAAMM.
RUTAS_LISTA = [
    "https://medidas.api.coordinador.cl/medidas-v2/measure-points",
    "https://medidas.api.coordinador.cl/medidas-v2/measurePoint",
    "https://medidas.api.coordinador.cl/medidas/api/puntos-medida/",
    "https://medidas.api.coordinador.cl/medidas/api/medidas/{periodo}/",
]
TIMEOUT = 120
MAXIMO_RESULTADOS = 60


# ------------------------------------------------------------
# Donde guardar / leer (sirve desde cualquier carpeta)
# ------------------------------------------------------------

def carpeta_config():
    """El __config__ mas cercano hacia arriba, o esta misma carpeta."""

    for carpeta in [AQUI, *AQUI.parents]:
        if (carpeta / "__config__").is_dir():
            return carpeta / "__config__"
    return AQUI


def token_guardado():
    try:
        datos = json.loads((carpeta_config() / "tks.json").read_text(encoding="utf-8"))
        token = datos.get("coordinador", {}).get("prmte", "")
        return "" if token in ("", "PEGAR_AQUI_LA_CLAVE") else token.strip()
    except (OSError, ValueError, AttributeError):
        return ""


def ruta_lista():
    return carpeta_config() / "puntos_prmte.json"


# ------------------------------------------------------------
# La lista de puntos
# ------------------------------------------------------------

def _registros(datos):
    """La lista de dicts de una respuesta, venga como venga envuelta."""

    if isinstance(datos, list):
        return [d for d in datos if isinstance(d, dict)]
    if isinstance(datos, dict):
        for valor in datos.values():
            if isinstance(valor, list) and valor and isinstance(valor[0], dict):
                return valor
    return []


def a_puntos(registros):
    """[{'id':..., 'detalle':...}] sin repetidos, de cualquier forma de respuesta."""

    puntos = {}
    for r in registros:
        clave_id = next(
            (k for k in r if re.sub(r"[^a-z]", "", k.lower()) in
             ("idpuntomedida", "measurepointid", "idpunto", "puntomedida", "id")),
            None,
        )
        if clave_id is None or not str(r[clave_id]).strip():
            continue
        ident = str(r[clave_id]).strip()
        detalle = " · ".join(
            str(v).strip() for k, v in r.items()
            if k != clave_id and isinstance(v, (str, int)) and str(v).strip()
            and not isinstance(v, bool) and len(str(v)) < 80
        )
        puntos.setdefault(ident, detalle)
    return [{"id": i, "detalle": d} for i, d in sorted(puntos.items())]


def bajar_puntos(token, avisar=print):
    """Prueba RUTAS_LISTA hasta que una devuelva puntos."""

    import requests

    hoy = date.today()
    anterior = (hoy.replace(day=1).toordinal() - 1)
    periodo = date.fromordinal(anterior).strftime("%Y%m")
    respuestas = []
    for ruta in RUTAS_LISTA:
        url = ruta.format(periodo=periodo)
        avisar(f"Probando {url} ...")
        try:
            r = requests.get(url, params={"user_key": token}, timeout=TIMEOUT,
                             headers={"accept": "application/json"})
        except requests.RequestException as error:
            respuestas.append(f"{url}: {error}")
            continue
        if r.status_code != 200:
            respuestas.append(f"{url}: HTTP {r.status_code} {r.text[:150]}")
            continue
        try:
            puntos = a_puntos(_registros(r.json()))
        except ValueError:
            respuestas.append(f"{url}: no es JSON")
            continue
        if puntos:
            return puntos, url
        respuestas.append(f"{url}: respondio sin puntos")
    raise RuntimeError("Ninguna ruta entrego la lista de puntos:\n\n" + "\n".join(respuestas))


# ------------------------------------------------------------
# La busqueda
# ------------------------------------------------------------

def limpiar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^0-9A-Za-z]+", " ", texto.upper()).split())


def buscar(puntos, texto, limite=MAXIMO_RESULTADOS):
    """[(punto, parecido 0..100)]: todas las palabras (en cualquier orden)
    arriba; despues lo que se parece aunque este mal escrito."""

    palabras = limpiar(texto).split()
    if not palabras:
        return []
    salida = []
    for p in puntos:
        suyas = limpiar(p["id"] + " " + p["detalle"]).split()
        notas = []
        for palabra in palabras:
            if palabra in suyas or (palabra.isdigit() and any(
                    s.isdigit() and int(s) == int(palabra) for s in suyas)):
                notas.append(1.0)
            elif any(palabra in s for s in suyas):
                notas.append(0.9)
            else:
                notas.append(max((difflib.SequenceMatcher(None, palabra, s).ratio()
                                  for s in suyas), default=0.0))
        nota = sum(notas) / len(notas)
        if nota >= 0.6:
            salida.append((p, round(100 * nota)))
    salida.sort(key=lambda x: (-x[1], x[0]["id"]))
    return salida[:limite]


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.title("Buscar puntos de medida PRMTE")
    root.geometry("900x580")

    estado = {"puntos": []}
    var_token = tk.StringVar(value=token_guardado())
    var_texto = tk.StringVar()
    var_info = tk.StringVar()

    arriba = tk.LabelFrame(root, text="Lista de puntos (API de medidas)", padx=10, pady=6)
    arriba.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(arriba, text="Token PRMTE:").pack(side="left")
    tk.Entry(arriba, textvariable=var_token, show="•", width=34).pack(side="left", padx=6)
    boton_bajar = tk.Button(arriba, text="Actualizar lista")
    boton_bajar.pack(side="left", padx=6)

    medio = tk.LabelFrame(root, text="Nombre aproximado del punto", padx=10, pady=6)
    medio.pack(fill="x", padx=14, pady=4)
    entrada = tk.Entry(medio, textvariable=var_texto, font=("Consolas", 11, "bold"))
    entrada.pack(side="left", fill="x", expand=True)
    tk.Label(medio, text="ej. arena 220, tocopilla", fg="#444444",
             font=("Segoe UI", 8)).pack(side="left", padx=8)

    marco = tk.Frame(root)
    marco.pack(fill="both", expand=True, padx=14, pady=4)
    tabla = ttk.Treeview(marco, columns=("id", "parecido", "detalle"),
                         show="headings", selectmode="browse")
    for col, titulo, ancho in (("id", "Punto de medida (id exacto)", 260),
                               ("parecido", "Parecido", 80),
                               ("detalle", "Detalle", 480)):
        tabla.heading(col, text=titulo)
        tabla.column(col, width=ancho, anchor="center" if col == "parecido" else "w")
    sb = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)

    abajo = tk.Frame(root)
    abajo.pack(fill="x", padx=14, pady=(4, 10))
    boton_copiar = tk.Button(abajo, text="Copiar id", bg="#1f5fa8", fg="white",
                             font=("Segoe UI", 9, "bold"))
    boton_copiar.pack(side="left")
    tk.Label(abajo, textvariable=var_info, anchor="w", fg="#444444",
             wraplength=700, justify="left").pack(side="left", padx=12)

    def mostrar(*_):
        tabla.delete(*tabla.get_children())
        puntos = estado["puntos"]
        if not puntos:
            return
        if not limpiar(var_texto.get()):
            var_info.set(f"{len(puntos):,} puntos en la lista. Escribe un nombre.")
            return
        resultados = buscar(puntos, var_texto.get())
        for p, nota in resultados:
            tabla.insert("", "end", values=(p["id"], f"{nota} %", p["detalle"]))
        if resultados:
            tabla.selection_set(tabla.get_children()[0])
        var_info.set(f"{len(resultados)} parecido(s) de {len(puntos):,}. "
                     f"Doble clic o 'Copiar id'." if resultados
                     else "Ningun punto se parece. Prueba con menos palabras.")

    def copiar(*_):
        sel = tabla.selection()
        if sel:
            ident = tabla.item(sel[0], "values")[0]
            root.clipboard_clear()
            root.clipboard_append(ident)
            var_info.set(f"Copiado: {ident}")

    def actualizar():
        token = var_token.get().strip()
        if not token:
            messagebox.showwarning("Falta el token", "Pega el token PRMTE.")
            return
        boton_bajar.config(state="disabled")

        def trabajo():
            try:
                puntos, url = bajar_puntos(
                    token, avisar=lambda m: root.after(0, var_info.set, m))
                ruta_lista().write_text(json.dumps(
                    {"origen": url, "fecha": date.today().isoformat(),
                     "puntos": puntos}, ensure_ascii=False, indent=1),
                    encoding="utf-8")
                root.after(0, listo, puntos, f"Lista bajada de {url}.", None)
            except Exception as error:  # noqa: BLE001 - se muestra
                root.after(0, listo, None, None, error)

        threading.Thread(target=trabajo, daemon=True).start()

    def listo(puntos, mensaje, error):
        boton_bajar.config(state="normal")
        if error is not None:
            var_info.set("No se pudo bajar la lista.")
            messagebox.showerror("API de medidas", str(error))
            return
        estado["puntos"] = puntos
        mostrar()
        var_info.set(mensaje + " " + var_info.get())
        entrada.focus_set()

    # La lista guardada de la vez anterior: se busca sin conectarse.
    try:
        guardada = json.loads(ruta_lista().read_text(encoding="utf-8"))
        estado["puntos"] = guardada["puntos"]
        var_info.set(f"{len(estado['puntos']):,} puntos (lista del "
                     f"{guardada.get('fecha', '?')}). Escribe un nombre.")
    except (OSError, ValueError, KeyError):
        var_info.set("Pega el token y aprieta 'Actualizar lista' (solo la primera vez).")

    boton_bajar.config(command=actualizar)
    boton_copiar.config(command=copiar)
    tabla.bind("<Double-1>", copiar)
    entrada.bind("<Return>", copiar)
    var_texto.trace_add("write", mostrar)
    root.mainloop()


if __name__ == "__main__":
    main()
