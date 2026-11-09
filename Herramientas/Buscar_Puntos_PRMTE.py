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

La lista sale de GET /medidas-v2/measurement-point/all (portal de la
API, 2026-11-09), que entrega los puntos por paginas:
{"page", "pageSize", "totalElements", "totalPages",
 "content": [{"measurePointId", "subStation", "region", "tension",
              "coordinatorId"}]}.
"""

import difflib
import json
import re
import threading
import unicodedata
from datetime import date
from pathlib import Path

AQUI = Path(__file__).resolve().parent

# La lista de puntos de medida, por paginas (portal de la API).
URL_PUNTOS = "https://medidas.api.coordinador.cl/medidas-v2/measurement-point/all"
TAMANO_PAGINA = 1000
MAXIMO_PAGINAS = 500   # tope de seguridad
# Los datos de cada punto que se muestran, en este orden.
CAMPOS = (("subStation", "Subestación"), ("tension", "Tensión"),
          ("region", "Región"), ("coordinatorId", "Coordinado"))
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
        campos = {c: str(r.get(c) or "").strip() for c, _ in CAMPOS}
        puntos.setdefault(ident, {"detalle": detalle, **campos})
    return [{"id": i, **datos} for i, datos in sorted(puntos.items())]


def bajar_puntos(token, avisar=print):
    """
    Todos los puntos de /measurement-point/all, pagina por pagina
    (page empieza en 0). Devuelve (puntos, url).
    """

    import requests

    registros, pagina, total, anterior = [], 0, None, None
    while pagina < MAXIMO_PAGINAS:
        avisar(f"Bajando la pagina {pagina + 1}"
               + (f" de {total}" if total else "") + "...")
        try:
            r = requests.get(
                URL_PUNTOS, timeout=TIMEOUT,
                params={"page": pagina, "pageSize": TAMANO_PAGINA,
                        "user_key": token},
                headers={"accept": "application/json"},
            )
        except requests.RequestException as error:
            raise RuntimeError(f"{URL_PUNTOS}: {error}") from error
        if r.status_code != 200:
            raise RuntimeError(
                f"{URL_PUNTOS} (pagina {pagina}): HTTP {r.status_code} "
                f"{r.text[:300]}"
            )
        try:
            datos = r.json()
        except ValueError as error:
            raise RuntimeError(f"{URL_PUNTOS}: la respuesta no es JSON") from error
        contenido = _registros(datos)
        # Sin contenido, o la misma pagina otra vez (la API ignoro 'page'):
        # no hay mas.
        if not contenido or contenido == anterior:
            break
        registros += contenido
        anterior = contenido
        if isinstance(datos, dict) and datos.get("totalPages") is not None:
            total = int(datos["totalPages"])
            if pagina + 1 >= total:
                break
        pagina += 1
    puntos = a_puntos(registros)
    if not puntos:
        raise RuntimeError(f"{URL_PUNTOS} respondio sin puntos de medida.")
    return puntos, URL_PUNTOS

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
# Cuadro de error con boton "Copiar" (copia de Script/dialogos.py: esta
# herramienta no importa nada del programa).
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.title("Buscar puntos de medida PRMTE")
    root.geometry("1000x580")

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
    columnas = (("id", "Punto de medida (id exacto)", 250),
                ("parecido", "Parecido", 70)) + tuple(
        (c, titulo, 130) for c, titulo in CAMPOS)
    tabla = ttk.Treeview(marco, columns=[c for c, _, _ in columnas],
                         show="headings", selectmode="browse")
    for col, titulo, ancho in columnas:
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
            tabla.insert("", "end", values=(
                p["id"], f"{nota} %", *[p.get(c, "") for c, _ in CAMPOS]))
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
            mostrar_error("API de medidas", str(error))
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
