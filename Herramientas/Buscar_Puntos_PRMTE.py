# -*- coding: utf-8 -*-
"""
Buscar puntos de medida PRMTE por nombre aproximado.

Funciona guardado en cualquier carpeta: no usa nada del resto del
programa. Solo necesita Python con 'requests'.

  1. Baja de la API de medidas del Coordinador la lista de puntos de
     medida y la guarda (en __config__ si la encuentra hacia arriba de
     esta carpeta; si no, al lado de este archivo). Despues busca sin
     conectarse, hasta que se aprieta 'Actualizar lista'. Si la API
     informa mas puntos de los que llegaron, lo avisa.
  3. 'Comprobar en la API': un id exacto que no esta en la lista se
     pregunta a la API de medidas (la misma de la descarga, mes
     anterior). Si tiene medidas, existe: se agrega a la lista y queda
     guardado aparte ("encontrados_por_medidas"), asi no se pierde al
     volver a actualizar.
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


def _ids(registros):
    """(id, registro) de cada registro que trae un id de punto reconocible."""

    for punto in a_puntos(registros):
        yield punto["id"], punto


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

    No se confia en 'totalPages' ni en 'totalElements' para parar: se
    sigue pidiendo la pagina siguiente hasta que la API conteste vacio,
    repita la pagina anterior o de error. Un error en la PRIMERA pagina
    es un error de verdad (token, red); despues, es el fin de la lista y
    se queda con lo bajado. Al final se compara con 'totalElements'.
    """

    import requests

    registros, pagina, total, anterior = [], 0, None, None
    total_elementos = None
    vistos = set()
    fin = "tope de paginas"
    while pagina < MAXIMO_PAGINAS:
        avisar(f"Bajando la pagina {pagina + 1}"
               + (f" (la API dice {total})" if total else "") + "...")
        problema = None
        try:
            r = requests.get(
                URL_PUNTOS, timeout=TIMEOUT,
                params={"page": pagina, "pageSize": TAMANO_PAGINA,
                        "user_key": token},
                headers={"accept": "application/json"},
            )
            if r.status_code != 200:
                problema = f"HTTP {r.status_code} {r.text[:300]}"
            else:
                try:
                    datos = r.json()
                except ValueError:
                    problema = "la respuesta no es JSON"
        except requests.RequestException as error:
            problema = str(error)

        if problema is not None:
            if pagina == 0:
                raise RuntimeError(f"{URL_PUNTOS} (pagina 0): {problema}")
            fin = f"la pagina {pagina + 1} dio error ({problema[:80]})"
            break

        contenido = _registros(datos)
        if not contenido:
            fin = f"la pagina {pagina + 1} vino vacia"
            break
        if contenido == anterior:
            fin = f"la pagina {pagina + 1} repitio la anterior"
            break
        nuevos = {str(i) for i, _ in _ids(contenido)} - vistos
        if not nuevos:
            # La API ignora 'page' pasado el final y devuelve algo ya
            # bajado (otra pagina, no la anterior): no hay mas.
            fin = f"la pagina {pagina + 1} solo traia puntos ya bajados"
            break
        vistos |= nuevos
        registros += contenido
        anterior = contenido
        if isinstance(datos, dict):
            if datos.get("totalPages") is not None:
                total = int(datos["totalPages"])
            if datos.get("totalElements") is not None:
                total_elementos = int(datos["totalElements"])
        pagina += 1

    puntos = a_puntos(registros)
    if not puntos:
        raise RuntimeError(f"{URL_PUNTOS} respondio sin puntos de medida.")
    resumen = (f"{len(puntos):,} puntos en {pagina} pagina(s); se paro "
               f"porque {fin}")
    if total_elementos is not None and len(registros) < total_elementos:
        avisar(f"OJO: la API informa {total_elementos:,} puntos y llegaron "
               f"{len(registros):,} ({resumen}): la lista quedo incompleta. "
               f"Usa 'Comprobar en la API' para un punto que no aparezca.")
    else:
        avisar(resumen + (f"; la API informa {total_elementos:,}."
                          if total_elementos is not None else "."))
    return puntos, URL_PUNTOS


# ------------------------------------------------------------
# Un punto que no esta en la lista: se pregunta por sus medidas
# ------------------------------------------------------------

# El mismo endpoint y la misma forma de pedir que la descarga de
# Medidas_SAE (Script/Medidas/Descarga_PRMTE.py): si el punto tiene
# medidas, existe aunque /measurement-point/all no lo traiga.
URL_MEDIDAS = "https://medidas.api.coordinador.cl/medidas-v2/measurement"
CANALES = (1, 3)


def mes_anterior(hoy=None):
    hoy = hoy or date.today()
    anio, mes = (hoy.year, hoy.month - 1) if hoy.month > 1 else (hoy.year - 1, 12)
    return f"{anio:04d}{mes:02d}"


def comprobar_punto(token, ident, periodo=None):
    """
    Pregunta a la API de medidas por 'ident' (id exacto), canal 1 y 3,
    en el mes 'periodo' (AAAAMM; por defecto el mes anterior). Devuelve
    (punto, mensaje): punto es el dict de la lista (con 'origen' =
    "medidas") si respondio con medidas, o None.
    """

    import requests

    ident = str(ident or "").strip()
    if not ident:
        return None, "Escribe el id exacto del punto."
    periodo = periodo or mes_anterior()

    ultimo = ""
    for canal in CANALES:
        try:
            r = requests.get(
                URL_MEDIDAS, timeout=TIMEOUT,
                params={"channelId": canal, "measurePointId": ident,
                        "period": f"{periodo}010000", "user_key": token},
                headers={"accept": "application/json"},
            )
        except requests.RequestException as error:
            ultimo = str(error)
            continue
        if r.status_code != 200:
            ultimo = f"HTTP {r.status_code} {r.text[:200]}"
            continue
        try:
            datos = r.json()
        except ValueError:
            ultimo = "la respuesta no es JSON"
            continue
        registro = datos[0] if isinstance(datos, list) and datos else None
        if not isinstance(registro, dict) or not registro.get("measurement"):
            ultimo = f"canal {canal} sin medidas en {periodo}"
            continue

        campos = {c: str(registro.get(c) or "").strip() for c, _ in CAMPOS}
        if not campos["subStation"]:
            campos["subStation"] = "(encontrado por medidas)"
        punto = {"id": ident, "detalle": "encontrado por medidas", **campos,
                 "origen": "medidas"}
        return punto, (
            f"'{ident}' existe: tiene medidas en {periodo} (canal {canal}), "
            f"pero /measurement-point/all no lo trae. Quedo agregado a la lista."
        )

    return None, f"'{ident}' no respondio con medidas en {periodo} ({ultimo})."


def unir_encontrados(puntos, encontrados):
    """La lista bajada mas los puntos hallados por medidas que no trae."""

    ids = {p["id"] for p in puntos}
    return puntos + [p for p in encontrados if p["id"] not in ids]


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

    estado = {"puntos": [], "encontrados": [], "fecha": "", "origen": ""}
    var_token = tk.StringVar(value=token_guardado())
    var_texto = tk.StringVar()
    var_info = tk.StringVar()

    arriba = tk.LabelFrame(root, text="Lista de puntos (API de medidas)", padx=10, pady=6)
    arriba.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(arriba, text="Token PRMTE:").pack(side="left")
    tk.Entry(arriba, textvariable=var_token, show="•", width=34).pack(side="left", padx=6)
    boton_bajar = tk.Button(arriba, text="Actualizar lista")
    boton_bajar.pack(side="left", padx=6)
    boton_comprobar = tk.Button(arriba, text="Comprobar en la API")
    boton_comprobar.pack(side="left", padx=6)
    tk.Label(arriba, text="(id exacto escrito abajo)",
             fg="#444444", font=("Segoe UI", 8)).pack(side="left", padx=4)

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

    def guardar():
        ruta_lista().write_text(json.dumps(
            {"origen": estado["origen"], "fecha": estado["fecha"],
             "puntos": [p for p in estado["puntos"] if p.get("origen") != "medidas"],
             "encontrados_por_medidas": estado["encontrados"]},
            ensure_ascii=False, indent=1), encoding="utf-8")

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
                root.after(0, listo, puntos, url, None)
            except Exception as error:  # noqa: BLE001 - se muestra
                root.after(0, listo, None, None, error)

        threading.Thread(target=trabajo, daemon=True).start()

    def listo(puntos, url, error):
        boton_bajar.config(state="normal")
        if error is not None:
            var_info.set("No se pudo bajar la lista.")
            mostrar_error("API de medidas", str(error))
            return
        estado.update(origen=url, fecha=date.today().isoformat())
        estado["puntos"] = unir_encontrados(puntos, estado["encontrados"])
        guardar()
        resumen = var_info.get()
        mostrar()
        var_info.set(f"Lista bajada de {url}. {resumen}")
        entrada.focus_set()

    def comprobar():
        token = var_token.get().strip()
        ident = var_texto.get().strip()
        if not token:
            messagebox.showwarning("Falta el token", "Pega el token PRMTE.")
            return
        if not ident:
            messagebox.showwarning("Falta el id", "Escribe abajo el id exacto del punto.")
            return
        boton_comprobar.config(state="disabled")
        var_info.set(f"Consultando las medidas de '{ident}'...")

        def trabajo():
            try:
                punto, mensaje = comprobar_punto(token, ident)
            except Exception as error:  # noqa: BLE001 - se muestra
                punto, mensaje = None, str(error)
            root.after(0, comprobado, punto, mensaje)

        threading.Thread(target=trabajo, daemon=True).start()

    def comprobado(punto, mensaje):
        boton_comprobar.config(state="normal")
        if punto is not None:
            ya = {p["id"] for p in estado["puntos"]}
            if punto["id"] in ya:
                mensaje = f"'{punto['id']}' tiene medidas y ya estaba en la lista."
            else:
                estado["encontrados"].append(punto)
                estado["puntos"].append(punto)
                try:
                    guardar()
                except OSError as error:
                    mensaje += f" (no se pudo guardar: {error})"
            mostrar()
        var_info.set(mensaje)

    # La lista guardada de la vez anterior: se busca sin conectarse.
    try:
        guardada = json.loads(ruta_lista().read_text(encoding="utf-8"))
        estado["encontrados"] = guardada.get("encontrados_por_medidas", [])
        estado["puntos"] = unir_encontrados(guardada["puntos"], estado["encontrados"])
        estado.update(origen=guardada.get("origen", ""), fecha=guardada.get("fecha", ""))
        var_info.set(f"{len(estado['puntos']):,} puntos (lista del "
                     f"{guardada.get('fecha', '?')}). Escribe un nombre.")
    except (OSError, ValueError, KeyError):
        var_info.set("Pega el token y aprieta 'Actualizar lista' (solo la primera vez).")

    boton_bajar.config(command=actualizar)
    boton_comprobar.config(command=comprobar)
    boton_copiar.config(command=copiar)
    tabla.bind("<Double-1>", copiar)
    entrada.bind("<Return>", copiar)
    var_texto.trace_add("write", mostrar)
    root.mainloop()


if __name__ == "__main__":
    main()
