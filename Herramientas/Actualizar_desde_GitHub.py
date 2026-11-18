# -*- coding: utf-8 -*-
"""
Actualizar_desde_GitHub.py

Trae los cambios del repositorio de GitHub a esta carpeta, sin Git y sin
copiar a mano desde el zip.

    python Herramientas/Actualizar_desde_GitHub.py

Actualiza la carpeta del proyecto (la de arriba de Herramientas/), y solo
lo que hace falta para correr (pedido del usuario, 2026-11-18): los .py
de las ventanas (Balance_BESS.py, Balance_CRA.py), requirements.txt, la
carpeta Herramientas/ y la carpeta Script/.

1. Se elige la rama en la lista desplegable (se carga sola al abrir; por
   defecto "main", donde quedan los cambios una vez aceptado el pull
   request; las "claude/..." son las de trabajo). Al abrir, y al cambiar
   de rama, se revisa sola.
2. La lista muestra SOLO lo que cambio, como un arbol de carpetas con
   casillas: las carpetas vienen cerradas (la flecha las abre) y todo
   viene marcado. Marcar o desmarcar una carpeta marca o desmarca lo que
   cambio adentro.
3. "Actualizar" copia lo marcado. Antes guarda la version anterior de cada
   archivo que reemplaza en __config__/respaldos_actualizacion/<fecha>/.

Nunca borra nada. Nunca toca __config__, tks.json, config.json ni las
salidas (no vienen en el zip). Solo usa la biblioteca estandar de Python
(urllib usa los certificados de Windows, que suele ser lo que pide el
proxy de la pega).
"""

import io
import json
import queue
import shutil
import ssl
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

DUENO = "italocarreran"
REPO = "Balance_SSCC"
RAMA_POR_DEFECTO = "main"

# La carpeta del proyecto: la de arriba si este archivo esta en Herramientas/.
_ESTA = Path(__file__).resolve().parent
AQUI = _ESTA.parent if _ESTA.name.lower() == "herramientas" else _ESTA

# Lo que se actualiza: lo necesario para correr los programas.
CARPETAS_ALCANCE = ("Script", "Herramientas")
EXTENSIONES_RAIZ = (".py", ".txt")

# Lo que nunca se escribe aunque viniera en el zip.
PROTEGIDOS = {"config.json", "tks.json"}
CARPETAS_PROTEGIDAS = {"__config__", "__pycache__", ".git", ".venv", "venv"}

MARCADO, VACIO, PARCIAL = "☑", "☐", "▣"


# ------------------------------------------------------------
# Configuracion (la rama elegida la ultima vez)
# ------------------------------------------------------------

def carpeta_config():
    """__config__ hermana de esta carpeta (o la mas cercana hacia arriba)."""

    for carpeta in [AQUI, *AQUI.parents]:
        if (carpeta / "__config__").is_dir():
            return carpeta / "__config__"
    return AQUI.parent / "__config__"


def ruta_ajustes():
    return carpeta_config() / "actualizar_github.json"


def leer_ajustes():
    try:
        return json.loads(ruta_ajustes().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_ajustes(datos):
    try:
        ruta = ruta_ajustes()
        ruta.parent.mkdir(parents=True, exist_ok=True)
        tmp = ruta.with_suffix(".tmp")
        tmp.write_text(json.dumps(datos, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(ruta)
    except OSError:
        pass


# ------------------------------------------------------------
# GitHub
# ------------------------------------------------------------

def _abrir(url, timeout=60):
    pedido = urllib.request.Request(url, headers={"User-Agent": "Actualizar-Balance-SSCC"})
    contexto = ssl.create_default_context()
    return urllib.request.urlopen(pedido, timeout=timeout, context=contexto)


def listar_ramas():
    """Las ramas del repo, con main primero."""

    ramas = []
    pagina = 1
    while pagina <= 10:
        url = (f"https://api.github.com/repos/{DUENO}/{REPO}/branches"
               f"?per_page=100&page={pagina}")
        with _abrir(url, timeout=30) as r:
            datos = json.loads(r.read().decode("utf-8"))
        if not datos:
            break
        ramas += [d["name"] for d in datos if isinstance(d, dict) and "name" in d]
        if len(datos) < 100:
            break
        pagina += 1
    ramas = sorted(set(ramas), key=str.lower)
    if RAMA_POR_DEFECTO in ramas:
        ramas.remove(RAMA_POR_DEFECTO)
        ramas.insert(0, RAMA_POR_DEFECTO)
    return ramas


def bajar_zip(rama):
    url = f"https://codeload.github.com/{DUENO}/{REPO}/zip/refs/heads/{rama}"
    with _abrir(url, timeout=120) as r:
        return r.read()


def archivos_del_zip(contenido):
    """{ruta relativa con '/': bytes}, sin la carpeta raiz del zip."""

    archivos = {}
    with zipfile.ZipFile(io.BytesIO(contenido)) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            partes = info.filename.split("/", 1)
            if len(partes) < 2 or not partes[1]:
                continue
            archivos[partes[1]] = z.read(info)
    return archivos


# ------------------------------------------------------------
# Comparacion
# ------------------------------------------------------------

def en_alcance(relativa):
    """Lo necesario para correr: .py/.txt de la raiz, Script/ y Herramientas/."""

    partes = relativa.split("/")
    if partes[-1] in PROTEGIDOS or any(p in CARPETAS_PROTEGIDAS for p in partes):
        return False
    if len(partes) == 1:
        return Path(relativa).suffix.lower() in EXTENSIONES_RAIZ
    return partes[0] in CARPETAS_ALCANCE


def _iguales(a, b):
    # Mismo contenido aunque uno tenga fin de linea Windows y el otro no.
    return a == b or a.replace(b"\r\n", b"\n") == b.replace(b"\r\n", b"\n")


def comparar(remotos, aqui=None):
    """{ruta: "Nuevo" | "Modificado"} de lo que cambio, dentro del alcance."""

    aqui = Path(aqui or AQUI)
    cambios = {}
    for relativa, datos in sorted(remotos.items()):
        if not en_alcance(relativa):
            continue
        local = aqui / relativa
        if not local.exists():
            cambios[relativa] = "Nuevo"
            continue
        try:
            if not _iguales(local.read_bytes(), datos):
                cambios[relativa] = "Modificado"
        except OSError:
            cambios[relativa] = "Modificado"
    return cambios


def aplicar(remotos, rutas, aqui=None):
    """Escribe las rutas; respalda antes lo que reemplaza. Devuelve la carpeta de respaldo."""

    aqui = Path(aqui or AQUI)
    respaldo = (carpeta_config() / "respaldos_actualizacion"
                / datetime.now().strftime("%Y-%m-%d_%H%M%S"))
    hubo_respaldo = False
    for relativa in rutas:
        destino = aqui / relativa
        if destino.exists():
            copia = respaldo / relativa
            copia.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(destino, copia)
            hubo_respaldo = True
        destino.parent.mkdir(parents=True, exist_ok=True)
        tmp = destino.with_name(destino.name + ".tmp_actualizar")
        tmp.write_bytes(remotos[relativa])
        tmp.replace(destino)
    return respaldo if hubo_respaldo else None


# ------------------------------------------------------------
# Ventana
# ------------------------------------------------------------

class Ventana:
    def __init__(self, raiz):
        self.raiz = raiz
        self.cola = queue.Queue()
        self.remotos = {}
        self.cambios = {}          # ruta -> "Nuevo" | "Modificado"
        self.marcados = set()      # rutas marcadas
        self.item_de = {}          # ruta (archivo o carpeta) -> item del arbol
        self.ruta_de = {}          # item -> ruta
        self.ocupado = False
        self.inicio = None

        raiz.title("Actualizar desde GitHub")
        raiz.geometry("720x560")
        raiz.minsize(560, 420)

        marco = ttk.Frame(raiz, padding=10)
        marco.pack(fill="both", expand=True)

        fila = ttk.Frame(marco)
        fila.pack(fill="x")
        ttk.Label(fila, text="Carpeta:").pack(side="left")
        etiqueta = ttk.Label(fila, text=str(AQUI), foreground="#1a5fb4", cursor="hand2")
        etiqueta.pack(side="left", padx=6)
        etiqueta.bind("<Button-1>", lambda e: self._abrir_carpeta(AQUI))

        fila = ttk.Frame(marco)
        fila.pack(fill="x", pady=(8, 0))
        ttk.Label(fila, text="Rama:").pack(side="left")
        self.rama = tk.StringVar(value=leer_ajustes().get("rama") or RAMA_POR_DEFECTO)
        self.combo = ttk.Combobox(fila, textvariable=self.rama, width=44,
                                  values=[self.rama.get()], state="readonly")
        self.combo.pack(side="left", padx=6)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self.revisar())
        self.btn_revisar = ttk.Button(fila, text="Volver a revisar", command=self.revisar)
        self.btn_revisar.pack(side="left")

        self.estado = ttk.Label(marco, text="")
        self.estado.pack(fill="x", pady=(8, 4))

        lista = ttk.Frame(marco)
        lista.pack(fill="both", expand=True)
        self.arbol = ttk.Treeview(lista, columns=("estado",), show="tree headings",
                                  selectmode="none")
        self.arbol.heading("#0", text="Lo que cambió")
        self.arbol.heading("estado", text="Estado")
        self.arbol.column("estado", width=170, anchor="w", stretch=False)
        barra_v = ttk.Scrollbar(lista, orient="vertical", command=self.arbol.yview)
        self.arbol.configure(yscrollcommand=barra_v.set)
        self.arbol.pack(side="left", fill="both", expand=True)
        barra_v.pack(side="right", fill="y")
        self.arbol.bind("<Button-1>", self._clic)

        fila = ttk.Frame(marco)
        fila.pack(fill="x", pady=(8, 0))
        self.btn_aplicar = ttk.Button(fila, text="Actualizar lo marcado",
                                      command=self.actualizar, state="disabled")
        self.btn_aplicar.pack(side="left")
        ttk.Button(fila, text="Salir", command=raiz.destroy).pack(side="right")

        self.log = tk.Text(marco, height=5, wrap="word", state="disabled")
        self.log.pack(fill="x", pady=(8, 0))

        raiz.after(100, self._procesar_cola)
        raiz.after(500, self._reloj)
        raiz.after(200, self.revisar)
        raiz.after(300, self.cargar_ramas)

    # --- utilidades ---

    def _abrir_carpeta(self, ruta):
        import os
        import subprocess
        import sys
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(ruta))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(ruta)])
            else:
                subprocess.Popen(["xdg-open", str(ruta)])
        except OSError:
            pass

    def _escribir(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _ocupar(self, si, texto=""):
        self.ocupado = si
        self.btn_revisar.configure(state="disabled" if si else "normal")
        self.combo.configure(state="disabled" if si else "readonly")
        self.inicio = time.time() if si else None
        self.estado.configure(text=texto)
        self._pintar_boton()

    def _reloj(self):
        if self.inicio is not None:
            texto = self.estado.cget("text").split("  (")[0]
            self.estado.configure(text=f"{texto}  ({time.time() - self.inicio:.0f} s)")
        self.raiz.after(500, self._reloj)

    def _en_hilo(self, funcion):
        def correr():
            try:
                funcion()
            except urllib.error.HTTPError as e:
                if e.code == 404:
                    self.cola.put(("error", "No existe esa rama en GitHub (404)."))
                else:
                    self.cola.put(("error", f"GitHub respondio {e.code}: {e.reason}"))
            except urllib.error.URLError as e:
                self.cola.put(("error", "No se pudo conectar con GitHub: "
                                        f"{e.reason}. Revisar internet / proxy."))
            except Exception as e:  # noqa: BLE001 - se muestra en la ventana
                self.cola.put(("error", f"{type(e).__name__}: {e}"))
        threading.Thread(target=correr, daemon=True).start()

    def _procesar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                getattr(self, "_al_" + tipo)(dato)
        except queue.Empty:
            pass
        self.raiz.after(100, self._procesar_cola)

    # --- ramas ---

    def cargar_ramas(self):
        """La lista desplegable de ramas, cargada en segundo plano al abrir."""

        def trabajo():
            try:
                self.cola.put(("ramas", listar_ramas()))
            except Exception as e:  # noqa: BLE001 - se avisa en el registro
                self.cola.put(("aviso", f"No se pudo cargar la lista de ramas: {e}"))
        threading.Thread(target=trabajo, daemon=True).start()

    def _al_ramas(self, ramas):
        actual = self.rama.get()
        if actual and actual not in ramas:
            ramas = [actual, *ramas]
        self.combo.configure(values=ramas)
        self._escribir(f"{len(ramas)} rama(s) en GitHub: elegir en la lista 'Rama'.")

    def _al_aviso(self, texto):
        self._escribir(texto)

    # --- revisar ---

    def revisar(self):
        rama = self.rama.get().strip()
        if not rama or self.ocupado:
            return
        self._ocupar(True, f"Revisando la rama '{rama}'...")
        self.arbol.delete(*self.arbol.get_children())

        def trabajo():
            remotos = archivos_del_zip(bajar_zip(rama))
            self.cola.put(("revision", (rama, remotos, comparar(remotos))))
        self._en_hilo(trabajo)

    def _al_revision(self, datos):
        rama, remotos, cambios = datos
        self.remotos = remotos
        self.cambios = cambios
        self.marcados = set(cambios)      # todo lo que cambio, marcado
        guardar_ajustes({**leer_ajustes(), "rama": rama})
        self._armar_arbol()
        nuevos = sum(1 for e in cambios.values() if e == "Nuevo")
        if cambios:
            texto = (f"Rama '{rama}': {len(cambios) - nuevos} modificado(s), "
                     f"{nuevos} nuevo(s).")
        else:
            texto = f"Rama '{rama}': ya está todo al día."
        self._ocupar(False, texto)

    def _armar_arbol(self):
        self.arbol.delete(*self.arbol.get_children())
        self.item_de, self.ruta_de = {}, {}
        for relativa in sorted(self.cambios, key=lambda r: (r.count("/") == 0, r.lower())):
            partes = relativa.split("/")
            padre = ""
            for i in range(1, len(partes)):
                carpeta = "/".join(partes[:i])
                if carpeta not in self.item_de:
                    item = self.arbol.insert(padre, "end", open=False, text="")
                    self.item_de[carpeta], self.ruta_de[item] = item, carpeta + "/"
                padre = self.item_de[carpeta]
            item = self.arbol.insert(padre, "end", text="",
                                     values=(self.cambios[relativa],))
            self.item_de[relativa], self.ruta_de[item] = item, relativa
        self._pintar()

    # --- casillas ---

    def _archivos_bajo(self, ruta):
        if not ruta.endswith("/"):
            return [ruta]
        return [r for r in self.cambios if r.startswith(ruta)]

    def _pintar(self):
        for item, ruta in self.ruta_de.items():
            archivos = self._archivos_bajo(ruta)
            marcados = sum(1 for r in archivos if r in self.marcados)
            casilla = (MARCADO if marcados == len(archivos)
                       else VACIO if marcados == 0 else PARCIAL)
            nombre = ruta.rstrip("/").split("/")[-1]
            if ruta.endswith("/"):
                self.arbol.item(item, text=f"{casilla}  {nombre}/",
                                values=(f"{len(archivos)} con cambios",))
            else:
                self.arbol.item(item, text=f"{casilla}  {nombre}")
        self._pintar_boton()

    def _pintar_boton(self):
        listo = bool(self.marcados) and not self.ocupado
        self.btn_aplicar.configure(
            state="normal" if listo else "disabled",
            text=f"Actualizar lo marcado ({len(self.marcados)})" if self.marcados
            else "Actualizar lo marcado",
        )

    def alternar(self, ruta):
        archivos = self._archivos_bajo(ruta)
        if all(r in self.marcados for r in archivos):
            self.marcados.difference_update(archivos)
        else:
            self.marcados.update(archivos)
        self._pintar()

    def _clic(self, evento):
        item = self.arbol.identify_row(evento.y)
        if not item or self.ocupado:
            return None
        # La flecha abre y cierra la carpeta (lo hace el arbol solo).
        if "indicator" in self.arbol.identify_element(evento.x, evento.y):
            return None
        self.alternar(self.ruta_de[item])
        return "break"

    # --- actualizar ---

    def actualizar(self):
        rutas = sorted(self.marcados)
        if not rutas:
            return
        if not messagebox.askyesno(
                "Actualizar",
                f"Se van a escribir {len(rutas)} archivo(s) en\n{AQUI}\n\n"
                "Los que se reemplazan quedan respaldados en __config__/"
                "respaldos_actualizacion.\n\nCerrar antes el programa si esta abierto. "
                "¿Seguir?"):
            return
        self._ocupar(True, "Escribiendo...")
        remotos = self.remotos
        self._en_hilo(lambda: self.cola.put(("aplicado", (rutas, aplicar(remotos, rutas)))))

    def _al_aplicado(self, datos):
        rutas, respaldo = datos
        for r in rutas:
            self.cambios.pop(r, None)
        self.marcados.clear()
        self._armar_arbol()
        quedan = len(self.cambios)
        self._ocupar(False, f"Listo: {len(rutas)} archivo(s) actualizado(s)."
                     + (f" Quedan {quedan} sin actualizar." if quedan else ""))
        self._escribir(f"Actualizados: {len(rutas)} archivo(s).")
        if respaldo:
            self._escribir(f"Respaldo de lo anterior: {respaldo}")

    def _al_error(self, texto):
        self._ocupar(False, "Error")
        self._escribir("ERROR: " + texto)
        messagebox.showerror("Actualizar desde GitHub", texto)


def main():
    raiz = tk.Tk()
    Ventana(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
