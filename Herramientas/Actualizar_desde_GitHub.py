# -*- coding: utf-8 -*-
"""
Actualizar_desde_GitHub.py

Trae los cambios del repositorio de GitHub a esta carpeta, sin Git y sin
copiar a mano desde el zip.

    python Herramientas/Actualizar_desde_GitHub.py

Actualiza la carpeta del proyecto (la de arriba de Herramientas/).

1. Se elige la rama (por defecto "main", que es donde quedan los cambios
   una vez aceptado el pull request; las "claude/..." son las de trabajo).
2. "Revisar cambios" baja el zip de esa rama y lo compara con esta
   carpeta: lista los archivos nuevos y los modificados.
3. "Actualizar" copia solo esos. Antes guarda la version anterior de cada
   archivo que reemplaza en __config__/respaldos_actualizacion/<fecha>/.

Nunca borra nada: los archivos que ya no estan en GitHub solo se avisan.
Nunca toca __config__, tks.json, config.json ni las salidas (no vienen en
el zip). Solo usa la biblioteca estandar de Python (urllib usa los
certificados de Windows, que suele ser lo que pide el proxy de la pega).
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

# Lo que nunca se escribe aunque viniera en el zip.
PROTEGIDOS = {"config.json", "tks.json"}
CARPETAS_PROTEGIDAS = {"__config__", "__pycache__", ".git", ".venv", "venv"}

# Para el aviso de "ya no esta en GitHub": solo archivos de codigo y texto.
EXTENSIONES_AVISO = {".py", ".md", ".txt"}


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
    """Las ramas del repo, con main primero y las mas nuevas despues."""

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


def bajar_zip(rama, avance=None):
    url = f"https://codeload.github.com/{DUENO}/{REPO}/zip/refs/heads/{rama}"
    with _abrir(url, timeout=120) as r:
        partes = []
        total = 0
        while True:
            trozo = r.read(256 * 1024)
            if not trozo:
                break
            partes.append(trozo)
            total += len(trozo)
            if avance:
                avance(total)
    return b"".join(partes)


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

def _protegido(relativa):
    partes = relativa.split("/")
    return partes[-1] in PROTEGIDOS or any(p in CARPETAS_PROTEGIDAS for p in partes)


def _iguales(a, b):
    # Mismo contenido aunque uno tenga fin de linea Windows y el otro no.
    return a == b or a.replace(b"\r\n", b"\n") == b.replace(b"\r\n", b"\n")


def comparar(remotos):
    nuevos, modificados = [], []
    for relativa, datos in sorted(remotos.items()):
        if _protegido(relativa):
            continue
        local = AQUI / relativa
        if not local.exists():
            nuevos.append(relativa)
        else:
            try:
                if not _iguales(local.read_bytes(), datos):
                    modificados.append(relativa)
            except OSError:
                modificados.append(relativa)

    # Archivos locales de codigo que ya no estan en GitHub (solo aviso).
    sobrantes = []
    carpetas_repo = {r.split("/")[0] for r in remotos if "/" in r}
    for local in AQUI.rglob("*"):
        if not local.is_file() or local.suffix.lower() not in EXTENSIONES_AVISO:
            continue
        relativa = local.relative_to(AQUI).as_posix()
        if _protegido(relativa) or relativa in remotos:
            continue
        primera = relativa.split("/")[0]
        if "/" in relativa and primera not in carpetas_repo:
            continue  # carpetas que no son del repo (salidas, etc.)
        sobrantes.append(relativa)
    return nuevos, modificados, sorted(sobrantes)


def aplicar(remotos, rutas):
    """Escribe las rutas; respalda antes lo que reemplaza. Devuelve la carpeta de respaldo."""

    respaldo = (carpeta_config() / "respaldos_actualizacion"
                / datetime.now().strftime("%Y-%m-%d_%H%M%S"))
    hubo_respaldo = False
    for relativa in rutas:
        destino = AQUI / relativa
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
        self.pendientes = []
        self.ocupado = False
        self.inicio = None

        raiz.title("Actualizar desde GitHub")
        raiz.geometry("760x560")
        raiz.minsize(600, 420)

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
        ajustes = leer_ajustes()
        self.rama = tk.StringVar(value=ajustes.get("rama") or RAMA_POR_DEFECTO)
        self.combo = ttk.Combobox(fila, textvariable=self.rama, width=48,
                                  values=[RAMA_POR_DEFECTO])
        self.combo.pack(side="left", padx=6)
        self.btn_ramas = ttk.Button(fila, text="Cargar ramas", command=self.cargar_ramas)
        self.btn_ramas.pack(side="left")

        fila = ttk.Frame(marco)
        fila.pack(fill="x", pady=8)
        self.btn_revisar = ttk.Button(fila, text="Revisar cambios", command=self.revisar)
        self.btn_revisar.pack(side="left")
        self.btn_aplicar = ttk.Button(fila, text="Actualizar", command=self.actualizar,
                                      state="disabled")
        self.btn_aplicar.pack(side="left", padx=6)
        self.estado = ttk.Label(fila, text="")
        self.estado.pack(side="left", padx=10)

        self.barra = ttk.Progressbar(marco, mode="indeterminate")
        self.barra.pack(fill="x")

        lista = ttk.Frame(marco)
        lista.pack(fill="both", expand=True, pady=(8, 0))
        self.arbol = ttk.Treeview(lista, columns=("estado",), show="tree headings", height=10)
        self.arbol.heading("#0", text="Archivo")
        self.arbol.heading("estado", text="Estado")
        self.arbol.column("estado", width=150, anchor="w", stretch=False)
        barra_v = ttk.Scrollbar(lista, orient="vertical", command=self.arbol.yview)
        self.arbol.configure(yscrollcommand=barra_v.set)
        self.arbol.pack(side="left", fill="both", expand=True)
        barra_v.pack(side="right", fill="y")

        self.log = tk.Text(marco, height=7, wrap="word", state="disabled")
        self.log.pack(fill="x", pady=(8, 0))

        self._escribir(f"Repositorio: github.com/{DUENO}/{REPO}")
        self._escribir("Elegir la rama y presionar 'Revisar cambios'.")
        raiz.after(100, self._procesar_cola)
        raiz.after(500, self._reloj)

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
        estado = "disabled" if si else "normal"
        self.btn_revisar.configure(state=estado)
        self.btn_ramas.configure(state=estado)
        self.combo.configure(state=estado)
        if si:
            self.btn_aplicar.configure(state="disabled")
            self.inicio = time.time()
            self.barra.start(12)
        else:
            self.inicio = None
            self.barra.stop()
        self.estado.configure(text=texto)

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

    # --- acciones ---

    def cargar_ramas(self):
        self._ocupar(True, "Cargando ramas...")
        self._en_hilo(lambda: self.cola.put(("ramas", listar_ramas())))

    def _al_ramas(self, ramas):
        self._ocupar(False, f"{len(ramas)} ramas")
        self.combo.configure(values=ramas)
        self._escribir(f"Ramas disponibles: {len(ramas)}.")

    def revisar(self):
        rama = self.rama.get().strip()
        if not rama:
            return
        self._ocupar(True, f"Bajando '{rama}'...")
        self.arbol.delete(*self.arbol.get_children())
        self._escribir(f"Bajando la rama '{rama}'...")

        def trabajo():
            contenido = bajar_zip(rama, lambda n: None)
            remotos = archivos_del_zip(contenido)
            nuevos, modificados, sobrantes = comparar(remotos)
            self.cola.put(("revision", (rama, remotos, nuevos, modificados, sobrantes,
                                        len(contenido))))
        self._en_hilo(trabajo)

    def _al_revision(self, datos):
        rama, remotos, nuevos, modificados, sobrantes, peso = datos
        self.remotos = remotos
        self.pendientes = modificados + nuevos
        guardar_ajustes({**leer_ajustes(), "rama": rama})

        for r in modificados:
            self.arbol.insert("", "end", text=r, values=("Modificado",))
        for r in nuevos:
            self.arbol.insert("", "end", text=r, values=("Nuevo",))
        for r in sobrantes:
            self.arbol.insert("", "end", text=r, values=("Ya no esta en GitHub",))

        self._ocupar(False, f"{len(self.pendientes)} archivo(s) por actualizar")
        self._escribir(f"Zip de '{rama}': {peso / 1024:.0f} KB, {len(remotos)} archivos. "
                       f"Modificados: {len(modificados)}, nuevos: {len(nuevos)}.")
        if sobrantes:
            self._escribir(f"{len(sobrantes)} archivo(s) local(es) ya no estan en GitHub "
                           "(no se borran; revisar si sobran).")
        if self.pendientes:
            self.btn_aplicar.configure(state="normal")
        else:
            self._escribir("La carpeta ya esta al dia.")

    def actualizar(self):
        if not self.pendientes:
            return
        if not messagebox.askyesno(
                "Actualizar",
                f"Se van a escribir {len(self.pendientes)} archivo(s) en\n{AQUI}\n\n"
                "Los que se reemplazan quedan respaldados en __config__/"
                "respaldos_actualizacion.\n\nCerrar antes el programa si esta abierto. "
                "¿Seguir?"):
            return
        self._ocupar(True, "Escribiendo...")
        pendientes, remotos = list(self.pendientes), self.remotos
        self._en_hilo(lambda: self.cola.put(
            ("aplicado", (pendientes, aplicar(remotos, pendientes)))))

    def _al_aplicado(self, datos):
        rutas, respaldo = datos
        self._ocupar(False, "Listo")
        for r in rutas:
            self._escribir(f"  actualizado: {r}")
        self._escribir(f"Listo: {len(rutas)} archivo(s) actualizado(s).")
        if respaldo:
            self._escribir(f"Respaldo de lo anterior: {respaldo}")
        self.pendientes = []
        for item in self.arbol.get_children():
            if self.arbol.set(item, "estado") in ("Modificado", "Nuevo"):
                self.arbol.set(item, "estado", "Actualizado")

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
