# -*- coding: utf-8 -*-
"""
Reemplazar hojas de un .xlsx sin abrir el libro entero.

Pedido del usuario: "la hoja Centrales se demora mucho en armar, deberia
ser rapido [...] que se copien, se cierre y se peguen en el documento de
salida".

Por que demoraba: para actualizar UNA hoja de Balance_BESS.xlsx la
escritura abria el libro entero con openpyxl (todas las hojas, celda por
celda: ~27.000 filas por hoja, decenas de MB), copiaba celda por celda
todas las que NO cambiaban y volvia a guardar el libro completo. Con el
libro lleno eso son minutos, aunque la hoja nueva tenga 12 filas.

Ahora: un .xlsx es un zip con un XML por hoja. Las hojas nuevas se
escriben solas en un libro temporal chico (con el mismo codigo y el
mismo formato de siempre), y despues se arma el libro final copiando
el zip viejo entrada por entrada -- sin descomprimir ni interpretar las
hojas que no cambian -- y poniendo el XML de las hojas nuevas en su
lugar. El costo pasa a depender del tamaño de lo que se escribe, no del
libro entero.

Lo que hay que ajustar al pasar una hoja de un libro a otro (y nada mas):

  - textos: el libro temporal los guarda en SU tabla de textos
    compartidos (sharedStrings.xml); en la hoja pegada se escriben
    "en linea" (inlineStr), asi no se toca la tabla del libro final;
  - estilos: cada celda apunta a un estilo por numero (s="3") de SU
    styles.xml; se busca el mismo estilo en el libro final (fuente,
    relleno, borde, formato de numero, alineacion) y, si no esta, se
    agrega. Los estilos se agregan como texto al final de cada lista: el
    resto del styles.xml queda byte por byte como estaba (importa si el
    libro lo guardo Excel, que declara espacios de nombres que un parser
    XML descartaria al reescribirlo);
  - el indice del libro (workbook.xml, sus relaciones y
    [Content_Types].xml) cuando la hoja es nueva, y el orden de las
    hojas.

El libro final se escribe en un archivo temporal al lado y recien al
terminar reemplaza al original (os.replace): si algo falla a la mitad,
el libro que habia queda intacto.
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_TIPO_HOJA = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml."
    "worksheet+xml"
)
_REL_HOJA = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "worksheet"
)

_WORKBOOK = "xl/workbook.xml"
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_CONTENT_TYPES = "[Content_Types].xml"
_STYLES = "xl/styles.xml"
_SHARED = "xl/sharedStrings.xml"


# ============================================================
# LEER EL INDICE DE UN LIBRO
# ============================================================

def _ruta_parte(target):
    """'worksheets/sheet1.xml' o '/xl/worksheets/sheet1.xml' -> ruta en el zip."""

    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return "xl/" + target


def _hojas(zip_libro):
    """
    Las hojas del libro en su orden: lista de dicts con name, sheetId,
    rid, parte (la ruta del XML dentro del zip) y elemento (el texto
    original del <sheet .../> en workbook.xml).
    """

    workbook = zip_libro.read(_WORKBOOK).decode("utf-8")
    rels = ET.fromstring(zip_libro.read(_WORKBOOK_RELS))

    destino = {
        rel.get("Id"): _ruta_parte(rel.get("Target"))
        for rel in rels.findall(f"{{{_REL_PKG}}}Relationship")
    }

    hojas = []

    for elemento in re.findall(r"<(?:\w+:)?sheet\s[^>]*/>", workbook):
        nodo = ET.fromstring(_con_espacios(elemento))
        rid = nodo.get(f"{{{_REL_DOC}}}id")
        hojas.append({
            "name": nodo.get("name"),
            "sheetId": int(nodo.get("sheetId")),
            "rid": rid,
            "parte": destino.get(rid),
            "elemento": elemento,
        })

    return hojas


def _con_espacios(elemento):
    """
    Un <sheet .../> suelto, copiado de workbook.xml, con los espacios de
    nombres declarados para que ET lo pueda leer (el r:id usa el prefijo
    que haya elegido quien escribio el libro).
    """

    # Los prefijos que usa el elemento y que no declara el mismo (openpyxl
    # repite xmlns:r en cada <sheet>; Excel lo declara solo en la raiz).
    declarados = set(re.findall(r"\sxmlns:(\w+)=", elemento))
    prefijos = set(re.findall(r"\s(\w+):\w+=", elemento)) - declarados
    prefijo_tag = re.match(r"<(\w+):", elemento)
    declaraciones = " ".join(
        f'xmlns:{p}="{_REL_DOC}"' for p in prefijos if p != "xmlns"
    )
    if prefijo_tag:
        declaraciones += f' xmlns:{prefijo_tag.group(1)}="{_MAIN}"'
        return elemento.replace(
            f"<{prefijo_tag.group(1)}:sheet",
            f"<{prefijo_tag.group(1)}:sheet {declaraciones}", 1,
        )
    return elemento.replace("<sheet", f"<sheet {declaraciones}", 1)


def nombres_de_hojas(ruta):
    """Los nombres de las hojas de un .xlsx, sin leer ninguna hoja."""

    with zipfile.ZipFile(ruta) as zip_libro:
        return [hoja["name"] for hoja in _hojas(zip_libro)]


# ============================================================
# TEXTOS COMPARTIDOS -> EN LINEA
# ============================================================

def _textos_compartidos(zip_libro):
    if _SHARED not in zip_libro.namelist():
        return []

    raiz = ET.fromstring(zip_libro.read(_SHARED))
    textos = []

    for si in raiz.findall(f"{{{_MAIN}}}si"):
        textos.append("".join(t.text or "" for t in si.iter(f"{{{_MAIN}}}t")))

    return textos


_CELDA_COMPARTIDA = re.compile(
    rb'<c ([^>]*?)t="s"([^>]*)><v>(\d+)</v></c>'
)


def _a_texto_en_linea(xml, textos):
    def reemplazo(match):
        texto = escape(textos[int(match.group(3))]).encode("utf-8")
        return (
            b"<c " + match.group(1) + b't="inlineStr"' + match.group(2)
            + b'><is><t xml:space="preserve">' + texto + b"</t></is></c>"
        )

    return _CELDA_COMPARTIDA.sub(reemplazo, xml)


# ============================================================
# ESTILOS
# ============================================================

class _Estilos:
    """
    El styles.xml del libro final, para agregarle (como texto) los
    estilos que traen las hojas nuevas.
    """

    _LISTAS = ("fonts", "fills", "borders", "cellXfs")

    def __init__(self, texto):
        self.texto = texto
        raiz = ET.fromstring(texto.encode("utf-8"))

        self.listas = {}
        for lista in self._LISTAS:
            nodo = raiz.find(f"{{{_MAIN}}}{lista}")
            self.listas[lista] = (
                [_canonico(hijo) for hijo in nodo] if nodo is not None else []
            )

        self.formatos = {}   # codigo de formato -> id
        nodo = raiz.find(f"{{{_MAIN}}}numFmts")
        if nodo is not None:
            for fmt in nodo:
                self.formatos[fmt.get("formatCode")] = int(fmt.get("numFmtId"))

        self.nuevos = {lista: [] for lista in self._LISTAS}
        self.formatos_nuevos = []

    def indice(self, lista, elemento):
        """La posicion de 'elemento' en 'lista', agregandolo si no esta."""

        canonico = _canonico(elemento)
        actuales = self.listas[lista]

        if canonico in actuales:
            return actuales.index(canonico)

        actuales.append(canonico)
        self.nuevos[lista].append(_canonico(elemento))
        return len(actuales) - 1

    def formato(self, codigo):
        if codigo in self.formatos:
            return self.formatos[codigo]

        nuevo = max([163] + list(self.formatos.values())) + 1
        self.formatos[codigo] = nuevo
        self.formatos_nuevos.append(
            f'<numFmt numFmtId="{nuevo}" formatCode="{escape(codigo, {chr(34): "&quot;"})}"/>'
        )
        return nuevo

    def texto_final(self):
        texto = self.texto

        if self.formatos_nuevos:
            texto = _agregar_a_lista(
                texto, "numFmts", self.formatos_nuevos,
                antes_de=("fonts",),
            )

        for lista in self._LISTAS:
            if self.nuevos[lista]:
                texto = _agregar_a_lista(texto, lista, self.nuevos[lista])

        return texto


def _canonico(elemento):
    return ET.tostring(elemento, encoding="unicode").replace(
        f' xmlns="{_MAIN}"', ""
    ).replace("ns0:", "").replace(f' xmlns:ns0="{_MAIN}"', "")


def _agregar_a_lista(texto, lista, elementos, antes_de=()):
    """
    Agrega 'elementos' al final de <lista> en el styles.xml (como texto)
    y actualiza su count. Si la lista no existe, la crea justo antes de
    la primera de 'antes_de' que encuentre.
    """

    patron = re.compile(
        rf"<(?:(\w+):)?{lista}(\s[^>]*)?(/?)>", re.S
    )
    match = patron.search(texto)
    nuevos = "".join(elementos)

    if match is None:
        for siguiente in antes_de:
            pos = re.search(rf"<(?:\w+:)?{siguiente}[\s>/]", texto)
            if pos:
                bloque = f'<{lista} count="{len(elementos)}">{nuevos}</{lista}>'
                return texto[: pos.start()] + bloque + texto[pos.start():]
        raise ValueError(f"styles.xml sin <{lista}>")

    prefijo = f"{match.group(1)}:" if match.group(1) else ""
    if prefijo:
        nuevos = re.sub(r"<(/?)(\w)", rf"<\1{prefijo}\2", nuevos)

    if match.group(3):  # <lista/> vacia
        atributos = re.sub(r'\scount="\d+"', "", match.group(2) or "")
        bloque = (
            f'<{prefijo}{lista}{atributos} count="{len(elementos)}">'
            f"{nuevos}</{prefijo}{lista}>"
        )
        return texto[: match.start()] + bloque + texto[match.end():]

    cierre = texto.index(f"</{prefijo}{lista}>", match.end())
    texto = texto[:cierre] + nuevos + texto[cierre:]

    apertura = match.group(0)
    contados = re.search(r'count="(\d+)"', apertura)
    if contados:
        total = int(contados.group(1)) + len(elementos)
        nueva_apertura = apertura.replace(
            contados.group(0), f'count="{total}"', 1
        )
        texto = (
            texto[: match.start()] + nueva_apertura
            + texto[match.start() + len(apertura):]
        )

    return texto


def _mapa_de_estilos(styles_origen, estilos_destino):
    """
    xf del libro temporal -> xf del libro final, para cada estilo de
    celda del temporal.
    """

    raiz = ET.fromstring(styles_origen)

    def lista(nombre):
        nodo = raiz.find(f"{{{_MAIN}}}{nombre}")
        return list(nodo) if nodo is not None else []

    fuentes, rellenos, bordes = lista("fonts"), lista("fills"), lista("borders")
    codigos = {
        int(fmt.get("numFmtId")): fmt.get("formatCode")
        for fmt in lista("numFmts")
    }

    mapa = {}

    for indice, xf in enumerate(lista("cellXfs")):

        nuevo = ET.fromstring(ET.tostring(xf))

        for atributo, nombre, elementos in (
            ("fontId", "fonts", fuentes),
            ("fillId", "fills", rellenos),
            ("borderId", "borders", bordes),
        ):
            valor = int(nuevo.get(atributo, 0))
            if valor < len(elementos):
                nuevo.set(
                    atributo,
                    str(estilos_destino.indice(nombre, elementos[valor])),
                )

        formato = int(nuevo.get("numFmtId", 0))
        if formato in codigos:
            nuevo.set("numFmtId", str(estilos_destino.formato(codigos[formato])))

        nuevo.set("xfId", "0")
        mapa[indice] = estilos_destino.indice("cellXfs", nuevo)

    return mapa


_ESTILO = re.compile(rb' s="(\d+)"')


def _remapear_estilos(xml, mapa):
    return _ESTILO.sub(
        lambda m: b' s="%d"' % mapa.get(int(m.group(1)), 0), xml
    )


# ============================================================
# PEGAR
# ============================================================

def pegar_hojas(ruta_destino, ruta_nuevas, orden=(), ignorar=()):
    """
    Pone en ruta_destino TODAS las hojas de ruta_nuevas (reemplazando
    las del mismo nombre o agregandolas) y deja las demas hojas de
    ruta_destino exactamente como estaban. Despues ordena las hojas
    segun 'orden' (las que no esten ahi quedan al final, en el orden que
    tenian). Si ruta_destino no existe, el libro nuevo pasa a serlo.

    ignorar: hojas de ruta_nuevas que no se pegan (un relleno).
    """

    ruta_destino = Path(ruta_destino)
    ruta_nuevas = Path(ruta_nuevas)

    if not ruta_destino.is_file():
        # Ignorar solo puede ser el relleno de un libro sin hojas: si
        # ademas no habia libro, no hay nada que escribir.
        if not ignorar:
            os.replace(ruta_nuevas, ruta_destino)
            if orden:
                _reordenar_en_el_lugar(ruta_destino, orden)
        return

    temporal = ruta_destino.with_name(f".~{ruta_destino.name}")

    with zipfile.ZipFile(ruta_destino) as viejo, \
            zipfile.ZipFile(ruta_nuevas) as nuevas:

        hojas_viejas = _hojas(viejo)
        hojas_nuevas = [
            hoja for hoja in _hojas(nuevas) if hoja["name"] not in ignorar
        ]
        textos = _textos_compartidos(nuevas)

        estilos = _Estilos(viejo.read(_STYLES).decode("utf-8"))
        mapa_estilos = _mapa_de_estilos(nuevas.read(_STYLES), estilos)

        por_nombre = {hoja["name"]: hoja for hoja in hojas_viejas}
        partes_usadas = {hoja["parte"] for hoja in hojas_viejas}
        reemplazos = {}       # parte -> bytes
        agregadas = []        # hojas nuevas que no existian

        rids = _ids_de_relaciones(viejo)
        siguiente_id = max([0] + [h["sheetId"] for h in hojas_viejas]) + 1

        for hoja in hojas_nuevas:

            xml = nuevas.read(hoja["parte"])
            xml = _a_texto_en_linea(xml, textos)
            xml = _remapear_estilos(xml, mapa_estilos)
            xml = xml.replace(b' tabSelected="1"', b"")

            existente = por_nombre.get(hoja["name"])

            if existente is not None:
                reemplazos[existente["parte"]] = xml
                continue

            numero = 1
            while f"xl/worksheets/sheet{numero}.xml" in partes_usadas:
                numero += 1
            parte = f"xl/worksheets/sheet{numero}.xml"
            partes_usadas.add(parte)

            rid = _nuevo_rid(rids)
            agregadas.append({
                "name": hoja["name"], "sheetId": siguiente_id,
                "rid": rid, "parte": parte, "xml": xml,
            })
            siguiente_id += 1

        workbook = viejo.read(_WORKBOOK).decode("utf-8")
        workbook = _workbook_final(workbook, hojas_viejas, agregadas, orden)

        rels = viejo.read(_WORKBOOK_RELS).decode("utf-8")
        tipos = viejo.read(_CONTENT_TYPES).decode("utf-8")

        for hoja in agregadas:
            destino = hoja["parte"][len("xl/"):]
            rels = rels.replace(
                "</Relationships>",
                f'<Relationship Id="{hoja["rid"]}" Type="{_REL_HOJA}" '
                f'Target="{destino}"/></Relationships>',
            )
            tipos = tipos.replace(
                "</Types>",
                f'<Override PartName="/{hoja["parte"]}" '
                f'ContentType="{_TIPO_HOJA}"/></Types>',
            )

        especiales = {
            _WORKBOOK: workbook.encode("utf-8"),
            _WORKBOOK_RELS: rels.encode("utf-8"),
            _CONTENT_TYPES: tipos.encode("utf-8"),
            _STYLES: estilos.texto_final().encode("utf-8"),
        }

        with zipfile.ZipFile(
            temporal, "w", compression=zipfile.ZIP_DEFLATED
        ) as final:

            for info in viejo.infolist():
                nombre = info.filename
                if nombre in especiales:
                    final.writestr(_info(nombre), especiales[nombre])
                elif nombre in reemplazos:
                    final.writestr(_info(nombre), reemplazos[nombre])
                else:
                    # Copia comprimida tal cual: ni se descomprime.
                    _copiar_crudo(viejo, info, final)

            for hoja in agregadas:
                final.writestr(_info(hoja["parte"]), hoja["xml"])

    os.replace(temporal, ruta_destino)


def _info(nombre):
    info = zipfile.ZipInfo(nombre, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    return info


def _copiar_crudo(origen, info, destino):
    """
    Copia una entrada de un zip a otro sin descomprimirla ni volver a
    comprimirla: es lo que hace que las hojas grandes que no cambian
    no cuesten nada. zipfile no tiene "copiar comprimido", asi que se
    lee el bloque comprimido tal cual del archivo de origen y se
    escribe con su misma cabecera (mismo CRC y tamaños), con los mismos
    pasos que hace ZipFile al abrir una entrada para escribir.
    """

    archivo = origen.fp
    archivo.seek(info.header_offset)
    cabecera = archivo.read(30)
    largo_nombre = int.from_bytes(cabecera[26:28], "little")
    largo_extra = int.from_bytes(cabecera[28:30], "little")
    archivo.seek(info.header_offset + 30 + largo_nombre + largo_extra)
    datos = archivo.read(info.compress_size)

    nueva = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    nueva.compress_type = info.compress_type
    nueva.external_attr = info.external_attr
    nueva.CRC = info.CRC
    nueva.compress_size = info.compress_size
    nueva.file_size = info.file_size
    # Sin "descriptor de datos" al final: el CRC y los tamaños ya van
    # en la cabecera. Se conserva la marca de nombre en UTF-8.
    nueva.flag_bits = info.flag_bits & 0x800

    destino.fp.seek(destino.start_dir)
    nueva.header_offset = destino.fp.tell()
    destino._writecheck(nueva)
    destino._didModify = True
    destino.fp.write(nueva.FileHeader(zip64=False))
    destino.fp.write(datos)
    destino.filelist.append(nueva)
    destino.NameToInfo[nueva.filename] = nueva
    destino.start_dir = destino.fp.tell()


def _ids_de_relaciones(zip_libro):
    rels = ET.fromstring(zip_libro.read(_WORKBOOK_RELS))
    return {
        rel.get("Id") for rel in rels.findall(f"{{{_REL_PKG}}}Relationship")
    }


def _nuevo_rid(usados):
    numero = 1
    while f"rId{numero}" in usados:
        numero += 1
    rid = f"rId{numero}"
    usados.add(rid)
    return rid


def _workbook_final(workbook, hojas_viejas, agregadas, orden):
    """
    workbook.xml con las hojas agregadas y en el orden pedido. Los
    <sheet .../> de las hojas que ya estaban se conservan tal cual
    (texto original); los nuevos usan el mismo prefijo de r:id.
    """

    # Cada <sheet> nuevo declara su propio prefijo de relaciones: openpyxl
    # no lo declara en la raiz sino en cada <sheet>, y repetirlo es
    # valido aunque la raiz ya lo tenga.
    prefijo_r = "r"
    prefijo_tag = re.search(r"<(\w+:)?sheets[\s>]", workbook).group(1) or ""

    elementos = {hoja["name"]: hoja["elemento"] for hoja in hojas_viejas}
    for hoja in agregadas:
        elementos[hoja["name"]] = (
            f'<{prefijo_tag}sheet xmlns:{prefijo_r}="{_REL_DOC}" '
            f'name="{escape(hoja["name"], {chr(34): "&quot;"})}" '
            f'sheetId="{hoja["sheetId"]}" {prefijo_r}:id="{hoja["rid"]}"/>'
        )

    actuales = [hoja["name"] for hoja in hojas_viejas] + [
        hoja["name"] for hoja in agregadas
    ]
    posicion = {nombre: i for i, nombre in enumerate(orden)}
    finales = sorted(
        actuales, key=lambda n: (posicion.get(n, len(posicion)), actuales.index(n))
    )

    bloque = "".join(elementos[nombre] for nombre in finales)
    workbook = re.sub(
        r"(<(?:\w+:)?sheets(?:\s[^>]*)?>).*?(</(?:\w+:)?sheets>)",
        lambda m: m.group(1) + bloque + m.group(2),
        workbook, count=1, flags=re.S,
    )

    # Los nombres definidos de una hoja (filtros, areas de impresion)
    # la nombran por POSICION: se acomodan al orden nuevo.
    viejas = [hoja["name"] for hoja in hojas_viejas]
    workbook = re.sub(
        r'localSheetId="(\d+)"',
        lambda m: (
            f'localSheetId="{finales.index(viejas[int(m.group(1))])}"'
            if int(m.group(1)) < len(viejas) else m.group(0)
        ),
        workbook,
    )

    # Se abre en la primera hoja (el Resumen).
    workbook = re.sub(r'\sactiveTab="\d+"', "", workbook)
    workbook = re.sub(r'\sfirstSheet="\d+"', "", workbook)

    return workbook


def _reordenar_en_el_lugar(ruta, orden):
    """Solo el orden de las hojas de un libro (reescribe workbook.xml)."""

    ruta = Path(ruta)
    temporal = ruta.with_name(f".~{ruta.name}")

    with zipfile.ZipFile(ruta) as viejo:
        workbook = viejo.read(_WORKBOOK).decode("utf-8")
        workbook = _workbook_final(workbook, _hojas(viejo), [], orden)

        with zipfile.ZipFile(
            temporal, "w", compression=zipfile.ZIP_DEFLATED
        ) as final:
            for info in viejo.infolist():
                if info.filename == _WORKBOOK:
                    final.writestr(_info(_WORKBOOK), workbook.encode("utf-8"))
                else:
                    _copiar_crudo(viejo, info, final)

    os.replace(temporal, ruta)
