"""
generador.py  ·  Motor del generador de informes "Armada de Fe"

Rellena plantilla/plantilla.docx con texto + fotos y (opcionalmente) lo exporta a PDF.
No depende de Flask ni de MCP: lo usan app.py (web) y server.py (MCP).

Reglas de diseño que implementa
-------------------------------
 1. Hoja 1: texto del informe + EXACTAMENTE 1 foto (la primera). La foto se escala
    sin deformarse para ocupar todo el alto que deja libre el texto (máx. 22.5 cm de
    alto y 16.5 cm de ancho; en esta plantilla el ancho real lo limita la columna
    lateral de contactos).
 2. Hoja 2 en adelante, con las fotos restantes (r):
        r = 1 -> 1 foto grande           r = 2 -> 2 lado a lado (50% / 50%)
        r = 3 -> 2 arriba + 1 abajo      r = 4 -> cuadrícula 2x2
        r > 4 -> bloques de 4; si el resto es 1 se rebalancea (4,1 -> 3,2)
        Ej.: 7 fotos en total -> [1, 4, 2]
 3. Cada bloque aprovecha el alto de la hoja (sin dejar huecos grandes): 2 filas de
    ~11.35 cm, 2 fotos lado a lado de 14 cm, 1 foto sola hasta 22.5 cm. En la última
    hoja la altura se reduce lo justo para que también quepan el texto de cierre y la
    firma sin crear una hoja extra.
 4. Todas las fotos llevan el MISMO contorno suavizado (softEdge) que la foto de la hoja 1.
 5. Las fotos de una cuadrícula llenan su celda (modo "cover", sin huecos blancos);
    una foto que ocupa toda la fila se muestra entera (modo "contain").
 6. Tablas con márgenes de celda = 0 (nivel tabla y nivel celda).
 7. {{FOTOS}} se sustituye por las tablas; {{CIERRE}} marca dónde empieza el cierre.
 8. Entre bloques hay un salto de página real (WD_BREAK.PAGE).
"""
from __future__ import annotations

import copy
import io
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from PIL import Image as PILImage
from PIL import ImageOps

try:  # fotos HEIC de iPhone (opcional)
    import pillow_heif

    pillow_heif.register_heif_opener()
except Exception:  # pragma: no cover
    pass

# ----------------------------- Configuración ------------------------------
BASE = Path(__file__).resolve().parent
PLANTILLA = BASE / "plantilla" / "plantilla.docx"
ENTRADA = BASE / "entrada"
SALIDAS = BASE / "salidas"
EXTENSIONES = {".jpg", ".jpeg", ".jfif", ".png", ".bmp",
               ".gif", ".webp", ".heic", ".heif", ".tif", ".tiff"}

# --- Medidas (cm) ---
ANCHO_UTIL_CM = 16.5        # Ancho máximo de foto / tabla
ALTO_MAX_CM = 22.5          # Alto máximo de una foto (hoja 1 o foto sola)
# Alto por fila cuando hay 2 filas (2 x 11.35 = 22.7 cm: llena la hoja)
ALTO_MAX_GRUPO_CM = 11.35
ALTO_MAX_2FOTOS_CM = 14.0   # Alto cuando son 2 fotos lado a lado (1 sola fila)
MARGEN_CM = 0.2             # Separación visual entre fotos
GAP_TEXTO_FOTO_CM = 0.25    # Aire entre el texto y la foto de la hoja 1
ALTO_MIN_FOTO1_CM = 4.0     # Si el texto es larguísimo, la foto no baja de esto

# --- Contorno: igual que la foto de la hoja 1 de la plantilla (a:softEdge rad=112500) ---
SOFT_EDGE_EMU = 112500

# --- Geometría de la hoja (Carta, márgenes 1.27 cm, marco decorativo) ---
CENTRO_MARCO_CM = 11.55     # Centro horizontal del marco naranja
MARGEN_IZQ_CM = 1.27
# Alto aprovechable en hojas 2+ (debajo del colchón superior)
PAG_ALTO_UTIL_CM = 24.0
COLCHON_SUPERIOR_CM = 0.5   # Aire bajo la cinta superior del marco
CIERRE_ALTO_CM = 7.6        # Alto del bloque "Agradecido..." + firma
# Reserva para el párrafo del salto de página (evita hojas en blanco)
MARGEN_SALTO_CM = 1.3
ALTO_MIN_FILA_CM = 6.5      # Menos que esto no se reduce: el cierre pasa a otra hoja

# --- Estimación del alto del cuadro de texto de la hoja 1 (Times 12 pt) ---
# Medidas reales (LibreOffice/Word): línea 14.9 pt, 8 pt de espacio tras cada párrafo
# caracteres por línea (12 cm de ancho útil, conservador)
CHARS_POR_LINEA = 70
LINEA_CM = 0.526
ESPACIO_PARRAFO_CM = 0.282
CUADRO_BASE_CM = 0.40       # márgenes internos del cuadro + holgura

# "cover"   = llena la celda recortando un poco (sin huecos blancos)  <- por defecto
# "contain" = nunca recorta, pero puede dejar espacio en blanco
MODO_AJUSTE_GRUPO = "cover"
# punto de la foto que se conserva al recortar (x, y); 0.4 favorece las caras
CENTRADO_COVER = (0.5, 0.4)

CAMPOS_TEXTO = ["REF", "ACTIVIDAD", "FECHA_EVENTO",
                "MODALIDAD", "PARTICIPANTES", "DESARROLLO"]
EMU_CM = 360000


@dataclass
class Foto:
    ruta: str
    descripcion: str = ""


@dataclass
class Resultado:
    docx: Path
    plan: list = field(default_factory=list)
    avisos: list = field(default_factory=list)


def _cm_emu(cm: float) -> int:
    return int(round(cm * EMU_CM))


def _cm_twips(cm: float) -> int:
    return int(round(cm * 567))


# ------------------------------- Imágenes ---------------------------------
def _resolver_imagen(ruta: str) -> Path:
    p = Path(ruta)
    if not p.is_absolute():
        p = ENTRADA / p
    if not p.exists():
        raise FileNotFoundError(f"No existe la imagen: {p}")
    if p.suffix.lower() not in EXTENSIONES:
        raise ValueError(f"Formato no soportado: {p.suffix} ({p.name})")
    return p


def _abrir_imagen(p: Path) -> PILImage.Image:
    """Abre, corrige rotación EXIF y aplana transparencias sobre blanco."""
    im = PILImage.open(p)
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        fondo = PILImage.new("RGB", im.size, (255, 255, 255))
        fondo.paste(im, mask=im.split()[-1])
        return fondo
    return im.convert("RGB")


def _preparar_imagen(p: Path, caja_w_cm: float, caja_h_cm: float, modo: str = "contain"):
    """
    Devuelve (BytesIO jpeg, ancho_cm, alto_cm) ya ajustada a la caja.
    contain: respeta proporción, cabe entera.  cover: llena la caja recortando.
    """
    im = _abrir_imagen(p)
    w, h = im.size
    if modo == "cover":
        ancho_cm, alto_cm = caja_w_cm, caja_h_cm
        px = (max(1, int(ancho_cm / 2.54 * 200)),
              max(1, int(alto_cm / 2.54 * 200)))
        im = ImageOps.fit(im, px, method=PILImage.LANCZOS,
                          centering=CENTRADO_COVER)
    else:
        escala = min(caja_w_cm / w, caja_h_cm / h)
        ancho_cm, alto_cm = w * escala, h * escala
        px = (max(1, int(ancho_cm / 2.54 * 200)),
              max(1, int(alto_cm / 2.54 * 200)))
        if px[0] < w:  # nunca ampliar, solo reducir
            im = im.resize(px, PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=88, optimize=True)
    buf.seek(0)
    return buf, ancho_cm, alto_cm


# --------------------------- Plan de distribución -------------------------
def _planificar_paginas(n: int) -> list[int]:
    """
    Fotos por hoja. La hoja 1 lleva 1 foto; el resto va en bloques de 4 y el sobrante
    se reparte para no dejar una foto sola:
        n=1 -> [1]            n=2 -> [1,1]        n=3 -> [1,2]       n=4 -> [1,3]
        n=5 -> [1,4]          n=6 -> [1,3,2]      n=7 -> [1,4,2]     n=8 -> [1,4,3]
        n=9 -> [1,4,4]        n=10 -> [1,4,3,2]   n=11 -> [1,4,4,2]
    Cada hoja ≥ 2 contiene como máximo 4 fotos (2x2).
    """
    if n <= 0:
        return []
    plan = [1]
    resto = n - 1
    if resto == 0:
        return plan
    while resto > 0:
        if resto == 5:          # 4+1 dejaría una foto huérfana -> 3+2
            plan += [3, 2]
            resto = 0
        elif resto >= 4:
            plan.append(4)
            resto -= 4
        else:                   # 1, 2 o 3
            plan.append(resto)
            resto = 0
    return plan


# Filas de cada bloque (índices de foto dentro del bloque)
_FILAS = {1: [[0]], 2: [[0, 1]], 3: [[0, 1], [2]], 4: [[0, 1], [2, 3]]}


def _alto_bloque(n_fotos: int, libre_cm: float) -> float:
    """Alto por fila de un bloque, aprovechando al máximo el espacio disponible."""
    filas = len(_FILAS[n_fotos])
    if n_fotos == 1:
        return min(ALTO_MAX_CM, libre_cm)
    if filas == 1:                                   # 2 fotos lado a lado
        return min(ALTO_MAX_2FOTOS_CM, libre_cm)
    return min(ALTO_MAX_GRUPO_CM, libre_cm / filas)  # 2 filas


def _calcular_alturas(plan_resto: list[int]):
    """
    Devuelve (alto_por_fila_cm: list[float] por bloque, cierre_en_hoja_nueva: bool).
    Hojas intermedias: ocupan la hoja. Última hoja: se reduce para que quepa el cierre.
    """
    if not plan_resto:
        return [], False
    libre_hoja = PAG_ALTO_UTIL_CM - MARGEN_SALTO_CM
    alturas = [_alto_bloque(n, libre_hoja) for n in plan_resto]

    libre_cierre = PAG_ALTO_UTIL_CM - CIERRE_ALTO_CM - 0.3
    h = _alto_bloque(plan_resto[-1], libre_cierre)
    if h >= ALTO_MIN_FILA_CM:
        alturas[-1] = h
        return alturas, False
    return alturas, True        # no cabe ni reduciendo: el cierre va a su propia hoja


# ------------------------------ XML / tablas -------------------------------
_ORDEN_TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
                "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
                "shd", "tblLayout", "tblCellMar", "tblLook"]
_ORDEN_TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
               "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]


def _insertar_ordenado(padre, hijo, orden):
    """Inserta `hijo` respetando el orden del esquema OOXML."""
    nombre = hijo.tag.split("}")[1]
    for existente in padre.findall(hijo.tag):
        padre.remove(existente)
    posterior = orden[orden.index(nombre) + 1:]
    for el in padre:
        if el.tag.split("}")[1] in posterior:
            el.addprevious(hijo)
            return hijo
    padre.append(hijo)
    return hijo


def _el(nombre, **attrs):
    e = OxmlElement(nombre)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def _margenes_cero(padre_pr, orden, nombre="w:tcMar"):
    mar = _el(nombre)
    for lado in ("top", "left", "bottom", "right"):
        mar.append(_el("w:" + lado, w=0, type="dxa"))
    _insertar_ordenado(padre_pr, mar, orden)


def _configurar_tabla(tabla, anchos_cm: list[float], ind_cm: float):
    tbl = tabla._tbl
    tblPr = tbl.tblPr
    for hijo in list(tblPr):
        tblPr.remove(hijo)
    total = sum(anchos_cm)
    _insertar_ordenado(tblPr, _el(
        "w:tblW", w=_cm_twips(total), type="dxa"), _ORDEN_TBLPR)
    _insertar_ordenado(tblPr, _el(
        "w:tblInd", w=_cm_twips(ind_cm), type="dxa"), _ORDEN_TBLPR)
    bordes = _el("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        bordes.append(_el("w:" + lado, val="nil"))
    _insertar_ordenado(tblPr, bordes, _ORDEN_TBLPR)
    _insertar_ordenado(tblPr, _el("w:tblLayout", type="fixed"), _ORDEN_TBLPR)
    # márgenes 0 a nivel tabla
    _margenes_cero(tblPr, _ORDEN_TBLPR, "w:tblCellMar")
    for gc, ancho in zip(tbl.tblGrid.findall(qn("w:gridCol")), anchos_cm):
        gc.set(qn("w:w"), str(_cm_twips(ancho)))


def _configurar_celda(celda, ancho_cm: float, vcentro=True):
    tcPr = celda._tc.get_or_add_tcPr()
    _insertar_ordenado(tcPr, _el("w:tcW", w=_cm_twips(
        ancho_cm), type="dxa"), _ORDEN_TCPR)
    # márgenes 0 a nivel celda
    _margenes_cero(tcPr, _ORDEN_TCPR)
    if vcentro:
        _insertar_ordenado(tcPr, _el("w:vAlign", val="center"), _ORDEN_TCPR)


def _fijar_alto_fila(fila, alto_cm: float):
    trPr = fila._tr.get_or_add_trPr()
    for hijo in list(trPr):
        trPr.remove(hijo)
    trPr.append(_el("w:cantSplit"))
    trPr.append(_el("w:trHeight", val=_cm_twips(alto_cm), hRule="exact"))


def _formato_parrafo_foto(par):
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = par.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pf.keep_together = True


def _contorno_suave(run):
    """
    Da a la foto el mismo contorno que la foto de la hoja 1 de la plantilla:
    sin línea (<a:ln><a:noFill/></a:ln>) y borde difuminado (<a:softEdge rad=.../>).
    """
    sppr = run._r.find(".//" + qn("pic:spPr"))
    if sppr is None:
        return
    for viejo in sppr.findall(qn("a:ln")) + sppr.findall(qn("a:effectLst")):
        sppr.remove(viejo)
    ln = OxmlElement("a:ln")
    ln.append(OxmlElement("a:noFill"))
    sppr.append(ln)
    efectos = OxmlElement("a:effectLst")
    soft = OxmlElement("a:softEdge")
    soft.set("rad", str(SOFT_EDGE_EMU))
    efectos.append(soft)
    sppr.append(efectos)


def _insertar_foto_en_celda(celda, foto: Foto, ancho_celda_cm: float, alto_caja_cm: float,
                            modo: str = MODO_AJUSTE_GRUPO):
    ruta = _resolver_imagen(foto.ruta)
    par = celda.paragraphs[0]
    _formato_parrafo_foto(par)
    reserva = 0.7 if foto.descripcion else 0.0
    buf, w, h = _preparar_imagen(ruta, ancho_celda_cm - MARGEN_CM,
                                 alto_caja_cm - reserva - MARGEN_CM, modo)
    run = par.add_run()
    run.add_picture(buf, width=Cm(w), height=Cm(h))
    _contorno_suave(run)
    if foto.descripcion:
        pie = celda.add_paragraph()
        _formato_parrafo_foto(pie)
        r = pie.add_run(foto.descripcion)
        r.italic = True
        r.font.size = Pt(9)


def _tabla_bloque(doc, bloque: list[Foto], alto_foto_cm: float):
    """Crea la tabla de un bloque (1-4 fotos) según _FILAS. Devuelve el elemento <w:tbl>."""
    filas = _FILAS[len(bloque)]
    ancho = ANCHO_UTIL_CM
    ind = CENTRO_MARCO_CM - ancho / 2 - MARGEN_IZQ_CM
    n_cols = max(len(f) for f in filas)
    ancho_col = ancho / n_cols if n_cols > 1 else ancho
    anchos = [ancho_col] * n_cols

    tabla = doc.add_table(rows=len(filas), cols=n_cols)
    tabla.autofit = False
    _configurar_tabla(tabla, anchos, ind)

    for i_fila, indices in enumerate(filas):
        fila = tabla.rows[i_fila]
        _fijar_alto_fila(fila, alto_foto_cm)
        celdas = list(fila.cells)
        fila_completa = len(indices) == n_cols
        if not fila_completa:                          # fila con 1 foto en tabla de 2 columnas
            celdas = [celdas[0].merge(celdas[-1])]
            ancho_celdas = [ancho]
        else:
            ancho_celdas = anchos
        # Una foto que ocupa toda la fila se muestra entera; en cuadrícula llena su celda.
        modo = MODO_AJUSTE_GRUPO if (
            fila_completa and n_cols > 1) else "contain"
        for celda, idx, anc in zip(celdas, indices, ancho_celdas):
            _configurar_celda(celda, anc)
            _insertar_foto_en_celda(
                celda, bloque[idx], anc, alto_foto_cm, modo)
    return tabla._tbl


def _parrafo_salto_pagina(doc):
    """Párrafo con salto de página real (WD_BREAK.PAGE); su resto hace de colchón superior."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Cm(COLCHON_SUPERIOR_CM)
    p.add_run().add_break(WD_BREAK.PAGE)
    return p._p


# ------------------------------ Marcadores ---------------------------------
def _todos_los_parrafos(doc):
    """Todos los <w:p> del cuerpo, incluidos los de cuadros de texto y tablas."""
    return list(doc.element.body.iter(qn("w:p")))


def _texto_parrafo(p_el) -> str:
    """Texto propio del párrafo (sin el de cuadros de texto que lleve anclados dentro)."""
    return "".join(t.text or "" for r in p_el.findall(qn("w:r")) for t in r.findall(qn("w:t")))


def _buscar_marcador(doc, nombre: str):
    for p in _todos_los_parrafos(doc):
        if _texto_parrafo(p).strip() == "{{" + nombre + "}}":
            return p
    return None


def _reemplazar_marcador(par: Paragraph, marcador: str, valor: str) -> bool:
    """Reemplaza `marcador` conservando el formato aunque Word lo haya partido en varios runs."""
    runs = par.runs
    textos = [r.text for r in runs]
    total = "".join(textos)
    pos = total.find(marcador)
    if pos < 0:
        return False
    fin = pos + len(marcador)
    acum, primero = 0, True
    for r, t in zip(runs, textos):
        ini, fi = acum, acum + len(t)
        acum = fi
        if fi <= pos or ini >= fin:
            continue
        a, b = max(pos, ini) - ini, min(fin, fi) - ini
        if primero:
            r.text = t[:a] + valor + t[b:]
            primero = False
        else:
            r.text = t[:a] + t[b:]
    return True


def _reemplazar_todos(doc, marcador: str, valor: str) -> int:
    n = 0
    for p_el in _todos_los_parrafos(doc):
        par = Paragraph(p_el, None)
        while _reemplazar_marcador(par, marcador, valor):
            n += 1
    return n


def _partes(texto: str) -> list[str]:
    return [l.strip() for l in re.split(r"\n+", (texto or "").strip()) if l.strip()]


def _expandir_parrafos(doc, nombre: str, texto: str) -> bool:
    """Una línea de `texto` por párrafo. El primero reemplaza al marcador; los demás clonan el párrafo."""
    marcador = "{{" + nombre + "}}"
    destino = None
    for p_el in _todos_los_parrafos(doc):
        if marcador in _texto_parrafo(p_el):
            destino = p_el
            break
    if destino is None:
        return False
    partes = _partes(texto) or [""]
    molde = copy.deepcopy(destino)
    # el clon conserva solo el run del valor
    for r in molde.findall(qn("w:r")):
        if marcador not in "".join(t.text or "" for t in r.iter(qn("w:t"))):
            molde.remove(r)
    _reemplazar_marcador(Paragraph(destino, None), marcador, partes[0])
    anterior = destino
    for extra in partes[1:]:
        nuevo = copy.deepcopy(molde)
        anterior.addnext(nuevo)
        _reemplazar_marcador(Paragraph(nuevo, None), marcador, extra)
        anterior = nuevo
    return True


def _aplicar_campos(doc, campos: dict, informacion: str = ""):
    campos = {k: (v or "") for k, v in campos.items()}
    if "ACTIVIDAD" in campos:
        campos["ACTIVIDAD"] = campos["ACTIVIDAD"].strip().upper()
    _expandir_parrafos(doc, "DESARROLLO", campos.pop("DESARROLLO", ""))
    if informacion.strip():
        _expandir_parrafos(doc, "INFORMACION", informacion)
    for k, v in campos.items():
        _reemplazar_todos(doc, "{{" + k + "}}",
                          re.sub(r"\s*\n+\s*", ", ", v.strip()))
    # Marcadores sin dato -> vacío (menos {{FOTOS}} y {{CIERRE}} que se tratan aparte)
    for p_el in _todos_los_parrafos(doc):
        for m in set(re.findall(r"\{\{([A-Za-z0-9_ÁÉÍÓÚÑáéíóúñ]+)\}\}", _texto_parrafo(p_el))):
            if m not in ("FOTOS", "CIERRE"):
                _reemplazar_todos(doc, "{{" + m + "}}", "")


# -------------------------- Hoja 1: texto + foto ----------------------------
def _estimar_alto_texto_cm(campos: dict) -> float:
    """Estima el alto (cm) que ocupa el texto del cuadro de datos (conservador)."""
    def lineas(txt):
        return max(1, math.ceil(len(txt) / CHARS_POR_LINEA))

    bloques = [
        1,                                                              # título
        lineas("Fecha de evento: " + (campos.get("FECHA_EVENTO") or "")),
        lineas("Modalidad de evento: " + (campos.get("MODALIDAD") or "")),
        lineas("Participantes: " + re.sub(r"\s*\n+\s*",
               ", ", campos.get("PARTICIPANTES") or "")),
    ]
    desarrollo = _partes(campos.get("DESARROLLO") or "") or [""]
    bloques.append(lineas("Desarrollo de la actividad:  " + desarrollo[0]))
    bloques += [lineas(x) for x in desarrollo[1:]]
    # "Evidencias:"
    bloques.append(1)
    n_lineas = sum(bloques)
    return CUADRO_BASE_CM + n_lineas * LINEA_CM + (len(bloques) - 1) * ESPACIO_PARRAFO_CM


def _buscar_ancla(doc, nombre: str):
    for a in doc.element.body.iter(qn("wp:anchor")):
        dp = a.find(qn("wp:docPr"))
        if dp is not None and dp.get("name") == nombre:
            return a
    return None


def _pos_offset(ancla, eje: str) -> int:
    nodo = ancla.find(qn("wp:position" + eje))
    return int(nodo.find(qn("wp:posOffset")).text)


def _set_pos(ancla, eje: str, emu: int):
    nodo = ancla.find(qn("wp:position" + eje))
    for hijo in list(nodo):
        nodo.remove(hijo)
    off = OxmlElement("wp:posOffset")
    off.text = str(int(emu))
    nodo.append(off)


def _set_tamano_cuadro(ancla, cy_emu: int):
    ancla.find(qn("wp:extent")).set("cy", str(int(cy_emu)))
    for ext in ancla.iter(qn("a:ext")):
        if ext.get("cy") is not None and ext.get("cx") is not None:
            ext.set("cy", str(int(cy_emu)))


def _colocar_foto_principal(doc, foto: Foto | None, campos: dict) -> list[str]:
    avisos = []
    ancla_foto = _buscar_ancla(doc, "FOTO_PRINCIPAL")
    cuadro = _buscar_ancla(doc, "CUADRO_DATOS")
    if ancla_foto is None:
        return ["La plantilla no tiene el ancla FOTO_PRINCIPAL (ejecuta tools/preparar_plantilla.py)."]
    if foto is None:                                   # sin fotos: se elimina la imagen de ejemplo
        run = ancla_foto.getparent().getparent()
        run.getparent().remove(run)
        return avisos

    ext = ancla_foto.find(qn("wp:extent"))
    x0, y0 = _pos_offset(ancla_foto, "H"), _pos_offset(ancla_foto, "V")
    caja_w, caja_h0 = int(ext.get("cx")), int(ext.get("cy"))
    # hasta dónde llega la zona de foto en la plantilla
    limite_inferior = y0 + caja_h0

    alto_texto = _cm_emu(_estimar_alto_texto_cm(campos))
    if cuadro is not None:
        _set_tamano_cuadro(cuadro, alto_texto)
        arriba = _pos_offset(cuadro, "V") + alto_texto + \
            _cm_emu(GAP_TEXTO_FOTO_CM)
    else:
        arriba = y0
    caja_h = min(_cm_emu(ALTO_MAX_CM), limite_inferior - arriba)
    if caja_h < _cm_emu(ALTO_MIN_FOTO1_CM):
        avisos.append(
            "El texto es muy largo: la foto de la hoja 1 quedó pequeña. Resume el desarrollo.")
        caja_h = _cm_emu(ALTO_MIN_FOTO1_CM)
    caja_w = min(caja_w, _cm_emu(ANCHO_UTIL_CM))

    buf, w_cm, h_cm = _preparar_imagen(_resolver_imagen(
        foto.ruta), caja_w / EMU_CM, caja_h / EMU_CM, "contain")
    cx, cy = _cm_emu(w_cm), _cm_emu(h_cm)

    rid, _ = doc.part.get_or_add_image(buf)
    blip = ancla_foto.find(".//" + qn("a:blip"))
    viejo = blip.get(qn("r:embed"))
    blip.set(qn("r:embed"), rid)
    if viejo != rid:
        doc.part.drop_rel(viejo)
    for rect in list(ancla_foto.iter(qn("a:srcRect"))):
        rect.getparent().remove(rect)

    ext.set("cx", str(cx))
    ext.set("cy", str(cy))
    for e in ancla_foto.find(".//" + qn("pic:spPr")).iter(qn("a:ext")):
        e.set("cx", str(cx))
        e.set("cy", str(cy))
    # Centrada dentro de su caja (horizontal y verticalmente)
    _set_pos(ancla_foto, "H", x0 + (caja_w - cx) // 2)
    _set_pos(ancla_foto, "V", arriba + (caja_h - cy) // 2)
    return avisos


# ----------------------------- Bloques de fotos ------------------------------
def _fijar_parrafo_exacto(p_el, alto_cm: float):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), str(max(20, _cm_twips(alto_cm))))
    sp.set(qn("w:lineRule"), "exact")


def _vaciar_parrafo(p_el):
    for r in p_el.findall(qn("w:r")):
        p_el.remove(r)


def _insertar_bloques(doc, fotos_resto: list[Foto], plan_resto: list[int]) -> list[str]:
    avisos = []
    marcador_cierre = _buscar_marcador(doc, "CIERRE")
    marcador_fotos = _buscar_marcador(doc, "FOTOS")
    if marcador_fotos is None:
        avisos.append(
            "No encontré {{FOTOS}} en la plantilla; las fotos se agregaron antes del cierre.")
        marcador_fotos = _parrafo_salto_pagina(doc)
        (marcador_cierre if marcador_cierre is not None else doc.element.body[-1]).addprevious(
            marcador_fotos)
        _vaciar_parrafo(marcador_fotos)

    # {{FOTOS}} se borra y queda como colchón superior de la hoja 2
    _vaciar_parrafo(marcador_fotos)
    _fijar_parrafo_exacto(marcador_fotos, COLCHON_SUPERIOR_CM)

    alturas, cierre_nueva = _calcular_alturas(plan_resto)
    cursor = marcador_fotos
    idx = 0
    for i, cantidad in enumerate(plan_resto):
        bloque = fotos_resto[idx:idx + cantidad]
        idx += cantidad
        tbl = _tabla_bloque(doc, bloque, alturas[i])
        cursor.addnext(tbl)
        cursor = tbl
        if i < len(plan_resto) - 1:                    # salto de página entre bloques
            salto = _parrafo_salto_pagina(doc)
            cursor.addnext(salto)
            cursor = salto

    if marcador_cierre is not None:
        _vaciar_parrafo(marcador_cierre)
        if plan_resto:
            # Un párrafo mínimo separa la tabla del cierre (Word exige un párrafo tras una tabla)
            _fijar_parrafo_exacto(
                marcador_cierre, COLCHON_SUPERIOR_CM if cierre_nueva else 0.04)
            if cierre_nueva:
                pPr = marcador_cierre.find(qn("w:pPr"))
                pPr.insert(0, OxmlElement("w:pageBreakBefore"))
        else:
            _fijar_parrafo_exacto(marcador_cierre, 0.04)
    return avisos


# ------------------------------ API pública --------------------------------
def revisar_plantilla() -> str:
    if not PLANTILLA.exists():
        return f"No encuentro la plantilla en: {PLANTILLA}"
    doc = Document(str(PLANTILLA))
    encontrados = set()
    for p in _todos_los_parrafos(doc):
        encontrados.update(re.findall(
            r"\{\{\s*([A-Za-z0-9_ÁÉÍÓÚÑáéíóúñ]+)\s*\}\}", _texto_parrafo(p)))
    if not encontrados:
        return "La plantilla no tiene marcadores {{...}}."
    return "Marcadores encontrados: " + ", ".join("{{" + m + "}}" for m in sorted(encontrados))


def _ruta_salida(nombre: str, ext: str) -> Path:
    nombre = re.sub(r'[\\/:*?"<>|]', "_", (nombre or "").strip()) or "informe"
    return SALIDAS / f"{Path(nombre).stem}{ext}"


def generar_informe(campos: dict | None, fotos: list[Foto], nombre_salida: str = "informe",
                    informacion: str = "") -> Resultado:
    """Crea salidas/<nombre_salida>.docx. Lanza excepción solo si no hay plantilla o falla una foto."""
    if not PLANTILLA.exists():
        raise FileNotFoundError(f"No encuentro la plantilla en: {PLANTILLA}")
    SALIDAS.mkdir(exist_ok=True)
    campos = dict(campos or {})
    fotos = list(fotos or [])
    for f in fotos:                                    # validar antes de tocar el documento
        _resolver_imagen(f.ruta)

    doc = Document(str(PLANTILLA))
    avisos: list[str] = []

    plan = _planificar_paginas(len(fotos))
    avisos += _colocar_foto_principal(doc, fotos[0] if fotos else None, campos)
    _aplicar_campos(doc, campos, informacion)
    avisos += _insertar_bloques(doc, fotos[1:], plan[1:])

    destino = _ruta_salida(nombre_salida, ".docx")
    doc.save(str(destino))
    return Resultado(docx=destino, plan=plan, avisos=avisos)


# ---------------------------------- PDF -------------------------------------
def _buscar_soffice() -> str | None:
    for nombre in ("soffice", "libreoffice"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(ruta).exists():
            return ruta
    return None


def _pdf_con_word(origen: Path, destino: Path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(origen), ReadOnly=True)
        doc.SaveAs(str(destino), FileFormat=17)     # 17 = PDF
        doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def _pdf_con_libreoffice(origen: Path, destino: Path, soffice: str):
    with tempfile.TemporaryDirectory() as perfil:
        subprocess.run(
            [soffice, f"-env:UserInstallation={Path(perfil).as_uri()}", "--headless",
             "--convert-to", "pdf", "--outdir", str(destino.parent), str(origen)],
            check=True, capture_output=True, timeout=240)


def exportar_pdf(docx: Path) -> tuple[Path | None, str]:
    """Devuelve (ruta_pdf | None, mensaje). Intenta Word (Windows) y luego LibreOffice."""
    docx = Path(docx)
    destino = docx.with_suffix(".pdf")
    errores = []
    if sys.platform == "win32":
        try:
            _pdf_con_word(docx, destino)
            if destino.exists():
                return destino, "PDF creado con Microsoft Word."
        except Exception as e:  # Word no instalado, etc.
            errores.append(f"Word: {e}")
    soffice = _buscar_soffice()
    if soffice:
        try:
            _pdf_con_libreoffice(docx, destino, soffice)
            if destino.exists():
                return destino, "PDF creado con LibreOffice."
        except Exception as e:
            errores.append(f"LibreOffice: {e}")
    detalle = " | ".join(
        errores) if errores else "No hay Word (Windows) ni LibreOffice instalados."
    return None, f"No se pudo crear el PDF. {detalle}"


def abrir_archivo(ruta: Path) -> str:
    if sys.platform == "win32":
        os.startfile(str(ruta))  # type: ignore[attr-defined]
        return f"Abierto: {ruta}"
    return f"Archivo listo en {ruta} (ábrelo manualmente)."
