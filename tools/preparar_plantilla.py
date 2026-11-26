"""
Convierte tu formato original (formato_Armada_de_fe_2026.docx) en una plantilla
con marcadores {{...}} lista para el generador.

Uso:
    python tools/preparar_plantilla.py [original.docx] [salida.docx]

Qué hace:
  1. Quita los duplicados "AlternateContent/Fallback" (los cuadros de texto
     aparecían dos veces en el XML).
  2. Pone marcadores en los cuadros de texto:
        {{REF}} {{ACTIVIDAD}} {{FECHA_EVENTO}} {{MODALIDAD}}
        {{PARTICIPANTES}} {{DESARROLLO}}
  3. Borra las fotos de ejemplo de la hoja 2 y los ~20 párrafos vacíos que se
     usaban como relleno.
  4. Deja un párrafo {{FOTOS}} (aquí se insertan las tablas de fotos) y un
     párrafo {{CIERRE}} (justo antes del texto "Agradecido..." y la firma).
  5. Marca la foto de la hoja 1 como "FOTO_PRINCIPAL" y el cuadro de datos como
     "CUADRO_DATOS" para que el generador los encuentre siempre.
"""
import copy
import io
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
BASE = Path(__file__).resolve().parent.parent


def _desenvolver_alternate_content(body):
    """Reemplaza <mc:AlternateContent> por el contenido de su <mc:Choice>."""
    for ac in list(body.iter(MC + "AlternateContent")):
        choice = ac.find(MC + "Choice")
        padre = ac.getparent()
        if choice is None or padre is None:
            continue
        idx = padre.index(ac)
        for hijo in list(choice):
            padre.insert(idx, hijo)
            idx += 1
        padre.remove(ac)


def _texto(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _runs(p):
    return [r for r in p.findall(qn("w:r"))]


def _set_run(r, texto):
    for t in r.findall(qn("w:t")):
        r.remove(t)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = texto
    r.append(t)


def _quitar_runs(p, desde):
    for r in _runs(p)[desde:]:
        p.remove(r)


def _poner_marcadores_cuadros(body):
    for tx in body.iter(qn("w:txbxContent")):
        for p in tx.findall(qn("w:p")):
            t = _texto(p)
            runs = _runs(p)
            if t.startswith("Ref.:"):
                _set_run(runs[2], "{{REF}}")
                _quitar_runs(p, 3)
            elif t.strip() == "PROYECTO MANÁ":
                _set_run(runs[0], "{{ACTIVIDAD}}")
                _quitar_runs(p, 1)
            elif t.startswith("Fecha de evento:"):
                _set_run(runs[2], "{{FECHA_EVENTO}}")
                _quitar_runs(p, 3)
            elif t.startswith("Modalidad de evento:"):
                _set_run(runs[2], "{{MODALIDAD}}")
                _quitar_runs(p, 3)
            elif t.startswith("Participantes:"):
                _set_run(runs[2], "{{PARTICIPANTES}}")
                _quitar_runs(p, 3)
            elif t.startswith("Desarrollo de la actividad"):
                _set_run(runs[4], "{{DESARROLLO}}")
                _quitar_runs(p, 5)


def _nombrar_anclas(body):
    """Pone nombre a la foto de la hoja 1 y al cuadro de datos."""
    for anchor in body.iter(qn("wp:anchor")):
        ext = anchor.find(qn("wp:extent"))
        docpr = anchor.find(qn("wp:docPr"))
        if ext is None or docpr is None:
            continue
        cx = int(ext.get("cx"))
        if list(anchor.iter(qn("pic:pic"))) and cx > 4_000_000 and cx < 5_000_000:
            docpr.set("name", "FOTO_PRINCIPAL")
        elif list(anchor.iter(qn("w:txbxContent"))):
            if "{{DESARROLLO}}" in _texto(anchor):
                docpr.set("name", "CUADRO_DATOS")


def _parrafo_marcador(texto):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "20")
    sp.set(qn("w:lineRule"), "exact")
    ppr.append(sp)
    p.append(ppr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = texto
    r.append(t)
    p.append(r)
    return p


def _foto_gris():
    """Imagen de relleno neutra para la foto de ejemplo (la real la pone el generador)."""
    buf = io.BytesIO()
    Image.new("RGB", (800, 600), (200, 200, 200)).save(buf, "JPEG", quality=60)
    buf.seek(0)
    return buf


def preparar(origen: Path, destino: Path):
    doc = Document(str(origen))
    body = doc.element.body

    _desenvolver_alternate_content(body)
    _poner_marcadores_cuadros(body)
    _nombrar_anclas(body)

    hijos = list(body)
    # Estructura del formato original:
    #  0 cabecera | 1 datos + foto hoja 1 + salto | 2 fotos hoja 2 | 3-23 vacíos
    #  24 "Agradecido..." | 25-31 firma | sectPr
    p_fotos = hijos[2]
    p_cierre = None
    for el in hijos:
        if el.tag == qn("w:p") and "Agradecido" in _texto(el):
            p_cierre = el
            break
    if p_cierre is None:
        raise RuntimeError("No encontré el párrafo 'Agradecido...' en el formato.")

    # Borrar todo lo que está entre el párrafo de fotos y el de cierre
    idx_f = hijos.index(p_fotos)
    idx_c = hijos.index(p_cierre)
    for el in hijos[idx_f:idx_c]:
        body.remove(el)

    p_fotos_nuevo = _parrafo_marcador("{{FOTOS}}")
    p_cierre_nuevo = _parrafo_marcador("{{CIERRE}}")
    p_cierre.addprevious(p_fotos_nuevo)
    p_cierre.addprevious(p_cierre_nuevo)

    # Foto de ejemplo -> imagen gris neutra (no dejamos fotos personales en la plantilla)
    for anchor in body.iter(qn("wp:anchor")):
        docpr = anchor.find(qn("wp:docPr"))
        if docpr is not None and docpr.get("name") == "FOTO_PRINCIPAL":
            blip = anchor.find(".//" + qn("a:blip"))
            viejo = blip.get(qn("r:embed"))
            nuevo, _ = doc.part.get_or_add_image(_foto_gris())
            blip.set(qn("r:embed"), nuevo)
            doc.part.drop_rel(viejo)

    doc.save(str(destino))
    print(f"Plantilla creada: {destino}")


if __name__ == "__main__":
    origen = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "plantilla" / "formato_original.docx"
    destino = Path(sys.argv[2]) if len(sys.argv) > 2 else BASE / "plantilla" / "plantilla.docx"
    preparar(origen, destino)
