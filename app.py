"""
app.py  ·  Interfaz web del generador de informes "Armada de Fe"

    python app.py        ->  abre http://127.0.0.1:5000

Endpoints
  GET  /                 interfaz (static/index.html)
  POST /api/redactar     notas sueltas  -> campos del informe (IA o reglas locales)
  POST /api/generar      campos + fotos -> .docx (+ .pdf si hay Word/LibreOffice)
  GET  /salidas/<f>      descarga / vista previa
"""
import json
import os
import re
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

import generador as g

BASE = Path(__file__).resolve().parent


def _cargar_env():
    """Lee .env (ANTHROPIC_API_KEY=...) sin depender de python-dotenv."""
    env = BASE / ".env"
    if env.exists():
        for linea in env.read_text(encoding="utf-8").splitlines():
            linea = linea.strip()
            if linea and not linea.startswith("#") and "=" in linea:
                k, v = linea.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_cargar_env()

app = Flask(__name__, static_folder=str(BASE / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024   # 300 MB por informe

MODALIDADES = {"Presencial", "Virtual", "Mixta"}


# ----------------------------- Redacción con IA -----------------------------
PROMPT = """Eres el secretario de un Club de Conquistadores (Iglesia Adventista) en Bolivia.
Convierte las notas sueltas del director en datos para un informe formal de actividad.
Devuelve SOLO un objeto JSON (sin markdown) con estas claves de texto:
REF: "INFORME DE ACTIVIDAD: <ACTIVIDAD EN MAYÚSCULAS>"
ACTIVIDAD: nombre corto de la actividad, en MAYÚSCULAS
FECHA_EVENTO: fecha tal como la dieron (ej. "Sábado 6 de marzo"); "" si no la dieron
MODALIDAD: exactamente "Presencial", "Virtual" o "Mixta"; "" si no se sabe
PARTICIPANTES: quiénes y cuántos participaron; "" si no se sabe
DESARROLLO: 1 a 3 párrafos en tercera persona, tono formal, buena ortografía, separados por \\n.
No inventes datos, nombres ni cifras que no estén en las notas."""


def _redactar_con_ia(notas: str) -> dict:
    import anthropic

    cliente = anthropic.Anthropic()                    # usa ANTHROPIC_API_KEY
    modelo = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5-5")
    r = cliente.messages.create(
        model=modelo, max_tokens=1500, system=PROMPT,
        messages=[{"role": "user", "content": notas}],
    )
    texto = "".join(b.text for b in r.content if getattr(b, "type", "") == "text")
    texto = re.sub(r"^```(?:json)?|```$", "", texto.strip(), flags=re.M).strip()
    return json.loads(texto)


def _oraciones(texto: str):
    partes = re.split(r"(?<=[.!?])\s+|\n+", texto.strip())
    for p in partes:
        p = p.strip(" -•\t")
        if p:
            p = p[0].upper() + p[1:]
            yield p if p[-1] in ".!?" else p + "."


def _redactar_local(notas: str) -> dict:
    """Plan B sin internet ni API: limpia el texto y extrae lo evidente."""
    t = " ".join(notas.split())
    m_fecha = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", t) or re.search(
        r"\b((?:lunes|martes|miércoles|miercoles|jueves|viernes|sábado|sabado|domingo)\s+\d{1,2}\s+de\s+\w+)", t, re.I)
    m_part = re.search(r"\b(\d+\s+(?:líderes|lideres|conquistadores|aventureros|jóvenes|jovenes|guías|guias|niños|ninos|participantes)[^.;\n]*)", t, re.I)
    low = t.lower()
    modalidad = "Mixta" if re.search(r"mixta|semi", low) else "Presencial" if "presencial" in low else "Virtual" if "virtual" in low else ""
    actividad = "ACTIVIDAD"
    for clave in ("estudio bíblico", "estudio biblico", "rincón de unidades", "rincon de unidades", "reunión", "reunion",
                  "campamento", "caminata", "investidura", "culto", "ensayo", "marcha"):
        if clave in low:
            actividad = clave.upper().replace("BIBLICO", "BÍBLICO").replace("RINCON", "RINCÓN").replace("REUNION", "REUNIÓN")
            break
    parrafos = list(_oraciones(notas))
    desarrollo = "\n".join(" ".join(parrafos[i:i + 3]) for i in range(0, len(parrafos), 3))
    return {
        "REF": f"INFORME DE ACTIVIDAD: {actividad}",
        "ACTIVIDAD": actividad,
        "FECHA_EVENTO": m_fecha.group(1) if m_fecha else "",
        "MODALIDAD": modalidad,
        "PARTICIPANTES": m_part.group(1).strip() if m_part else "",
        "DESARROLLO": desarrollo,
    }


@app.post("/api/redactar")
def api_redactar():
    notas = (request.form.get("instrucciones") or "").strip()
    if not notas:
        return "Faltan las notas.", 400
    aviso = ""
    datos = None
    if os.getenv("ANTHROPIC_API_KEY"):
        try:
            datos = _redactar_con_ia(notas)
        except Exception as e:                          # sin internet, clave mala, etc.
            aviso = f"No se pudo usar la IA ({type(e).__name__}); se usó el redactor básico."
    else:
        aviso = "Sin ANTHROPIC_API_KEY: se usó el redactor básico (revisa el texto)."
    if datos is None:
        datos = _redactar_local(notas)
    salida = {k: str(datos.get(k, "") or "") for k in g.CAMPOS_TEXTO}
    if salida["MODALIDAD"] not in MODALIDADES:
        salida["MODALIDAD"] = ""
    salida["_aviso"] = aviso
    return jsonify(salida)


# ------------------------------ Generar informe -----------------------------
def _nombre_archivo(campos: dict) -> str:
    base = re.sub(r"[^\w\- ]+", "", campos.get("ACTIVIDAD", "") or "informe", flags=re.U).strip().replace(" ", "_")
    return f"Informe_{base or 'actividad'}_{time.strftime('%Y%m%d_%H%M%S')}"


@app.post("/api/generar")
def api_generar():
    try:
        campos = json.loads(request.form.get("campos", "{}"))
    except json.JSONDecodeError:
        return "Campos inválidos.", 400

    nombre = _nombre_archivo(campos)
    carpeta = g.ENTRADA / nombre
    carpeta.mkdir(parents=True, exist_ok=True)
    fotos = []
    for i, f in enumerate(request.files.getlist("fotos"), start=1):   # el navegador las envía en orden
        ext = Path(f.filename or "").suffix.lower()
        if ext not in g.EXTENSIONES:
            return f"Formato de foto no soportado: {f.filename}", 400
        destino = carpeta / f"{i:02d}{ext}"
        f.save(destino)
        fotos.append(g.Foto(str(destino)))

    try:
        res = g.generar_informe(campos, fotos, nombre)
    except Exception as e:
        return f"No se pudo generar el informe: {e}", 500

    pdf, msg_pdf = g.exportar_pdf(res.docx)
    partes = [f"Informe creado ({len(fotos)} foto(s); distribución por hoja: {res.plan or 'sin fotos'})."]
    partes.append(msg_pdf)
    partes += res.avisos
    return jsonify({
        "mensaje": " ".join(partes),
        "docx": f"/salidas/{res.docx.name}",
        "pdf": f"/salidas/{pdf.name}" if pdf else None,
    })


@app.get("/salidas/<path:nombre>")
def salidas(nombre):
    return send_from_directory(g.SALIDAS, nombre)


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


if __name__ == "__main__":
    puerto = int(os.getenv("PORT", "5000"))
    threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{puerto}")).start()
    app.run(host="127.0.0.1", port=puerto, debug=False)
