"""
Prueba rápida sin navegador: genera un informe con N fotos de mentira.

    python tools/probar.py 7            -> salidas/prueba_7_fotos.docx (+ PDF si se puede)
    python tools/probar.py 7 largo      -> con texto muy largo
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
import generador as g  # noqa: E402

FORMAS = [(1600, 1067), (1067, 1600), (1600, 900), (1200, 1200), (900, 1600), (1600, 1200)]
COLORES = [(214, 90, 60), (60, 120, 200), (80, 170, 90), (230, 170, 40), (150, 90, 190), (60, 170, 170)]


def fotos_falsas(n):
    g.ENTRADA.mkdir(exist_ok=True)
    fotos = []
    for i in range(n):
        w, h = FORMAS[i % len(FORMAS)]
        im = Image.new("RGB", (w, h), COLORES[i % len(COLORES)])
        d = ImageDraw.Draw(im)
        d.rectangle([10, 10, w - 10, h - 10], outline="white", width=12)
        d.text((w // 2 - 60, h // 2 - 20), f"FOTO {i + 1}  {w}x{h}", fill="white")
        ruta = g.ENTRADA / f"_test_{i + 1:02d}.jpg"
        im.save(ruta, quality=85)
        fotos.append(g.Foto(str(ruta)))
    return fotos


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 5
    largo = len(sys.argv) > 2 and sys.argv[2] == "largo"
    campos = {
        "REF": "INFORME DE ACTIVIDAD: ESTUDIO BÍBLICO",
        "ACTIVIDAD": "Estudio bíblico",
        "FECHA_EVENTO": "Sábado 6 de marzo",
        "MODALIDAD": "Presencial",
        "PARTICIPANTES": "25 líderes, directiva Ministerio Joven Ebenezer, Regional Aventureros Vanesa Aguilar, "
                         "Regional Conquistadores – Magdiela Berrios, Regional Jóvenes – Jonathan Maná",
        "DESARROLLO": "El club se reunió el sábado para el estudio bíblico, dirigido por edades y al mando de cada consejero. "
                      "Se estudiaron las clases 10 a 12 con participación de todas las unidades.",
    }
    if largo:
        campos["DESARROLLO"] = ("\n".join(
            ["El club se reunió el sábado para el estudio bíblico, dirigido por edades y al mando de cada consejero. "
             "Se estudiaron las clases 10 a 12 con participación de todas las unidades y se hizo un repaso general."] * 3))
    res = g.generar_informe(campos, fotos_falsas(n), f"prueba_{n}_fotos" + ("_largo" if largo else ""))
    print("Plan de fotos por hoja:", res.plan)
    print("Avisos:", res.avisos or "ninguno")
    print("DOCX:", res.docx)
    pdf, msg = g.exportar_pdf(res.docx)
    print(msg, pdf or "")
