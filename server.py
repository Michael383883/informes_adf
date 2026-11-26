"""
server.py  ·  Servidor MCP (OPCIONAL) para usar el generador desde Claude Desktop, Copilot, etc.

La web (app.py) NO lo necesita: usa generador.py directamente.
Para usarlo:  pip install mcp   y luego   python server.py
"""
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel

import generador as g

mcp = FastMCP("informes-word")


class Foto(BaseModel):
    ruta: str
    descripcion: str = ""


@mcp.tool()
def revisar_plantilla() -> str:
    """Lista los marcadores {{...}} que tiene la plantilla."""
    return g.revisar_plantilla()


@mcp.tool()
def crear_informe(informacion: str = "", fotos: list[Foto] | None = None,
                  campos: dict[str, str] | None = None, nombre_salida: str = "informe") -> str:
    """Crea el .docx. `fotos` van en orden: la 1.ª a la hoja 1, el resto agrupado en hojas 2+."""
    res = g.generar_informe(campos, [g.Foto(f.ruta, f.descripcion) for f in (fotos or [])],
                            nombre_salida, informacion)
    msg = f"Informe guardado en: {res.docx}  (fotos por hoja: {res.plan})"
    return msg + ("\nAvisos: " + " ".join(res.avisos) if res.avisos else "")


@mcp.tool()
def abrir_en_word(nombre_archivo: str) -> str:
    destino = g._ruta_salida(nombre_archivo, ".docx")
    return g.abrir_archivo(destino) if destino.exists() else f"No existe: {destino}"


@mcp.tool()
def exportar_pdf(nombre_archivo: str) -> str:
    origen = g._ruta_salida(nombre_archivo, ".docx")
    if not origen.exists():
        return f"No existe: {origen}"
    pdf, msg = g.exportar_pdf(origen)
    return f"{msg} {pdf or ''}".strip()


if __name__ == "__main__":
    mcp.run()
