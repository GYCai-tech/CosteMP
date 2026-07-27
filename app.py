"""
Interfaz web ligera (Flask) para buscar articulos del ERP y ver su despiece.

Arrancar:
    py app.py
Luego abrir http://127.0.0.1:5000
"""
import io
import math

import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template

from desglose import buscar_articulos, desglose, nombre_articulo, exportar_excel

app = Flask(__name__)


def _records(df):
    """Convierte el DataFrame a lista de dicts reemplazando NaN por None.
    (NaN no es JSON válido y rompe JSON.parse en el navegador.)"""
    recs = df.to_dict(orient="records")
    for r in recs:
        for k, v in r.items():
            if isinstance(v, float) and math.isnan(v):
                r[k] = None
    return recs


def _num(v):
    return None if v is None or (isinstance(v, float) and math.isnan(v)) else float(v)


def construir_arbol(df, codigo, nombre):
    """Construye el árbol jerárquico del escandallo a partir de la columna Ruta,
    con el coste acumulado (rollup) en cada nodo (fase/subconjunto)."""
    root = {"id": codigo, "nombre": nombre, "cant": None, "unidad": None,
            "tipo": None, "coste": 0.0, "hijos": []}
    nodos = {f"|{codigo}|": root}

    for r in df.sort_values("Ruta").to_dict(orient="records"):
        ruta = r["Ruta"]
        segs = [s for s in ruta.split("|") if s]
        # fila del propio articulo (compra directa): el root ES la hoja
        if segs == [codigo]:
            root["coste"] = _num(r["Coste"]) or 0.0
            root["cant"] = _num(r["Cant"]); root["unidad"] = r["Unidad"]
            root["tipo"] = r["TipoCompra"]
            continue
        nodo = {"id": r["IdArticulo"], "nombre": r["Componente"],
                "cant": _num(r["Cant"]), "unidad": r["Unidad"],
                "tipo": r["TipoCompra"], "precio": _num(r["PrecioCompra"]),
                "coste": _num(r["Coste"]) or 0.0, "hijos": []}
        nodos[ruta] = nodo
        padre_key = ("|" + "|".join(segs[:-1]) + "|") if len(segs) > 1 else f"|{codigo}|"
        nodos.get(padre_key, root)["hijos"].append(nodo)

    def rollup(n):
        total = n["coste"] + sum(rollup(h) for h in n["hijos"])
        n["coste_total"] = round(total, 4)
        n["es_hoja"] = not n["hijos"]
        return total

    rollup(root)
    return root


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/buscar")
def api_buscar():
    q = request.args.get("q", "")
    if len(q.strip()) < 2:
        return jsonify([])
    df = buscar_articulos(q)
    return jsonify(_records(df))


@app.route("/api/desglose")
def api_desglose():
    codigo = request.args.get("codigo", "")
    df = desglose(codigo)
    return jsonify({
        "codigo": codigo,
        "lineas": len(df),
        "niveles": int(df["Nivel"].max()) if not df.empty else 0,
        "coste_total": round(float(df["Coste"].sum()), 4) if not df.empty else 0,
        "sin_precio": int(df["SinPrecio"].sum()) if not df.empty else 0,
        "filas": _records(df),
        "arbol": construir_arbol(df, codigo, nombre_articulo(codigo)) if not df.empty else None,
    })


@app.route("/api/excel")
def api_excel():
    codigo = request.args.get("codigo", "")
    df = desglose(codigo)
    if df.empty:
        return "Sin desglose", 404
    nombre = nombre_articulo(codigo)
    buf = io.BytesIO()
    exportar_excel(df, codigo, nombre, buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"desglose_{codigo}.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
