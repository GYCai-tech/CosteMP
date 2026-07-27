"""
Interfaz web ligera (Flask) para buscar articulos del ERP y ver su despiece.

Arrancar:
    py app.py
Luego abrir http://127.0.0.1:5000
"""
import io
import math

import pandas as pd
from openpyxl.styles import PatternFill
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


# ---------------------------------------------------------------------------
# Flujo por LOTE: subir un Excel con IDs -> Excel con el coste total de cada uno
# ---------------------------------------------------------------------------

def extraer_ids(file):
    """Extrae los IDs de artículo de un Excel subido. Busca una columna que
    se llame algo tipo 'articulo'/'codigo'/'id'; si no, usa la primera columna."""
    df = pd.read_excel(file, dtype=str)
    objetivo = None
    for col in df.columns:
        cl = str(col).strip().lower()
        if "articulo" in cl or "codigo" in cl or "código" in cl or cl in ("id", "ref", "referencia"):
            objetivo = col
            break
    if objetivo is None:
        objetivo = df.columns[0]

    ids, vistos = [], set()
    for v in df[objetivo].dropna().tolist():
        s = str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]                       # números leídos como 1234.0
        if s.isdigit() and len(s) < 8:
            s = s.zfill(8)                   # códigos ERP de 8 dígitos con ceros
        if s and s not in vistos:
            vistos.add(s)
            ids.append(s)
    return ids


def resumen_lote(codigo):
    """Coste total de un artículo + datos faltantes (sin tipo / sin precio)."""
    df = desglose(codigo)
    nombre = nombre_articulo(codigo)
    if df.empty:
        fila = {"IdArticulo": codigo, "Descripcion": nombre, "CosteTotal": None,
                "Sin tipo": 0, "Sin precio": 0, "Estado": "No encontrado / sin datos"}
        return fila, []

    total = round(float(df["Coste"].sum()), 4)
    padres = set(df["Articulo"])
    hojas = df[~df["IdArticulo"].isin(padres)]
    sin_tipo = hojas[hojas["TipoCompra"].isna()].drop_duplicates("IdArticulo")
    sin_precio = df[df["SinPrecio"] == 1].drop_duplicates("IdArticulo")

    faltantes = []
    for r in sin_tipo.to_dict(orient="records"):
        faltantes.append({"Articulo": codigo, "IdComponente": r["IdArticulo"],
                          "Descripcion": r["Componente"], "Motivo": "Sin tipo de aprovisionamiento"})
    for r in sin_precio.to_dict(orient="records"):
        faltantes.append({"Articulo": codigo, "IdComponente": r["IdArticulo"],
                          "Descripcion": r["Componente"], "Motivo": "Sin precio de compra"})

    n_sin_tipo, n_sin_precio = len(sin_tipo), len(sin_precio)
    fila = {"IdArticulo": codigo, "Descripcion": nombre, "CosteTotal": total,
            "Sin tipo": n_sin_tipo, "Sin precio": n_sin_precio,
            "Estado": "OK" if (n_sin_tipo + n_sin_precio) == 0 else "Incompleto"}
    return fila, faltantes


@app.route("/lote", methods=["GET", "POST"])
def lote():
    if request.method == "GET":
        return render_template("lote.html")

    file = request.files.get("archivo")
    if not file or not file.filename:
        return "Sube un archivo Excel (.xlsx)", 400
    try:
        ids = extraer_ids(file)
    except Exception as ex:
        return f"No se pudo leer el Excel: {ex}", 400
    if not ids:
        return "No se encontraron IDs de artículo en el Excel", 400

    costes, faltantes = [], []
    for cod in ids[:2000]:                   # tope de seguridad
        fila, falt = resumen_lote(cod)
        costes.append(fila)
        faltantes.extend(falt)

    df_costes = pd.DataFrame(costes)
    df_falt = pd.DataFrame(faltantes) if faltantes else \
        pd.DataFrame([{"Articulo": "", "IdComponente": "", "Descripcion": "", "Motivo": "(sin faltantes)"}])

    buf = io.BytesIO()
    rojo = PatternFill("solid", fgColor="FFC7CE")
    with pd.ExcelWriter(buf, engine="openpyxl") as xls:
        df_costes.to_excel(xls, sheet_name="Costes", index=False)
        df_falt.to_excel(xls, sheet_name="Faltantes", index=False)
        # resaltar en rojo los artículos incompletos / no encontrados
        ws = xls.sheets["Costes"]
        ncols = len(df_costes.columns)
        for pos, estado in enumerate(df_costes["Estado"].tolist()):
            if estado != "OK":
                for c in range(1, ncols + 1):
                    ws.cell(row=pos + 2, column=c).fill = rojo
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="costes_lote.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
