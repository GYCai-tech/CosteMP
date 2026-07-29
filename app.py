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

from desglose import (buscar_articulos, desglose, nombre_articulo, exportar_excel,
                      tiempo_operacion, escandallo_directo, sin_operacion,
                      servicio_externo)

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


def _s(v):
    """Como _num pero para campos de texto (evita NaN suelto e invalido en el JSON)."""
    return None if v is None or (isinstance(v, float) and math.isnan(v)) else v


RATE_OP = 17.0 / 60.0   # 17 €/h operario -> €/min


def construir_arbol(df, codigo, nombre, tiempo_raiz=None, sin_op_raiz=0,
                    servicio_raiz=None):
    """Árbol del escandallo con coste de MATERIAL (hojas compradas) y de OPERACIÓN
    (nodos fabricados = tiempo × 17 €/h), y el rollup material+operación por nodo.

    servicio_raiz: coste del trabajo externo del propio artículo buscado, que se
    suma a sus materiales (ver SQL_SERVICIO_EXTERNO en desglose.py)."""
    root = {"id": codigo, "nombre": nombre, "cant": 1.0, "unidad": None, "tipo": None,
            "precio": None, "de_conjunto": 0, "sin_escandallo": 0,
            "sin_operacion": int(sin_op_raiz or 0), "servicio": servicio_raiz,
            "tiempo": tiempo_raiz, "coste": servicio_raiz or 0.0, "hijos": []}
    nodos = {f"|{codigo}|": root}

    for r in df.sort_values("Ruta").to_dict(orient="records"):
        ruta = r["Ruta"]
        segs = [s for s in ruta.split("|") if s]
        # fila del propio articulo (compra directa): el root ES la hoja
        if segs == [codigo]:
            root["coste"] = _num(r["Coste"]) or 0.0
            root["cant"] = _num(r["Cant"]) or 1.0; root["unidad"] = _s(r["Unidad"])
            root["tipo"] = _s(r["TipoCompra"]); root["precio"] = _num(r["PrecioCompra"])
            continue
        nodo = {"id": r["IdArticulo"], "nombre": r["Componente"],
                "cant": _num(r["Cant"]), "unidad": _s(r["Unidad"]),
                "tipo": _s(r["TipoCompra"]), "precio": _num(r["PrecioCompra"]),
                "fuente": _s(r.get("PrecioFuente")),
                "de_conjunto": int(r.get("DeConjunto", 0) or 0),
                "sin_escandallo": int(r.get("SinEscandallo", 0) or 0),
                "sin_operacion": int(r.get("SinOperacion", 0) or 0),
                "tiempo": _num(r.get("TiempoOp")),
                "coste": _num(r["Coste"]) or 0.0, "hijos": []}
        nodos[ruta] = nodo
        padre_key = ("|" + "|".join(segs[:-1]) + "|") if len(segs) > 1 else f"|{codigo}|"
        nodos.get(padre_key, root)["hijos"].append(nodo)

    def rollup(n):
        n["es_hoja"] = not n["hijos"]
        # coste de operacion: solo nodos FABRICADOS (ramas)
        if n["es_hoja"]:
            n["coste_op"] = 0.0; n["sin_tiempo"] = 0
        elif n.get("tiempo") is not None:
            # hay tiempo medido en bonos: manda el dato real
            n["coste_op"] = round(n["tiempo"] * (n["cant"] or 0) * RATE_OP, 4); n["sin_tiempo"] = 0
        elif n.get("sin_operacion"):
            # su fase no declara operación ("Sin operación"): 0 € de mano de obra
            # es CORRECTO, no es un dato que falte -> no se avisa
            n["coste_op"] = 0.0; n["sin_tiempo"] = 0
        else:
            n["coste_op"] = 0.0; n["sin_tiempo"] = 1     # fabricado pero sin tiempo de bono
        mat = n["coste"] or 0.0
        op = n["coste_op"]
        for h in n["hijos"]:
            rollup(h)
            mat += h["coste_mat"]; op += h["coste_op_total"]
        # 6 decimales: con 4, un coste de 8,5e-06 EUR se redondeaba a 0,0 exacto
        # y desaparecia del arbol (ademas de dispararse como "sin coste").
        n["coste_mat"] = round(mat, 6)
        n["coste_op_total"] = round(op, 6)
        n["coste_total"] = round(mat + op, 6)

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
    # solo articulos activos: los descatalogados (Estado = 1) no se buscan
    df = buscar_articulos(q)
    return jsonify(_records(df))


@app.route("/api/desglose")
def api_desglose():
    codigo = request.args.get("codigo", "")
    df = desglose(codigo)
    arbol = (construir_arbol(df, codigo, nombre_articulo(codigo), tiempo_operacion(codigo),
                             sin_operacion(codigo), servicio_externo(codigo))
             if not df.empty else None)
    return jsonify({
        "codigo": codigo,
        "lineas": len(df),
        "niveles": int(df["Nivel"].max()) if not df.empty else 0,
        "coste_material": arbol["coste_mat"] if arbol else 0,
        "coste_operacion": arbol["coste_op_total"] if arbol else 0,
        "coste_total": arbol["coste_total"] if arbol else 0,
        "sin_precio": int(df["SinPrecio"].sum()) if not df.empty else 0,
        "filas": _records(df),
        "arbol": arbol,
    })


@app.route("/api/escandallo")
def api_escandallo():
    codigo = request.args.get("codigo", "")
    df = escandallo_directo(codigo)
    if df.empty:
        return "Sin escandallo directo (¿es compra o conjunto?)", 404
    df = df.rename(columns={"Descripcion": "Descripción"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xls:
        df.to_excel(xls, sheet_name="Escandallo", index=False)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"escandallo_{codigo}.xlsx",
    )


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

def _norm(s):
    """minúsculas sin tildes, para comparar nombres de columna."""
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", str(s).lower())
                   if not unicodedata.combining(c)).strip()


def extraer_ids(file):
    """Extrae los IDs de artículo de un Excel subido. Busca una columna que
    se llame algo tipo 'articulo'/'codigo'/'id'; si no, usa la primera columna."""
    df = pd.read_excel(file, dtype=str)
    objetivo = None
    for col in df.columns:
        cl = _norm(col)
        if "articul" in cl or "codig" in cl or cl in ("id", "ref", "referencia"):
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
    """Coste total de un artículo + datos faltantes (sin tipo / sin precio),
    calculando coste de material y operación mediante el árbol."""
    df = desglose(codigo)
    nombre = nombre_articulo(codigo)
    if df.empty:
        fila = {"IdArticulo": codigo, "Descripcion": nombre,
                "CosteMaterial": None, "CosteOperacion": None, "CosteTotal": None,
                "De conjunto": 0, "Sin escandallo": 0, "Sin tipo": 0, "Sin precio": 0,
                "Estado": "No encontrado / sin datos"}
        return fila, []

    arbol = construir_arbol(df, codigo, nombre, tiempo_operacion(codigo), sin_operacion(codigo),
                            servicio_externo(codigo))
    coste_mat = arbol["coste_mat"] if arbol else 0.0
    coste_op = arbol["coste_op_total"] if arbol else 0.0
    coste_tot = arbol["coste_total"] if arbol else 0.0

    padres = set(df["Articulo"])
    hojas = df[~df["IdArticulo"].isin(padres)]
    de_conj = hojas[hojas["DeConjunto"] == 1].drop_duplicates("IdArticulo")
    sin_esc = hojas[(hojas["SinEscandallo"] == 1) & (hojas["DeConjunto"] == 0)].drop_duplicates("IdArticulo")
    sin_tipo = hojas[(hojas["TipoCompra"].isna()) & (hojas["SinEscandallo"] == 0)].drop_duplicates("IdArticulo")
    sin_precio = df[df["SinPrecio"] == 1].drop_duplicates("IdArticulo")

    faltantes = []
    for motivo, sub in [("Proviene de conjunto", de_conj),
                        ("Sin escandallo activo", sin_esc),
                        ("Sin tipo de aprovisionamiento", sin_tipo),
                        ("Sin precio de compra", sin_precio)]:
        for r in sub.to_dict(orient="records"):
            faltantes.append({"Articulo": codigo, "IdComponente": r["IdArticulo"],
                              "Descripcion": r["Componente"], "Motivo": motivo})

    n_conj, n_esc, n_st, n_sp = len(de_conj), len(sin_esc), len(sin_tipo), len(sin_precio)
    fila = {"IdArticulo": codigo, "Descripcion": nombre,
            "CosteMaterial": coste_mat, "CosteOperacion": coste_op, "CosteTotal": coste_tot,
            "De conjunto": n_conj, "Sin escandallo": n_esc, "Sin tipo": n_st, "Sin precio": n_sp,
            "Estado": "OK" if (n_conj + n_esc + n_st + n_sp) == 0 else "Incompleto"}
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
