"""
Interfaz web ligera (Flask) para buscar articulos del ERP y ver su despiece.

Arrancar:
    py app.py
Luego abrir http://127.0.0.1:5000
"""
import io
import math

import openpyxl
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template

from desglose import (buscar_articulos, desglose, es_externo, nombre_articulo,
                      exportar_excel,
                      tiempo_operacion, escandallo_directo, sin_operacion,
                      coste_propio)

app = Flask(__name__)

# La pagina de costes del catalogo va en su propio modulo: tiene estado
# (el recalculo en curso) y no tiene nada que ver con el buscador.
# El import va aqui abajo y no arriba porque web_costes importa de este
# modulo (construir_arbol, avisos_arbol) y arriba seria circular.
from web_costes import bp as bp_costes  # noqa: E402
app.register_blueprint(bp_costes)


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
                    servicio_raiz=None, medio_raiz=0, externo_raiz=0):
    """Árbol del escandallo con coste de MATERIAL (hojas compradas) y de OPERACIÓN
    (nodos fabricados = tiempo × 17 €/h), y el rollup material+operación por nodo.

    servicio_raiz: coste del trabajo externo del propio artículo buscado, que se
    suma a sus materiales (ver SQL_SERVICIO_EXTERNO en desglose.py).
    medio_raiz: 1 si el tiempo de la raíz sale de la media de bonos en vez de
    la mano de obra imputada en el ERP."""
    root = {"id": codigo, "nombre": nombre, "cant": 1.0, "unidad": None, "tipo": None,
            "precio": None, "de_conjunto": 0, "sin_escandallo": 0,
            "sin_operacion": int(sin_op_raiz or 0), "servicio": servicio_raiz,
            "tiempo": tiempo_raiz, "tiempo_medio": int(medio_raiz or 0),
            # marcado a mano en el ERP (Trabajos_Operacion.Externo): solo etiqueta,
            # no cambia ni el coste ni los avisos
            "externo": int(externo_raiz or 0),
            "coste": servicio_raiz or 0.0, "hijos": []}
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
                "tiempo_medio": int(r.get("TiempoMedio", 0) or 0),
                "externo": int(r.get("EsExterno", 0) or 0),
                "coste": _num(r["Coste"]) or 0.0, "hijos": []}
        nodos[ruta] = nodo
        padre_key = ("|" + "|".join(segs[:-1]) + "|") if len(segs) > 1 else f"|{codigo}|"
        nodos.get(padre_key, root)["hijos"].append(nodo)

    def rollup(n):
        n["es_hoja"] = not n["hijos"]
        # coste de operacion: solo nodos FABRICADOS (ramas).
        # "tiempo" es min/PIEZA; "tiempo_op" son los minutos que consume esta
        # linea (min/pieza x cantidad usada), que es lo que se cobra a 17 €/h.
        if n["es_hoja"]:
            n["coste_op"] = 0.0; n["tiempo_op"] = 0.0
            n["sin_tiempo"] = 0; n["tiempo_efectivo"] = None
        elif n.get("tiempo") is not None and not n.get("tiempo_medio"):
            # mano de obra IMPUTADA en el ERP: es una declaración explícita de que
            # la pieza lleva trabajo, así que manda incluso sobre la casilla.
            n["tiempo_op"] = round(n["tiempo"] * (n["cant"] or 0), 4)
            n["coste_op"] = round(n["tiempo"] * (n["cant"] or 0) * RATE_OP, 4)
            n["sin_tiempo"] = 0; n["tiempo_efectivo"] = n["tiempo"]
        elif n.get("sin_operacion"):
            # su fase no declara operación ("Sin operación"): 0 € de mano de obra
            # es CORRECTO, no es un dato que falte -> no se avisa.
            #
            # MANDA SOBRE LA MEDIA DE PARTES. Hay 3 piezas declaradas sin operación
            # que aun así tienen partes fichados (10902103, 23101014, 60104182), y
            # antes ese tiempo se colaba: el ANGULO DER. salía a 0,27 min mientras
            # su espejo el ANGULO IZQ., declarado igual pero sin partes, salía "sin
            # operación". Dos piezas idénticas con distinto coste según si alguien
            # llegó a fichar. La casilla del ERP es el criterio (ver c012ed1).
            n["coste_op"] = 0.0; n["tiempo_op"] = 0.0
            n["sin_tiempo"] = 0; n["tiempo_efectivo"] = None
        elif n.get("tiempo") is not None:
            # media de los bonos, para las piezas que sí declaran operación.
            # Se usa igual, marcada como "medio" en la tabla: es la media REAL
            # de lo que se tardó, no un tiempo teórico. El teórico sería el
            # estándar de producción, que es la rama de arriba.
            n["tiempo_op"] = round(n["tiempo"] * (n["cant"] or 0), 4)
            n["coste_op"] = round(n["tiempo"] * (n["cant"] or 0) * RATE_OP, 4)
            n["sin_tiempo"] = 0; n["tiempo_efectivo"] = n["tiempo"]
        elif n.get("servicio") is not None or n.get("precio") is not None:
            # TRABAJO EXTERNO (lacado, zincado, remachado...): el nodo tiene
            # escandallo Y ADEMÁS se compra (ver coste_propio/SQL_SERVICIO_EXTERNO
            # para la raíz, o la rama "c" de EsValorable para un hijo). Ese trabajo
            # lo hace un tercero fuera de GYC, así que 0 min de mano de obra propia
            # es CORRECTO, no un dato que falte -> no se avisa como "sin tiempo"
            # (18103604 BASE...LACADA avisaba en falso al buscarlo directamente).
            n["coste_op"] = 0.0; n["tiempo_op"] = 0.0
            n["sin_tiempo"] = 0; n["tiempo_efectivo"] = None
        else:
            # fabricado pero sin ningún tiempo con el que costear
            n["coste_op"] = 0.0; n["tiempo_op"] = 0.0
            n["sin_tiempo"] = 1; n["tiempo_efectivo"] = None
        mat = n["coste"] or 0.0
        op = n["coste_op"]
        tmin = n["tiempo_op"]
        for h in n["hijos"]:
            rollup(h)
            mat += h["coste_mat"]; op += h["coste_op_total"]; tmin += h["tiempo_op_total"]
        # 6 decimales: con 4, un coste de 8,5e-06 EUR se redondeaba a 0,0 exacto
        # y desaparecia del arbol (ademas de dispararse como "sin coste").
        n["coste_mat"] = round(mat, 6)
        n["coste_op_total"] = round(op, 6)
        # minutos acumulados del nodo: los suyos mas los de todo lo que cuelga
        n["tiempo_op_total"] = round(tmin, 6)
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
    tiempo_raiz, medio_raiz = tiempo_operacion(codigo)
    arbol = (construir_arbol(df, codigo, nombre_articulo(codigo), tiempo_raiz,
                             sin_operacion(codigo), coste_propio(codigo), medio_raiz,
                             es_externo(codigo))
             if not df.empty else None)
    return jsonify({
        "codigo": codigo,
        "lineas": len(df),
        "niveles": int(df["Nivel"].max()) if not df.empty else 0,
        "coste_material": arbol["coste_mat"] if arbol else 0,
        "coste_operacion": arbol["coste_op_total"] if arbol else 0,
        "tiempo_total": arbol["tiempo_op_total"] if arbol else 0,
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


def primera_hoja_visible(file):
    """Nombre de la primera hoja NO oculta del libro, o None si no se puede leer.

    pd.read_excel() sin sheet_name coge la primera hoja del libro aunque esté
    oculta. Los libros de análisis que llegan de Finanzas arrastran decenas de
    hojas ocultas de trabajo, así que la primera del fichero no suele ser la que
    el usuario ve al abrirlo: se costeaba un juego de artículos distinto del que
    creía haber subido, y sin ningún aviso."""
    try:
        wb = openpyxl.load_workbook(file, read_only=True)
        try:
            for ws in wb.worksheets:
                if ws.sheet_state == "visible":
                    return ws.title
        finally:
            wb.close()
    except Exception:
        return None                          # que no rompa: se cae al comportamiento previo
    return None


def extraer_ids(file):
    """Extrae los IDs de artículo de un Excel subido. Busca una columna que
    se llame algo tipo 'articulo'/'codigo'/'id'; si no, usa la primera columna."""
    hoja = primera_hoja_visible(file)
    if hasattr(file, "seek"):
        file.seek(0)                         # openpyxl deja consumido el stream subido
    df = pd.read_excel(file, dtype=str, **({"sheet_name": hoja} if hoja else {}))
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


# Cada motivo de detalle cuelga de la alerta con la que la interfaz lo agrupa
# (ver renderAlert() en templates/index.html): "sin tipo" y "sin precio" son dos
# lecturas del mismo agujero —la pieza no se ha podido valorar— y por eso el
# Excel las suma bajo "Sin coste" sin perder el motivo concreto de cada línea.
ALERTA_DE_MOTIVO = {
    "Sin tipo de aprovisionamiento": "Sin coste",
    "Sin precio de compra": "Sin coste",
    "Sin escandallo activo": "Sin escandallo",
    "Sin tiempo de operación": "Sin tiempo de operación",
}


def avisos_arbol(arbol):
    """Los mismos huecos que el panel de avisos de la interfaz, recorriendo el
    árbol con su mismo criterio (ver el bloque flat() de templates/index.html).

    Se recorre el árbol y no el DataFrame a propósito: el lote tenía sus propios
    filtros y se desviaba de lo que se ve en pantalla. Los dos casos que lo
    delataron: una pieza FABRICADA sin precio de compra no es un hueco (su coste
    sube de sus componentes) y se contaba, y una compra directa —cuyo despiece es
    una sola fila autorreferencial— se quedaba fuera del filtro de hojas y no se
    miraba nunca, saliendo Completo con 0 €.

    Criterio por nodo, en este orden:
      - tiene hijos  -> es fabricada: solo puede faltarle el tiempo de operación
      - de_conjunto  -> no es un hueco: su coste entra por el conjunto
      - sin_escandallo -> pieza fabricada sin fase activa con la que costearla
      - precio nulo  -> hueco de coste real; es "sin tipo" si además no tiene tipo
    """
    huecos = {"Sin escandallo activo": {}, "Sin tipo de aprovisionamiento": {},
              "Sin precio de compra": {}, "Sin tiempo de operación": {}}

    def visita(n):
        if n["hijos"]:
            if n.get("sin_tiempo"):
                huecos["Sin tiempo de operación"].setdefault(n["id"], n)
            for h in n["hijos"]:
                visita(h)
        # Una pieza de conjunto NO es un hueco: el material entra una sola vez al
        # explotar el conjunto, y las demás piezas que salen de esa misma chapa
        # llegan sin coste propio a propósito. El 11401042 lo enseña: el CONJUNTO
        # PONEDERO 2 HUECOS agrupa techo, barreta y piso; el piso arrastra la
        # chapa (3,79 €) y techo y barreta se avisaban como "sin escandallo"
        # cuando su coste ya estaba contado. Esta rama va ANTES de la cadena
        # porque si solo se quitara de "sin escandallo" caería en "precio nulo"
        # y el falso aviso reaparecería como "Sin coste", que es peor.
        elif n.get("de_conjunto"):
            pass
        elif n.get("sin_escandallo"):
            huecos["Sin escandallo activo"].setdefault(n["id"], n)
        elif n.get("precio") is None:
            motivo = ("Sin precio de compra" if n.get("tipo")
                      else "Sin tipo de aprovisionamiento")
            huecos[motivo].setdefault(n["id"], n)

    # la raíz solo se examina cuando ES la hoja (compra directa), igual que la
    # interfaz, que arranca en arbol.hijos y cae a [arbol] si no tiene ninguno.
    for n in (arbol["hijos"] or [arbol]):
        visita(n)
    return {k: list(v.values()) for k, v in huecos.items()}


def resumen_lote(codigo):
    """Coste total de un artículo + datos faltantes (sin tipo / sin precio),
    calculando coste de material y operación mediante el árbol.

    La columna Completo (Sí/No) resume si al artículo le falta algún dato. El
    detalle de lo que falta va SIEMPRE a la hoja Faltantes, así que un "No" se
    puede rastrear allí sin depender de ningún color."""
    df = desglose(codigo)
    nombre = nombre_articulo(codigo)
    if df.empty:
        # sin despiece no hay nada que contar, pero el motivo tiene que quedar
        # registrado en Faltantes o el "No" se quedaría sin explicación
        fila = {"IdArticulo": codigo, "Descripcion": nombre,
                "CosteMaterial": None, "CosteOperacion": None, "CosteTotal": None,
                "Sin coste": 0, "Sin escandallo": 0, "Sin tiempo": 0,
                "Sin tipo": 0, "Sin precio": 0,
                "Completo": "No"}
        return fila, [{"Articulo": codigo, "IdComponente": codigo,
                       "Descripcion": nombre or "(no existe en el ERP)",
                       "Alerta": "Sin coste",
                       "Motivo": "Artículo no encontrado o sin despiece"}]

    tiempo_raiz, medio_raiz = tiempo_operacion(codigo)
    arbol = construir_arbol(df, codigo, nombre, tiempo_raiz, sin_operacion(codigo),
                            coste_propio(codigo), medio_raiz)
    coste_mat = arbol["coste_mat"] if arbol else 0.0
    coste_op = arbol["coste_op_total"] if arbol else 0.0
    coste_tot = arbol["coste_total"] if arbol else 0.0

    huecos = avisos_arbol(arbol)
    sin_esc = huecos["Sin escandallo activo"]
    sin_tipo = huecos["Sin tipo de aprovisionamiento"]
    sin_precio = huecos["Sin precio de compra"]
    sin_tiempo = huecos["Sin tiempo de operación"]

    faltantes = []
    for motivo in ("Sin tipo de aprovisionamiento", "Sin precio de compra",
                   "Sin escandallo activo", "Sin tiempo de operación"):
        for n in huecos[motivo]:
            faltantes.append({"Articulo": codigo, "IdComponente": n["id"],
                              "Descripcion": n["nombre"],
                              "Alerta": ALERTA_DE_MOTIVO[motivo], "Motivo": motivo})

    n_esc, n_st, n_sp, n_tie = len(sin_esc), len(sin_tipo), len(sin_precio), len(sin_tiempo)
    fila = {"IdArticulo": codigo, "Descripcion": nombre,
            "CosteMaterial": coste_mat, "CosteOperacion": coste_op, "CosteTotal": coste_tot,
            # las dos alertas que la interfaz separa, y debajo el desglose de
            # "Sin coste" en sus dos motivos, que se conservan por compatibilidad
            # con los Excel ya repartidos.
            "Sin coste": n_st + n_sp, "Sin escandallo": n_esc, "Sin tiempo": n_tie,
            "Sin tipo": n_st, "Sin precio": n_sp,
            # "Sin tiempo" NO tumba el Completo: en la interfaz es el aviso azul
            # (informativo, la pieza está costeada en material), no la alerta roja
            # de "Sin coste". El 10902001 se lee como completo teniéndolo.
            "Completo": "Sí" if (n_esc + n_st + n_sp) == 0 else "No"}
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
        pd.DataFrame([{"Articulo": "", "IdComponente": "", "Descripcion": "",
                       "Alerta": "", "Motivo": "(sin faltantes)"}])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xls:
        # sin resaltado de color: la columna Completo (Sí/No) es la que marca los
        # artículos con datos incompletos, y el detalle está en la hoja Faltantes.
        # Así el fichero se puede filtrar y ordenar sin depender del formato.
        df_costes.to_excel(xls, sheet_name="Costes", index=False)
        df_falt.to_excel(xls, sheet_name="Faltantes", index=False)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="costes_lote.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
