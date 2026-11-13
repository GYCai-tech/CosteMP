"""
Interfaz web ligera (Flask) para buscar articulos del ERP y ver su despiece.

Arrancar:
    py app.py
Luego abrir http://127.0.0.1:5000
"""
import datetime as dt
import io
import math
import re
import threading
import time
import uuid

import openpyxl
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from sqlalchemy import text

import db_pg
import exportar_costes
# construir_arbol y avisos_arbol se importan tambien para quien los siga
# pidiendo a este modulo; el calculo vive en desglose.py, solo una vez
from desglose import (avisos_arbol, buscar_articulos, con_orden,  # noqa: F401
                      construir_arbol, desglose, es_externo,
                      nombre_articulo, olvidar_articulo, leido_a,
                      exportar_excel,
                      tiempo_operacion, escandallo_directo, sin_operacion,
                      coste_propio, anulado)
# La pagina de costes del catalogo va en su propio modulo: tiene estado
# (el recalculo en curso) y no tiene nada que ver con el buscador.
from web_costes import bp as bp_costes
# Y la de configuracion (el coste/hora de operario), igual de aparte.
from web_config import bp as bp_config
# Y las que leen el detalle por pieza: /evolucion y /piezas.
from web_analisis import bp as bp_analisis

app = Flask(__name__)
app.register_blueprint(bp_costes)
app.register_blueprint(bp_config)
app.register_blueprint(bp_analisis)


def _records(df):
    """Convierte el DataFrame a lista de dicts reemplazando NaN por None.
    (NaN no es JSON válido y rompe JSON.parse en el navegador.)"""
    recs = df.to_dict(orient="records")
    for r in recs:
        for k, v in r.items():
            if isinstance(v, float) and math.isnan(v):
                r[k] = None
    return recs


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
                             es_externo(codigo), con_orden(codigo), anulado(codigo))
             if not df.empty else None)
    return jsonify({
        "codigo": codigo,
        "lineas": len(df),
        "niveles": int(df["Nivel"].max()) if not df.empty else 0,
        "coste_material": arbol["coste_mat"] if arbol else 0,
        "coste_operacion": arbol["coste_op_total"] if arbol else 0,
        "tiempo_total": arbol["tiempo_op_total"] if arbol else 0,
        "coste_total": arbol["coste_total"] if arbol else 0,
        # EUR/h con los que se ha valorado la operacion (del ERP, Empleados_Categorias)
        "coste_hora": arbol["coste_hora"] if arbol else None,
        "sin_precio": int(df["SinPrecio"].sum()) if not df.empty else 0,
        "filas": _records(df),
        "arbol": arbol,
        # de cuando son los datos: se recuerdan unos minutos (ver desglose.py)
        "leido": _hora(leido_a(codigo)),
    })


def _hora(momento):
    return momento.strftime("%H:%M") if momento else None


@app.route("/api/recalcular", methods=["POST"])
def api_recalcular():
    """El boton "Recalcular" de la pantalla del articulo.

    Vuelve a leer el articulo del ERP en ese momento, sin esperar a que caduque
    lo recordado, y si es del catalogo deja su coste al dia en Postgres, en
    /catalogo y en el libro de Excel sin esperar a la pasada de cada pocos
    dias. Despues la pantalla pide /api/desglose, que ya sale de lo recien
    leido: pulsar el boton cuesta UNA lectura del ERP, no dos.
    """
    codigo = re.sub(r"[^A-Za-z0-9]", "", request.args.get("codigo", ""))
    if not codigo:
        return jsonify(error="Falta el codigo del articulo"), 400
    olvidar_articulo(codigo)
    try:
        pg = db_pg.get_pg_engine()
        catalogo = exportar_costes.en_catalogo(pg, codigo)
        if catalogo:
            exportar_costes.recalcular_articulo(pg, codigo)
        else:
            desglose(codigo)        # solo leerlo, para que la pantalla lo pinte
    except Exception as ex:                                   # noqa: BLE001
        return jsonify(error=str(ex)[:300]), 500
    return jsonify(ok=True, codigo=codigo, en_catalogo=catalogo,
                   leido=_hora(leido_a(codigo)))


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


LOTE_MAX = 2000          # articulos por lote; tope de seguridad
LOTE_VIDA_S = 3600       # lo que se guarda un lote terminado para descargarlo

# Lotes en curso o recien terminados, por id. Se calculan en un hilo aparte:
# antes la peticion no volvia hasta tener el Excel, varios minutos sin ninguna
# senal en pantalla (el aviso de "calculando" se apagaba a los 4 s) y con un
# hilo del servidor ocupado todo ese rato.
_lotes = {}
_lotes_lock = threading.Lock()

MOTIVOS_LOTE = ("Sin tipo de aprovisionamiento", "Sin precio de compra",
                "Sin escandallo activo", "Sin tiempo de operación")


def costes_publicados(ids):
    """El coste que ya esta publicado en Postgres de los articulos del catalogo
    que esten en `ids`: {codigo: (fila, faltantes)}, con el mismo formato que
    resumen_lote().

    Es el mismo calculo (exportar_costes.calcular usa construir_arbol y
    avisos_arbol igual que resumen_lote), hecho en la ultima pasada o al pulsar
    Recalcular. Leerlo de ahi ahorra ir al ERP por cada articulo del catalogo,
    que en un lote suelen ser casi todos. La columna Calculado dice de cuando es.

    Los que se publicaron con error se dejan fuera para calcularlos al momento.
    Si Postgres no responde se devuelve {} y el lote entero va al ERP, como
    antes: no hace falta Postgres para costear.
    """
    if not ids:
        return {}
    try:
        with db_pg.get_pg_engine().connect() as c:
            filas = c.execute(text("""
                SELECT c.idarticulo, c.descripcion, c.coste_material,
                       c.coste_operacion, c.coste_total,
                       COALESCE(c.piezas_sin_escandallo,0), COALESCE(c.piezas_sin_tipo,0),
                       COALESCE(c.piezas_sin_precio,0), COALESCE(c.piezas_sin_tiempo,0),
                       c.completo, c.fecha_calculo
                FROM core.dim_coste_escandallo c
                INNER JOIN core.cfg_catalogo_coste g
                        ON g.idarticulo = c.idarticulo AND g.activo
                WHERE c.error IS NULL
                  AND c.idarticulo = ANY(CAST(:ids AS varchar[]))"""),
                {"ids": ids}).all()
            huecos = c.execute(text("""
                SELECT idarticulo, idpieza, descripcion, motivo
                FROM core.fact_coste_hueco
                WHERE idarticulo = ANY(CAST(:ids AS varchar[]))
                ORDER BY idarticulo, idpieza"""), {"ids": ids}).all()
    except Exception as ex:                                   # noqa: BLE001
        print(f"  aviso: lote sin el coste publicado, todo al ERP ({str(ex)[:120]})",
              flush=True)
        return {}

    por_articulo = {}
    for art, pieza, desc, motivo in huecos:
        por_articulo.setdefault(art, []).append((pieza, desc, motivo))

    out = {}
    for (cod, desc, mat, op, tot, n_esc, n_st, n_sp, n_tie, completo, fecha) in filas:
        fila = {"IdArticulo": cod, "Descripcion": desc,
                "CosteMaterial": float(mat) if mat is not None else None,
                "CosteOperacion": float(op) if op is not None else None,
                "CosteTotal": float(tot) if tot is not None else None,
                "Sin coste": n_st + n_sp, "Sin escandallo": n_esc, "Sin tiempo": n_tie,
                "Sin tipo": n_st, "Sin precio": n_sp,
                "Completo": "Sí" if completo else "No",
                "Calculado": fecha}
        # mismo orden de motivos que resumen_lote()
        huecos_art = sorted(por_articulo.get(cod, []),
                            key=lambda h: MOTIVOS_LOTE.index(h[2])
                            if h[2] in MOTIVOS_LOTE else len(MOTIVOS_LOTE))
        faltantes = [{"Articulo": cod, "IdComponente": pieza, "Descripcion": d,
                      "Alerta": ALERTA_DE_MOTIVO.get(motivo, motivo), "Motivo": motivo}
                     for pieza, d, motivo in huecos_art]
        out[cod] = (fila, faltantes)
    return out


def excel_lote(costes, faltantes):
    """El libro costes_lote.xlsx: hojas Costes y Faltantes."""
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
    return buf.getvalue()


def _calcular_lote(lote_id, ids):
    """Hilo de un lote. Va dejando el avance en _lotes[lote_id]."""
    def avance(**cambios):
        with _lotes_lock:
            _lotes[lote_id].update(cambios)

    try:
        publicados = costes_publicados(ids)
        avance(del_catalogo=len(publicados))
        costes, faltantes = [], []
        for i, cod in enumerate(ids, 1):
            if cod in publicados:
                fila, falt = publicados[cod]
            else:
                try:
                    fila, falt = resumen_lote(cod)
                except Exception as ex:                       # noqa: BLE001
                    # un articulo que falla no tumba el lote entero
                    fila = {"IdArticulo": cod, "Descripcion": None,
                            "CosteMaterial": None, "CosteOperacion": None,
                            "CosteTotal": None,
                            "Sin coste": 0, "Sin escandallo": 0, "Sin tiempo": 0,
                            "Sin tipo": 0, "Sin precio": 0, "Completo": "No"}
                    falt = [{"Articulo": cod, "IdComponente": cod, "Descripcion": None,
                             "Alerta": "Sin coste",
                             "Motivo": f"Error al calcular: {str(ex)[:200]}"}]
                fila["Calculado"] = dt.datetime.now().replace(microsecond=0)
            costes.append(fila)
            faltantes.extend(falt)
            avance(hechos=i)
        avance(excel=excel_lote(costes, faltantes), listo=True)
    except Exception as ex:                                   # noqa: BLE001
        avance(error=str(ex)[:300], listo=True)


@app.route("/lote", methods=["GET", "POST"])
def lote():
    """GET: la pagina. POST: recibe el Excel, lanza el calculo en segundo plano
    y devuelve el id con el que la pagina sigue el avance y luego descarga."""
    if request.method == "GET":
        return render_template("lote.html", maximo=LOTE_MAX)

    file = request.files.get("archivo")
    if not file or not file.filename:
        return "Sube un archivo Excel (.xlsx)", 400
    try:
        ids = extraer_ids(file)
    except Exception as ex:
        return f"No se pudo leer el Excel: {ex}", 400
    if not ids:
        return "No se encontraron IDs de artículo en el Excel", 400

    lote_id = uuid.uuid4().hex
    ahora = time.monotonic()
    with _lotes_lock:
        for k in [k for k, v in _lotes.items() if ahora - v["creado"] > LOTE_VIDA_S]:
            del _lotes[k]
        _lotes[lote_id] = {"creado": ahora, "total": min(len(ids), LOTE_MAX),
                           "hechos": 0, "del_catalogo": None,
                           # antes se cortaba en silencio; ahora se dice
                           "recortados": max(0, len(ids) - LOTE_MAX),
                           "listo": False, "error": None, "excel": None}
    threading.Thread(target=_calcular_lote, args=(lote_id, ids[:LOTE_MAX]),
                     daemon=True).start()
    return jsonify(id=lote_id, total=min(len(ids), LOTE_MAX),
                   recortados=max(0, len(ids) - LOTE_MAX))


@app.route("/lote/<lote_id>/estado")
def lote_estado(lote_id):
    with _lotes_lock:
        e = _lotes.get(lote_id)
        if e is None:
            return jsonify(error="Ese lote no existe o ya caducó"), 404
        return jsonify({k: v for k, v in e.items() if k not in ("excel", "creado")})


@app.route("/lote/<lote_id>/excel")
def lote_excel(lote_id):
    with _lotes_lock:
        e = _lotes.get(lote_id)
        contenido = e and e["excel"]
    if not contenido:
        return "Ese lote no existe, no ha terminado o ya caducó", 404
    return send_file(
        io.BytesIO(contenido),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="costes_lote.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
