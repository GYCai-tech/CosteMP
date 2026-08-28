"""Pagina /costes: recalcular el coste del catalogo y servirlo a Excel.

Sustituye a los .bat sueltos. El recalculo se lanza desde la propia aplicacion
y el resultado se sirve como libro de Excel en /costes/datos.xlsx, que es lo
que consume Power Query.

Se sirve por HTTP y no como fichero en disco a proposito: Excel no puede
conectar a PostgreSQL sin instalar Npgsql en cada puesto, y un fichero en el
escritorio de una persona no lo puede refrescar nadie mas.
"""
import io
import threading
import datetime as dt

import pandas as pd
from flask import Blueprint, jsonify, render_template, request, send_file
from sqlalchemy import text

from db_pg import get_pg_engine

bp = Blueprint("costes", __name__)

# Estado del recalculo en curso. Solo puede haber uno: dos a la vez se pisarian
# al publicar, y ademas cargarian el ERP el doble.
_estado = {"activo": False, "hechos": 0, "total": 0, "modo": None,
           "inicio": None, "resultado": None, "error": None}
_lock = threading.Lock()

HOJAS = {"Costes": "core.v_coste_hoja10",
         "Pendientes": "core.v_coste_pendiente",
         "Resumen": "core.v_coste_resumen"}


def _resumen():
    pg = get_pg_engine()
    with pg.connect() as c:
        filas = list(c.execute(text(
            'SELECT "Indicador", "Valor" FROM core.v_coste_resumen')))
        ultimo = list(c.execute(text(
            "SELECT inicio, articulos, errores, duracion_seg, origen "
            "FROM core.log_coste_recalculo WHERE resultado='ok' "
            "ORDER BY id DESC LIMIT 1")))
    return {"indicadores": [{"nombre": n, "valor": v} for n, v in filas],
            "ultimo": dict(zip(("inicio", "articulos", "errores", "duracion",
                                "origen"), ultimo[0])) if ultimo else None}


def _trabajar(solo_pendientes):
    """Hilo del recalculo. Actualiza _estado para que la pagina pinte el avance."""
    # el import va aqui dentro: exportar_costes importa de app.py, y app.py
    # registra este blueprint al arrancar -> arriba seria circular
    import exportar_costes

    def progreso(hechos, total):
        with _lock:
            _estado["hechos"], _estado["total"] = hechos, total
    try:
        r = exportar_costes.ejecutar(get_pg_engine(),
                                     solo_pendientes=solo_pendientes,
                                     origen="manual", progreso=progreso)
        with _lock:
            _estado["resultado"] = r
    except Exception as ex:                                   # noqa: BLE001
        with _lock:
            _estado["error"] = str(ex)[:500]
    finally:
        with _lock:
            _estado["activo"] = False


# strict_slashes=False para que /costes y /costes/ valgan las dos: escribir la
# barra final al teclear la direccion es lo normal y devolvia un 404 seco.
@bp.route("/costes", strict_slashes=False)
def pagina():
    return render_template("costes.html", **_resumen())


@bp.route("/costes/recalcular", methods=["POST"])
def recalcular():
    modo = request.args.get("modo", "pendientes")
    with _lock:
        if _estado["activo"]:
            return jsonify(error="Ya hay un recalculo en curso"), 409
        _estado.update(activo=True, hechos=0, total=0, modo=modo,
                       inicio=dt.datetime.now().isoformat(timespec="seconds"),
                       resultado=None, error=None)
    threading.Thread(target=_trabajar, args=(modo == "pendientes",),
                     daemon=True).start()
    return jsonify(ok=True, modo=modo)


@bp.route("/costes/estado")
def estado():
    with _lock:
        e = dict(_estado)
    if not e["activo"]:
        e.update(_resumen())
    return jsonify(e)


@bp.route("/costes/pendientes")
def pendientes():
    """La lista de piezas que faltan, para pintarla en la propia pagina.

    Se sirve entera (~200 filas) y se filtra en el navegador: paginar en el
    servidor para doscientas filas seria complicarlo sin ganar nada.
    """
    pg = get_pg_engine()
    with pg.connect() as c:
        filas = [dict(r._mapping) for r in c.execute(text(
            'SELECT "IdPieza", "Descripcion", "Tipo aprovisionamiento", '
            '       "Familia", "Motivo", "Articulos que desbloquea", '
            '       "Articulos afectados" '
            'FROM core.v_coste_pendiente'))]
    return jsonify(filas)


@bp.route("/catalogo", strict_slashes=False)
def catalogo():
    return render_template("catalogo.html")


@bp.route("/catalogo/datos")
def catalogo_datos():
    """Los articulos del catalogo con su coste desglosado."""
    with get_pg_engine().connect() as c:
        filas = [{"id": r[0], "desc": r[1], "familia": r[2], "mp": r[3],
                  "op": r[4], "total": r[5], "sin_escandallo": r[6],
                  "sin_tipo": r[7], "sin_precio": r[8], "sin_tiempo": r[9],
                  "error": r[10]}
                 for r in c.execute(text("""
            SELECT idarticulo, descripcion, familia,
                   coste_material, coste_operacion, coste_total,
                   COALESCE(piezas_sin_escandallo,0), COALESCE(piezas_sin_tipo,0),
                   COALESCE(piezas_sin_precio,0), COALESCE(piezas_sin_tiempo,0),
                   error
            FROM core.dim_coste_escandallo c
            ORDER BY idarticulo"""))]
    return jsonify(filas)


@bp.route("/catalogo/<idarticulo>/huecos")
def catalogo_huecos(idarticulo):
    """Los problemas de UN articulo. Se pide al desplegar la fila, no antes:
    cargar los ~700 huecos del catalogo entero de golpe no lo mira nadie."""
    with get_pg_engine().connect() as c:
        filas = [{"pieza": r[0], "desc": r[1], "motivo": r[2],
                  "tipo": r[3], "familia": r[4]}
                 for r in c.execute(text("""
            SELECT idpieza, descripcion, motivo,
                   COALESCE(tipo_aprovisionamiento,'SIN TIPO'),
                   COALESCE(familia,'(Sin definir)')
            FROM core.fact_coste_hueco WHERE idarticulo = :a
            ORDER BY CASE motivo
                       WHEN 'Sin escandallo activo' THEN 1
                       WHEN 'Sin tipo de aprovisionamiento' THEN 2
                       WHEN 'Sin precio de compra' THEN 3
                       ELSE 4 END, idpieza"""), {"a": idarticulo})]
    return jsonify(filas)


@bp.route("/costes/datos.xlsx")
def datos():
    """El libro que lee Power Query. Se genera al vuelo desde las vistas."""
    pg = get_pg_engine()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xls:
        for hoja, vista in HOJAS.items():
            df = pd.read_sql(text(f"SELECT * FROM {vista}"), pg)
            df = df.drop(columns=["orden"], errors="ignore")
            df.to_excel(xls, sheet_name=hoja, index=False)
            ws = xls.sheets[hoja]
            for i, col in enumerate(df.columns, start=1):
                ancho = max([len(str(col))] +
                            [len(str(v)) for v in df[col].head(200)]) if len(df) \
                        else len(str(col))
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = \
                    min(55, ancho + 2)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="costes_catalogo.xlsx",
    )
