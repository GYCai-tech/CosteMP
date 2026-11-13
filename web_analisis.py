"""Paginas que sacan informacion del detalle por pieza (sql/07_detalle_coste.sql).

    /evolucion   como ha cambiado el coste de cada articulo y que pieza lo explica
    /piezas      en que articulos se usa cada pieza y que pasa si sube su precio

Solo leen Postgres: nada de esto va al ERP. El detalle lo escribe la pasada
(exportar_costes._guardar_detalle), asi que hablan del catalogo, no de
cualquier articulo del ERP, y el historico empieza el dia que se desplego.

Las ventas (unidades de los ultimos 12 meses, sin rectificativas) solo sirven
para pasar un euro por unidad a euros al ano. Es una aproximacion: el coste de
hoy aplicado a lo vendido el ultimo ano.
"""
import datetime as dt

from flask import Blueprint, jsonify, render_template, request
from sqlalchemy import text

import db_pg

bp = Blueprint("analisis", __name__)

# lo que se ha vendido de cada articulo en el ultimo ano
SQL_VENTAS_12M = """
    ventas AS (
        SELECT vl.idarticulo, SUM(vl.cantidad) AS uds
        FROM core.fact_ventas_lineas vl
        JOIN core.fact_ventas f ON f.idfactura = vl.idfactura
        WHERE NOT f.es_rectificativa AND f.fechafact >= current_date - 365
        GROUP BY vl.idarticulo)"""


def _f(v):
    return float(v) if v is not None else None


def _es(v, fmt="g"):
    """Numero con coma decimal, como se escribe aqui."""
    return format(v, fmt).replace(".", ",")


def _hay_historico(c):
    return c.execute(text("SELECT to_regclass('core.hist_coste_articulo')")).scalar() \
        is not None


def _margen(precio, coste, completo):
    """Margen sobre tarifa, en %. None si no se puede afirmar: sin tarifa, o con
    el coste incompleto (le faltan euros, asi que el margen real seria menor;
    misma regla que core.v_rentabilidad)."""
    if not completo or not precio or coste is None:
        return None
    return round(100.0 * (precio - coste) / precio, 1)


# ---------------------------------------------------------------------------
# /evolucion
# ---------------------------------------------------------------------------

@bp.route("/evolucion", strict_slashes=False)
def evolucion():
    return render_template("evolucion.html")


@bp.route("/evolucion/datos")
def evolucion_datos():
    """Coste de cada articulo del catalogo a fecha `desde` y ahora.

    Si un articulo no tiene calculo tan antiguo, se compara con el primero que
    tenga (y se dice de cuando es)."""
    try:
        desde = dt.date.fromisoformat(request.args.get("desde", ""))
    except ValueError:
        desde = dt.date.today() - dt.timedelta(days=30)
    with db_pg.get_pg_engine().connect() as c:
        if not _hay_historico(c):
            return jsonify(primero=None, filas=[])
        primero = c.execute(text(
            "SELECT min(fecha_calculo) FROM core.hist_coste_articulo")).scalar()
        filas = c.execute(text(f"""
            WITH ahora AS (
                SELECT DISTINCT ON (idarticulo) * FROM core.hist_coste_articulo
                ORDER BY idarticulo, fecha_calculo DESC),
            antes AS (
                SELECT DISTINCT ON (idarticulo) * FROM core.hist_coste_articulo
                WHERE fecha_calculo < CAST(:desde AS date) + 1
                ORDER BY idarticulo, fecha_calculo DESC),
            primero AS (
                SELECT DISTINCT ON (idarticulo) * FROM core.hist_coste_articulo
                ORDER BY idarticulo, fecha_calculo),
            {SQL_VENTAS_12M}
            SELECT a.idarticulo, e.descripcion, COALESCE(e.familia, '(Sin definir)'),
                   COALESCE(b.fecha_calculo, p.fecha_calculo),
                   COALESCE(b.coste_total, p.coste_total), COALESCE(b.completo, p.completo),
                   a.fecha_calculo, a.coste_total, a.completo, v.uds
            FROM ahora a
            JOIN core.dim_coste_escandallo e ON e.idarticulo = a.idarticulo
            JOIN primero p ON p.idarticulo = a.idarticulo
            LEFT JOIN antes b ON b.idarticulo = a.idarticulo
            LEFT JOIN ventas v ON v.idarticulo = a.idarticulo"""), {"desde": desde}).all()
    out = []
    for (art, desc, fam, f0, c0, ok0, f1, c1, ok1, uds) in filas:
        c0, c1, uds = _f(c0), _f(c1), _f(uds) or 0.0
        dif = (c1 - c0) if c0 is not None and c1 is not None else None
        out.append({"id": art, "desc": desc, "familia": fam,
                    "fecha_antes": f0.isoformat(timespec="minutes"), "antes": c0,
                    "completo_antes": ok0,
                    "fecha_ahora": f1.isoformat(timespec="minutes"), "ahora": c1,
                    "completo": ok1, "dif": dif,
                    "pct": round(100.0 * dif / c0, 2) if dif is not None and c0 else None,
                    "uds_12m": uds,
                    "impacto_anual": round(dif * uds, 2) if dif is not None else None})
    return jsonify(primero=primero.isoformat(timespec="minutes") if primero else None,
                   filas=out)


@bp.route("/evolucion/<idarticulo>")
def evolucion_articulo(idarticulo):
    """Que piezas explican el cambio de un articulo entre `desde` y ahora."""
    try:
        desde = dt.date.fromisoformat(request.args.get("desde", ""))
    except ValueError:
        desde = dt.date.today() - dt.timedelta(days=30)
    with db_pg.get_pg_engine().connect() as c:
        if not _hay_historico(c):
            return jsonify(piezas=[])
        p = {"a": idarticulo, "desde": desde}
        # la foto de "antes" es la del articulo a esa fecha (o la primera)...
        f_antes = c.execute(text("""
            SELECT COALESCE(
                (SELECT max(fecha_calculo) FROM core.hist_coste_articulo
                  WHERE idarticulo = :a AND fecha_calculo < CAST(:desde AS date) + 1),
                (SELECT min(fecha_calculo) FROM core.hist_coste_articulo
                  WHERE idarticulo = :a))"""), p).scalar()
        if f_antes is None:
            return jsonify(piezas=[])
        horas = dict(c.execute(text("""
            SELECT 'antes', coste_hora FROM core.hist_coste_articulo
             WHERE idarticulo = :a AND fecha_calculo = :f
            UNION ALL
            SELECT 'ahora', coste_hora FROM (
                SELECT coste_hora FROM core.hist_coste_articulo WHERE idarticulo = :a
                ORDER BY fecha_calculo DESC LIMIT 1) x"""), {"a": idarticulo, "f": f_antes}).all())

        # ...y su detalle por pieza es el ultimo guardado en o antes de esa foto
        # (el detalle solo se guarda cuando cambia)
        def detalle(hasta):
            return {r[0]: r for r in c.execute(text("""
                SELECT idpieza, descripcion, cantidad, coste_material, tiempo_min,
                       coste_operacion
                FROM core.hist_coste_linea
                WHERE idarticulo = :a AND fecha_calculo = (
                    SELECT max(fecha_calculo) FROM core.hist_coste_linea
                    WHERE idarticulo = :a AND fecha_calculo <= :h)"""),
                {"a": idarticulo, "h": hasta})}
        antes = detalle(f_antes)
        ahora = detalle(dt.datetime(9999, 1, 1))

    h0, h1 = _f(horas.get("antes")), _f(horas.get("ahora"))
    piezas = []
    for pieza in sorted(set(antes) | set(ahora)):
        a, b = antes.get(pieza), ahora.get(pieza)
        cant0, mat0, t0, op0 = (_f(a[2]) or 0, _f(a[3]) or 0, _f(a[4]) or 0, _f(a[5]) or 0) \
            if a else (0, 0, 0, 0)
        cant1, mat1, t1, op1 = (_f(b[2]) or 0, _f(b[3]) or 0, _f(b[4]) or 0, _f(b[5]) or 0) \
            if b else (0, 0, 0, 0)
        dif = (mat1 + op1) - (mat0 + op0)
        if abs(dif) < 0.00005:
            continue
        motivos = []
        if not a:
            motivos.append("pieza nueva en el escandallo")
        elif not b:
            motivos.append("ya no está en el escandallo")
        else:
            if abs(cant1 - cant0) > 1e-6:
                motivos.append(f"cantidad {_es(cant0)} → {_es(cant1)}")
            if mat0 and mat1 and cant0 and cant1:
                u0, u1 = mat0 / cant0, mat1 / cant1
                if abs(u1 - u0) > 1e-9:
                    motivos.append(f"precio {_es(100.0 * (u1 - u0) / u0, '+.1f')} %")
            if abs(t1 - t0) > 1e-6:
                motivos.append(f"tiempo {_es(t0)} → {_es(t1)} min")
            if op0 and op1 and abs(t1 - t0) <= 1e-6 and h0 and h1 and h0 != h1:
                motivos.append(f"coste/hora {_es(h0)} → {_es(h1)} €/h")
        piezas.append({"pieza": pieza, "desc": (b or a)[1], "antes": mat0 + op0,
                       "ahora": mat1 + op1, "dif": round(dif, 6),
                       "motivo": ", ".join(motivos) or "redondeo"})
    piezas.sort(key=lambda p: -abs(p["dif"]))
    return jsonify(piezas=piezas)


# ---------------------------------------------------------------------------
# /piezas
# ---------------------------------------------------------------------------

@bp.route("/piezas", strict_slashes=False)
def piezas():
    return render_template("piezas.html")


@bp.route("/piezas/buscar")
def piezas_buscar():
    """Piezas compradas del catalogo, ordenadas por lo que pesan en lo vendido
    el ultimo ano (coste de la pieza en cada articulo x unidades vendidas)."""
    q = request.args.get("q", "").strip()
    with db_pg.get_pg_engine().connect() as c:
        if c.execute(text("SELECT to_regclass('core.fact_coste_linea')")).scalar() is None:
            return jsonify([])
        filas = c.execute(text(f"""
            WITH {SQL_VENTAS_12M}
            SELECT l.idpieza, max(l.descripcion), max(l.tipo), max(l.unidad),
                   max(l.precio), count(*),
                   SUM(l.coste_material * COALESCE(v.uds, 0))
            FROM core.fact_coste_linea l
            LEFT JOIN ventas v ON v.idarticulo = l.idarticulo
            WHERE l.coste_material > 0
              AND (:q = '' OR l.idpieza ILIKE :patron OR l.descripcion ILIKE :patron)
            GROUP BY l.idpieza
            ORDER BY 7 DESC NULLS LAST, 6 DESC
            LIMIT 50"""), {"q": q, "patron": f"%{q}%"}).all()
    return jsonify([{"pieza": r[0], "desc": r[1], "tipo": r[2], "unidad": r[3],
                     "precio": _f(r[4]), "articulos": r[5],
                     "en_ventas": round(_f(r[6]) or 0, 2)} for r in filas])


def _usos(c, piezas):
    """Cada (articulo, pieza) de `piezas`, con lo necesario para simular."""
    return c.execute(text(f"""
        WITH {SQL_VENTAS_12M}
        SELECT l.idarticulo, e.descripcion, l.idpieza, l.cantidad, l.unidad,
               l.coste_material, e.coste_total, e.precio_tarifa, e.completo,
               COALESCE(v.uds, 0)
        FROM core.fact_coste_linea l
        JOIN core.dim_coste_escandallo e ON e.idarticulo = l.idarticulo
        LEFT JOIN ventas v ON v.idarticulo = l.idarticulo
        WHERE l.idpieza = ANY(CAST(:p AS varchar[])) AND l.coste_material > 0
        ORDER BY l.coste_material * COALESCE(v.uds, 0) DESC, l.idarticulo"""),
        {"p": list(piezas)}).all()


@bp.route("/piezas/<idpieza>")
def pieza_usos(idpieza):
    """En que articulos entra la pieza y cuanto pesa en cada uno."""
    with db_pg.get_pg_engine().connect() as c:
        filas = _usos(c, [idpieza])
    return jsonify([{"id": a, "desc": d, "cantidad": _f(cant), "unidad": u,
                     "coste_pieza": _f(mat), "coste_total": _f(tot),
                     "peso": round(100.0 * float(mat) / float(tot), 1) if tot else None,
                     "margen": _margen(_f(pv), _f(tot), ok), "uds_12m": _f(uds)}
                    for a, d, _, cant, u, mat, tot, pv, ok, uds in filas])


@bp.route("/piezas/simular", methods=["POST"])
def simular():
    """Sube (o baja) un % el coste de unas piezas y recalcula los articulos que
    las usan: coste, margen sobre tarifa e impacto en euros al ano.

    Escala el coste de material que ya tiene cada pieza en cada articulo, asi
    que respeta descuentos, mermas y unidades sin tener que repetir el calculo.
    """
    cambios = {}
    for ch in (request.get_json(silent=True) or {}).get("cambios", []):
        try:
            pct = float(str(ch.get("pct", "")).replace(",", "."))
        except ValueError:
            return jsonify(error="Porcentaje no válido"), 400
        if ch.get("pieza") and -100 < pct <= 1000:
            cambios[str(ch["pieza"])] = pct
    if not cambios:
        return jsonify(error="Añade al menos una pieza con su %"), 400

    with db_pg.get_pg_engine().connect() as c:
        filas = _usos(c, cambios)

    arts = {}
    for a, d, pieza, _, _, mat, tot, pv, ok, uds in filas:
        r = arts.setdefault(a, {"id": a, "desc": d, "coste": _f(tot), "precio": _f(pv),
                                "completo": ok, "uds_12m": _f(uds), "delta": 0.0})
        r["delta"] += float(mat) * cambios[pieza] / 100.0

    out, impacto, cruzan = [], 0.0, 0
    for r in arts.values():
        nuevo = (r["coste"] or 0) + r["delta"]
        m0, m1 = _margen(r["precio"], r["coste"], r["completo"]), \
            _margen(r["precio"], nuevo, r["completo"])
        anual = r["delta"] * (r["uds_12m"] or 0)
        impacto += anual
        if m0 is not None and m1 is not None and m0 >= 0 > m1:
            cruzan += 1
        out.append(dict(r, delta=round(r["delta"], 4), nuevo=round(nuevo, 4),
                        margen=m0, margen_nuevo=m1, impacto_anual=round(anual, 2)))
    out.sort(key=lambda r: -abs(r["impacto_anual"]))
    return jsonify(articulos=out, resumen={
        "articulos": len(out), "impacto_anual": round(impacto, 2),
        "bajo_coste": cruzan})
