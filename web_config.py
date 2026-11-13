"""Pagina /configuracion: los parametros del calculo del coste.

Cada parametro se declara en PARAMETROS y la pagina los pinta todos, agrupados.
Para anadir uno nuevo basta con una entrada mas ahi (y leerlo donde se use);
la pagina, el guardado y el historial no cambian.

Un parametro con "erp" no se guarda aqui: se lee del ERP y la pagina solo lo
ensena, con quien lo cambio alli y cuando. Es el caso del coste/hora de
operario (Empleados_Categorias, ver desglose.coste_hora_erp): se cambia en el
ERP y la aplicacion lo sigue sola.

Los valores viven en Postgres (core.cfg_coste_parametro, ver
sql/06_parametros_coste.sql) y no en el .env a proposito: asi los leen igual la
web y la pasada del vigilante, que son dos contenedores, y cambiarlos no pide
reiniciar nada.

La aplicacion no tiene usuarios. Si el .env define CONFIG_CLAVE, guardar pide
esa clave; si no, cualquiera que abra la pagina puede cambiar los valores. Cada
cambio queda en core.log_coste_parametro con el equipo desde el que se hizo.
"""
import hmac
import os
import pathlib

from flask import Blueprint, jsonify, render_template, request
from sqlalchemy import text

import db_pg
import desglose

bp = Blueprint("configuracion", __name__)

SQL_PARAMETROS = pathlib.Path(__file__).with_name("sql") / "06_parametros_coste.sql"

# min/max no son reglas de negocio: solo frenan un error al teclear (170 por
# 17,0, o un 0), que cambiaria el coste de todo el catalogo en la siguiente
# pasada. al_cambiar: lo que hay que olvidar para que la web lo aplique ya.
# erp: los que mandan desde el ERP: donde se cambian alli y como leerlos.
PARAMETROS = {
    "coste_hora_operario": {
        "grupo": "Cálculo del coste",
        "nombre": "Coste de la hora de operario",
        "unidad": "€/h",
        "defecto": desglose.COSTE_HORA_DEFECTO,
        "erp": {"donde": "Empleados > Categorías, campo Coste de la categoría "
                         f"{desglose.ID_CATEGORIA_OPERARIO}",
                "leer": desglose.coste_hora_erp,
                "olvidar": desglose.olvidar_coste_hora},
    },
}


def _editables():
    return {k: p for k, p in PARAMETROS.items() if "erp" not in p}


def _clave_config():
    return os.getenv("CONFIG_CLAVE") or None


def leer_config():
    """El valor vigente de cada parametro y los ultimos cambios de todos."""
    valores = {k: {"valor": p["defecto"], "fecha_cambio": None, "por_defecto": True}
               for k, p in PARAMETROS.items()}
    out = {"valores": valores, "historial": [], "error": None}
    # los del ERP se vuelven a leer al abrir la pagina: quien acaba de cambiarlo
    # alli lo ve aqui, y la web lo aplica ya, sin esperar los 10 minutos
    for clave, p in PARAMETROS.items():
        if "erp" in p:
            p["erp"]["olvidar"]()
            d = p["erp"]["leer"]()
            valores[clave] = {"valor": d["valor"], "fecha_cambio": d["fecha"],
                              "por_defecto": d["error"] is not None,
                              "usuario": d["usuario"], "categoria": d["categoria"],
                              "error_erp": d["error"]}
    try:
        with db_pg.get_pg_engine().connect() as c:
            if c.execute(text("SELECT to_regclass('core.cfg_coste_parametro')")).scalar():
                for clave, valor, fecha in c.execute(text(
                        "SELECT clave, valor, fecha_cambio FROM core.cfg_coste_parametro")):
                    if clave in _editables():
                        valores[clave] = {"valor": float(valor), "fecha_cambio": fecha,
                                          "por_defecto": False}
            if c.execute(text("SELECT to_regclass('core.log_coste_parametro')")).scalar():
                out["historial"] = [
                    {"fecha": r[0], "clave": r[1],
                     "nombre": PARAMETROS.get(r[1], {}).get("nombre", r[1]),
                     "unidad": PARAMETROS.get(r[1], {}).get("unidad", ""),
                     "anterior": float(r[2]) if r[2] is not None else None,
                     "nuevo": float(r[3]), "desde": r[4]}
                    for r in c.execute(text(
                        "SELECT fecha, clave, valor_anterior, valor_nuevo, desde "
                        "FROM core.log_coste_parametro ORDER BY id DESC LIMIT 20"))]
    except Exception as ex:                                   # noqa: BLE001
        out["error"] = str(ex)[:300]
    return out


def guardar_parametro(clave, valor, desde):
    """Cambia un parametro y deja el cambio en el historial, en una transaccion.

    Aplica antes sql/06_parametros_coste.sql (idempotente): asi la primera vez
    que alguien guarda se crean las tablas, sin tener que acordarse de lanzar
    exportar_costes.py --crear en el servidor. Devuelve el valor anterior.
    """
    p = PARAMETROS[clave]
    with db_pg.get_pg_engine().begin() as c:
        c.execute(text(SQL_PARAMETROS.read_text(encoding="utf-8")))
        anterior = c.execute(text(
            "SELECT valor FROM core.cfg_coste_parametro WHERE clave = :k FOR UPDATE"),
            {"k": clave}).scalar()
        # upsert: un parametro nuevo no tiene fila hasta que alguien lo guarda
        c.execute(text(
            "INSERT INTO core.cfg_coste_parametro (clave, valor, unidad) "
            "VALUES (:k, :v, :u) "
            "ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor, "
            "                                  fecha_cambio = now()"),
            {"k": clave, "v": valor, "u": p["unidad"]})
        c.execute(text(
            "INSERT INTO core.log_coste_parametro "
            "(clave, valor_anterior, valor_nuevo, desde) VALUES (:k, :a, :v, :d)"),
            {"k": clave, "a": anterior, "v": valor, "d": desde})
    if p.get("al_cambiar"):
        p["al_cambiar"]()
    return float(anterior) if anterior is not None else None


@bp.route("/configuracion", strict_slashes=False)
def pagina():
    grupos = {}
    for clave, p in PARAMETROS.items():
        grupos.setdefault(p["grupo"], []).append(dict(p, clave=clave))
    return render_template("configuracion.html", grupos=grupos, cfg=leer_config(),
                           hay_editables=bool(_editables()),
                           pide_clave=_clave_config() is not None and bool(_editables()))


@bp.route("/configuracion/<clave>", methods=["POST"])
def cambiar(clave):
    p = PARAMETROS.get(clave)
    if p is None:
        return jsonify(error="Ese parámetro no existe"), 404
    if "erp" in p:
        return jsonify(error="Este valor se cambia en el ERP: " + p["erp"]["donde"]), 409
    datos = request.get_json(silent=True) or {}
    secreto = _clave_config()
    if secreto and not hmac.compare_digest(str(datos.get("clave", "")), secreto):
        return jsonify(error="La clave no es correcta"), 403

    # se acepta la coma decimal, que es como lo va a escribir todo el mundo
    try:
        valor = round(float(str(datos.get("valor", "")).strip().replace(",", ".")), 4)
    except ValueError:
        return jsonify(error="Escribe un número"), 400
    if not p["min"] <= valor <= p["max"]:
        return jsonify(error=f"Tiene que estar entre {p['min']:g} y "
                             f"{p['max']:g} {p['unidad']}"), 400

    try:
        anterior = guardar_parametro(clave, valor, request.remote_addr)
    except Exception as ex:                                   # noqa: BLE001
        return jsonify(error="No se pudo guardar: " + str(ex)[:300]), 500
    return jsonify(ok=True, clave=clave, valor=valor, anterior=anterior)
