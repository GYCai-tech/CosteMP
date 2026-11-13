"""El coste/hora de operario (leido del ERP) y el lote en segundo plano.

Como test_menos_erp.py, NO conectan con el ERP ni con Postgres: se sustituyen
por dobles para comprobar de donde sale cada dato.
"""
import datetime as dt
import io
import time
from decimal import Decimal

import pandas as pd
import pytest

import app as app_mod
import db_pg
import desglose


# --- coste de la hora de operario (del ERP) ----------------------------------

class _Conexion:
    def __init__(self, fila, llamadas):
        self.fila, self.llamadas = fila, llamadas

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def execute(self, sql, params=None):
        self.llamadas.append(params)
        fila = self.fila

        class Res:
            def first(self):
                return fila
        return Res()


FECHA_ERP = dt.datetime(2026, 11, 10, 9, 30)


def _erp_con(monkeypatch, coste, fila=True):
    """El ERP sustituido: Empleados_Categorias devuelve `coste` para Operario
    (o ninguna fila, con fila=False)."""
    llamadas = []
    datos = (coste, "Operario", "gilberto", FECHA_ERP) if fila else None

    class Motor:
        def connect(self):
            return _Conexion(datos, llamadas)
    monkeypatch.setattr(desglose, "get_engine", lambda: Motor())
    return llamadas


@pytest.fixture(autouse=True)
def _olvidar():
    desglose.olvidar_cache()
    yield
    desglose.olvidar_cache()


def _arbol_de_30_min():
    """Un articulo fabricado (30 min/pieza, media de bonos) con una hoja de 3 EUR."""
    df = pd.DataFrame([{"Ruta": "|A|B|", "IdArticulo": "B", "Componente": "CHAPA",
                        "Cant": 2.0, "Coste": 3.0, "Unidad": "kg",
                        "TipoCompra": "MATERIA PRIMA", "PrecioCompra": 1.5}])
    return desglose.construir_arbol(df, "A", "ART A", tiempo_raiz=30.0, medio_raiz=1)


def test_la_hora_sale_de_la_categoria_operario_del_erp(monkeypatch):
    llamadas = _erp_con(monkeypatch, Decimal("20.00000000000000"))
    arbol = _arbol_de_30_min()
    assert arbol["coste_hora"] == 20.0
    assert arbol["coste_op_total"] == pytest.approx(10.0)       # 30 min a 20 EUR/h
    assert arbol["coste_mat"] == pytest.approx(3.0)             # el material no cambia
    assert llamadas == [{"cat": desglose.ID_CATEGORIA_OPERARIO}]
    info = desglose.coste_hora_erp()
    assert (info["categoria"], info["usuario"], info["fecha"], info["error"]) == (
        "Operario", "gilberto", FECHA_ERP, None)


def test_sin_erp_la_hora_vale_17(monkeypatch):
    def caido():
        raise ConnectionError("el ERP no responde")
    monkeypatch.setattr(desglose, "get_engine", caido)
    arbol = _arbol_de_30_min()
    assert arbol["coste_hora"] == 17.0
    assert arbol["coste_op_total"] == pytest.approx(8.5)        # 30 min a 17 EUR/h
    assert arbol["coste_total"] == pytest.approx(11.5)
    assert "ERP" in desglose.coste_hora_erp()["error"]


@pytest.mark.parametrize("coste", [None, 0])
def test_si_el_erp_no_tiene_coste_la_hora_vale_17(monkeypatch, coste):
    _erp_con(monkeypatch, coste)
    assert desglose.coste_hora_operario() == 17.0
    assert "no tiene coste" in desglose.coste_hora_erp()["error"]


def test_sin_la_categoria_en_el_erp_la_hora_vale_17(monkeypatch):
    _erp_con(monkeypatch, None, fila=False)
    assert desglose.coste_hora_operario() == 17.0
    assert "No existe" in desglose.coste_hora_erp()["error"]


def test_la_hora_se_lee_una_vez_y_no_por_articulo(monkeypatch):
    llamadas = _erp_con(monkeypatch, 20)
    for _ in range(5):
        _arbol_de_30_min()
    assert len(llamadas) == 1


def test_un_cambio_en_el_erp_se_ve_al_olvidar_lo_recordado(monkeypatch):
    _erp_con(monkeypatch, 17)
    assert desglose.coste_hora_operario() == 17.0
    _erp_con(monkeypatch, 19.5)                     # alguien lo cambia en el ERP
    assert desglose.coste_hora_operario() == 17.0   # recordado hasta 10 minutos
    desglose.olvidar_coste_hora()
    assert desglose.coste_hora_operario() == 19.5


# --- el lote, en segundo plano ------------------------------------------------

def _excel_con(ids):
    buf = io.BytesIO()
    pd.DataFrame({"IdArticulo": ids}).to_excel(buf, index=False)
    buf.seek(0)
    return buf


def _fila(cod, total):
    return {"IdArticulo": cod, "Descripcion": f"ART {cod}", "CosteMaterial": total,
            "CosteOperacion": 0.0, "CosteTotal": total, "Sin coste": 0,
            "Sin escandallo": 0, "Sin tiempo": 0, "Sin tipo": 0, "Sin precio": 0,
            "Completo": "Sí"}


def _esperar(cliente, lote_id):
    for _ in range(200):
        e = cliente.get(f"/lote/{lote_id}/estado").get_json()
        if e["listo"]:
            return e
        time.sleep(0.02)
    raise AssertionError("el lote no termino")


@pytest.fixture
def lote(monkeypatch):
    """La ruta /lote con el coste publicado y el ERP sustituidos."""
    al_erp = []
    publicado = {"10201001": (dict(_fila("10201001", 5.0), Calculado=pd.Timestamp("2026-10-23 05:04")),
                              [{"Articulo": "10201001", "IdComponente": "X1",
                                "Descripcion": "TORNILLO", "Alerta": "Sin tiempo de operación",
                                "Motivo": "Sin tiempo de operación"}])}

    def resumen(cod):
        al_erp.append(cod)
        if cod == "00000666":
            raise RuntimeError("el ERP se ha quejado")
        return _fila(cod, 7.0), []

    monkeypatch.setattr(app_mod, "costes_publicados",
                        lambda ids: {c: v for c, v in publicado.items() if c in ids})
    monkeypatch.setattr(app_mod, "resumen_lote", resumen)
    return app_mod.app.test_client(), al_erp


def test_el_lote_coge_el_catalogo_publicado_y_solo_va_al_erp_por_el_resto(lote):
    cliente, al_erp = lote
    r = cliente.post("/lote", data={"archivo": (_excel_con(["10201001", "12101021"]),
                                                "lote.xlsx")})
    assert r.status_code == 200
    e = _esperar(cliente, r.get_json()["id"])
    assert e["error"] is None and e["hechos"] == 2 and e["del_catalogo"] == 1
    assert al_erp == ["12101021"]

    x = cliente.get(f"/lote/{r.get_json()['id']}/excel")
    assert x.status_code == 200
    costes = pd.read_excel(io.BytesIO(x.data), sheet_name="Costes", dtype={"IdArticulo": str})
    faltan = pd.read_excel(io.BytesIO(x.data), sheet_name="Faltantes", dtype=str)
    assert costes["IdArticulo"].tolist() == ["10201001", "12101021"]   # el orden del Excel
    assert costes["CosteTotal"].tolist() == [5.0, 7.0]
    assert costes["Calculado"].notna().all()            # siempre se dice de cuando es
    assert faltan["IdComponente"].tolist() == ["X1"]


def test_un_articulo_que_falla_no_tumba_el_lote(lote):
    cliente, _ = lote
    r = cliente.post("/lote", data={"archivo": (_excel_con(["666", "12101021"]), "l.xlsx")})
    e = _esperar(cliente, r.get_json()["id"])
    assert e["error"] is None
    x = cliente.get(f"/lote/{r.get_json()['id']}/excel")
    costes = pd.read_excel(io.BytesIO(x.data), sheet_name="Costes", dtype={"IdArticulo": str})
    faltan = pd.read_excel(io.BytesIO(x.data), sheet_name="Faltantes", dtype=str)
    assert costes["Completo"].tolist() == ["No", "Sí"]
    assert faltan["Motivo"].str.startswith("Error al calcular").tolist() == [True]


def test_el_lote_avisa_si_se_recorta(lote, monkeypatch):
    cliente, _ = lote
    monkeypatch.setattr(app_mod, "LOTE_MAX", 2)
    r = cliente.post("/lote", data={"archivo": (_excel_con(["1", "2", "3"]), "l.xlsx")})
    assert r.get_json()["total"] == 2 and r.get_json()["recortados"] == 1
    _esperar(cliente, r.get_json()["id"])


def test_un_lote_que_no_existe_da_404(lote):
    cliente, _ = lote
    assert cliente.get("/lote/nada/estado").status_code == 404
    assert cliente.get("/lote/nada/excel").status_code == 404


# --- la pagina /configuracion -------------------------------------------------

@pytest.fixture
def config(monkeypatch):
    """/configuracion con el ERP sustituido (17 EUR/h), sin Postgres, y con un
    parametro editable de prueba, porque hoy los de verdad vienen todos del ERP.
    Se apunta lo que se guardaria."""
    import web_config
    guardados, aplicados = [], []
    _erp_con(monkeypatch, 17)

    def sin_postgres():
        raise ConnectionError("Postgres no responde")
    monkeypatch.setattr(db_pg, "get_pg_engine", sin_postgres)
    monkeypatch.setitem(web_config.PARAMETROS, "prueba", {
        "grupo": "Pruebas", "nombre": "Parámetro de prueba", "unidad": "€/h",
        "defecto": 10.0, "min": 1.0, "max": 200.0,
        "al_cambiar": lambda: aplicados.append(1)})

    def guardar(clave, valor, desde):
        guardados.append((clave, valor))
        web_config.PARAMETROS[clave]["al_cambiar"]()
        return 10.0
    monkeypatch.setattr(web_config, "guardar_parametro", guardar)
    monkeypatch.delenv("CONFIG_CLAVE", raising=False)
    return app_mod.app.test_client(), guardados, aplicados, monkeypatch


def test_la_pagina_pinta_todos_los_parametros(config):
    import web_config
    cliente = config[0]
    html = cliente.get("/configuracion").get_data(as_text=True)
    for p in web_config.PARAMETROS.values():
        assert p["nombre"] in html


def test_la_pagina_ensena_la_hora_del_erp_y_quien_la_cambio(config, monkeypatch):
    cliente = config[0]
    assert desglose.coste_hora_operario() == 17.0          # recordado
    _erp_con(monkeypatch, 19.5)                              # cambiado en el ERP
    html = cliente.get("/configuracion").get_data(as_text=True)
    assert "19,5" in html and "gilberto" in html and "10/11/2026" in html
    assert 'id="in-coste_hora_operario"' not in html        # aqui no se edita
    assert desglose.coste_hora_operario() == 19.5          # y la web ya lo aplica


def test_la_pagina_avisa_si_el_erp_no_da_la_hora(config, monkeypatch):
    cliente = config[0]
    _erp_con(monkeypatch, None)
    html = cliente.get("/configuracion").get_data(as_text=True)
    assert "no tiene coste en el ERP" in html


def test_la_hora_del_erp_no_se_puede_guardar_aqui(config):
    cliente, guardados, _, _ = config
    r = cliente.post("/configuracion/coste_hora_operario", json={"valor": "18"})
    assert r.status_code == 409 and "ERP" in r.get_json()["error"]
    assert guardados == []


def test_se_guarda_con_coma_decimal_y_la_web_lo_aplica_ya(config):
    cliente, guardados, aplicados, _ = config
    r = cliente.post("/configuracion/prueba", json={"valor": "18,5"})
    assert r.status_code == 200 and r.get_json()["valor"] == 18.5
    assert guardados == [("prueba", 18.5)] and aplicados == [1]


@pytest.mark.parametrize("valor", ["", "abc", "0", "-3", "250"])
def test_no_se_guarda_un_valor_imposible(config, valor):
    cliente, guardados, _, _ = config
    r = cliente.post("/configuracion/prueba", json={"valor": valor})
    assert r.status_code == 400 and guardados == []


def test_un_parametro_que_no_existe_da_404(config):
    cliente, guardados, _, _ = config
    assert cliente.post("/configuracion/inventado", json={"valor": "1"}).status_code == 404
    assert guardados == []


def test_con_clave_en_el_env_hay_que_darla(config):
    cliente, guardados, _, mp = config
    mp.setenv("CONFIG_CLAVE", "secreta")
    assert cliente.post("/configuracion/prueba",
                        json={"valor": "18", "clave": "mala"}).status_code == 403
    assert guardados == []
    assert cliente.post("/configuracion/prueba",
                        json={"valor": "18", "clave": "secreta"}).status_code == 200
    assert guardados == [("prueba", 18.0)]
