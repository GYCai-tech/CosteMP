"""Lo que evita cargar el ERP: la pasada cada pocos dias y la memoria por articulo.

A diferencia de test_regresion_costes.py, estos NO conectan con el ERP: lo
que comprueban es justamente cuando se le pregunta y cuando no.
"""
import datetime as dt

import pandas as pd
import pytest

import desglose
from vigilar_costes import toca_pasada


# --- la pasada entera, cada 3 dias a las 05:00 ------------------------------

LUNES_5 = dt.datetime(2026, 10, 5, 5, 0)


@pytest.mark.parametrize("ultima, ahora, toca", [
    # sin ninguna pasada registrada, toca en cuanto sea la hora
    (None, LUNES_5, True),
    (None, LUNES_5.replace(hour=4, minute=59), False),
    # pasada el lunes: ni el martes ni el miercoles
    (LUNES_5, dt.datetime(2026, 10, 6, 5, 0), False),
    (LUNES_5, dt.datetime(2026, 10, 7, 23, 0), False),
    # el jueves si, pero no antes de las 05:00
    (LUNES_5, dt.datetime(2026, 10, 8, 4, 59), False),
    (LUNES_5, dt.datetime(2026, 10, 8, 5, 0), True),
    # se cuentan dias de calendario: una pasada que acabo a las 05:04 no
    # retrasa la siguiente a las 05:04 del jueves
    (dt.datetime(2026, 10, 5, 5, 4), dt.datetime(2026, 10, 8, 5, 0), True),
    # si el servidor estuvo apagado, la atrasada se hace en cuanto vuelve
    (LUNES_5, dt.datetime(2026, 10, 12, 10, 0), True),
])
def test_la_pasada_toca_cada_tres_dias_a_partir_de_la_hora(ultima, ahora, toca):
    assert toca_pasada(ultima, ahora, 3, "05:00") is toca


# --- lo leido de un articulo se recuerda unos minutos ------------------------

@pytest.fixture
def erp_falso(monkeypatch):
    """Sustituye las lecturas del ERP por contadores."""
    desglose.olvidar_cache()
    llamadas = {"raiz": 0, "desglose": 0}

    def raiz(codigo):
        llamadas["raiz"] += 1
        return {"nombre": f"ART {codigo}", "tiempo": (1.5, 1), "sin_operacion": 0,
                "con_orden": 1, "es_externo": 0, "coste_propio": None}

    class Conexion:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    class Motor:
        def connect(self):
            return Conexion()

    def read_sql(sql, cn, params):
        llamadas["desglose"] += 1
        return pd.DataFrame({"Nivel": [1], "IdArticulo": [params["codigo"]]})

    monkeypatch.setattr(desglose, "_leer_raiz", raiz)
    monkeypatch.setattr(desglose, "get_engine", lambda: Motor())
    monkeypatch.setattr(desglose.pd, "read_sql", read_sql)
    yield llamadas
    desglose.olvidar_cache()


def test_las_seis_consultas_del_articulo_van_juntas_y_una_sola_vez(erp_falso):
    # lo que hace /api/desglose: las seis piezas del articulo raiz
    c = "10201001"
    desglose.nombre_articulo(c)
    desglose.tiempo_operacion(c)
    desglose.sin_operacion(c)
    desglose.con_orden(c)
    desglose.es_externo(c)
    desglose.coste_propio(c)
    assert erp_falso["raiz"] == 1


def test_ver_y_descargar_el_mismo_articulo_no_repite_el_despiece(erp_falso):
    desglose.desglose("10201001")        # verlo
    desglose.desglose("10201001")        # descargar su Excel
    desglose.desglose("10201-001")       # mismo codigo, escrito distinto
    assert erp_falso["desglose"] == 1
    desglose.desglose("12101021")        # otro articulo si se lee
    assert erp_falso["desglose"] == 2


def test_lo_recordado_caduca(erp_falso, monkeypatch):
    desglose.nombre_articulo("10201001")
    monkeypatch.setattr(desglose, "CACHE_TTL_S", 0)
    desglose.nombre_articulo("10201001")
    assert erp_falso["raiz"] == 2


def test_tocar_el_despiece_recibido_no_estropea_el_recordado(erp_falso):
    df = desglose.desglose("10201001")
    df["Nivel"] = 99
    assert desglose.desglose("10201001")["Nivel"].tolist() == [1]


def test_la_memoria_no_crece_sin_limite(erp_falso, monkeypatch):
    monkeypatch.setattr(desglose, "CACHE_MAX", 10)
    for i in range(25):
        desglose.nombre_articulo(f"A{i}")
    assert len(desglose._cache) <= 10
    # lo ultimo consultado sigue recordado
    antes = erp_falso["raiz"]
    desglose.nombre_articulo("A24")
    assert erp_falso["raiz"] == antes


# --- el boton "Recalcular" de la pantalla del articulo ------------------------

def test_olvidar_un_articulo_obliga_a_releerlo_y_solo_a_ese(erp_falso):
    desglose.desglose("10201001")
    desglose.desglose("12101021")
    assert desglose.leido_a("10201001") is not None
    desglose.olvidar_articulo("10201001")
    assert desglose.leido_a("10201001") is None
    desglose.desglose("10201001")
    desglose.desglose("12101021")
    assert erp_falso["desglose"] == 3          # releido el olvidado, no el otro


@pytest.fixture
def boton(erp_falso, monkeypatch):
    """La ruta /api/recalcular con Postgres y la publicacion sustituidos."""
    import exportar_costes
    import db_pg
    from app import app
    publicados = []
    catalogo = {"10201001"}
    monkeypatch.setattr(db_pg, "get_pg_engine", lambda: object())
    monkeypatch.setattr(exportar_costes, "en_catalogo", lambda pg, c: c in catalogo)
    monkeypatch.setattr(exportar_costes, "recalcular_articulo",
                        lambda pg, c: (desglose.desglose(c), publicados.append(c)))
    return app.test_client(), publicados


def test_recalcular_un_articulo_del_catalogo_lo_relee_y_lo_publica(boton, erp_falso):
    cliente, publicados = boton
    desglose.desglose("10201001")                       # ya estaba recordado
    r = cliente.post("/api/recalcular?codigo=10201001")
    assert r.status_code == 200 and r.get_json()["en_catalogo"] is True
    assert publicados == ["10201001"]
    assert erp_falso["desglose"] == 2                   # el boton fue al ERP
    cliente.get("/api/desglose?codigo=10201001")         # y la pantalla no repite
    assert erp_falso["desglose"] == 2


def test_recalcular_un_articulo_fuera_del_catalogo_lo_relee_sin_publicar(boton, erp_falso):
    cliente, publicados = boton
    r = cliente.post("/api/recalcular?codigo=99999999")
    assert r.status_code == 200 and r.get_json()["en_catalogo"] is False
    assert publicados == [] and erp_falso["desglose"] == 1


def test_recalcular_sin_codigo_no_va_al_erp(boton, erp_falso):
    cliente, _ = boton
    assert cliente.post("/api/recalcular?codigo=").status_code == 400
    assert erp_falso["desglose"] == 0
