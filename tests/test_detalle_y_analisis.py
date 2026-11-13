"""El detalle por pieza y lo que se saca de el (/evolucion, /piezas).

Sin ERP ni Postgres, como test_menos_erp.py. La parte que escribe en Postgres
(exportar_costes._guardar_detalle) se probo contra un Postgres desechable; aqui
va lo que se puede comprobar sin base de datos.
"""
import pandas as pd
import pytest

import app as app_mod
import db_pg
import desglose
import exportar_costes
import web_analisis


@pytest.fixture(autouse=True)
def _sin_postgres(monkeypatch):
    """El coste/hora, a su valor por defecto (17) sin intentar conectar: sale
    del ERP, y sin ERP vale 17. Postgres tampoco se toca."""
    def caido():
        raise ConnectionError("sin bases de datos en los tests")
    monkeypatch.setattr(desglose, "get_engine", caido)
    monkeypatch.setattr(db_pg, "get_pg_engine", caido)
    desglose.olvidar_cache()
    yield
    desglose.olvidar_cache()


def _arbol():
    """A (30 min) lleva dos B fabricados (10 min); cada B lleva CHAPA. Ademas A
    lleva CHAPA directa: la misma pieza en dos ramas."""
    filas = [
        {"Ruta": "|A|B|", "IdArticulo": "B", "Componente": "SUB", "Cant": 2.0, "Coste": 0.0,
         "Unidad": "ud", "TipoCompra": None, "PrecioCompra": None, "TiempoOp": 10.0,
         "TiempoMedio": 1},
        {"Ruta": "|A|B|CHAPA|", "IdArticulo": "CHAPA", "Componente": "CHAPA", "Cant": 4.0,
         "Coste": 8.0, "Unidad": "kg", "TipoCompra": "MATERIA PRIMA", "PrecioCompra": 2.0},
        {"Ruta": "|A|CHAPA|", "IdArticulo": "CHAPA", "Componente": "CHAPA", "Cant": 1.0,
         "Coste": 2.0, "Unidad": "kg", "TipoCompra": "MATERIA PRIMA", "PrecioCompra": 2.0},
        {"Ruta": "|A|TORN|", "IdArticulo": "TORN", "Componente": "TORNILLO", "Cant": 6.0,
         "Coste": 0.6, "Unidad": "ud", "TipoCompra": "MERCADERIA", "PrecioCompra": 0.1},
    ]
    for f in filas:                  # el ERP siempre trae la columna
        f.setdefault("TiempoMedio", 0)
    return desglose.construir_arbol(pd.DataFrame(filas), "A", "ART A",
                                    tiempo_raiz=30.0, medio_raiz=1)


def test_el_detalle_suma_lo_mismo_que_el_arbol():
    arbol = _arbol()
    lineas = desglose.lineas_coste(arbol)
    assert sum(l["coste_material"] for l in lineas) == pytest.approx(arbol["coste_mat"])
    assert sum(l["coste_operacion"] for l in lineas) == pytest.approx(arbol["coste_op_total"])


def test_una_pieza_en_varias_ramas_sale_una_vez_sumada():
    por_pieza = {l["idpieza"]: l for l in desglose.lineas_coste(_arbol())}
    assert set(por_pieza) == {"A", "B", "CHAPA", "TORN"}
    assert por_pieza["CHAPA"]["cantidad"] == pytest.approx(5.0)
    assert por_pieza["CHAPA"]["coste_material"] == pytest.approx(10.0)
    # B solo aporta operacion: 10 min x 2 piezas a 17 EUR/h (a 4 decimales,
    # como redondea construir_arbol el coste de operacion de cada linea)
    assert por_pieza["B"]["coste_material"] == 0
    assert por_pieza["B"]["coste_operacion"] == pytest.approx(20 * 17 / 60, abs=1e-4)


def test_la_huella_solo_cambia_si_cambia_el_coste():
    lineas = desglose.lineas_coste(_arbol())
    misma = [dict(l, descripcion="otro nombre") for l in lineas]
    otra = [dict(l, coste_material=l["coste_material"] * 1.01) for l in lineas]
    assert exportar_costes._huella(lineas) == exportar_costes._huella(misma)
    assert exportar_costes._huella(lineas) != exportar_costes._huella(otra)


# --- el simulador --------------------------------------------------------------

@pytest.fixture
def simulador(monkeypatch):
    """/piezas/simular con los usos de las piezas sustituidos."""
    usos = [
        # articulo, desc, pieza, cant, unidad, coste pieza, coste art, PV, completo, uds 12m
        ("X1", "JAULA", "CHAPA", 5, "kg", 10.0, 40.0, 100.0, True, 1000),
        ("X1", "JAULA", "TORN", 6, "ud", 0.6, 40.0, 100.0, True, 1000),
        ("X2", "COMEDERO", "CHAPA", 1, "kg", 2.0, 9.5, 10.0, True, 50),
        ("X3", "BEBEDERO", "CHAPA", 1, "kg", 2.0, 5.0, 8.0, False, 10),
    ]

    class Conexion:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    class Motor:
        def connect(self):
            return Conexion()

    monkeypatch.setattr(web_analisis.db_pg, "get_pg_engine", lambda: Motor())
    monkeypatch.setattr(web_analisis, "_usos",
                        lambda c, piezas: [u for u in usos if u[2] in piezas])
    return app_mod.app.test_client()


def test_subir_una_pieza_escala_su_coste_en_cada_articulo(simulador):
    r = simulador.post("/piezas/simular", json={"cambios": [{"pieza": "CHAPA", "pct": "50"}]})
    d = r.get_json()
    arts = {a["id"]: a for a in d["articulos"]}
    assert arts["X1"]["delta"] == pytest.approx(5.0)          # 10 EUR de chapa +50 %
    assert arts["X1"]["nuevo"] == pytest.approx(45.0)
    assert arts["X1"]["margen"] == 60.0 and arts["X1"]["margen_nuevo"] == 55.0
    assert arts["X1"]["impacto_anual"] == pytest.approx(5000)  # x 1000 vendidas
    # X2 pasa de ganar 0,5 a perder 0,5 por unidad
    assert arts["X2"]["margen"] == 5.0 and arts["X2"]["margen_nuevo"] == -5.0
    assert d["resumen"]["bajo_coste"] == 1
    # con el coste incompleto el margen no se afirma (regla de v_rentabilidad)
    assert arts["X3"]["margen"] is None and arts["X3"]["margen_nuevo"] is None


def test_varias_piezas_a_la_vez_se_suman_en_el_articulo(simulador):
    r = simulador.post("/piezas/simular", json={"cambios": [
        {"pieza": "CHAPA", "pct": "10"}, {"pieza": "TORN", "pct": "100"}]})
    x1 = next(a for a in r.get_json()["articulos"] if a["id"] == "X1")
    assert x1["delta"] == pytest.approx(1.0 + 0.6)


@pytest.mark.parametrize("cuerpo", [{}, {"cambios": []},
                                    {"cambios": [{"pieza": "CHAPA", "pct": "x"}]}])
def test_el_simulador_rechaza_lo_que_no_es_un_cambio(simulador, cuerpo):
    assert simulador.post("/piezas/simular", json=cuerpo).status_code == 400
