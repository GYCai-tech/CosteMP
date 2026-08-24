"""Tests de regresion de la logica de negocio (calculo de coste).

Son tests de INTEGRACION: conectan contra el ERP real (GOMEZYCRESPO) en modo
solo lectura, igual que hace la propia app (ver db.py). No hay mocks ni base
de datos de prueba -- el trabajo real de este script es interpretar datos
reales y a veces contradictorios del ERP, asi que las reglas solo se pueden
verificar contra articulos reales.

Cada test fija un articulo concreto que sirvio para encontrar y arreglar un
bug real (ver el historial de commits de desglose.py / app.py). Si el precio
de compra de alguno de estos articulos cambia mucho en el ERP con el tiempo,
puede hacer falta ajustar el margen o el valor esperado -- lo importante no
es el numero exacto, sino que seguir cumpliendose la RELACION que describe
cada test (no se duplica un precio, no desaparece un componente, etc). Por
eso casi todos comparan contra un umbral o una relacion, no un euro exacto.

Ejecutar (necesita alcanzar el SQL Server del ERP, igual que la app):
    pip install -r requirements-dev.txt
    pytest tests/ -v

O dentro del contenedor en marcha:
    docker exec -w /app costemp python -m pytest tests/ -v
"""
import pytest

from app import resumen_lote
from desglose import coste_propio, desglose


def test_no_duplica_precio_propio_si_esta_en_no_suma_propia():
    """11601001 BEBEDERO AUTOMATICO COLGANTE: su propio precio de compra
    coincide con el de 11601067 BEBEDERO AUTOMATICO Arion, que ya viene
    dentro de su escandallo (esta en NO_SUMA_PROPIA). Si esta exclusion se
    rompe, el coste vuelve a dispararse por sumar el mismo bebedero dos
    veces (bug original: 24,78 EUR en vez de ~16,5 EUR)."""
    fila, _ = resumen_lote("11601001")
    assert fila["CosteTotal"] < 20.0, (
        f"CosteTotal={fila['CosteTotal']}: parece que 11601067 se esta "
        "sumando dos veces (precio propio de la raiz + escandallo). "
        "Revisar NO_SUMA_PROPIA en desglose.py."
    )
    assert coste_propio("11601001") is None, (
        "coste_propio('11601001') deberia seguir suprimido por NO_SUMA_PROPIA."
    )


@pytest.mark.parametrize("codigo,minimo", [
    ("18608010", 1.5),   # NIDO MADERA PERIQUITOS
    ("18608012", 2.0),   # NIDO MADERA AGAPORNIS
    ("18608013", 4.0),   # NIDO MADERA NINFAS (Completo=No: le falta el
                          # precio de la caja, pero el nido si debe contar)
])
def test_ciclo_fase_conjunto_degenerado_no_borra_el_precio(codigo, minimo):
    """Estos 3 nidos tienen una relacion circular fase<->conjunto de UN SOLO
    componente: el "conjunto" esta hecho de exactamente el mismo articulo
    que ya lo lleva dentro por su fase (dato contradictorio en el ERP). Antes
    del fix, el freno anti-duplicados se comia el precio del componente
    "comprado" sin avisar y el coste caia a una fraccion de centimo."""
    fila, _ = resumen_lote(codigo)
    assert fila["CosteTotal"] >= minimo, (
        f"{codigo}: CosteTotal={fila['CosteTotal']}, por debajo de {minimo}. "
        "Puede que el corte de ciclos fase/conjunto (ver 'componentes' CTE "
        "en desglose.py) se haya vuelto a tragar el precio del componente "
        "'comprado'."
    )


def test_ciclo_fase_conjunto_con_varios_componentes_no_se_toca():
    """61000165 CONJUNTO TAPAS PERRO GRANDE Y SUPER tiene el mismo tipo de
    bucle fase<->conjunto que los nidos, pero con 2 componentes REALES (no
    es el caso degenerado: el conjunto no esta hecho de 1 sola unidad de si
    mismo). La primera version del fix rompio este caso por un error de
    precedencia de operadores (AND en vez de OR) y lo dejo sin escandallo.
    No debe volver a pasar."""
    df = desglose("61000165")
    ids = set(df["IdArticulo"])
    assert {"60703023", "60703024"} <= ids, (
        f"61000165 solo tiene estos componentes: {ids}. Deberia conservar "
        "sus 2 piezas reales (60703023 TAPA SUP. y 60703024 TAPA INF.)."
    )


def test_linea_de_compra_con_100_por_cien_descuento_se_ignora():
    """25701001 GRAPADORA MANUAL tenia 2 lineas del mismo pedido y misma
    fecha: una real (0% descuento, 6,95 EUR) y una de regalo/promocion (100%
    descuento). El desempate por IdPedido elegia arbitrariamente la de
    regalo, dejando el coste en 0 EUR sin ningun aviso (el precio no era
    NULL, asi que "Sin precio" tampoco saltaba)."""
    fila, _ = resumen_lote("25701001")
    assert fila["CosteTotal"] == pytest.approx(6.95, abs=0.01)
    assert fila["Completo"] == "Sí"


def test_trabajo_externo_no_dispara_aviso_de_sin_tiempo():
    """18103604 BASE COMEDERO/BEBEDERO SUELO...LACADA tiene escandallo Y
    ADEMAS precio propio (trabajo externo de lacado). GYC no mide ni ficha
    el tiempo de un proceso que hace un tercero fuera de la empresa, asi que
    0 minutos de mano de obra propia es correcto y no debe avisar "sin
    tiempo"."""
    fila, _ = resumen_lote("18103604")
    assert fila["Sin tiempo"] == 0
    assert fila["Completo"] == "Sí"


def test_pieza_fabricada_sin_partes_SI_avisa_sin_tiempo():
    """Guarda de no-regresion al reves del test anterior: 11602053 AGUJA
    PLASTICO (componente de 11601035) es una pieza fabricada NORMAL -sin
    precio de compra propio, no es trabajo externo- que no tiene ningun
    dato de tiempo (ni mano de obra imputada ni media de bonos). Esta SI es
    un hueco real y debe seguir avisando: si el fix del trabajo externo se
    hace demasiado amplio, este aviso legitimo desaparece en silencio."""
    _, faltantes = resumen_lote("11601035")
    ids_con_aviso_tiempo = {
        f["IdComponente"] for f in faltantes
        if f["Motivo"] == "Sin tiempo de operación"
    }
    assert "11602053" in ids_con_aviso_tiempo


def test_articulo_inexistente_no_revienta():
    """Un codigo que no existe en el ERP debe devolver un hueco explicado en
    Faltantes, no un error ni un coste inventado."""
    fila, faltantes = resumen_lote("XXNOEXISTE99")
    assert fila["Completo"] == "No"
    assert fila["CosteTotal"] is None
    assert any(f["Motivo"] == "Artículo no encontrado o sin despiece"
               for f in faltantes)


@pytest.mark.parametrize("codigo", ["11601001", "15101023", "18608010"])
def test_material_mas_operacion_es_igual_al_total(codigo):
    """Invariante basica del rollup: para cualquier articulo con coste,
    Material + Operacion tiene que cuadrar con el Total."""
    fila, _ = resumen_lote(codigo)
    assert fila["CosteMaterial"] + fila["CosteOperacion"] == pytest.approx(
        fila["CosteTotal"], abs=0.001
    )
