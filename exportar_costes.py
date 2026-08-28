"""Recalcula el coste de los articulos del catalogo y lo publica en gyc_analytics.

NO calcula nada por su cuenta: reutiliza desglose.py y app.py, el mismo codigo
que alimenta la web. Es a proposito. Una segunda implementacion del calculo
acabaria divergiendo en silencio y nadie se enteraria hasta que un numero no
cuadrase.

Lo usan dos sitios y comparten ejecutar():
    - la pagina /costes de la aplicacion (boton Recalcular)
    - la linea de comandos, para la pasada programada de madrugada

        py exportar_costes.py                  # el catalogo entero (~4 min)
        py exportar_costes.py --solo-pendientes  # solo los sin cerrar (~45 s)
        py exportar_costes.py --crear          # ademas crea las tablas si faltan
        py exportar_costes.py --origen programado

La escritura es ATOMICA: se calcula todo en memoria y solo al final se sustituye
el contenido de las tablas en una sola transaccion. Asi nadie lee una foto a
medias, y si el proceso falla el dato anterior queda intacto.
"""
import argparse
import datetime as dt
import os
import pathlib
import time

from sqlalchemy import text

from db import get_engine
from db_pg import get_pg_engine
from desglose import (coste_propio, desglose, nombre_articulo, sin_operacion,
                      tiempo_operacion)
from app import avisos_arbol, construir_arbol

DDL = pathlib.Path(__file__).with_name("sql") / "01_coste_objetos.sql"
# respaldo por si la tabla de catalogo esta vacia (primer arranque)
CATALOGO_TXT = r"C:\Users\santiago.arce\Desktop\costes\catalogo.txt"

MOTIVOS = ("Sin tipo de aprovisionamiento", "Sin precio de compra",
           "Sin escandallo activo", "Sin tiempo de operación")

# Libro que lee Power Query desde el Excel de analisis. Se genera en cada
# recalculo. Va aparte del libro de analisis a proposito: aquel tiene 32 hojas
# con tablas dinamicas y openpyxl no sabe conservarlas.
#
# La ruta se configura con COSTES_SALIDA_XLSX porque en el servidor no hay
# ningun escritorio de Windows: alli apunta a un volumen montado. Si la carpeta
# no existe no pasa nada -- los datos siguen publicados en Postgres y servidos
# por /costes/datos.xlsx, que es de donde puede leer cualquiera.
SALIDA = pathlib.Path(os.getenv(
    "COSTES_SALIDA_XLSX",
    r"C:\Users\santiago.arce\Desktop\costes\datos_costes.xlsx"))

HOJAS = {"Resumen": "core.v_coste_resumen",       # cuantos articulos y piezas quedan
         "Articulos": "core.v_coste_articulo",    # coste partido en MP y operacion
         "Pendientes": "core.v_coste_pendiente",  # que falta, pieza a pieza
         "Costes": "core.v_coste_hoja10"}         # alimenta el VLOOKUP del libro


def exportar_fichero(pg):
    """Vuelca las vistas al libro que consume Excel.

    Devuelve la ruta escrita, o None si no se pudo. Un fallo aqui NO tumba el
    recalculo: lo que importa ya esta publicado en Postgres. Los dos casos
    habituales son tener el .xlsx abierto en Excel (lo bloquea) y que la ruta
    no exista, que es lo que pasa en el contenedor si no se monta el volumen.
    """
    import pandas as pd                                   # solo hace falta aqui

    try:
        SALIDA.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(SALIDA, engine="openpyxl") as xls:
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
    except OSError as ex:
        print(f"  aviso: no se pudo escribir {SALIDA}: {ex}", flush=True)
        return None
    return SALIDA


def crear_objetos(pg):
    with pg.begin() as c:
        c.execute(text(DDL.read_text(encoding="utf-8")))


def leer_pendientes(pg):
    """Solo los articulos que la ultima pasada dejo sin cerrar.

    Sirve para seguir el avance en caliente: son ~113 de 547, asi que baja de
    4 minutos a menos de uno. NO sustituye a la pasada completa: un articulo
    hoy completo puede volver a tener huecos si alguien activa una fase dentro
    de el y detras aparecen piezas sin precio (paso el 27/08 con el 10201001,
    que gano 4 huecos nuevos al abrirse una rama). Esas regresiones solo las ve
    la pasada entera.
    """
    with pg.connect() as c:
        return [r[0] for r in c.execute(text(
            "SELECT c.idarticulo FROM core.dim_coste_escandallo c "
            "INNER JOIN core.cfg_catalogo_coste g ON g.idarticulo = c.idarticulo "
            "WHERE g.activo AND (NOT c.completo OR c.error IS NOT NULL) "
            "ORDER BY c.idarticulo"))]


def leer_catalogo(pg, limite=None):
    with pg.connect() as c:
        ids = [r[0] for r in c.execute(text(
            "SELECT idarticulo FROM core.cfg_catalogo_coste "
            "WHERE activo ORDER BY idarticulo"))]
    if not ids and pathlib.Path(CATALOGO_TXT).exists():
        ids = [l.strip() for l in open(CATALOGO_TXT, encoding="utf-8") if l.strip()]
        with pg.begin() as c:
            c.execute(text("INSERT INTO core.cfg_catalogo_coste (idarticulo) "
                           "SELECT unnest(CAST(:ids AS varchar[])) "
                           "ON CONFLICT (idarticulo) DO NOTHING"), {"ids": ids})
    return ids[:limite] if limite else ids


def calcular(cod, ahora):
    """Una fila de coste + sus huecos. Mismo criterio que resumen_lote()."""
    nombre = nombre_articulo(cod)
    df = desglose(cod)
    if df.empty:
        return dict(idarticulo=cod, descripcion=nombre, coste_material=None,
                    coste_operacion=None, coste_total=None, piezas_sin_coste=1,
                    piezas_sin_escandallo=0, piezas_sin_tipo=0,
                    piezas_sin_precio=0, piezas_sin_tiempo=0, completo=False,
                    error="Sin despiece o articulo inexistente",
                    fecha_calculo=ahora), []

    t, es_medio = tiempo_operacion(cod)
    arbol = construir_arbol(df, cod, nombre, t, sin_operacion(cod),
                            coste_propio(cod), es_medio)
    h = avisos_arbol(arbol)
    n_esc = len(h["Sin escandallo activo"])
    n_coste = len(h["Sin tipo de aprovisionamiento"]) + len(h["Sin precio de compra"])

    fila = dict(idarticulo=cod, descripcion=nombre,
                coste_material=round(arbol["coste_mat"], 4),
                coste_operacion=round(arbol["coste_op_total"], 4),
                coste_total=round(arbol["coste_total"], 4),
                piezas_sin_coste=n_coste, piezas_sin_escandallo=n_esc,
                # "sin coste" se guarda ademas partido en sus dos motivos, que
                # es lo que permite filtrar por tipo de problema en la pagina
                piezas_sin_tipo=len(h["Sin tipo de aprovisionamiento"]),
                piezas_sin_precio=len(h["Sin precio de compra"]),
                piezas_sin_tiempo=len(h["Sin tiempo de operación"]),
                # "Sin tiempo" NO tumba el completo: el material si esta
                # costeado. Mismo criterio que resumen_lote() y que la web.
                completo=(n_esc + n_coste) == 0,
                error=None, fecha_calculo=ahora)

    huecos, vistos = [], set()
    for motivo in MOTIVOS:
        for n in h[motivo]:
            clave = (n["id"], motivo)
            if clave in vistos:            # la PK no admite el par repetido
                continue
            vistos.add(clave)
            huecos.append(dict(idarticulo=cod, idpieza=n["id"],
                               descripcion=n["nombre"], motivo=motivo,
                               fecha_calculo=ahora))
    return fila, huecos


def enriquecer(huecos):
    """Anade a cada hueco el tipo, la familia y las senales de compra/fabricacion.

    Se consulta en BLOQUE, no pieza a pieza: son ~500 piezas distintas y una
    consulta por cada una multiplicaria la duracion del proceso.
    """
    if not huecos:
        return
    ids = sorted({h["idpieza"] for h in huecos})
    datos = {}
    with get_engine().connect() as c:
        for i in range(0, len(ids), 500):           # el IN de SQL Server tiene tope
            lote = ids[i:i + 500]
            marcas = ", ".join(f":p{j}" for j in range(len(lote)))
            params = {f"p{j}": v for j, v in enumerate(lote)}
            for r in c.execute(text(f"""
                SELECT a.IdArticulo, t.Descrip,
                       COALESCE(f.Descrip, '(Sin familia)'),
                       CASE WHEN EXISTS (SELECT 1 FROM dbo.Fases_Salidas fs
                                         WHERE fs.IdArticulo = a.IdArticulo)
                            THEN 1 ELSE 0 END,
                       (SELECT COUNT(*) FROM dbo.Pedidos_Prov_Lineas pl
                         WHERE pl.IdArticulo = a.IdArticulo
                           AND pl.FechaAlbaran IS NOT NULL
                           AND pl.Precio_EURO > 0 AND pl.Descuento < 100),
                       CASE WHEN EXISTS (
                             SELECT 1 FROM dbo.Listas_Precios_Prov_Art lpa
                             INNER JOIN dbo.Listas_Precios_Prov lp ON lp.IdLista = lpa.IdLista
                             WHERE lpa.IdArticulo = a.IdArticulo
                               AND lp.IdProveedor <> '0' AND lpa.Precio > 0)
                            THEN 1 ELSE 0 END
                FROM dbo.Articulos a
                LEFT JOIN dbo.Articulos_Tipos_Aprovisionamiento t
                       ON t.IdTipoAprovisionamiento = a.IdTipoAprovisionamiento
                LEFT JOIN dbo.Articulos_Familias f ON f.IdFamilia = a.IdFamilia
                WHERE a.IdArticulo IN ({marcas})"""), params):
                datos[r[0]] = r
    for h in huecos:
        r = datos.get(h["idpieza"])
        h["tipo_aprovisionamiento"] = r[1] if r else None
        h["familia"] = r[2] if r else None
        h["tuvo_fase"] = bool(r[3]) if r else None
        h["veces_comprada"] = int(r[4]) if r else None
        h["tiene_tarifa"] = bool(r[5]) if r else None


def enriquecer_articulos(filas):
    """Anade la familia del ERP a cada articulo.

    No se coge de core.dim_articulo porque alli nombrefamilia esta practicamente
    vacia: 3 de 547 (comprobado 27/08/2026). La fuente buena es
    dbo.Articulos_Familias del ERP.
    """
    if not filas:
        return
    ids = sorted({f["idarticulo"] for f in filas})
    fam = {}
    with get_engine().connect() as c:
        for i in range(0, len(ids), 500):
            lote = ids[i:i + 500]
            marcas = ", ".join(f":p{j}" for j in range(len(lote)))
            for r in c.execute(text(f"""
                SELECT a.IdArticulo, COALESCE(f.Descrip, '(Sin definir)')
                FROM dbo.Articulos a
                LEFT JOIN dbo.Articulos_Familias f ON f.IdFamilia = a.IdFamilia
                WHERE a.IdArticulo IN ({marcas})"""),
                {f"p{j}": v for j, v in enumerate(lote)}):
                fam[r[0]] = r[1]
    for f in filas:
        f["familia"] = fam.get(f["idarticulo"])


def publicar(pg, filas, huecos, parcial=False):
    """Sustituye el contenido de las dos tablas en UNA transaccion.

    parcial=True cambia solo los articulos recalculados y deja intactos los
    demas. Es lo que permite el recalculo rapido sin perder la foto completa.
    """
    ids = [f["idarticulo"] for f in filas]
    with pg.begin() as c:
        if parcial:
            c.execute(text("DELETE FROM core.dim_coste_escandallo "
                           "WHERE idarticulo = ANY(CAST(:ids AS varchar[]))"), {"ids": ids})
            c.execute(text("DELETE FROM core.fact_coste_hueco "
                           "WHERE idarticulo = ANY(CAST(:ids AS varchar[]))"), {"ids": ids})
        else:
            c.execute(text("TRUNCATE core.dim_coste_escandallo"))
            c.execute(text("TRUNCATE core.fact_coste_hueco"))
        if filas:
            c.execute(text("""
                INSERT INTO core.dim_coste_escandallo
                  (idarticulo, descripcion, familia, coste_material, coste_operacion,
                   coste_total, piezas_sin_coste, piezas_sin_escandallo,
                   piezas_sin_tipo, piezas_sin_precio,
                   piezas_sin_tiempo, completo, error, fecha_calculo)
                VALUES
                  (:idarticulo, :descripcion, :familia, :coste_material, :coste_operacion,
                   :coste_total, :piezas_sin_coste, :piezas_sin_escandallo,
                   :piezas_sin_tipo, :piezas_sin_precio,
                   :piezas_sin_tiempo, :completo, :error, :fecha_calculo)"""), filas)
        if huecos:
            c.execute(text("""
                INSERT INTO core.fact_coste_hueco
                  (idarticulo, idpieza, descripcion, motivo, fecha_calculo,
                   tipo_aprovisionamiento, familia, tuvo_fase,
                   veces_comprada, tiene_tarifa)
                VALUES (:idarticulo, :idpieza, :descripcion, :motivo, :fecha_calculo,
                        :tipo_aprovisionamiento, :familia, :tuvo_fase,
                        :veces_comprada, :tiene_tarifa)
                ON CONFLICT DO NOTHING"""), huecos)


def ejecutar(pg, solo_pendientes=False, limite=None, origen="manual",
             progreso=None):
    """El recalculo completo, de principio a fin. Devuelve un resumen.

    progreso: funcion opcional que recibe (hechos, total) para pintar el avance
    en la pagina web. La linea de comandos no la usa.
    """
    inicio = dt.datetime.now()
    with pg.begin() as c:
        log_id = c.execute(text(
            "INSERT INTO core.log_coste_recalculo (inicio, origen, resultado) "
            "VALUES (:i, :o, 'en curso') RETURNING id"),
            {"i": inicio, "o": origen}).scalar()

    ids = leer_pendientes(pg) if solo_pendientes else leer_catalogo(pg, limite)
    filas, huecos, errores = [], [], 0
    for i, cod in enumerate(ids, 1):
        try:
            f, hs = calcular(cod, inicio)
            filas.append(f)
            huecos.extend(hs)
            if f["error"]:
                errores += 1
        except Exception as ex:                       # noqa: BLE001
            errores += 1
            filas.append(dict(idarticulo=cod, descripcion=None, coste_material=None,
                              coste_operacion=None, coste_total=None,
                              piezas_sin_coste=0, piezas_sin_escandallo=0,
                              piezas_sin_tipo=0, piezas_sin_precio=0,
                              piezas_sin_tiempo=0, completo=False,
                              error=str(ex)[:500], fecha_calculo=inicio))
        if progreso:
            progreso(i, len(ids))

    enriquecer_articulos(filas)
    enriquecer(huecos)
    publicar(pg, filas, huecos, parcial=solo_pendientes)

    fin = dt.datetime.now()
    dur = round((fin - inicio).total_seconds(), 1)
    with pg.begin() as c:
        c.execute(text("""UPDATE core.log_coste_recalculo
                             SET fin=:f, articulos=:a, errores=:e,
                                 duracion_seg=:d, resultado='ok'
                           WHERE id=:id"""),
                  {"f": fin, "a": len(filas), "e": errores, "d": dur, "id": log_id})
    exportar_fichero(pg)
    return {"articulos": len(filas), "errores": errores, "duracion": dur,
            "completos": sum(1 for f in filas if f["completo"])}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--crear", action="store_true", help="crea las tablas si faltan")
    ap.add_argument("--limite", type=int, help="solo los N primeros (pruebas)")
    ap.add_argument("--solo-pendientes", action="store_true",
                    help="recalcula solo los articulos sin cerrar (rapido)")
    ap.add_argument("--origen", default="manual", choices=("manual", "programado"))
    args = ap.parse_args()

    pg = get_pg_engine()
    if args.crear:
        crear_objetos(pg)
        print("Objetos verificados/creados en core.*")

    t0 = time.time()
    print("Recalculando" + (" SOLO PENDIENTES" if args.solo_pendientes else
                            " el catalogo entero") + "...", flush=True)
    r = ejecutar(pg, solo_pendientes=args.solo_pendientes, limite=args.limite,
                 origen=args.origen)
    print(f"OK  {r['completos']}/{r['articulos']} completos  |  "
          f"{r['errores']} errores  |  {time.time() - t0:.0f}s")
    print("Fichero para Excel ->", SALIDA)


if __name__ == "__main__":
    main()
