"""Comprueba que lo publicado en gyc_analytics coincide con lo que calcula la app.

Es la red de seguridad de todo el montaje: si algun dia el proceso de
publicacion se desincroniza del calculo, esto lo canta. Conviene pasarlo
despues de cualquier cambio en desglose.py o app.py.

    py validar_costes.py                # muestra 30 articulos al azar
    py validar_costes.py --todos        # los 548 (unos 4 minutos)

Tolerancia: 0,0001 EUR. Cualquier diferencia por encima es un fallo.
"""
import argparse
import random

from sqlalchemy import text

from db_pg import get_pg_engine
from desglose import (coste_propio, desglose, nombre_articulo, sin_operacion,
                      tiempo_operacion)
from app import construir_arbol

TOL = 0.0001


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--todos", action="store_true")
    ap.add_argument("--muestra", type=int, default=30)
    args = ap.parse_args()

    pg = get_pg_engine()
    with pg.connect() as c:
        pub = {r[0]: (r[1], r[2], r[3]) for r in c.execute(text(
            "SELECT idarticulo, coste_material, coste_operacion, coste_total "
            "FROM core.dim_coste_escandallo WHERE error IS NULL"))}
    if not pub:
        print("No hay nada publicado todavia. Lanza antes exportar_costes.py")
        return

    codigos = sorted(pub)
    if not args.todos:
        random.seed()
        codigos = random.sample(codigos, min(args.muestra, len(codigos)))
    print(f"Comparando {len(codigos)} articulos (tolerancia {TOL} EUR)\n")

    fallos = 0
    for cod in codigos:
        df = desglose(cod)
        if df.empty:
            continue
        t, es_medio = tiempo_operacion(cod)
        arbol = construir_arbol(df, cod, nombre_articulo(cod), t,
                                sin_operacion(cod), coste_propio(cod), es_medio)
        vivo = (arbol["coste_mat"], arbol["coste_op_total"], arbol["coste_total"])
        guardado = tuple(float(x) if x is not None else 0.0 for x in pub[cod])
        difs = [abs(a - b) for a, b in zip(vivo, guardado)]
        if max(difs) > TOL:
            fallos += 1
            print(f"  DIFERENCIA {cod}")
            for etiqueta, a, b in zip(("material", "operacion", "total"), vivo, guardado):
                print(f"     {etiqueta:10} app={a:12.4f}  publicado={b:12.4f}  dif={a - b:+.4f}")

    if fallos:
        print(f"\n{fallos} de {len(codigos)} NO cuadran. Revisar antes de fiarse del Excel.")
        raise SystemExit(1)
    print(f"Todo cuadra: {len(codigos)}/{len(codigos)} identicos.")


if __name__ == "__main__":
    main()
