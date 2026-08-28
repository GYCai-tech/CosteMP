"""Recalcula los costes SOLO cuando alguien cambia algo en el ERP.

Mira la marca de tiempo mas alta de las tablas que afectan al coste. Si no se
ha movido desde la ultima vez, no hace nada: son unos milisegundos. Si se ha
movido, lanza el recalculo rapido (~40 s, solo los articulos sin cerrar).

Asi el fichero de Excel esta al dia a los pocos minutos de que gilberto active
una fase o Abel cargue una tarifa, sin machacar el ERP con recalculos a ciegas.

    py vigilar_costes.py                 # una comprobacion y sale (tarea programada)
    py vigilar_costes.py --bucle 300     # se queda vigilando cada 5 minutos
    py vigilar_costes.py --completo      # si detecta cambios, pasada entera

Un articulo hoy completo puede volver a tener huecos si se abre una rama dentro
de el, y el modo rapido no lo ve porque ni lo mira. Por eso conviene mantener
ademas la pasada entera diaria: esta vigilancia la complementa, no la sustituye.
"""
import argparse
import datetime as dt
import time

from sqlalchemy import text

from db import get_engine
from db_pg import get_pg_engine
import exportar_costes

# Tablas cuyo cambio puede alterar un coste. Fases_Entradas es la que mas
# importa (el escandallo), pero activar una fase toca Fases, y un precio nuevo
# toca las listas de proveedor.
VIGILADAS = [
    ("dbo.Fases", "FechaInsertUpdate"),
    ("dbo.Fases_Entradas", "FechaInsertUpdate"),
    ("dbo.Fases_Salidas", "FechaInsertUpdate"),
    ("dbo.Articulos_Conjuntos", "FechaInsertUpdate"),
    ("dbo.Listas_Precios_Prov_Art", "FechaInsertUpdate"),
    ("dbo.Trabajos_ManoObra", "FechaInsertUpdate"),
]

DDL = """
CREATE TABLE IF NOT EXISTS core.cfg_coste_vigilancia (
    id             boolean PRIMARY KEY DEFAULT true CHECK (id),   -- una sola fila
    ultimo_cambio  timestamp,
    ultima_revision timestamp
);
INSERT INTO core.cfg_coste_vigilancia (id) VALUES (true) ON CONFLICT DO NOTHING;
"""


def marca_erp():
    """La fecha de cambio mas alta de todo lo que afecta al coste."""
    union = " UNION ALL ".join(
        f"SELECT MAX({col}) AS m FROM {tabla}" for tabla, col in VIGILADAS)
    with get_engine().connect() as c:
        return c.execute(text(f"SELECT MAX(m) FROM ({union}) t")).scalar()


def revisar(pg, completo=False):
    """Compara la marca del ERP con la guardada. Recalcula si cambio."""
    with pg.begin() as c:
        c.execute(text(DDL))
        guardada = c.execute(text(
            "SELECT ultimo_cambio FROM core.cfg_coste_vigilancia")).scalar()

    actual = marca_erp()
    ahora = dt.datetime.now().strftime("%H:%M:%S")

    if guardada is not None and actual is not None and actual <= guardada:
        with pg.begin() as c:
            c.execute(text("UPDATE core.cfg_coste_vigilancia SET ultima_revision = now()"))
        print(f"{ahora}  sin cambios en el ERP")
        return False

    print(f"{ahora}  CAMBIOS detectados (ultimo: {actual}) -> recalculando...", flush=True)
    r = exportar_costes.ejecutar(pg, solo_pendientes=not completo, origen="programado")
    with pg.begin() as c:
        c.execute(text("UPDATE core.cfg_coste_vigilancia "
                       "SET ultimo_cambio = :m, ultima_revision = now()"), {"m": actual})
    print(f"          {r['completos']}/{r['articulos']} completos en {r['duracion']}s"
          f"  ->  {exportar_costes.SALIDA.name} actualizado", flush=True)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--bucle", type=int, metavar="SEG",
                    help="vigilar cada SEG segundos en vez de salir")
    ap.add_argument("--completo", action="store_true",
                    help="al detectar cambios, recalcular el catalogo entero")
    args = ap.parse_args()

    pg = get_pg_engine()
    if not args.bucle:
        revisar(pg, args.completo)
        return
    print(f"Vigilando el ERP cada {args.bucle}s. Ctrl+C para parar.", flush=True)
    while True:
        try:
            revisar(pg, args.completo)
        except Exception as ex:                                # noqa: BLE001
            # una caida de red no debe tumbar la vigilancia: se reintenta
            print(f"  ERROR: {str(ex)[:160]}", flush=True)
        time.sleep(args.bucle)


if __name__ == "__main__":
    main()
