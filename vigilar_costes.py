"""Recalcula el coste del catalogo entero cada pocos dias (3 por defecto).

Antes vigilaba el ERP cada 10 minutos (seis MAX() sobre las tablas del
escandallo) y, si algo se habia movido, recalculaba los pendientes; ademas
hacia una pasada entera cada manana. Con el ERP en uso eso eran unas 30
pasadas al dia y ~4.000 despieces: el 17 % de la CPU del SQL Server del ERP
(diagnostico del 05/10/2026). Ahora el ERP solo se toca en la pasada, y entre
pasada y pasada la vigilancia mira unicamente Postgres.

El libro de Excel, /catalogo y /costes leen de Postgres, nunca del ERP, asi
que siguen funcionando igual entre pasadas, con el coste de la ultima. Quien
necesite el coste al momento puede consultar el articulo en el buscador o
pulsar "Recalcular" en /costes.

    py vigilar_costes.py                      # mira si toca, la hace si toca y sale
    py vigilar_costes.py --bucle 900          # se queda vigilando (solo Postgres)
    py vigilar_costes.py --cada-dias 3 --a 05:00
"""
import argparse
import datetime as dt
import time

from sqlalchemy import text

from db_pg import get_pg_engine
import exportar_costes

DDL = """
CREATE TABLE IF NOT EXISTS core.cfg_coste_vigilancia (
    id             boolean PRIMARY KEY DEFAULT true CHECK (id),   -- una sola fila
    ultimo_cambio  timestamp,
    ultima_revision timestamp
);
ALTER TABLE core.cfg_coste_vigilancia ADD COLUMN IF NOT EXISTS ultima_completa timestamp;
INSERT INTO core.cfg_coste_vigilancia (id) VALUES (true) ON CONFLICT DO NOTHING;
"""


def ultima_pasada(pg):
    """Cuando fue la ultima pasada entera programada.

    Se guarda en Postgres y no en memoria: si el contenedor se reinicia, no
    tiene que volver a pasar por el ERP solo porque haya olvidado la fecha. La
    primera vez, sin fecha guardada, vale la ultima pasada programada del log,
    para que desplegar esto no dispare una pasada nada mas arrancar.
    """
    with pg.begin() as c:
        c.execute(text(DDL))
        ultima = c.execute(text(
            "SELECT ultima_completa FROM core.cfg_coste_vigilancia")).scalar()
        if ultima is None:
            ultima = c.execute(text(
                "SELECT MAX(inicio) FROM core.log_coste_recalculo "
                "WHERE resultado = 'ok' AND origen = 'programado'")).scalar()
    return ultima


def toca_pasada(ultima, ahora, cada_dias, hhmm):
    """True si han pasado `cada_dias` dias (de calendario) desde la ultima
    pasada y ya es la hora `hhmm`.

    Se cuentan dias de calendario y no horas para que la pasada caiga siempre
    a la misma hora y no se vaya retrasando lo que tarda cada una.
    """
    h, m = (int(x) for x in hhmm.split(":"))
    if ahora.time() < dt.time(h, m):
        return False
    return ultima is None or (ahora.date() - ultima.date()).days >= cada_dias


def pasada(pg):
    """La pasada entera. Unico punto en que la vigilancia toca el ERP."""
    print(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  pasada COMPLETA...", flush=True)
    r = exportar_costes.ejecutar(pg, solo_pendientes=False, origen="programado")
    with pg.begin() as c:
        c.execute(text("UPDATE core.cfg_coste_vigilancia "
                       "SET ultima_completa = now(), ultima_revision = now()"))
    print(f"          {r['completos']}/{r['articulos']} completos en {r['duracion']}s"
          f"  ->  {exportar_costes.SALIDA.name} actualizado", flush=True)


def revisar(pg, cada_dias, hhmm):
    """Hace la pasada si toca. Devuelve True si la ha hecho."""
    if not toca_pasada(ultima_pasada(pg), dt.datetime.now(), cada_dias, hhmm):
        return False
    pasada(pg)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--bucle", type=int, metavar="SEG",
                    help="vigilar cada SEG segundos en vez de salir")
    ap.add_argument("--cada-dias", type=int, default=3, metavar="N",
                    help="dias entre pasadas enteras (3 por defecto)")
    ap.add_argument("--a", dest="hora", default="05:00", metavar="HH:MM",
                    help="hora a partir de la que se hace la pasada (05:00)")
    args = ap.parse_args()

    pg = get_pg_engine()
    if not args.bucle:
        revisar(pg, args.cada_dias, args.hora)
        return

    print(f"Pasada completa cada {args.cada_dias} dias a partir de las {args.hora}."
          f" Entre medias no se consulta el ERP.", flush=True)
    while True:
        try:
            revisar(pg, args.cada_dias, args.hora)
        except Exception as ex:                                # noqa: BLE001
            # una caida de red no debe tumbar la vigilancia: se reintenta
            print(f"  ERROR: {str(ex)[:160]}", flush=True)
        time.sleep(args.bucle)


if __name__ == "__main__":
    main()
