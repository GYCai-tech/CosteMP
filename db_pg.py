"""Conexion a la replica analitica gyc_analytics (Postgres).

Es donde se publican los costes calculados para que los consuman Excel y
Power BI. El ERP (db.py) sigue siendo de SOLO LECTURA y solo lo usa el calculo.
"""
import os

import sqlalchemy
from dotenv import load_dotenv

load_dotenv()

_engine = None


def get_pg_engine():
    """Devuelve un engine SQLAlchemy (cacheado) contra gyc_analytics."""
    global _engine
    if _engine is None:
        url = (
            f"postgresql+psycopg2://{os.getenv('PG_USER')}:{os.getenv('PG_PASS')}"
            f"@{os.getenv('PG_HOST')}:{os.getenv('PG_PORT', '5432')}/{os.getenv('PG_DB')}"
        )
        _engine = sqlalchemy.create_engine(url, pool_pre_ping=True, pool_size=5)
    return _engine
