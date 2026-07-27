"""Conexion al ERP GOMEZYCRESPO (SQL Server) en solo lectura."""
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

load_dotenv()  # lee el .env de este directorio

_engine = None


def get_engine():
    """Devuelve un engine SQLAlchemy (cacheado) contra el ERP."""
    global _engine
    if _engine is None:
        url = URL.create(
            "mssql+pyodbc",
            username=os.getenv("SQLSERVER_USER"),
            password=os.getenv("SQLSERVER_PASS"),
            host=os.getenv("SQLSERVER_HOST"),
            database=os.getenv("SQLSERVER_DB"),
            query={
                "driver": "ODBC Driver 18 for SQL Server",
                "TrustServerCertificate": "yes",
            },
        )
        _engine = create_engine(url, pool_pre_ping=True)
    return _engine
