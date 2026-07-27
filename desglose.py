"""
Despiece (lista de materiales) multinivel de un articulo del ERP GOMEZYCRESPO.

Explota TODOS los niveles con un CTE recursivo, multiplicando las cantidades
en cascada, y marca que lineas son materia prima (hoja del arbol).

Uso:
    py desglose.py 12101021
    py desglose.py 12101021 -o salidas/mi_desglose.xlsx
"""
import argparse
import os
import re

import pandas as pd
from openpyxl.styles import PatternFill
from sqlalchemy import text

from db import get_engine

# CTE recursivo. Costeo de compras (criterio: TIPO DE APROVISIONAMIENTO):
#   - se despieza TODO lo que tiene escandallo, a todos los niveles, por DOS vias:
#       (1) FASE activa (fabricacion) -> Fases_Entradas.
#       (2) CONJUNTO (Articulos_Conjuntos) -> SOLO si el articulo no tiene fase activa.
#     Un conjunto sin fase se explota por sus componentes x Unidades (misma logica
#     recursiva que las fases). Prioridad: si hay fase activa manda la fase.
#   - se valora a ultimo precio de compra cada articulo con IdTipoAprovisionamiento
#     (=se compra), en cualquier nivel, y se clasifica por su tipo (MATERIA PRIMA /
#     MERCADERIA / ...). NOTA: un semielaborado que ademas tenga tipo se cuenta a su
#     precio y por sus materiales (caso "trabajos externos" / fabricar-o-comprar).
SQL = text("""
WITH
-- Ultimo precio de compra por articulo: la compra mas reciente con FechaAlbaran
-- (equivale al Last() de Access, pero determinista: ROW_NUMBER por fecha desc).
precio_ultimo AS (
    SELECT IdArticulo, Precio, Descuento FROM (
        SELECT
            IdArticulo,
            Precio_EURO AS Precio,
            Descuento,
            ROW_NUMBER() OVER (PARTITION BY IdArticulo
                               ORDER BY FechaAlbaran DESC, IdPedido DESC) AS rn
        FROM dbo.Pedidos_Prov_Lineas
        -- ultimo albaran con precio REAL: se ignoran lineas sin fecha y con precio 0
        WHERE FechaAlbaran IS NOT NULL AND Precio_EURO > 0
    ) t
    WHERE rn = 1
),
-- Componentes de cada articulo, unificando las dos vias de escandallo.
componentes AS (
    -- (1) via FASE ACTIVA (fabricacion)
    SELECT
        fs.IdArticulo AS Padre,
        fe.IdArticulo AS Hijo,
        fe.Descrip    AS Descripcion,
        fe.IdTipoUnidad AS Unidad,
        TRY_CONVERT(FLOAT, fe.Cantidad) AS Cantidad
    FROM dbo.Fases f
    INNER JOIN dbo.Fases_Salidas  fs ON fs.IdFase = f.IdFase AND f.Activa <> 0
    INNER JOIN dbo.Fases_Entradas fe ON fe.IdFase = f.IdFase

    UNION ALL

    -- (2) via CONJUNTO, SOLO si el articulo padre NO tiene fase activa
    SELECT
        ac.IdArticuloPadre AS Padre,
        ac.IdArticulo      AS Hijo,
        a.Descrip          AS Descripcion,
        CAST(NULL AS VARCHAR(20)) AS Unidad,
        TRY_CONVERT(FLOAT, ac.Unidades) AS Cantidad
    FROM dbo.Articulos_Conjuntos ac
    INNER JOIN dbo.Articulos a ON a.IdArticulo = ac.IdArticulo
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Fases_Salidas fs2
                      INNER JOIN dbo.Fases f2 ON f2.IdFase = fs2.IdFase AND f2.Activa <> 0
                      WHERE fs2.IdArticulo = ac.IdArticuloPadre)
),
desglose AS (
    -- NIVEL 0: componentes directos del articulo (por fase o por conjunto)
    SELECT
        CAST(:codigo AS VARCHAR(50)) AS Articulo,
        c.Hijo        AS IdArticulo,
        c.Cantidad    AS Cant,
        c.Descripcion AS Componente,
        0 AS Nivel,
        c.Unidad AS Unidad,
        CAST('|' + CAST(:codigo AS VARCHAR(50)) + '|' + c.Hijo + '|' AS VARCHAR(4000)) AS Ruta
    FROM componentes c
    WHERE c.Padre = :codigo

    UNION ALL

    -- NIVEL 0 alternativo: el PROPIO articulo si NO tiene escandallo (compra directa).
    -- Se muestra a si mismo (cantidad 1) y se valora a su ultimo precio de compra.
    SELECT
        CAST(:codigo AS VARCHAR(50)) AS Articulo,
        a.IdArticulo AS IdArticulo,
        CAST(1 AS FLOAT) AS Cant,
        a.Descrip AS Componente,
        0 AS Nivel,
        CAST(NULL AS VARCHAR(20)) AS Unidad,
        CAST('|' + a.IdArticulo + '|' AS VARCHAR(4000)) AS Ruta
    FROM dbo.Articulos a
    WHERE a.IdArticulo = :codigo
      AND NOT EXISTS (SELECT 1 FROM componentes c WHERE c.Padre = :codigo)

    UNION ALL

    -- NIVELES SIGUIENTES: explota cada componente por su propio escandallo
    SELECT
        d.IdArticulo,
        c.Hijo,
        d.Cant * c.Cantidad,           -- cantidad acumulada (x Unidades en conjuntos)
        c.Descripcion,
        d.Nivel + 1,
        c.Unidad,
        CAST(d.Ruta + c.Hijo + '|' AS VARCHAR(4000))
    FROM desglose d
    INNER JOIN componentes c ON c.Padre = d.IdArticulo
    WHERE d.Ruta NOT LIKE '%|' + c.Hijo + '|%'          -- corta ciclos A->B->A
),
-- Marca cada fila: cantidad convertida (gr->kg) y si es hoja (sin escandallo)
marcado AS (
    SELECT d.*,
        CASE WHEN d.Unidad = 'gr' THEN d.Cant / 1000.0 ELSE d.Cant END AS CantConv,
        CASE WHEN EXISTS (SELECT 1 FROM componentes c2 WHERE c2.Padre = d.IdArticulo)
             THEN 0 ELSE 1 END AS EsHoja
    FROM desglose d
)
SELECT
    m.Nivel,
    m.Articulo,
    m.IdArticulo,
    m.Componente,
    m.CantConv AS Cant,
    m.Unidad,
    t.Descrip AS TipoCompra,                       -- MATERIA PRIMA / MERCADERIA / ...
    -- Criterio de valoracion: TIPO DE APROVISIONAMIENTO (decidido por negocio).
    -- Se valora todo articulo con IdTipoAprovisionamiento (=se compra) que tenga precio.
    CASE WHEN art.IdTipoAprovisionamiento IS NOT NULL THEN p.Precio END    AS PrecioCompra,
    CASE WHEN art.IdTipoAprovisionamiento IS NOT NULL THEN p.Descuento END AS Descuento,
    CASE WHEN art.IdTipoAprovisionamiento IS NOT NULL AND p.Precio IS NOT NULL
         THEN m.CantConv * (p.Precio - p.Precio * COALESCE(p.Descuento, 0) / 100.0)
    END AS Coste,
    CASE WHEN art.IdTipoAprovisionamiento IS NOT NULL AND p.Precio IS NOT NULL THEN 1 ELSE 0 END AS Valorado,
    -- comprable (con tipo) pero SIN precio de compra (hueco de coste)
    CASE WHEN art.IdTipoAprovisionamiento IS NOT NULL AND p.Precio IS NULL
         THEN 1 ELSE 0 END AS SinPrecio,
    -- 1 si es una PIEZA FABRICADA (tiene fases) sin escandallo activo para
    -- costearla, sin valorar, y que es componente de un conjunto: su coste viene
    -- por el conjunto, no aqui. (NO marca tornilleria/compras que solo aparecen
    -- en conjuntos: esas se costean o son sin-precio/sin-tipo normales.)
    CASE WHEN m.EsHoja = 1
              AND NOT (art.IdTipoAprovisionamiento IS NOT NULL AND p.Precio IS NOT NULL)
              AND EXISTS (SELECT 1 FROM dbo.Fases_Salidas fss WHERE fss.IdArticulo = m.IdArticulo)
              AND EXISTS (SELECT 1 FROM dbo.Articulos_Conjuntos ac WHERE ac.IdArticulo = m.IdArticulo)
         THEN 1 ELSE 0 END AS DeConjunto,
    m.Ruta AS Ruta                                 -- ruta jerarquica para el arbol
FROM marcado m
LEFT JOIN dbo.Articulos  art ON art.IdArticulo = m.IdArticulo
LEFT JOIN precio_ultimo  p   ON p.IdArticulo   = m.IdArticulo
LEFT JOIN dbo.Articulos_Tipos_Aprovisionamiento t
       ON t.IdTipoAprovisionamiento = art.IdTipoAprovisionamiento
ORDER BY m.Ruta
OPTION (MAXRECURSION 0);   -- sin limite de niveles
""")


SQL_BUSCAR = text("""
SELECT TOP (:limite)
    a.IdArticulo,
    a.Descrip,
    CASE WHEN EXISTS (
        SELECT 1 FROM dbo.Fases_Salidas fs
        INNER JOIN dbo.Fases f ON f.IdFase = fs.IdFase AND f.Activa <> 0
        WHERE fs.IdArticulo = a.IdArticulo
    ) OR EXISTS (
        SELECT 1 FROM dbo.Articulos_Conjuntos ac WHERE ac.IdArticuloPadre = a.IdArticulo
    ) THEN 1 ELSE 0 END AS Fabricable
FROM dbo.Articulos a
WHERE a.IdArticulo LIKE :q OR a.Descrip LIKE :q
ORDER BY Fabricable DESC, a.IdArticulo
""")


def buscar_articulos(q: str, limite: int = 50) -> pd.DataFrame:
    """Busca articulos por codigo o descripcion. Los fabricables van primero."""
    q = (q or "").strip()
    limite = max(1, min(int(limite), 200))
    with get_engine().connect() as cn:
        return pd.read_sql(SQL_BUSCAR, cn, params={"q": f"%{q}%", "limite": limite})


SQL_NOMBRE = text("SELECT Descrip FROM dbo.Articulos WHERE IdArticulo = :codigo")


def nombre_articulo(codigo: str) -> str:
    """Devuelve la descripcion del articulo (cadena vacia si no existe)."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))
    with get_engine().connect() as cn:
        r = cn.execute(SQL_NOMBRE, {"codigo": codigo}).fetchone()
    return r[0] if r and r[0] else ""


def desglose(codigo: str) -> pd.DataFrame:
    """Devuelve el despiece multinivel del articulo como DataFrame."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))  # sanea el input
    with get_engine().connect() as cn:
        return pd.read_sql(SQL, cn, params={"codigo": codigo})


def articulos_comprados(df: pd.DataFrame) -> pd.DataFrame:
    """Lista plana SOLO de los articulos comprados del escandallo, agregados por
    articulo (suma la cantidad y el coste de todas las ramas donde aparecen)."""
    comprables = df[(df["PrecioCompra"].notna()) | (df["SinPrecio"] == 1)]
    if comprables.empty:
        return pd.DataFrame(columns=["CODIGO", "Descripcion", "Unidad", "Cantidad",
                                     "PrecioCompra", "Descuento", "Coste", "SinPrecio"])
    agg = (comprables.groupby("IdArticulo", as_index=False)
           .agg(Descripcion=("Componente", "first"),
                Tipo=("TipoCompra", "first"),
                Unidad=("Unidad", "first"),
                Cantidad=("Cant", "sum"),
                PrecioCompra=("PrecioCompra", "first"),
                Descuento=("Descuento", "first"),
                Coste=("Coste", "sum"),
                SinPrecio=("SinPrecio", "max"))
           .rename(columns={"IdArticulo": "CODIGO"}))
    return agg.sort_values("Coste", ascending=False, na_position="last")


def _resaltar_sin_precio(ws, flags, ncols):
    """Rellena en rojo las filas cuyo flag SinPrecio es 1 (cabecera en fila 1)."""
    rojo = PatternFill("solid", fgColor="FFC7CE")
    for pos, sinp in enumerate(flags):
        if sinp:
            for col in range(1, ncols + 1):
                ws.cell(row=pos + 2, column=col).fill = rojo


def exportar_excel(df: pd.DataFrame, codigo: str, nombre: str, destino) -> None:
    """Escribe el Excel (tres hojas) en `destino` (ruta o buffer en memoria).

    - Desglose: el despiece completo multinivel.
    - Comprados: solo los materiales comprados, agregados por articulo.
    - Resumen: cabecera con totales.
    Resalta en rojo las filas de comprables sin precio (huecos de coste).
    """
    coste_total = round(float(df["Coste"].sum()), 4)
    sin_precio = int(df["SinPrecio"].sum())

    # Subtotales por tipo de compra
    val = df[df["Valorado"] == 1]
    coste_mp = round(float(val.loc[val["TipoCompra"] == "MATERIA PRIMA", "Coste"].sum()), 4)
    coste_merc = round(float(val.loc[val["TipoCompra"] == "MERCADERIA", "Coste"].sum()), 4)
    coste_otros = round(coste_total - coste_mp - coste_merc, 4)

    df_vis = df.drop(columns=["SinPrecio", "Ruta", "DeConjunto"], errors="ignore")  # auxiliares
    comprados = articulos_comprados(df)
    comprados_vis = comprados.drop(columns=["SinPrecio"])

    resumen = pd.DataFrame({
        "Concepto": ["Artículo", "Nombre", "Líneas totales", "Niveles",
                     "Líneas valoradas (compradas)", "Artículos comprados",
                     "Comprables sin precio",
                     "Coste MATERIA PRIMA (€)", "Coste MERCADERÍA (€)",
                     "Coste otras compras (€)", "COSTE TOTAL (€)"],
        "Valor": [codigo, nombre, len(df), int(df["Nivel"].max()),
                  int((df["Valorado"] == 1).sum()), len(comprados),
                  sin_precio, coste_mp, coste_merc, coste_otros, coste_total],
    })

    with pd.ExcelWriter(destino, engine="openpyxl") as xls:
        df_vis.to_excel(xls, sheet_name="Desglose", index=False)
        comprados_vis.to_excel(xls, sheet_name="Comprados", index=False)
        resumen.to_excel(xls, sheet_name="Resumen", index=False)

        _resaltar_sin_precio(xls.sheets["Desglose"], df["SinPrecio"].tolist(), len(df_vis.columns))
        _resaltar_sin_precio(xls.sheets["Comprados"], comprados["SinPrecio"].tolist(), len(comprados_vis.columns))


def main():
    ap = argparse.ArgumentParser(description="Despiece multinivel de un articulo del ERP")
    ap.add_argument("codigo", help="Codigo de articulo (ej: 12101021)")
    ap.add_argument("-o", "--output", help="Ruta del Excel de salida")
    args = ap.parse_args()

    df = desglose(args.codigo)
    if df.empty:
        print(f"El articulo '{args.codigo}' no tiene fases activas / desglose.")
        return

    nombre = nombre_articulo(args.codigo)
    ruta = args.output or os.path.join("salidas", f"desglose_{args.codigo}.xlsx")
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    exportar_excel(df, args.codigo, nombre, ruta)
    print(f"Articulo {args.codigo} ({nombre}): {len(df)} lineas, {df['Nivel'].max()} niveles.")
    print(f"Coste total materia prima: {df['Coste'].sum():.4f} EUR")
    print(f"Excel generado en: {os.path.abspath(ruta)}")


if __name__ == "__main__":
    main()
