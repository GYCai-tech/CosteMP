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
from openpyxl.styles import PatternFill, Font
from sqlalchemy import text

from db import get_engine

# CTE recursivo. Costeo de compras (criterio: TIPO DE APROVISIONAMIENTO):
#   - se despieza TODO lo que tiene escandallo, a todos los niveles, por DOS vias:
#       (1) FASE activa (fabricacion) -> Fases_Entradas.
#       (2) CONJUNTO (Articulos_Conjuntos) -> SOLO si el articulo no tiene fase activa.
#     Un conjunto sin fase se explota por sus componentes x Unidades (misma logica
#     recursiva que las fases). Prioridad: si hay fase activa manda la fase.
#   - se valora a ultimo precio de compra cada articulo con IdTipoAprovisionamiento
#     (=se compra) que sea HOJA del arbol, y se clasifica por su tipo (MATERIA PRIMA /
#     MERCADERIA / ...). Si el articulo TIENE escandallo no se valora a precio: su
#     coste sale de explotar sus componentes; valorarlo ademas a precio lo contaria
#     dos veces (caso COMEDERO 2B MIXTO, tipo ALMACEN con compras historicas).
#     UNICA EXCEPCION: TRABAJOS EXTERNOS (tipo 5), donde el precio es el coste del
#     servicio subcontratado y SI se suma al de sus propios materiales.
SQL = text("""
WITH
-- Precio de compra por articulo: UNICAMENTE el ULTIMO ALBARAN con precio real
-- (Pedidos_Prov_Lineas, la compra mas reciente), con su descuento.
--
-- NO se usa la lista de precios de proveedor (Listas_Precios_Prov_Art) como
-- fallback: una tarifa es lo que el proveedor PIDE, no lo que se ha pagado.
-- Costear con ella mete importes que nadie ha desembolsado y que ademas pueden
-- estar en otra unidad que el albaran (p.ej. cordon a 0,026 EUR/m en tarifa
-- frente a 24,79 EUR/rollo en albaran). Si no hay compra, no hay coste: la
-- linea sale marcada como "sin precio" para que se vea el hueco.
precio AS (
    SELECT IdArticulo, Precio, Descuento, 'albaran' AS Fuente FROM (
        SELECT IdArticulo, Precio_EURO AS Precio, Descuento,
               ROW_NUMBER() OVER (PARTITION BY IdArticulo
                                  ORDER BY FechaAlbaran DESC, IdPedido DESC) AS rn
        FROM dbo.Pedidos_Prov_Lineas
        WHERE FechaAlbaran IS NOT NULL AND Precio_EURO > 0
    ) t
    WHERE rn = 1
),
-- Tiempo de operacion (min/pieza) por articulo: media de los bonos de produccion
-- (TotalMinutos/TotalPiezas) con fase activa. Es la base del coste de mano de obra.
tiempo_op AS (
    SELECT obs.IdArticulo,
           AVG(CAST(ob.TotalMinutos AS float) / NULLIF(ob.TotalPiezas, 0)) AS TiempoMin
    FROM dbo.Ordenes_Bonos_Salidas obs
    INNER JOIN dbo.Ordenes_Articulos oa ON oa.IdArticulo = obs.IdArticulo AND oa.IdOrden = obs.IdOrden
    INNER JOIN dbo.Fases f ON f.IdFase = oa.IdFase AND f.Activa <> 0
    INNER JOIN dbo.Ordenes_Bonos ob ON ob.IdBono = obs.IdBono AND ob.IdOrden = obs.IdOrden
    WHERE ob.TotalPiezas > 0
    GROUP BY obs.IdArticulo
),
-- Articulos cuya fase activa NO declara ninguna operacion de fabricacion:
-- todos sus trabajos tienen OrdenTrabajo = 0 (los "Sin operacion" del ERP).
-- La fase existe solo para llevar los materiales. Que no tengan tiempo NO es
-- un hueco de datos, es lo correcto: no hay mano de obra que imputar.
-- Se exige que NO haya ningun trabajo con OrdenTrabajo = 1 (hay 5 fases mixtas).
sin_operacion AS (
    SELECT DISTINCT fs.IdArticulo
    FROM dbo.Fases_Salidas fs
    INNER JOIN dbo.Fases f ON f.IdFase = fs.IdFase AND f.Activa <> 0
    WHERE EXISTS (SELECT 1 FROM dbo.Trabajos_Fases tf
                  WHERE tf.IdFase = fs.IdFase AND tf.OrdenTrabajo = 0)
      AND NOT EXISTS (SELECT 1 FROM dbo.Trabajos_Fases tf
                      WHERE tf.IdFase = fs.IdFase AND tf.OrdenTrabajo = 1)
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
-- Marca cada fila: cantidad convertida (gr->kg), si es hoja (sin escandallo) y si
-- se puede valorar a precio de compra.
marcado AS (
    SELECT d.*,
        CASE WHEN d.Unidad = 'gr' THEN d.Cant / 1000.0 ELSE d.Cant END AS CantConv,
        CASE WHEN EXISTS (SELECT 1 FROM componentes c2 WHERE c2.Padre = d.IdArticulo)
             THEN 0 ELSE 1 END AS EsHoja,
        -- Valorable a precio de albaran. Dos vias:
        --  (a) tiene TIPO DE APROVISIONAMIENTO (=alguien confirmo que se compra)
        --      y es hoja, o es TRABAJO EXTERNO (tipo 5, el precio es el servicio
        --      y se suma a sus materiales). Un articulo con escandallo se costea
        --      por sus componentes; valorarlo ademas a precio lo contaria dos veces.
        --  (b) NO tiene tipo pero tampoco ha tenido NUNCA fase de fabricacion:
        --      solo puede ser una compra a la que le falta rellenar el tipo en el
        --      ERP. Se valora igual (si hay albaran) en vez de contarla como 0 EUR
        --      en silencio. Las piezas con fases desactivadas quedan fuera a
        --      proposito: esas siguen avisando como "sin escandallo".
        CASE WHEN a2.IdTipoAprovisionamiento IS NOT NULL
                  AND (a2.IdTipoAprovisionamiento = 5
                       OR NOT EXISTS (SELECT 1 FROM componentes c3 WHERE c3.Padre = d.IdArticulo))
             THEN 1
             WHEN a2.IdTipoAprovisionamiento IS NULL
                  AND NOT EXISTS (SELECT 1 FROM componentes c4 WHERE c4.Padre = d.IdArticulo)
                  AND NOT EXISTS (SELECT 1 FROM dbo.Fases_Salidas fs3
                                  WHERE fs3.IdArticulo = d.IdArticulo)
             THEN 1
             ELSE 0 END AS EsValorable
    FROM desglose d
    LEFT JOIN dbo.Articulos a2 ON a2.IdArticulo = d.IdArticulo
)
SELECT
    m.Nivel,
    m.Articulo,
    m.IdArticulo,
    m.Componente,
    m.CantConv AS Cant,
    m.Unidad,
    t.Descrip AS TipoCompra,                       -- MATERIA PRIMA / MERCADERIA / ...
    -- Precio = ultimo ALBARAN de compra, solo en las filas valorables
    -- (ver EsValorable en 'marcado'). Nunca tarifa.
    CASE WHEN m.EsValorable = 1 THEN p.Precio END    AS PrecioCompra,
    CASE WHEN m.EsValorable = 1 THEN p.Descuento END AS Descuento,
    CASE WHEN m.EsValorable = 1 AND p.Precio IS NOT NULL
         THEN p.Fuente END AS PrecioFuente,       -- siempre 'albaran'
    CASE WHEN m.EsValorable = 1 AND p.Precio IS NOT NULL
         THEN m.CantConv * (p.Precio - p.Precio * COALESCE(p.Descuento, 0) / 100.0)
    END AS Coste,
    CASE WHEN m.EsValorable = 1 AND p.Precio IS NOT NULL THEN 1 ELSE 0 END AS Valorado,
    -- valorable pero SIN albaran de compra: hueco de coste real, hay que verlo
    CASE WHEN m.EsValorable = 1 AND p.Precio IS NULL
         THEN 1 ELSE 0 END AS SinPrecio,
    -- 1 si es una PIEZA FABRICADA (tiene fases) sin escandallo ACTIVO para
    -- costearla y sin valorar: su fase esta desactivada/vacia -> no se puede
    -- costear aqui. (NO marca tornilleria/compras: esas son sin-precio/sin-tipo.)
    CASE WHEN m.EsHoja = 1
              AND NOT (m.EsValorable = 1 AND p.Precio IS NOT NULL)
              AND EXISTS (SELECT 1 FROM dbo.Fases_Salidas fss WHERE fss.IdArticulo = m.IdArticulo)
         THEN 1 ELSE 0 END AS SinEscandallo,
    -- subconjunto de SinEscandallo que ademas es componente de un conjunto
    -- (su coste viene por el conjunto).
    CASE WHEN m.EsHoja = 1
              AND NOT (m.EsValorable = 1 AND p.Precio IS NOT NULL)
              AND EXISTS (SELECT 1 FROM dbo.Fases_Salidas fss WHERE fss.IdArticulo = m.IdArticulo)
              AND EXISTS (SELECT 1 FROM dbo.Articulos_Conjuntos ac WHERE ac.IdArticulo = m.IdArticulo)
         THEN 1 ELSE 0 END AS DeConjunto,
    tp.TiempoMin AS TiempoOp,                       -- min/pieza de operacion (mano de obra)
    -- 1 = su fase no declara operacion ("Sin operacion"): no lleva mano de obra
    -- y por tanto no debe avisarse como "sin tiempo".
    CASE WHEN so.IdArticulo IS NOT NULL THEN 1 ELSE 0 END AS SinOperacion,
    m.Ruta AS Ruta                                 -- ruta jerarquica para el arbol
FROM marcado m
LEFT JOIN dbo.Articulos  art ON art.IdArticulo = m.IdArticulo
LEFT JOIN precio         p   ON p.IdArticulo   = m.IdArticulo
LEFT JOIN tiempo_op      tp  ON tp.IdArticulo  = m.IdArticulo
LEFT JOIN sin_operacion  so  ON so.IdArticulo  = m.IdArticulo
LEFT JOIN dbo.Articulos_Tipos_Aprovisionamiento t
       ON t.IdTipoAprovisionamiento = art.IdTipoAprovisionamiento
ORDER BY m.Ruta
OPTION (MAXRECURSION 0);   -- sin limite de niveles
""")


# Articulos.Estado: 0 = activo, 1 = dado de baja / descatalogado.
# (FechaBajaArt existe pero esta practicamente sin usar: 1 registro en toda la tabla,
#  asi que Estado es el unico indicador fiable de baja.)
SQL_BUSCAR = text("""
SELECT TOP (:limite)
    a.IdArticulo,
    a.Descrip,
    a.Estado AS Baja,                          -- 1 = descatalogado
    CASE WHEN EXISTS (
        SELECT 1 FROM dbo.Fases_Salidas fs
        INNER JOIN dbo.Fases f ON f.IdFase = fs.IdFase AND f.Activa <> 0
        WHERE fs.IdArticulo = a.IdArticulo
    ) OR EXISTS (
        SELECT 1 FROM dbo.Articulos_Conjuntos ac WHERE ac.IdArticuloPadre = a.IdArticulo
    ) THEN 1 ELSE 0 END AS Fabricable
FROM dbo.Articulos a
WHERE (a.IdArticulo LIKE :q OR a.Descrip LIKE :q)
  AND (:incluir_bajas = 1 OR a.Estado = 0)     -- por defecto oculta los de baja
ORDER BY Fabricable DESC, a.IdArticulo
""")


def buscar_articulos(q: str, limite: int = 50, incluir_bajas: bool = False) -> pd.DataFrame:
    """Busca articulos por codigo o descripcion. Los fabricables van primero.

    Por defecto solo devuelve articulos ACTIVOS (Estado = 0); con incluir_bajas=True
    tambien salen los descatalogados, marcados con la columna Baja."""
    q = (q or "").strip()
    limite = max(1, min(int(limite), 200))
    with get_engine().connect() as cn:
        return pd.read_sql(SQL_BUSCAR, cn, params={"q": f"%{q}%", "limite": limite,
                                                   "incluir_bajas": 1 if incluir_bajas else 0})


SQL_NOMBRE = text("SELECT Descrip FROM dbo.Articulos WHERE IdArticulo = :codigo")


def nombre_articulo(codigo: str) -> str:
    """Devuelve la descripcion del articulo (cadena vacia si no existe)."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))
    with get_engine().connect() as cn:
        r = cn.execute(SQL_NOMBRE, {"codigo": codigo}).fetchone()
    return r[0] if r and r[0] else ""


SQL_TIEMPO = text("""
SELECT AVG(CAST(ob.TotalMinutos AS float) / NULLIF(ob.TotalPiezas, 0)) AS TiempoMin
FROM dbo.Ordenes_Bonos_Salidas obs
INNER JOIN dbo.Ordenes_Articulos oa ON oa.IdArticulo = obs.IdArticulo AND oa.IdOrden = obs.IdOrden
INNER JOIN dbo.Fases f ON f.IdFase = oa.IdFase AND f.Activa <> 0
INNER JOIN dbo.Ordenes_Bonos ob ON ob.IdBono = obs.IdBono AND ob.IdOrden = obs.IdOrden
WHERE ob.TotalPiezas > 0 AND obs.IdArticulo = :codigo
""")


def tiempo_operacion(codigo: str):
    """Tiempo medio de operacion (min/pieza) del articulo (para el nodo raiz)."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))
    with get_engine().connect() as cn:
        r = cn.execute(SQL_TIEMPO, {"codigo": codigo}).fetchone()
    return float(r[0]) if r and r[0] is not None else None


SQL_SIN_OPERACION = text("""
SELECT CASE WHEN EXISTS (
    SELECT 1 FROM dbo.Fases_Salidas fs
    INNER JOIN dbo.Fases f ON f.IdFase = fs.IdFase AND f.Activa <> 0
    WHERE fs.IdArticulo = :codigo
      AND EXISTS (SELECT 1 FROM dbo.Trabajos_Fases tf
                  WHERE tf.IdFase = fs.IdFase AND tf.OrdenTrabajo = 0)
      AND NOT EXISTS (SELECT 1 FROM dbo.Trabajos_Fases tf
                      WHERE tf.IdFase = fs.IdFase AND tf.OrdenTrabajo = 1)
) THEN 1 ELSE 0 END
""")


def sin_operacion(codigo: str) -> int:
    """1 si la fase activa del articulo no declara operacion de fabricacion
    (todos sus trabajos son 'Sin operacion'). Para el nodo raiz del arbol."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))
    with get_engine().connect() as cn:
        r = cn.execute(SQL_SIN_OPERACION, {"codigo": codigo}).fetchone()
    return int(r[0]) if r else 0


SQL_ESCANDALLO = text("""
SELECT TRY_CONVERT(float, fe.Cantidad) AS Cantidad,
       fe.IdArticulo,
       fe.Descrip AS Descripcion,
       fe.IdTipoUnidad AS Unidad
FROM dbo.Fases f
INNER JOIN dbo.Fases_Salidas  fs ON fs.IdFase = f.IdFase AND f.Activa <> 0
INNER JOIN dbo.Fases_Entradas fe ON fe.IdFase = f.IdFase
WHERE fs.IdArticulo = :codigo
ORDER BY fe.Descrip
""")


def escandallo_directo(codigo: str) -> pd.DataFrame:
    """Componentes DIRECTOS (nivel 0) del articulo, de su fase activa."""
    codigo = re.sub(r"[^A-Za-z0-9]", "", str(codigo))
    with get_engine().connect() as cn:
        return pd.read_sql(SQL_ESCANDALLO, cn, params={"codigo": codigo})


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


RATE_OP = 17.0 / 60.0   # 17 EUR/h operario -> EUR/min


def _num(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) else float(v)


def _s(v):
    """Campos de texto: NaN -> None. (NaN es un float TRUTHY, asi que colarlo en
    un `x or defecto` cortocircuita el defecto y deja la celda vacia.)"""
    return None if v is None or (isinstance(v, float) and pd.isna(v)) else v


def construir_arbol(df, codigo, nombre, tiempo_raiz=None, sin_op_raiz=0):
    """Arbol del escandallo con coste de MATERIAL (hojas) y de OPERACION (ramas
    fabricadas = tiempo x 17 EUR/h) y el rollup material+operacion por nodo."""
    root = {"id": codigo, "nombre": nombre, "cant": 1.0, "unidad": None, "tipo": None,
            "precio": None, "fuente": None, "de_conjunto": 0, "sin_escandallo": 0,
            "sin_operacion": int(sin_op_raiz or 0),
            "tiempo": tiempo_raiz, "coste": 0.0, "hijos": []}
    nodos = {f"|{codigo}|": root}

    for r in df.sort_values("Ruta").to_dict(orient="records"):
        ruta = r["Ruta"]
        segs = [s for s in ruta.split("|") if s]
        if segs == [codigo]:
            root["coste"] = _num(r["Coste"]) or 0.0
            root["cant"] = _num(r["Cant"]) or 1.0; root["unidad"] = _s(r["Unidad"])
            root["tipo"] = _s(r["TipoCompra"]); root["precio"] = _num(r["PrecioCompra"])
            root["fuente"] = _s(r.get("PrecioFuente"))
            continue
        nodo = {"id": r["IdArticulo"], "nombre": r["Componente"],
                "cant": _num(r["Cant"]), "unidad": _s(r["Unidad"]),
                "tipo": _s(r["TipoCompra"]), "precio": _num(r["PrecioCompra"]),
                "fuente": _s(r.get("PrecioFuente")),
                "de_conjunto": int(r.get("DeConjunto", 0) or 0),
                "sin_escandallo": int(r.get("SinEscandallo", 0) or 0),
                "sin_operacion": int(r.get("SinOperacion", 0) or 0),
                "tiempo": _num(r.get("TiempoOp")),
                "coste": _num(r["Coste"]) or 0.0, "hijos": []}
        nodos[ruta] = nodo
        padre_key = ("|" + "|".join(segs[:-1]) + "|") if len(segs) > 1 else f"|{codigo}|"
        nodos.get(padre_key, root)["hijos"].append(nodo)

    def rollup(n):
        n["es_hoja"] = not n["hijos"]
        if n["es_hoja"]:
            n["coste_op"] = 0.0; n["sin_tiempo"] = 0
        elif n.get("tiempo") is not None:
            # hay tiempo medido en bonos: manda el dato real aunque la fase
            # este declarada "sin operacion".
            n["coste_op"] = round(n["tiempo"] * (n["cant"] or 0) * RATE_OP, 4); n["sin_tiempo"] = 0
        elif n.get("sin_operacion"):
            # su fase no declara operacion: 0 EUR de mano de obra es CORRECTO,
            # no es un dato que falte -> no se avisa como "sin tiempo".
            n["coste_op"] = 0.0; n["sin_tiempo"] = 0
        else:
            n["coste_op"] = 0.0; n["sin_tiempo"] = 1
        mat = n["coste"] or 0.0
        op = n["coste_op"]
        for h in n["hijos"]:
            rollup(h)
            mat += h["coste_mat"]; op += h["coste_op_total"]
        # 6 decimales: con 4, un coste de 8,5e-06 EUR se redondeaba a 0,0 exacto
        # y desaparecia del arbol (ademas de dispararse como "sin coste").
        n["coste_mat"] = round(mat, 6)
        n["coste_op_total"] = round(op, 6)
        n["coste_total"] = round(mat + op, 6)

    rollup(root)
    return root


def aplanar_arbol(arbol):
    """Aplana el arbol en orden pre-order: lista de (profundidad, nodo)."""
    filas = []

    def walk(n, depth):
        filas.append((depth, n))
        for h in n["hijos"]:
            walk(h, depth + 1)

    walk(arbol, 0)
    return filas


def _categoria(n):
    """Categoria de una fila para colorear (misma que en la web)."""
    if n["hijos"]:
        return "noesc" if n.get("sin_tiempo") else "rama"
    if n.get("de_conjunto"): return "conjunto"
    if n.get("sin_escandallo"): return "noesc"
    # hueco = no hay precio de albaran, no que falte el tipo ni que el importe
    # sea minimo (el tipo ya no se exige para valorar)
    return "mp" if n.get("precio") is not None else "sinprecio"


_FILLS = {"mp": "FEF3C7", "sinprecio": "FDE0E0", "conjunto": "EDE9FE",
          "noesc": "DBEAFE", "total": "E2E8F0"}


def _resaltar_sin_precio(ws, flags, ncols):
    """Rellena en rojo las filas cuyo flag SinPrecio es 1 (cabecera en fila 1)."""
    rojo = PatternFill("solid", fgColor="FFC7CE")
    for pos, sinp in enumerate(flags):
        if sinp:
            for col in range(1, ncols + 1):
                ws.cell(row=pos + 2, column=col).fill = rojo


def exportar_excel(df: pd.DataFrame, codigo: str, nombre: str, destino) -> None:
    """Escribe el Excel (tres hojas): Desglose (arbol indentado con material y
    operacion, coloreado por categoria), Comprados y Resumen (material+operacion+total)."""
    arbol = construir_arbol(df, codigo, nombre, tiempo_operacion(codigo), sin_operacion(codigo))
    filas = aplanar_arbol(arbol)

    # ---- Hoja Desglose: el arbol indentado, con material y operacion ----
    reg, cats = [], []
    for depth, n in filas:
        hoja = not n["hijos"]
        tipo_txt = ("de conjunto" if n.get("de_conjunto")
                    else "sin escandallo" if n.get("sin_escandallo")
                    else (n.get("tipo") or
                          ("" if hoja
                           else "fabricado sin operación" if n.get("sin_operacion")
                           else "fabricado")))
        reg.append({
            "Nivel": depth,
            "Componente": ("    " * depth) + (n.get("nombre") or ""),
            "IdArticulo": n["id"],
            "Cant": n.get("cant"),
            "Ud": n.get("unidad"),
            "Tipo": tipo_txt,
            "Precio €": n.get("precio") if hoja else None,
            "Coste MP €": n.get("coste") if hoja else None,
            "Tiempo min": (None if hoja else n.get("tiempo")),
            "Coste operación €": (None if hoja else n.get("coste_op")),
            "Coste total €": n.get("coste_total"),
        })
        cats.append(_categoria(n))

    coste_mat = arbol["coste_mat"]; coste_op = arbol["coste_op_total"]; coste_tot = arbol["coste_total"]
    reg.append({"Nivel": None, "Componente": "TOTAL", "IdArticulo": "", "Cant": None,
                "Ud": "", "Tipo": "", "Precio €": None, "Coste MP €": coste_mat,
                "Tiempo min": None, "Coste operación €": coste_op, "Coste total €": coste_tot})
    cats.append("total")
    desg = pd.DataFrame(reg)

    # ---- Comprados y Resumen ----
    val = df[df["Valorado"] == 1]
    coste_mp = round(float(val.loc[val["TipoCompra"] == "MATERIA PRIMA", "Coste"].sum()), 4)
    coste_merc = round(float(val.loc[val["TipoCompra"] == "MERCADERIA", "Coste"].sum()), 4)
    coste_otros = round(coste_mat - coste_mp - coste_merc, 4)

    comprados = articulos_comprados(df)
    comprados_vis = comprados.drop(columns=["SinPrecio"])

    resumen = pd.DataFrame({
        "Concepto": ["Artículo", "Nombre", "Líneas", "Niveles",
                     "Coste MATERIA PRIMA (€)", "Coste MERCADERÍA (€)", "Coste otras compras (€)",
                     "COSTE MATERIAL (€)", "COSTE OPERACIÓN (€)", "COSTE TOTAL (€)",
                     "Comprables sin precio"],
        "Valor": [codigo, nombre, len(df), int(df["Nivel"].max()),
                  coste_mp, coste_merc, coste_otros,
                  coste_mat, coste_op, coste_tot, int(df["SinPrecio"].sum())],
    })

    with pd.ExcelWriter(destino, engine="openpyxl") as xls:
        desg.to_excel(xls, sheet_name="Desglose", index=False)
        comprados_vis.to_excel(xls, sheet_name="Comprados", index=False)
        resumen.to_excel(xls, sheet_name="Resumen", index=False)

        ws = xls.sheets["Desglose"]
        ncols = len(desg.columns)
        fills = {k: PatternFill("solid", fgColor=v) for k, v in _FILLS.items()}
        for i, cat in enumerate(cats):
            fill = fills.get(cat)
            for c in range(1, ncols + 1):
                cell = ws.cell(row=i + 2, column=c)
                if fill:
                    cell.fill = fill
                if cat == "total":
                    cell.font = Font(bold=True)
        for c in range(1, ncols + 1):     # cabecera en negrita
            ws.cell(row=1, column=c).font = Font(bold=True)

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
