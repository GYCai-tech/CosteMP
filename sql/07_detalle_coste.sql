-- ---------------------------------------------------------------------------
-- El coste de cada articulo partido por pieza, y su historico.
--
-- Hasta ahora cada recalculo publicaba solo los totales y sustituia los del
-- anterior: no habia forma de saber cuanto costaba un articulo hace un mes, ni
-- que pieza explicaba una subida, ni en que articulos entra una chapa. Con el
-- detalle por pieza (desglose.lineas_coste) salen las tres cosas sin volver
-- al ERP.
--
-- Lo escribe exportar_costes.publicar en la misma transaccion que el resto, y
-- aplica este fichero antes de escribir (es idempotente): asi un despliegue
-- nuevo no rompe la pasada aunque nadie haya lanzado --crear.
-- ---------------------------------------------------------------------------

-- El detalle VIGENTE: una fila por (articulo, pieza). Se sustituye en cada
-- recalculo, igual que dim_coste_escandallo.
CREATE TABLE IF NOT EXISTS core.fact_coste_linea (
    idarticulo       varchar(20) NOT NULL,
    idpieza          varchar(20) NOT NULL,
    descripcion      text,
    tipo             text,
    unidad           text,
    cantidad         numeric(18,6),
    precio           numeric(14,4),
    fuente           text,
    coste_material   numeric(16,6) NOT NULL DEFAULT 0,
    tiempo_min       numeric(16,6) NOT NULL DEFAULT 0,
    coste_operacion  numeric(16,6) NOT NULL DEFAULT 0,
    fecha_calculo    timestamp NOT NULL,
    PRIMARY KEY (idarticulo, idpieza)
);
CREATE INDEX IF NOT EXISTS ix_coste_linea_pieza ON core.fact_coste_linea (idpieza);
COMMENT ON TABLE core.fact_coste_linea IS
  'Coste vigente de cada articulo del catalogo partido por pieza (material propio y operacion). Suma por articulo = coste de dim_coste_escandallo.';
COMMENT ON COLUMN core.fact_coste_linea.cantidad IS
  'Por unidad del articulo, multiplicada en cascada y sumada en todas las ramas donde aparece la pieza.';

-- Los totales de cada articulo en CADA recalculo: pocas filas (~550 por
-- pasada), asi que se guardan siempre.
CREATE TABLE IF NOT EXISTS core.hist_coste_articulo (
    idarticulo       varchar(20) NOT NULL,
    fecha_calculo    timestamp   NOT NULL,
    coste_material   numeric(14,4),
    coste_operacion  numeric(14,4),
    coste_total      numeric(14,4),
    completo         boolean,
    coste_hora       numeric(14,4),
    -- md5 del detalle por pieza: si no cambia, el detalle no se vuelve a guardar
    huella           text,
    PRIMARY KEY (idarticulo, fecha_calculo)
);
COMMENT ON TABLE core.hist_coste_articulo IS
  'Coste de cada articulo en cada recalculo. La serie empieza el dia que se desplego (no hay datos anteriores).';

-- El detalle por pieza, SOLO cuando cambia respecto al ultimo guardado del
-- articulo (~11.000 filas por foto del catalogo; guardarlo cada 3 dias sin
-- cambios seria engordar la tabla para nada). El detalle de un articulo en una
-- fecha es el ultimo guardado en o antes de esa fecha.
CREATE TABLE IF NOT EXISTS core.hist_coste_linea (
    idarticulo       varchar(20) NOT NULL,
    fecha_calculo    timestamp   NOT NULL,
    idpieza          varchar(20) NOT NULL,
    descripcion      text,
    cantidad         numeric(18,6),
    precio           numeric(14,4),
    coste_material   numeric(16,6) NOT NULL DEFAULT 0,
    tiempo_min       numeric(16,6) NOT NULL DEFAULT 0,
    coste_operacion  numeric(16,6) NOT NULL DEFAULT 0,
    PRIMARY KEY (idarticulo, fecha_calculo, idpieza)
);
COMMENT ON TABLE core.hist_coste_linea IS
  'Detalle por pieza de un articulo, guardado solo cuando cambia. El de una fecha es el ultimo con fecha_calculo <= esa fecha.';
