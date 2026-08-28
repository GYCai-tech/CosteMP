-- ---------------------------------------------------------------------------
-- Costes de escandallo publicados en gyc_analytics.
--
-- Los calcula el proceso exportar_costes.py con el MISMO codigo que la app web
-- (desglose.py + app.py), asi que aqui no hay ninguna logica de calculo: estas
-- tablas son la foto del ultimo recalculo. Una sola verdad, la de Python.
--
-- NO se toca core.dim_coste_articulo: esa tabla esta muerta (ultima carga
-- 2022-09-30, con claves corruptas del tipo ',,'). Se deja como estaba.
-- ---------------------------------------------------------------------------

-- Lista de articulos a costear. Editable sin tocar codigo: para incorporar un
-- articulo nuevo al catalogo basta con insertarlo aqui.
CREATE TABLE IF NOT EXISTS core.cfg_catalogo_coste (
    idarticulo   varchar(20) PRIMARY KEY,
    activo       boolean      NOT NULL DEFAULT true,
    fecha_alta   timestamp    NOT NULL DEFAULT now(),
    nota         text
);
COMMENT ON TABLE core.cfg_catalogo_coste IS
  'Articulos del catalogo cuyo coste se recalcula. Poner activo=false para excluir uno sin perder el registro.';

-- Una fila por articulo: el resultado del recalculo.
CREATE TABLE IF NOT EXISTS core.dim_coste_escandallo (
    idarticulo            varchar(20) PRIMARY KEY,
    descripcion           text,
    coste_material        numeric(14,4),
    coste_operacion       numeric(14,4),
    coste_total           numeric(14,4),
    piezas_sin_coste      integer NOT NULL DEFAULT 0,
    piezas_sin_escandallo integer NOT NULL DEFAULT 0,
    piezas_sin_tiempo     integer NOT NULL DEFAULT 0,
    completo              boolean NOT NULL DEFAULT false,
    error                 text,
    fecha_calculo         timestamp NOT NULL
);
COMMENT ON TABLE core.dim_coste_escandallo IS
  'Coste multinivel por articulo, calculado por Coste-MP (exportar_costes.py). completo=false significa que le falta algun dato de coste.';
COMMENT ON COLUMN core.dim_coste_escandallo.piezas_sin_tiempo IS
  'Informativo: NO tumba completo. El material si esta costeado.';

-- Detalle: que pieza concreta bloquea a que articulo, y por que.
CREATE TABLE IF NOT EXISTS core.fact_coste_hueco (
    idarticulo    varchar(20) NOT NULL,
    idpieza       varchar(20) NOT NULL,
    descripcion   text,
    motivo        text        NOT NULL,
    fecha_calculo timestamp   NOT NULL,
    PRIMARY KEY (idarticulo, idpieza, motivo)
);
CREATE INDEX IF NOT EXISTS ix_coste_hueco_pieza  ON core.fact_coste_hueco (idpieza);
CREATE INDEX IF NOT EXISTS ix_coste_hueco_motivo ON core.fact_coste_hueco (motivo);
COMMENT ON TABLE core.fact_coste_hueco IS
  'Un hueco = (articulo, pieza que lo bloquea, motivo). Motivos: Sin escandallo activo / Sin tipo de aprovisionamiento / Sin precio de compra / Sin tiempo de operacion.';

-- Traza de ejecuciones, para saber si el dato es de hoy o de la semana pasada.
CREATE TABLE IF NOT EXISTS core.log_coste_recalculo (
    id             serial PRIMARY KEY,
    inicio         timestamp NOT NULL,
    fin            timestamp,
    articulos      integer,
    errores        integer,
    duracion_seg   numeric(10,1),
    origen         text,               -- 'manual' | 'programado'
    resultado      text                -- 'ok' | 'error'
);

-- ---------------------------------------------------------------------------
-- VISTAS: lo que consumen Excel y Power BI. Nadie deberia leer las tablas.
-- Van en core (no en analytics) porque analytics lo gestiona dbt y podria
-- eliminar objetos creados a mano.
-- ---------------------------------------------------------------------------
CREATE OR REPLACE VIEW core.v_coste_catalogo AS
SELECT
    c.idarticulo                                   AS "Articulo",
    COALESCE(c.descripcion, a.descrip)             AS "Descripcion",
    a.nombrefamilia                                AS "Familia",
    c.coste_material                               AS "Coste material",
    c.coste_operacion                              AS "Coste operacion",
    c.coste_total                                  AS "Coste total",
    CASE WHEN c.completo THEN 'Completo'
         ELSE 'Faltan datos' END                   AS "Estado",
    c.piezas_sin_coste                             AS "Piezas sin coste",
    c.piezas_sin_escandallo                        AS "Piezas sin escandallo",
    c.piezas_sin_tiempo                            AS "Piezas sin tiempo",
    c.error                                        AS "Error",
    c.fecha_calculo                                AS "Calculado",
    -- EXTRACT devuelve double y round(double,int) no existe en Postgres
    ROUND((EXTRACT(EPOCH FROM (now() - c.fecha_calculo)) / 3600.0)::numeric, 1)
                                                   AS "Horas de antiguedad"
FROM core.dim_coste_escandallo c
LEFT JOIN core.dim_articulo a ON a.idarticulo = c.idarticulo;

CREATE OR REPLACE VIEW core.v_coste_huecos AS
SELECT
    h.idpieza                    AS "Pieza",
    h.descripcion                AS "Descripcion",
    h.motivo                     AS "Motivo",
    COUNT(*)                     AS "Articulos que desbloquea",
    STRING_AGG(h.idarticulo, ', ' ORDER BY h.idarticulo) AS "Articulos afectados",
    MAX(h.fecha_calculo)         AS "Calculado"
FROM core.fact_coste_hueco h
GROUP BY h.idpieza, h.descripcion, h.motivo;
