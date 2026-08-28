-- ---------------------------------------------------------------------------
-- Trabajo pendiente: piezas que impiden cerrar el coste de un articulo.
--
-- Mismas 7 columnas, mismo orden y mismos nombres que la hoja
-- "Piezas sin coste (prioridad)" del informe catalogo_huecos_coste, mas la
-- fecha del calculo.
--
-- Entran TAMBIEN las "Sin escandallo activo", que en el informe iban en una
-- hoja aparte: la columna Motivo las separa y en una sola lista se ve de un
-- vistazo cuanto falta.
--
-- Se excluye "Sin tiempo de operacion": no impide valorar el material, son
-- ~312 piezas, y hoy el margen del catalogo no usa mano de obra.
--
-- NO se intenta adivinar de quien es cada hueco. Se probo y no sale: 165 de
-- las 204 piezas no tienen albaran, ni tarifa, ni rastro de fase, asi que
-- cualquier regla automatica responde "no se" en 4 de cada 5 filas. Se
-- publican los hechos y la decision es humana.
-- ---------------------------------------------------------------------------
ALTER TABLE core.fact_coste_hueco ADD COLUMN IF NOT EXISTS tipo_aprovisionamiento text;
ALTER TABLE core.fact_coste_hueco ADD COLUMN IF NOT EXISTS familia               text;
ALTER TABLE core.fact_coste_hueco ADD COLUMN IF NOT EXISTS tuvo_fase             boolean;
ALTER TABLE core.fact_coste_hueco ADD COLUMN IF NOT EXISTS veces_comprada        integer;
ALTER TABLE core.fact_coste_hueco ADD COLUMN IF NOT EXISTS tiene_tarifa          boolean;

DROP VIEW IF EXISTS core.v_coste_resumen;
DROP VIEW IF EXISTS core.v_coste_pendiente;

CREATE VIEW core.v_coste_pendiente AS
SELECT
    h.idpieza                                          AS "IdPieza",
    MIN(h.descripcion)                                 AS "Descripcion",
    COALESCE(MIN(h.tipo_aprovisionamiento),'SIN TIPO') AS "Tipo aprovisionamiento",
    COALESCE(MIN(h.familia),'(Sin definir)')           AS "Familia",
    MIN(h.motivo)                                      AS "Motivo",
    COUNT(DISTINCT h.idarticulo)                       AS "Articulos que desbloquea",
    STRING_AGG(DISTINCT h.idarticulo, ', ')            AS "Articulos afectados",
    MAX(h.fecha_calculo)                               AS "Calculado"
FROM core.fact_coste_hueco h
WHERE h.motivo <> 'Sin tiempo de operación'
GROUP BY h.idpieza
ORDER BY 6 DESC, 5, 1;

COMMENT ON VIEW core.v_coste_pendiente IS
  'Piezas que impiden cerrar el coste del catalogo. Copia la hoja "Piezas sin coste (prioridad)" del informe.';

CREATE VIEW core.v_coste_resumen AS
WITH a AS (SELECT * FROM core.dim_coste_escandallo)
SELECT 1 AS orden, 'Articulos del catalogo'  AS "Indicador", COUNT(*)::text AS "Valor" FROM a
UNION ALL SELECT 2, 'Completos',             COUNT(*) FILTER (WHERE completo)::text FROM a
UNION ALL SELECT 3, 'Pendientes',            COUNT(*) FILTER (WHERE NOT completo)::text FROM a
UNION ALL SELECT 4, 'Piezas por arreglar',   COUNT(*)::text FROM core.v_coste_pendiente
UNION ALL SELECT 5, '  ...sin escandallo',   COUNT(*)::text FROM core.v_coste_pendiente WHERE "Motivo"='Sin escandallo activo'
UNION ALL SELECT 6, '  ...sin tipo',         COUNT(*)::text FROM core.v_coste_pendiente WHERE "Motivo"='Sin tipo de aprovisionamiento'
UNION ALL SELECT 7, '  ...sin precio',       COUNT(*)::text FROM core.v_coste_pendiente WHERE "Motivo"='Sin precio de compra'
UNION ALL SELECT 8, 'Ultimo calculo',        TO_CHAR(MAX(fecha_calculo),'DD/MM/YYYY HH24:MI') FROM a
ORDER BY 1;
