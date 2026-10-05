-- ---------------------------------------------------------------------------
-- Rentabilidad: el coste de escandallo contra el precio de tarifa.
--
-- El precio de venta vive en el ERP (dbo.Listas_Precios_Cli_Art), no en
-- gyc_analytics, asi que aqui NO se puede juntar con el coste. Lo trae el
-- recalculo (exportar_costes.enriquecer_tarifa) y lo guarda en las columnas
-- que anade este fichero. Igual que con el coste: el calculo vive en Python,
-- estas vistas son la foto.
--
-- Se publica en Postgres y no solo en la pagina web a proposito. Si el margen
-- viviera unicamente en una pantalla, Finanzas acabaria copiandolo a mano a un
-- Excel y habria dos verdades -- que es justo el problema que este proyecto
-- lleva evitando desde el principio.
-- ---------------------------------------------------------------------------

-- Aditivo: no se toca ninguna columna existente, las vistas actuales no se
-- enteran. Se puede aplicar con la tabla ya poblada.
ALTER TABLE core.dim_coste_escandallo
    ADD COLUMN IF NOT EXISTS precio_tarifa    numeric(14,4),
    ADD COLUMN IF NOT EXISTS tarifa_nombre    text,
    ADD COLUMN IF NOT EXISTS tarifa_desde_uds numeric(14,4);

COMMENT ON COLUMN core.dim_coste_escandallo.precio_tarifa IS
  'Precio de venta de la lista de cliente VIGENTE en el ERP el dia del recalculo. NULL si el articulo no esta en ninguna.';
COMMENT ON COLUMN core.dim_coste_escandallo.tarifa_nombre IS
  'Nombre de esa lista (p.ej. "Tarifa Catalogo 2026"), para que el dato diga siempre de que precio habla.';
COMMENT ON COLUMN core.dim_coste_escandallo.tarifa_desde_uds IS
  'Cantidad minima a partir de la cual aplica ese precio. Si es > 1, el margen calculado NO vale para venta unitaria.';


-- ---------------------------------------------------------------------------
-- La vista de rentabilidad.
--
-- REGLA QUE NO SE PUEDE SALTAR: un articulo con el coste incompleto NO recibe
-- tramo. Al coste incompleto solo le pueden faltar euros, nunca sobrarle, asi
-- que su margen sale SIEMPRE mejor de lo que es. Meterlo en un tramo seria
-- publicar un numero que sabemos falso, y falso hacia el lado que enganha.
--
-- El 11601001 (BEBEDERO AUTOMATICO COLGANTE) es el caso de manual: sale a
-- 30,4% con 9 piezas sin valorar (campana, manguera, tapa deposito...), asi
-- que su margen real es menor y nadie sabe cuanto. Va a 'Sin determinar'.
-- ---------------------------------------------------------------------------
CREATE OR REPLACE VIEW core.v_rentabilidad AS
WITH base AS (
    SELECT
        c.idarticulo,
        c.descripcion,
        COALESCE(c.familia, '(Sin definir)') AS familia,
        c.coste_material,
        c.coste_operacion,
        c.coste_total,
        c.precio_tarifa,
        c.tarifa_nombre,
        c.tarifa_desde_uds,
        c.completo,
        c.error,
        c.fecha_calculo,
        COALESCE(c.piezas_sin_escandallo,0)
          + COALESCE(c.piezas_sin_tipo,0)
          + COALESCE(c.piezas_sin_precio,0) AS piezas_con_problema,
        -- analizable = tenemos las dos puntas y el coste esta cerrado
        (c.error IS NULL AND c.completo
         AND c.precio_tarifa IS NOT NULL AND c.precio_tarifa > 0) AS analizable
    FROM core.dim_coste_escandallo c
)
SELECT
    b.idarticulo                          AS "IdArticulo",
    b.descripcion                         AS "Descripcion",
    b.familia                             AS "Familia",
    b.precio_tarifa                       AS "PV tarifa (EUR)",
    b.tarifa_nombre                       AS "Tarifa",
    b.tarifa_desde_uds                    AS "Desde unidades",
    b.coste_material                      AS "Coste MP (EUR)",
    b.coste_operacion                     AS "Coste operacion (EUR)",
    b.coste_total                         AS "Coste total (EUR)",

    -- Margen: solo si es analizable. Para el resto va NULL, no un 0 que
    -- alguien pueda sumar por error en una tabla dinamica.
    CASE WHEN b.analizable
         THEN ROUND(b.precio_tarifa - b.coste_total, 4) END
                                          AS "Margen bruto (EUR)",
    CASE WHEN b.analizable
         THEN ROUND(100.0 * (b.precio_tarifa - b.coste_total) / b.precio_tarifa, 1) END
                                          AS "Margen (%)",

    -- Peso de cada componente sobre el precio: responde a "este articulo va
    -- justo por el material o por el tiempo?", que es la pregunta accionable.
    CASE WHEN b.analizable
         THEN ROUND(100.0 * b.coste_material / b.precio_tarifa, 1) END
                                          AS "MP sobre PV (%)",
    CASE WHEN b.analizable
         THEN ROUND(100.0 * b.coste_operacion / b.precio_tarifa, 1) END
                                          AS "Operacion sobre PV (%)",

    CASE
        WHEN b.error IS NOT NULL              THEN 'Error'
        WHEN b.precio_tarifa IS NULL
          OR b.precio_tarifa = 0              THEN 'Sin tarifa'
        WHEN NOT b.completo                   THEN 'Sin determinar'
        WHEN b.precio_tarifa < b.coste_total  THEN 'Por debajo de coste'
        WHEN 100.0 * (b.precio_tarifa - b.coste_total) / b.precio_tarifa < 20
                                              THEN '0-20%'
        WHEN 100.0 * (b.precio_tarifa - b.coste_total) / b.precio_tarifa < 40
                                              THEN '20-40%'
        WHEN 100.0 * (b.precio_tarifa - b.coste_total) / b.precio_tarifa < 60
                                              THEN '40-60%'
        WHEN 100.0 * (b.precio_tarifa - b.coste_total) / b.precio_tarifa < 80
                                              THEN '60-80%'
        ELSE                                       '>80%'
    END                                   AS "Tramo",

    -- Por que no se puede analizar, en cristiano. Vacio si si se puede.
    CASE
        WHEN b.error IS NOT NULL   THEN b.error
        WHEN b.precio_tarifa IS NULL OR b.precio_tarifa = 0
                                   THEN 'No esta en la tarifa vigente'
        WHEN NOT b.completo        THEN 'Coste incompleto: ' || b.piezas_con_problema
                                        || ' pieza(s) sin valorar. El margen real es MENOR.'
    END                                   AS "Motivo",

    b.piezas_con_problema                 AS "Piezas con problema",
    b.analizable                          AS "Analizable",
    b.fecha_calculo                       AS "Calculado"
FROM base b;

COMMENT ON VIEW core.v_rentabilidad IS
  'Margen bruto por articulo: coste de escandallo contra el precio de la tarifa vigente del ERP. Los articulos con coste incompleto salen con margen NULL y tramo "Sin determinar" a proposito: su margen real siempre es menor que el calculable.';


-- ---------------------------------------------------------------------------
-- Reparto por tramo. Alimenta la banda superior de la pagina /rentabilidad.
-- El orden es el del eje, no alfabetico: el tramo es una escala, no una
-- categoria suelta.
-- ---------------------------------------------------------------------------
CREATE OR REPLACE VIEW core.v_rentabilidad_tramos AS
SELECT
    t."Tramo"                                       AS "Tramo",
    CASE t."Tramo" WHEN 'Por debajo de coste' THEN 1
                   WHEN '0-20%'  THEN 2 WHEN '20-40%' THEN 3
                   WHEN '40-60%' THEN 4 WHEN '60-80%' THEN 5
                   WHEN '>80%'   THEN 6
                   WHEN 'Sin determinar' THEN 7
                   WHEN 'Sin tarifa'     THEN 8
                   ELSE 9 END                       AS orden,
    COUNT(*)                                        AS "Articulos",
    ROUND(AVG(t."Margen (%)"), 1)                   AS "Margen medio (%)",
    ROUND(SUM(t."Margen bruto (EUR)"), 2)           AS "Margen total (EUR)"
FROM core.v_rentabilidad t
GROUP BY t."Tramo"
ORDER BY orden;

COMMENT ON VIEW core.v_rentabilidad_tramos IS
  'Cuantos articulos hay en cada tramo de margen. Los tres ultimos (Sin determinar / Sin tarifa / Error) NO son tramos de rentabilidad: son articulos que no se pueden analizar.';
