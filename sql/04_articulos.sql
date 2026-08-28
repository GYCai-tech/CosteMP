-- ---------------------------------------------------------------------------
-- Hoja "Articulos" del libro de datos: un articulo por fila, con el coste
-- partido en materia prima y operacion, y el detalle de que le falta.
--
-- Es la vista que sustituye a la pagina /catalogo de la aplicacion. Va al
-- Excel porque es donde se trabaja: el autofiltro de Excel ya da el filtro por
-- columna, la ordenacion y la busqueda sin programar nada.
--
-- NO reemplaza a v_coste_hoja10: aquella tiene las 4 columnas exactas que
-- espera el VLOOKUP del libro de analisis y no se puede tocar.
-- ---------------------------------------------------------------------------
CREATE OR REPLACE VIEW core.v_coste_articulo AS
SELECT
    c.idarticulo                                  AS "IdArticulo",
    c.descripcion                                 AS "Descripcion",
    COALESCE(c.familia,'(Sin definir)')           AS "Familia",
    c.coste_material                              AS "Materia prima (EUR)",
    c.coste_operacion                             AS "Operacion (EUR)",
    c.coste_total                                 AS "Coste total (EUR)",
    -- un articulo esta MAL si le falta coste de MATERIAL. "Sin tiempo" no
    -- cuenta: el material esta costeado y solo falta afinar la mano de obra.
    -- Mismo criterio que el informe y que la interfaz.
    CASE WHEN c.error IS NOT NULL THEN 'Error'
         WHEN COALESCE(c.piezas_sin_escandallo,0)
            + COALESCE(c.piezas_sin_tipo,0)
            + COALESCE(c.piezas_sin_precio,0) > 0 THEN 'Faltan datos'
         ELSE 'Completo' END                      AS "Estado",
    COALESCE(c.piezas_sin_escandallo,0)
      + COALESCE(c.piezas_sin_tipo,0)
      + COALESCE(c.piezas_sin_precio,0)           AS "Piezas con problema",
    COALESCE(c.piezas_sin_escandallo,0)           AS "Sin escandallo",
    COALESCE(c.piezas_sin_tipo,0)                 AS "Sin tipo",
    COALESCE(c.piezas_sin_precio,0)               AS "Sin precio",
    COALESCE(c.piezas_sin_tiempo,0)               AS "Sin tiempo",
    c.error                                       AS "Error",
    c.fecha_calculo                               AS "Calculado"
FROM core.dim_coste_escandallo c
ORDER BY c.idarticulo;

COMMENT ON VIEW core.v_coste_articulo IS
  'Un articulo del catalogo por fila con el coste partido en MP y operacion. Alimenta la hoja "Articulos" del libro de datos.';
