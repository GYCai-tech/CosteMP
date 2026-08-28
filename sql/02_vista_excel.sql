-- ---------------------------------------------------------------------------
-- Vista hecha a la medida de la Hoja10 del libro
-- "ANALISIS COSTES CATALOGO AGOSTO 26".
--
-- Replica EXACTAMENTE sus 4 columnas, en el mismo orden y con los mismos
-- nombres, para que Power Query la vuelque encima sin transformar nada y la
-- formula   costes santiago!G = -VLOOKUP(A2; Hoja10!$A:$C; 3; 0)
-- siga funcionando sin tocarla.
--
-- OJO CON EL TIPO: en la hoja el codigo es NUMERO, no texto. Si se devuelve
-- como varchar, Excel no casa los VLOOKUP y salen todos #N/A. Por eso el CAST
-- a integer. Es seguro: los 548 codigos del catalogo son puramente numericos
-- y ninguno lleva ceros a la izquierda (comprobado 27/08/2026).
--
-- El coste que se publica aqui es SOLO MATERIAL, como hace hoy la hoja. La
-- mano de obra (core.dim_coste_escandallo.coste_operacion) queda fuera a
-- proposito. Si algun dia entra en el margen, cambiar aqui la columna y el
-- VLOOKUP pasa de 3 a la nueva posicion.
-- ---------------------------------------------------------------------------
CREATE OR REPLACE VIEW core.v_coste_hoja10 AS
SELECT
    CAST(c.idarticulo AS integer)              AS "IdArticulo",
    c.descripcion                              AS "Descripcion",
    c.coste_material                           AS "Coste material (EUR)",
    CASE WHEN c.completo THEN 'Sí' ELSE 'No' END AS "Completo"
FROM core.dim_coste_escandallo c
WHERE c.error IS NULL                -- los codigos inexistentes no se vuelcan
  AND c.idarticulo ~ '^[0-9]+$'      -- red de seguridad para el CAST
ORDER BY 1;

COMMENT ON VIEW core.v_coste_hoja10 IS
  'Espejo de la Hoja10 del Excel de analisis de catalogo. Solo coste material. No cambiar el orden ni el nombre de las columnas sin avisar: hay VLOOKUP apuntando por posicion.';
