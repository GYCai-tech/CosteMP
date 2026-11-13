-- ---------------------------------------------------------------------------
-- Parametros del calculo del coste, editables sin tocar codigo ni desplegar.
--
-- Uno por fila, por clave (los declara web_config.PARAMETROS). Aqui solo se
-- guardan los que se cambian en la aplicacion: los que vienen del ERP no. El
-- coste de la hora de operario, que fue el primero en vivir aqui, se lee ahora
-- del ERP (Empleados_Categorias, ver desglose.coste_hora_erp) y esta tabla ya
-- no lo guarda.
--
-- Se cambia desde la pagina /configuracion de la aplicacion, que deja cada
-- cambio en core.log_coste_parametro. La web lo aplica al momento y el
-- catalogo publicado en la siguiente pasada, o al pulsar Recalcular en /costes.
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS core.cfg_coste_parametro (
    clave         text PRIMARY KEY,
    valor         numeric(14,4) NOT NULL CHECK (valor > 0),
    unidad        text,
    fecha_cambio  timestamp NOT NULL DEFAULT now()
);
COMMENT ON TABLE core.cfg_coste_parametro IS
  'Parametros del calculo de coste de Coste-MP. Se cambian en /configuracion; ver sql/06_parametros_coste.sql.';

-- Historial: la tabla de arriba solo guarda el valor vigente, y un coste que
-- cambia sin dejar rastro hace que nadie sepa explicar por que el catalogo
-- subio de un dia para otro. La aplicacion no tiene usuarios, asi que lo unico
-- que se puede apuntar de quien lo cambio es el equipo (IP) desde el que se hizo.
CREATE TABLE IF NOT EXISTS core.log_coste_parametro (
    id              serial PRIMARY KEY,
    clave           text          NOT NULL,
    valor_anterior  numeric(14,4),
    valor_nuevo     numeric(14,4) NOT NULL,
    desde           text,
    fecha           timestamp     NOT NULL DEFAULT now()
);
COMMENT ON TABLE core.log_coste_parametro IS
  'Cada cambio hecho en /configuracion: valor anterior, nuevo y equipo desde el que se hizo.';
