# Guía de soporte — Coste MP

Esta guía es para quien tenga que **atender incidencias** de la aplicación de costes sin ser
programador: Finanzas, Producción, Compras, sistemas o cualquier persona con acceso al
servidor. Explica qué es la aplicación, de dónde saca los datos, cómo saber si funciona y qué
hacer ante los problemas más habituales.

Si vas a mantener o modificar el código, lee también
[DOCUMENTACION_TECNICA.md](DOCUMENTACION_TECNICA.md). Las reglas exactas con las que se
calcula cada euro están en [reglas-del-coste.html](reglas-del-coste.html) (también en PDF).

---

## 1. Qué es

Una web interna de Gómez y Crespo que calcula **cuánto cuesta fabricar un artículo**:

- **Coste de material**: despieza el artículo a todos los niveles (lo que lleva dentro, y lo
  que lleva dentro cada pieza…) y valora cada pieza comprada a su último precio de compra.
- **Coste de operación**: los minutos de mano de obra de cada pieza fabricada por el coste de
  la hora de operario (17 €/h por defecto, editable).
- **Avisos**: qué piezas no se han podido valorar y por qué (sin precio, sin escandallo…).

Además mantiene calculado el **catálogo** (unos 550 artículos) y lo publica para que el libro
de análisis de Finanzas (Excel / Power Query) y Power BI lo lean sin tocar el ERP.

La aplicación **solo lee el ERP**: nunca modifica nada en él. Si un coste sale mal, casi
siempre es porque falta o está mal un dato en el ERP (un precio, una fase, un escandallo), no
porque la aplicación falle.

## 2. Pantallas

| Dirección | Para qué sirve |
|---|---|
| `/` (inicio) | Buscador: escribe un código o una descripción, abre el artículo y ve su despiece con el coste de cada pieza. Botones para descargar el Excel y para **Recalcular** (volver a leer el ERP en ese momento) |
| `/lote` | Coste por lote: subes un Excel con códigos de artículo y te devuelve otro Excel con el coste de cada uno (hojas *Costes* y *Faltantes*). Máximo 2.000 artículos |
| `/costes` | Estado del catálogo: cuándo fue el último cálculo, cuántos artículos están completos, lista de piezas pendientes. Botones *Recalcular pendientes* y *Recalcular todo* |
| `/costes/datos.xlsx` | El libro de datos que lee Power Query |
| `/catalogo` | Tabla de los artículos del catálogo con su coste y sus problemas |
| `/configuracion` | Coste de la hora de operario, con historial de cambios. Puede pedir una clave |
| `/evolucion` | Cómo ha cambiado el coste de cada artículo y qué pieza lo explica |
| `/piezas` | En qué artículos se usa una pieza y simulador: «¿qué pasa si sube su precio?» |

> A 28/10/2026, `/configuracion`, `/evolucion`, `/piezas` y el lote en segundo plano están
> terminados en el equipo de desarrollo pero **pendientes de subir al repositorio y de
> desplegar**. Si no aparecen en el servidor, es por eso.

## 3. Dónde está y de dónde saca los datos

```
                ERP GOMEZYCRESPO (SQL Server, 10.0.0.6)  — solo lectura
                              │
          ┌───────────────────┴────────────────────┐
          │                                        │
   costemp (la web)                    costemp-vigilante
   http://10.0.0.12:5010               recalcula el catálogo entero
   consultas al momento                cada 3 días a las 05:00
          │                                        │
          └──────────────► gyc_analytics ◄─────────┘
                    (Postgres, 10.0.0.12:5434, esquema core)
                              │
              ┌───────────────┴────────────────┐
              ▼                                ▼
   M:\06 Recursos técnicos\IT\costes\      /costes/datos.xlsx
        datos_costes.xlsx                  (lo mismo, por la web)
              │
              ▼
   Libro de análisis de Finanzas (Power Query) · Power BI
```

| Qué | Dónde |
|---|---|
| Dirección web | `http://10.0.0.12:5010` |
| Servidor | `10.0.0.12` (el mismo del ETL, Prefect y `gyc-api`) |
| Contenedores | `costemp` (la web) y `costemp-vigilante` (el cálculo programado) |
| ERP | SQL Server `GOMEZYCRESPO` en `10.0.0.6`, usuario de **solo lectura** |
| Base donde se publica | Postgres `gyc_analytics` en `10.0.0.12:5434`, tablas `core.*coste*` |
| Libro para Excel | `M:\06 Recursos técnicos\IT\costes\datos_costes.xlsx` (en el servidor, `/mnt/costes`) |
| Código | GitHub `GYCai-tech/CosteMP`, rama `main` |

Cosas importantes para entender los avisos de la gente:

- **El buscador y el lote consultan el ERP al momento** (recuerdan cada artículo 10 minutos).
- **`/catalogo`, `/costes`, el libro de Excel y Power BI leen la foto del último cálculo**, que
  se hace cada 3 días a las 05:00. Entre medias no cambian aunque cambie el ERP. Es a
  propósito: antes se recalculaba sin parar y llegó a usar el 17 % de la CPU del ERP.
- Por eso es normal que **el buscador y el catálogo den cifras distintas** para el mismo
  artículo si alguien ha tocado el ERP después del último cálculo. Se igualan pulsando
  *Recalcular* en el artículo o en `/costes`.

## 4. Cómo comprobar si funciona

1. Abre `http://10.0.0.12:5010`, busca un artículo cualquiera del catálogo y ábrelo. Debe
   salir su despiece con costes.
2. Abre `/costes`: el «último cálculo» no debe tener más de **3–4 días**.
3. Comprueba que el fichero `M:\06 Recursos técnicos\IT\costes\datos_costes.xlsx` tiene fecha
   de modificación reciente (la del último cálculo).

## 5. Problemas habituales

### 5.1 «Un coste está mal» (lo más frecuente)

Antes de pensar en un fallo, abre el artículo en el buscador y mira los avisos y las etiquetas
de cada fila. Casi siempre explican el número:

| Lo que se ve | Qué significa | Quién lo arregla (en el ERP) |
|---|---|---|
| **Sin precio de compra** | La pieza se compra (tiene tipo de aprovisionamiento) pero no hay ningún albarán, tarifa ni subcontrata con precio | Compras |
| **Sin tipo de aprovisionamiento** | La aplicación no sabe si la pieza se compra o se fabrica: no tiene precio ni fase. Hay que decidir qué es | Por decidir (Compras / Producción) |
| **Sin escandallo activo** | La pieza se fabrica aquí (tuvo fase) pero no tiene una fase activa con sus materiales | Producción |
| **Sin tiempo de operación** (aviso azul) | Pieza fabricada sin bonos ni mano de obra imputada. El material sí está costeado; solo falta la mano de obra | Producción |
| Precio marcado como **lista** | No hay compras: se ha usado la tarifa del proveedor. Ojo con las unidades (puede venir por metro y comprarse por rollo) | Compras |
| Tiempo marcado como **medio** / **teórico** | Medio = media real de los bonos (preferente). Teórico = la mano de obra imputada en el ERP | — |
| Fila **de conjunto** a 0 € | Correcto: el material lo paga su «hermana» que sale de la misma chapa | — |
| Artículo marcado **anulado** | Está dado de baja en el ERP | — |

Tras corregir el dato en el ERP, pulsa **Recalcular** en la pantalla del artículo: lee el ERP
de nuevo y, si el artículo es del catálogo, actualiza también Postgres y el libro de Excel.

Arreglar un «sin escandallo» puede **aumentar** el número de avisos: al abrirse esa rama
aparecen piezas de debajo que antes no se veían. No es un fallo.

Si después de todo esto el número sigue sin tener sentido, recoge el código del artículo, una
captura y lo que esperabas, y escala (apartado 9).

### 5.2 Fallos de la aplicación

Para los comandos hay que entrar en el servidor (apartado 6).

| Lo que se ve | Causa probable | Qué hacer |
|---|---|---|
| La página no abre | Contenedor `costemp` parado o servidor apagado | Ver el estado (6.1) y arrancar (6.3) |
| La página abre pero al buscar o abrir un artículo da error o se queda cargando | No llega al ERP (red, SQL Server caído o credenciales caducadas) | Registros de `costemp` (6.2). Si sale `Login failed`, han cambiado la contraseña del usuario de solo lectura: actualizar `SQLSERVER_PASS` en el `.env` y reiniciar |
| `/costes` o `/catalogo` dan error, pero el buscador funciona | No llega a Postgres (`gyc_analytics`) | Registros de `costemp`. Comprobar que el contenedor de Postgres del servidor está en marcha |
| El «último cálculo» de `/costes` tiene más de 4 días | El vigilante está parado o su pasada falla | Registros de `costemp-vigilante` (6.2). Si está parado, arrancarlo. Se puede lanzar a mano desde `/costes` → *Recalcular todo* (unos 4 minutos) |
| *Recalcular* dice «Ya hay un recalculo en curso» | Otra persona ha lanzado un recálculo desde `/costes` y aún no ha terminado | Esperar unos minutos y volver a mirar `/costes` |
| El Excel de `M:` no se actualiza, pero `/costes` sí | La carpeta compartida no está montada en el servidor, o el fichero estaba **abierto en Excel** durante el cálculo (Excel lo bloquea) | Cerrar el fichero en todos los equipos y pulsar *Recalcular pendientes*. Si sigue igual, comprobar el montaje (6.5). Mientras tanto, Power Query puede leer `http://10.0.0.12:5010/costes/datos.xlsx` |
| En el libro de análisis el `VLOOKUP` da `#N/A` | El artículo no está en el catálogo, tiene error de cálculo, o se ha cambiado la columna o el tipo del código | Ver si el artículo está en `/catalogo`. No cambiar la vista `v_coste_hoja10` ni el formato de la columna de código del libro (debe ser número) |
| El lote tarda mucho | Es normal con muchos artículos que no son del catálogo: cada uno se consulta en el ERP | Esperar; la página muestra el avance. El resultado se puede descargar durante 1 hora |
| El lote ha costeado artículos que no eran | El Excel subido tenía la columna de códigos con otro nombre o había varias columnas | La aplicación busca una columna llamada «artículo», «código», «id» o «referencia»; si no la hay, usa la primera. Dejar los códigos en una columna con ese nombre, en la primera hoja visible |
| `/configuracion` pide una clave que nadie sabe | Está definida en `CONFIG_CLAVE` del `.env` del servidor | Mirarla en el `.env` (6.4) |
| `/configuracion` dice que falta `CONFIG_CLAVE` y no deja guardar | No está definida (o está vacía) en el `.env` del servidor | Definirla en el `.env` y reiniciar el contenedor de la web (6.4) |
| El coste de todo el catálogo ha cambiado de golpe | Alguien ha cambiado el coste/hora de operario | `/configuracion` muestra el historial de cambios con fecha y equipo |
| Aparece un artículo nuevo que no está en el catálogo | El catálogo es una lista fija en Postgres | Ver «Añadir o quitar artículos del catálogo» en la documentación técnica |

## 6. Comandos del servidor

Se entra por SSH al servidor `10.0.0.12` y se va a la carpeta de la aplicación (en el servidor
es `/home/goyco/apps/costes-articulo`, la misma de la que lee el `.env` la integración
continua). Para confirmarlo:

```bash
docker inspect costemp --format '{{ index .Config.Labels "com.docker.compose.project.working_dir" }}'
```

 Todos los comandos se ejecutan **desde esa carpeta** y ninguno borra datos.

### 6.1 Ver si está funcionando

```bash
docker compose ps
```

Deben salir `costemp` y `costemp-vigilante` con estado `Up`.

### 6.2 Ver los registros

```bash
docker compose logs --tail 100 costemp              # la web
docker compose logs --tail 100 costemp-vigilante    # el cálculo programado
docker compose logs -f costemp-vigilante            # en directo (salir con Ctrl+C)
```

En el vigilante es normal ver «Pasada completa cada 3 dias…» al arrancar y una línea por
pasada. Las líneas con `ERROR` o `Traceback` son las importantes.

### 6.3 Reiniciar

```bash
docker compose restart                 # las dos piezas
docker compose up -d                   # arranca lo que esté parado
```

Reiniciar no pierde nada: los datos están en Postgres. La pasada programada **no** se repite
por reiniciar: la fecha de la última se guarda en Postgres.

### 6.4 El fichero `.env`

Contiene las contraseñas del ERP y de Postgres y la clave de `/configuracion`. **No está en
GitHub** y no se debe enviar por correo. Si se pierde, hay que pedir las credenciales a
sistemas y rehacerlo a partir de `.env.example`. Tras cambiarlo:

```bash
docker compose up -d          # recrea los contenedores con los valores nuevos
```

### 6.5 Comprobar la carpeta compartida

```bash
mount | grep costes                         # debe salir //gcdc001/goyco/... en /mnt/costes
ls -l /mnt/costes/datos_costes.xlsx         # fecha de la última escritura
sudo mount -a                               # volver a montarla si no aparece
```

Si no se monta, el cálculo sigue funcionando (avisa en los registros y publica en Postgres);
solo deja de actualizarse el fichero de `M:`.

### 6.6 Desplegar una versión nueva

No hay despliegue automático: se hace a mano desde la carpeta de la aplicación.

```bash
git pull
docker compose up -d --build
docker compose ps
```

## 7. Datos y copias de seguridad

- Todo lo que guarda la aplicación está en Postgres `gyc_analytics` (tablas `core.*coste*`):
  el catálogo, los costes publicados, el histórico, la configuración y el registro de
  cálculos. Entra en las copias de esa base; no hay nada más que guardar.
- El ERP no se toca nunca.
- El `.env` del servidor conviene tenerlo guardado en un sitio seguro (gestor de contraseñas
  de sistemas).

## 8. Comprobaciones útiles en Postgres

Para quien tenga acceso a `gyc_analytics` (DBeaver, pgAdmin…):

```sql
-- Los últimos cálculos: cuándo, cuántos artículos, errores y si fue programado o manual
SELECT * FROM core.log_coste_recalculo ORDER BY id DESC LIMIT 5;

-- Fecha de la última pasada programada
SELECT ultima_completa FROM core.cfg_coste_vigilancia;

-- Coste publicado de un artículo
SELECT * FROM core.dim_coste_escandallo WHERE idarticulo = '12101021';

-- Qué le falta a un artículo
SELECT * FROM core.fact_coste_hueco WHERE idarticulo = '12101021';
```

## 9. Qué datos recoger antes de escalar

- Código del artículo, qué coste sale y cuál esperabas.
- Si lo has visto en el buscador, en el catálogo o en el Excel (salen de sitios distintos,
  apartado 3), y la hora.
- Captura de pantalla con los avisos.
- Si es un fallo de la aplicación: salida de `docker compose ps` y de
  `docker compose logs --tail 200 costemp costemp-vigilante`.

## 10. Glosario

| Término | Significado |
|---|---|
| **Escandallo / despiece** | Lista de lo que lleva un artículo. Se explota a todos los niveles |
| **Fase** | Operación de fabricación del ERP: produce un artículo (*salidas*) y consume materiales (*entradas*) |
| **Conjunto** | Agrupación de artículos en el ERP; se usa como escandallo solo si no hay fase activa |
| **Hoja** | Pieza sin escandallo: se valora a precio de compra |
| **Materia prima / mercadería** | Tipo de aprovisionamiento de una pieza comprada, según la ficha del ERP |
| **Albarán** | Compra real. Su precio es la primera fuente del coste |
| **Tarifa (lista)** | Precio que pide el proveedor. Se usa solo si no hay albarán |
| **Subcontrata** | Trabajo externo (lacado, zincado…) facturado contra la orden de trabajo |
| **Bonos** | Partes de producción con minutos y piezas. Su media es el tiempo preferente |
| **Coste/hora de operario** | Euros por hora con los que se valora la mano de obra (17 € por defecto) |
| **Catálogo** | Lista de unos 550 artículos que se recalculan y publican de forma automática |
| **Pasada** | Recálculo del catálogo entero (unos 4 minutos). Programada cada 3 días a las 05:00 |
| **Completo** | Artículo al que no le falta ningún coste de material (puede faltarle tiempo) |
| **Hueco / pendiente** | Pieza que impide cerrar el coste de un artículo |
| **Power Query** | Herramienta de Excel que carga `datos_costes.xlsx` en el libro de análisis |
