# Costes Artículo — Despiece de materiales del ERP

Genera el **despiece (lista de materiales) multinivel** de un artículo del ERP
GÓMEZ Y CRESPO (SQL Server `GOMEZYCRESPO`) y lo exporta a Excel.

Explota todos los niveles con un CTE recursivo, multiplicando las cantidades en
cascada, marca qué líneas son materia prima (hojas del árbol) y calcula el
**coste de materia prima por línea** y el **coste total** del artículo.

### Coste
- El módulo de costes del ERP está vacío (`CosteEstandar` NULL). El coste se
  calcula con el **último precio de compra**: la compra más reciente
  (`Pedidos_Prov_Lineas` con `FechaAlbaran`) de cada artículo → `Precio_EURO`
  y `Descuento`.
- `Coste_línea = Cant × (Precio − Precio × Descuento/100)`.
- Solo se valoran los artículos **comprables** (`Articulos.IdTipoAprovisionamiento`
  no nulo). Los semielaborados no llevan precio (su coste ya está en sus materias
  primas), evitando doble conteo.
- Las líneas en `gr` se pasan a kg (÷1000) antes de valorar.

## Requisitos

- Python con el launcher `py`
- `ODBC Driver 18 for SQL Server` (ya instalado en la máquina)
- Dependencias: `py -m pip install -r requirements.txt`

## Configuración

Copia `.env.example` a `.env` y rellena las credenciales de **solo lectura** del
ERP. La cuenta usada no tiene permisos de escritura: el script solo lee.

## Docker (recomendado para desplegar)

La imagen instala el `ODBC Driver 18 for SQL Server`, así que no hace falta
tenerlo en la máquina. Las credenciales se pasan por `.env` en tiempo de
ejecución (no se copian a la imagen).

```bash
cp .env.example .env      # y rellena tus credenciales
docker compose up -d --build
```
Abre http://localhost:5000

Sin compose:
```bash
docker build -t costemp .
docker run -d -p 5000:5000 --env-file .env --name costemp costemp
```

> El contenedor debe estar en una red que **alcance el SQL Server** del ERP
> (`SQLSERVER_HOST`).

## Uso — interfaz web (recomendado)

```bash
py app.py
```
Abre http://127.0.0.1:5000 : busca un artículo por código o descripción,
haz clic para ver su despiece multinivel (la materia prima va resaltada) y
descarga el Excel con un botón.

## Uso — línea de comandos

```bash
py desglose.py 12101021
py desglose.py 12101021 -o salidas/jaula_perdices.xlsx
```

Genera un Excel con dos hojas:
- **Desglose**: una fila por componente (Nivel, Artículo, IdArticulo,
  Componente, Cant, Unidad, EsMateriaPrima).
- **Resumen**: artículo, nº de líneas, niveles y líneas de materia prima.

## Estructura

| Archivo | Qué hace |
|---|---|
| `app.py` | Interfaz web Flask (buscador + despiece + descarga Excel) |
| `templates/index.html` | Frontend de la interfaz |
| `db.py` | Conexión SQLAlchemy al ERP (solo lectura) |
| `desglose.py` | Búsqueda de artículos + CTE recursivo + exportación a Excel |
| `.env` | Credenciales (no se versiona) |
| `salidas/` | Excels generados (no se versiona) |

## Notas del modelo (ERP)

- `Fases` = operación de fabricación; `Fases_Salidas` = lo que produce;
  `Fases_Entradas` = lo que consume (con `Cantidad`).
- `Articulos.IdArticulo` es **VARCHAR** (códigos de 8 dígitos con ceros a la
  izquierda) — tratar siempre como texto.
- Cantidades convertidas con `TRY_CONVERT(FLOAT, ...)`. Si aparecen cantidades
  con coma decimal darían NULL; revisar si ocurre.
