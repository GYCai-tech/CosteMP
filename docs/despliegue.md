# Despliegue en el servidor

El destino es **10.0.0.12**, el mismo host donde ya corren el ETL, Prefect y
`gyc-api`. Ahi el stack se levanta solo y deja de depender de que un PC este
encendido.

## Lo que hay que preparar antes

### 1. La carpeta compartida

El resultado se publica en la carpeta comun de la empresa, para que Excel lo
lea sin instalar drivers ni tener credenciales de base de datos:

```
\\gcdc001\goyco\06 Recursos técnicos\IT\costes
```

Vista desde Windows es `M:\06 Recursos técnicos\IT\costes` (la unidad `M:` ya
esta mapeada en los equipos). La carpeta existe y admite escritura.

### 2. Montarla en el servidor

Necesita el paquete `cifs-utils`. Credenciales de un usuario de dominio con
permiso de escritura, en un fichero que **solo pueda leer root**:

```bash
sudo tee /etc/samba/gyc.cred >/dev/null <<'EOF'
username=<usuario>
password=<contraseña>
domain=<dominio>
EOF
sudo chmod 600 /etc/samba/gyc.cred
```

El montaje, en `/etc/fstab` para que sobreviva a un reinicio. Ojo con los
espacios y el acento del nombre: en fstab los espacios van como `\040`.

```
//gcdc001/goyco/06\040Recursos\040técnicos/IT/costes  /mnt/costes  cifs  credentials=/etc/samba/gyc.cred,uid=1000,gid=1000,iocharset=utf8,vers=3.0,nofail  0  0
```

`nofail` es importante: si el servidor de ficheros no responde al arrancar, el
servidor sigue arrancando en vez de quedarse colgado.

```bash
sudo mkdir -p /mnt/costes && sudo mount -a
touch /mnt/costes/prueba && rm /mnt/costes/prueba    # comprobar escritura
```

### 3. El `.env`

No va en git. Copiar `.env.example` y rellenar. Ademas de las credenciales del
ERP y de Postgres, en el servidor hace falta:

```
COSTES_DATOS_DIR=/mnt/costes
```

Es la carpeta compartida **vista desde el host**. Dentro de los contenedores se
ve siempre como `/datos`, lo pone el compose.

## Levantar

```bash
git pull
docker compose up -d --build
```

Quedan dos contenedores:

| | |
|---|---|
| `costemp` | la aplicacion web, puerto 5010 |
| `costemp-vigilante` | recalcula: cada 10 min si hubo cambios, y entero a las 9:00 |

## Comprobar que funciona

```bash
docker logs -f costemp-vigilante          # deberia decir "sin cambios" o recalcular
ls -l /mnt/costes/datos_costes.xlsx       # se regenera en cada recalculo
curl -s -o /dev/null -w '%{http_code}\n' http://localhost:5010/catalogo
```

Y en Postgres, la traza de ejecuciones:

```sql
SELECT * FROM core.log_coste_recalculo ORDER BY id DESC LIMIT 5;
```

## Despues del despliegue

**Quitar las tareas programadas del PC de Santiago**, o habra dos procesos
recalculando lo mismo contra la misma base:

```
schtasks /delete /tn "GYC Costes vigilancia" /f
schtasks /delete /tn "GYC Costes catalogo" /f
```

Y cambiar el origen de las consultas de Power Query del libro de analisis para
que apunten a `M:\06 Recursos técnicos\IT\costes\datos_costes.xlsx` en vez de
al escritorio.

## Si el montaje CIFS se atasca

No bloquea nada: sin carpeta compartida el proceso avisa y sigue. Los datos se
publican igual en `gyc_analytics` y se sirven en
`http://10.0.0.12:5010/costes/datos.xlsx`, que Power Query lee con
*Datos -> Obtener datos -> Desde otras fuentes -> Desde la web*.
