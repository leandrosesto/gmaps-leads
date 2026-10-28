# gmaps-leads

Genera una planilla de negocios de una ciudad (nombre, dirección, teléfono, link de WhatsApp, web, rating) a partir de Google Maps. Sirve para armar listas de prospectos por rubro y contactarlos **a mano, de forma personalizada**.

**100% gratis: sin API key y sin tarjeta.** Abre un navegador (Chromium, vía Playwright) y lee Google Maps como lo haría una persona.

## Qué obtenés

Un Excel (o CSV) con una fila por negocio:

| Columna | Qué es |
|---|---|
| Nombre, Dirección, Teléfono | Datos públicos de la ficha en Google Maps |
| WhatsApp (verificar) | Clic en "Abrir WhatsApp" y abre el chat (en WhatsApp Web si estás en la PC). **Solo funciona si el número es celular y tiene WhatsApp**; los fijos no |
| Instagram | Clic en el `@usuario` y abre el perfil. Solo aparece si el negocio lo cargó como sitio web en Google Maps |
| Web | Clic en el dominio y abre el sitio |
| Rating | Estrellas de la ficha |
| Google Maps | Clic en "Ver en Maps" y abre la ficha |
| Contactado, Fecha contacto, Respuesta, Notas | Vacías: son para que lleves el seguimiento |

Se descartan los negocios cerrados definitivamente y los repetidos.

## Paso a paso

### 1. Instalar (una sola vez por computadora)

Necesitás Python 3.10 o superior.

```bash
git clone https://github.com/leandrosesto/gmaps-leads.git
cd gmaps-leads
python -m venv .venv
```

Activar el entorno:

```bash
# Windows (PowerShell)
.\.venv\Scripts\Activate.ps1
# Mac / Linux
source .venv/bin/activate
```

Instalar dependencias y el navegador (descarga unos 150 MB):

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

El repositorio es privado: para clonarlo tenés que estar logueado en GitHub (`gh auth login`).

### 2. Usarlo

Un rubro:

```bash
python gmaps_leads.py "kinesiólogo"
```

Varios rubros, en otra ciudad:

```bash
python gmaps_leads.py "estudio contable" "inmobiliaria" --city "Necochea"
```

El resultado queda en `output/leads_<fecha>_<hora>.xlsx`.

| Opción | Para qué |
|---|---|
| `--city "Mar del Plata"` | Ciudad (es el valor por defecto) |
| `--max 40` | Máximo de negocios por rubro (default 40, tope 120) |
| `--only-phone` | Descarta negocios sin teléfono |
| `--csv` | Exporta CSV en vez de Excel |
| `--out ruta/archivo.xlsx` | Elegir dónde guardar |
| `--show` | Muestra el navegador mientras trabaja (útil si algo falla) |
| `--instagram` | Además busca el Instagram en el sitio web de cada negocio (ver más abajo). Es más lento |

Velocidad: unos 2 a 3 segundos por negocio. 40 negocios tardan aproximadamente 2 minutos.

Tip: para cubrir más negocios de un rubro, hacé búsquedas más específicas (`"kinesiólogo deportivo"`, `"kinesiología respiratoria"`). Los repetidos se descartan solos.

### 3. Usarlo desde otra computadora

Repetí el paso 1 en esa máquina. No hay claves ni archivos de configuración que copiar.

## Guardar las planillas en Google Drive (carpeta LEADS)

Con esto, cada planilla que genere el programa se sube sola a una carpeta **LEADS** de tu Drive, y la abrís desde cualquier PC en <https://drive.google.com>. Es gratis y **no pide tarjeta**. Se configura una sola vez.

El programa pide el permiso mínimo (`drive.file`): solo ve los archivos y carpetas que él mismo crea, nunca el resto de tu Drive.

### Configuración (una sola vez)

Los nombres de los menús de Google cambian de vez en cuando; si algo no coincide, buscá el equivalente.

1. Entrá a <https://console.cloud.google.com/> con tu cuenta de Google y creá un proyecto (por ejemplo `leads`). **No hace falta activar facturación.**
2. **APIs y servicios > Biblioteca**: buscá **Google Drive API** y tocá **Habilitar**.
3. **APIs y servicios > Pantalla de consentimiento de OAuth** (o *Google Auth Platform*): tipo de usuario **Externo**, poné un nombre de app (por ejemplo `gmaps-leads`) y tu mail. En **Usuarios de prueba**, agregá tu mismo Gmail.
4. En la misma pantalla, pasá la app a **"En producción"** (*Publish app*). Si la dejás en "Prueba", Google te pide volver a iniciar sesión cada 7 días. Como el programa solo usa el permiso `drive.file`, no necesita verificación de Google.
5. **Credenciales > Crear credenciales > ID de cliente de OAuth**, tipo **Aplicación de escritorio**. Descargá el JSON y guardalo como `credentials.json` en la carpeta del proyecto (o en `~/.gmaps-leads/credentials.json`). Este archivo no se sube a GitHub.
6. Iniciá sesión una vez:

```bash
python drive_upload.py --login
```

Se abre el navegador, elegís tu cuenta y aceptás. Si Google avisa "app no verificada", tocá *Configuración avanzada > Ir a gmaps-leads*. Al terminar, se crea la carpeta **LEADS** en tu Drive.

### Uso

Nada cambia: corrés el scraper y, si Drive está configurado, la planilla se sube sola y el programa muestra el link.

```bash
python gmaps_leads.py "kinesiólogo" --instagram
```

| Opción | Para qué |
|---|---|
| `--no-drive` | No subir a Drive (queda solo en `output/`) |
| `--drive-folder "OTRA"` | Usar otra carpeta en vez de LEADS |

Los archivos se llaman `leads_<rubro>_<fecha>_<hora>.xlsx`. Si la subida falla (sin internet, por ejemplo), la planilla igual queda guardada en `output/`.

### Detalles a tener en cuenta

- **Carpeta LEADS existente:** como el programa solo ve lo que él creó, **no va a encontrar una carpeta LEADS que hayas creado a mano** y creará otra con el mismo nombre. Para evitar duplicados, borrá o renombrá la tuya antes de la primera subida, o mové los archivos viejos a la que cree el programa.
- **Editar desde cualquier PC:** en Drive, clic derecho sobre el archivo > *Abrir con > Hojas de cálculo de Google*. Los hipervínculos siguen funcionando. Si lo editás ahí, el cambio queda en Drive, no en la copia de `output/`.
- **Otra computadora:** repetí la instalación (paso 1), copiá tu `credentials.json` y corré `python drive_upload.py --login` una vez en esa PC.

## Encontrar el Instagram (`--instagram`)

Por defecto el Instagram sale solo si el negocio lo cargó como sitio web en Google Maps. Con `--instagram`, para los que tienen sitio web, el programa abre esa página y busca el link de Instagram (casi siempre está en el pie de página).

```bash
python gmaps_leads.py "kinesiólogo" --instagram
```

- No se loguea en Instagram ni usa tu cuenta. Solo lee el sitio web del negocio.
- **Revisá el perfil antes de escribir.** A veces el sitio enlaza al perfil de otra persona (el diseñador web, por ejemplo) en vez del negocio. En una prueba con 15 kinesiólogos, uno de los perfiles hallados parecía personal.
- Los negocios sin sitio web ni Instagram en Maps quedan sin Instagram: buscalos a mano por el nombre. Probé buscadores (DuckDuckGo, Bing) y no son confiables para esto: bloquean al programa o devuelven resultados que no corresponden.
- Para escribirles por Instagram no necesitás que tu cuenta sea pública. Igual conviene que tu perfil esté completo y se vea profesional, porque lo van a mirar antes de responder.

## Cómo usar la lista para contactar

- Mirá cada negocio (web, Instagram) antes de escribir. Un mensaje que menciona algo concreto del local funciona mucho mejor que uno genérico.
- Empezá con una pregunta corta, sin links ni adjuntos.
- Escribí a mano y pocos por día. El envío masivo desde WhatsApp es la forma más rápida de que bloqueen tu número.
- Anotá en las columnas de seguimiento a quién escribiste y qué respondió.

## Límites y avisos

- **Puede dejar de funcionar.** Lee la página de Google Maps, no una API. Si Google cambia su diseño, hay que actualizar los selectores en `gmaps_leads.py` (`scrape_place`).
- **Términos de uso.** Google no permite automatizar Maps. Usalo con volúmenes bajos y para tu uso propio. Con pocas decenas de negocios por vez no tuve problemas, pero si Google muestra un captcha, esperá un rato o bajá `--max`.
- Los datos son los de Google Maps y pueden estar desactualizados.
- Los teléfonos fijos no tienen WhatsApp: el link se arma para todos los números y hay que verificarlo.
- Los resultados dependen de la sesión de Google (ubicación, idioma): pueden variar entre corridas.

## Desarrollo

```bash
python -m unittest discover -s tests -v
```
