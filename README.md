# gmaps-leads

Genera una planilla de negocios de una ciudad (nombre, dirección, teléfono, link de WhatsApp, web, rating) a partir de Google Maps. Sirve para armar listas de prospectos por rubro y contactarlos **a mano, de forma personalizada**.

**100% gratis: sin API key y sin tarjeta.** Abre un navegador (Chromium, vía Playwright) y lee Google Maps como lo haría una persona.

## Qué obtenés

Un Excel (o CSV) con una fila por negocio:

| Columna | Qué es |
|---|---|
| Nombre, Dirección, Teléfono, Web | Datos públicos de la ficha en Google Maps |
| WhatsApp (verificar) | Link `wa.me` armado con el teléfono. **Solo funciona si el número es celular y tiene WhatsApp**; los fijos no |
| Rating | Estrellas de la ficha |
| Google Maps | Link a la ficha |
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

Velocidad: unos 2 a 3 segundos por negocio. 40 negocios tardan aproximadamente 2 minutos.

Tip: para cubrir más negocios de un rubro, hacé búsquedas más específicas (`"kinesiólogo deportivo"`, `"kinesiología respiratoria"`). Los repetidos se descartan solos.

### 3. Usarlo desde otra computadora

Repetí el paso 1 en esa máquina. No hay claves ni archivos de configuración que copiar.

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
