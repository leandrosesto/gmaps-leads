# gmaps-leads

Genera una planilla de negocios de una ciudad (nombre, dirección, teléfono, link de WhatsApp, web, rating) a partir de Google Maps. Sirve para armar listas de prospectos por rubro y contactarlos **a mano, de forma personalizada**.

Usa la API oficial de Google Places (New), no scrapea el HTML de Maps. Es más estable y no te bloquean la IP, pero necesita una API key (se crea una sola vez, paso 2).

## Qué obtenés

Un Excel (o CSV) con una fila por negocio y estas columnas:

| Columna | Qué es |
|---|---|
| Nombre, Dirección, Teléfono, Web | Datos públicos del negocio en Google Maps |
| WhatsApp (verificar) | Link `wa.me` armado con el teléfono. **Solo funciona si el número es celular y tiene WhatsApp**; los fijos no |
| Rating, Reseñas | Para priorizar |
| Google Maps | Link a la ficha |
| Contactado, Fecha contacto, Respuesta, Notas | Vacías: son para que lleves el seguimiento |

Se descartan los negocios cerrados definitivamente y los repetidos entre rubros.

## Paso a paso

### 1. Instalar

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

Instalar dependencias:

```bash
pip install -r requirements.txt
```

### 2. Conseguir la API key de Google (una sola vez)

1. Entrá a <https://console.cloud.google.com/> con tu cuenta de Google y creá un proyecto (por ejemplo `leads`).
2. Activá la facturación del proyecto (te piden una tarjeta; Google da un uso gratuito mensual por producto y esta herramienta consume muy poco, pero **revisá los precios vigentes** antes de usarla mucho).
3. Menú **APIs y servicios > Biblioteca**, buscá **Places API (New)** y tocá **Habilitar**.
4. Menú **APIs y servicios > Credenciales > Crear credenciales > Clave de API**.
5. Recomendado: en la clave, **Restricciones de API**, elegí solo *Places API (New)*.
6. En **Facturación > Presupuestos y alertas**, creá una alerta de presupuesto chica (por ejemplo USD 5) para que no haya sorpresas.

### 3. Configurar la clave

```bash
# Windows
copy .env.example .env
# Mac / Linux
cp .env.example .env
```

Abrí `.env` y pegá la clave después del `=`:

```
GOOGLE_MAPS_API_KEY=AIza...
```

El archivo `.env` está en `.gitignore`: nunca se sube a GitHub. No pegues la clave en el código ni en el chat.

### 4. Usarlo

Un rubro:

```bash
python gmaps_leads.py "kinesiólogo"
```

Varios rubros a la vez, en otra ciudad:

```bash
python gmaps_leads.py "estudio contable" "inmobiliaria" --city "Necochea"
```

El resultado queda en `output/leads_<fecha>_<hora>.xlsx`.

Opciones:

| Opción | Para qué |
|---|---|
| `--city "Mar del Plata"` | Ciudad (es el valor por defecto) |
| `--max 40` | Máximo de resultados por rubro. Google entrega hasta 60 por búsqueda |
| `--only-phone` | Descarta negocios sin teléfono |
| `--csv` | Exporta CSV en vez de Excel |
| `--out ruta/archivo.xlsx` | Elegir dónde guardar |

Tip: para cubrir más de 60 negocios de un rubro, hacé búsquedas más específicas (`"kinesiólogo deportivo"`, `"kinesiología respiratoria"`). El resultado se deduplica solo.

### 5. Usarlo desde otra computadora

Repetí los pasos 1 y 3 en esa máquina (clonar, crear el entorno, instalar, copiar tu `.env`). El repositorio es privado: para clonarlo tenés que estar logueado en GitHub.

## Cómo usar la lista para contactar

- Mirá cada negocio (web, Instagram, reseñas) antes de escribir. Un mensaje que menciona algo concreto del local funciona mucho mejor que uno genérico.
- Empezá con una pregunta corta, sin links ni adjuntos.
- Escribí a mano y pocos por día. El envío masivo desde WhatsApp es la forma más rápida de que bloqueen tu número.
- Anotá en las columnas de seguimiento a quién escribiste y qué respondió.

## Desarrollo

```bash
python -m unittest discover -s tests -v
```

## Límites y avisos

- Los datos son los que figuran en Google Maps. Pueden estar desactualizados.
- Los teléfonos fijos no tienen WhatsApp. El link se arma para todos los números y hay que verificarlo.
- Revisá los términos de uso de la API de Google y la normativa de datos personales aplicable antes de usar la lista a gran escala.
