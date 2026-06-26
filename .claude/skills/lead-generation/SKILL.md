---
name: lead-generation
description: Use when the user asks to search for prospects, find leads, update the prospects list, run a lead search, or automate lead generation for the web agency targeting local businesses in Buenos Aires without a web presence.
---

# Lead Generation — Agencia Web CABA

## Visión general

Búsqueda sistemática de comercios locales en Buenos Aires que no tienen presencia web propia, extracción de sus datos de contacto y guardado en `Prospectos_CABA.xlsx`. Puede ejecutarse manualmente o programarse de forma periódica.

---

## Paso 1 — Búsqueda de negocios locales

Usar `firecrawl_search` y `firecrawl_scrape` para encontrar comercios en:

| Fuente | Estrategia de búsqueda |
|--------|------------------------|
| Google Maps (scraping) | `"restaurante barrio palermo sin sitio web"`, `"peluquería villa crespo"`, etc. |
| Páginas Amarillas AR | `paginasamarillas.com.ar` por categoría y zona |
| Guía Oleo | `guiaoleo.com.ar` — restaurantes y gastronomía CABA |
| Yelp Argentina | `yelp.com.ar` — negocios con pocos reviews o sin web |
| Facebook / Instagram | Buscar páginas de negocios locales sin link a web propia |
| Directorios de barrios | Páginas de asociaciones vecinales o municipales |

**Queries de búsqueda recomendadas:**
```
site:paginasamarillas.com.ar "Buenos Aires" <rubro>
"<rubro> <barrio> Buenos Aires" -site:*.com -site:*.com.ar
firecrawl_search: "<rubro> <barrio> CABA teléfono"
```

Variar los rubros: carpintería, herrería, plomería, electricista, peluquería, veterinaria, kiosco, panadería, dietética, ferretería, taller mecánico, etc.

---

## Paso 2 — Filtro: "sin presencia digital"

Incluir un negocio en la lista SOLO si cumple al menos **uno** de estos criterios:

| Criterio | Cómo verificarlo |
|----------|-----------------|
| No tiene dominio propio | No hay URL de web en el perfil/directorio |
| Solo tiene redes sociales | Tiene Facebook o Instagram pero no web |
| Web rota o caída | `firecrawl_scrape` devuelve error o página en blanco |
| Web con más de 5 años sin actualizar | Copyright año viejo, tecnología Flash/HTML estático sin CMS |
| No aparece en Google con su nombre | `firecrawl_search "<nombre negocio> <barrio>"` no devuelve su web |

**Excluir** negocios con web moderna, e-commerce activo o franquicias con sitio corporativo.

---

## Paso 3 — Datos a extraer por prospecto

| Campo | Descripción |
|-------|-------------|
| `nombre` | Nombre comercial del negocio |
| `rubro` | Categoría (ej: peluquería, ferretería, veterinaria) |
| `direccion` | Calle y número o barrio / localidad |
| `telefono` | Teléfono o WhatsApp si está disponible |
| `email` | Email de contacto si está disponible |
| `facebook` | URL de la página de Facebook |
| `instagram` | URL del perfil de Instagram |
| `fuente` | Directorio o plataforma donde fue encontrado |
| `fecha_busqueda` | Fecha en formato YYYY-MM-DD |
| `observaciones` | Ej: "solo tiene Facebook", "web rota", "no aparece en Google" |
| `prioridad` | 1 = baja, 2 = media, 3 = alta (ver criterios abajo) |

### Criterios de prioridad

- **3 (alta):** Solo tiene redes sociales o directamente no tiene nada digital
- **2 (media):** Tiene web muy vieja, rota o sin actualizar
- **1 (baja):** Tiene web básica pero podría mejorar mucho

---

## Paso 4 — Guardado en Excel

Archivo: `Prospectos_CABA.xlsx` en la raíz del proyecto.

### Reglas de escritura

1. **Si el archivo no existe:** crearlo con los encabezados en el orden exacto de la tabla del Paso 3.
2. **Si el archivo ya existe:** cargar el contenido actual y AGREGAR filas al final — nunca borrar datos previos.
3. **Deduplicación:** antes de agregar un prospecto, verificar que no exista ya una fila con el mismo `nombre` + `direccion`. Si ya existe, omitir sin error.

### Código de referencia (Python / openpyxl)

```python
import openpyxl
from pathlib import Path
from datetime import date

COLUMNS = ["nombre","rubro","direccion","telefono","email",
           "facebook","instagram","fuente","fecha_busqueda",
           "observaciones","prioridad"]

xlsx_path = Path("Prospectos_CABA.xlsx")

if xlsx_path.exists():
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    existing = {(r[0], r[2]) for r in ws.iter_rows(min_row=2, values_only=True)}
else:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(COLUMNS)
    existing = set()

nuevos = 0
for lead in leads:  # lista de dicts con las claves de COLUMNS
    key = (lead["nombre"], lead["direccion"])
    if key not in existing:
        ws.append([lead.get(c, "") for c in COLUMNS])
        existing.add(key)
        nuevos += 1

wb.save(xlsx_path)
print(f"✅ {nuevos} leads nuevos agregados a {xlsx_path}")
```

---

## Paso 5 — Automatización y programación

### Ejecución manual

El agente debe responder a frases como:
- "buscá leads nuevos"
- "actualizá los prospectos"
- "corré la búsqueda de clientes"
- "encontrá comercios sin web"

Al recibir estas frases: ejecutar los Pasos 1–4 completos y reportar cuántos leads nuevos se agregaron.

### Búsqueda periódica

Para configurar ejecuciones automáticas recurrentes (ej: semanal, quincenal), usar la skill `schedule`:

```
/schedule
→ Tarea: "Búsqueda semanal de leads para agencia web"
→ Cron: "0 9 * * 1"  (todos los lunes a las 9 AM)
→ Prompt: "Ejecutá la skill lead-generation: buscá leads nuevos en CABA y actualizá Prospectos_CABA.xlsx"
```

### Reporte al finalizar

Al terminar cada búsqueda (manual o programada), mostrar siempre:

```
Búsqueda finalizada — [FECHA]
• Leads nuevos agregados: X
• Total en el archivo: Y
• Fuentes revisadas: [lista]
• Próxima búsqueda programada: [fecha si aplica]
```

---

## Errores comunes

| Problema | Solución |
|----------|----------|
| `firecrawl_scrape` falla en una URL | Registrar la fuente como no disponible y continuar |
| El Excel está abierto en otro programa | Avisar al usuario que cierre el archivo antes de guardar |
| Prospecto sin teléfono ni email | Agregar igual — el campo queda vacío, la observación lo aclara |
| Resultado duplicado de distintas fuentes | La deduplicación por nombre + dirección lo filtra automáticamente |
