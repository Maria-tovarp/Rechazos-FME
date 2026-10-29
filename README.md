# Gestión de Rechazos FME — V12 (Supabase)

V12 elimina la dependencia diaria de los archivos maestros y usa Supabase como almacenamiento persistente.

## Cambios principales V12

- Fuente oficial de coordinadores: `DATA COORDINADORES - NUEVA.xlsx`.
- Hoja principal: `Libro3`.
- Prioridad de búsqueda equivalente al BUSCARX original:
  1. `Libro3[CODIGO]`
  2. `Libro3[CodigoNetco]`
  3. `Hoja1[SITIO]`
- `#N/A`, `N/A`, `NA`, `NAN` y `-` no se aceptan como coordinador válido.
- La lectura se hace por **nombre de encabezado**, no por posición de columna.
- Se guarda un maestro detallado en `coordinadores` y un mapa rápido en `site_coordinator_map`.
- El uso diario solo requiere cargar `WO Result Query`.
- Se conserva el ciclo de auditoría/feedback: `Feedback Qty = 0` es rechazo inicial, `1` primer feedback, `2` segundo, etc.
- En ciclos de feedback, `Feedback Remark = N/A` se interpreta como **APROBADO**; si trae texto, queda como **RECHAZADO NUEVAMENTE**.
- Estados de SMS/WhatsApp, feedback OWS y gestiones Closed permanecen en Supabase.

## 1. Actualizar el esquema de Supabase

Abre **Supabase > SQL Editor**, pega el contenido de `schema_v12.sql` y ejecútalo.

El script es compatible con la V11: crea/actualiza las tablas necesarias y la función `buscar_coordinador(p_site_id)`.

## 2. Secrets de Streamlit

En Streamlit Community Cloud > Settings > Secrets:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_KEY = "sb_secret_TU_CLAVE"
APP_PASSWORD = "opcional"
```

No subas la secret key a GitHub.

## Despliegue gratuito en Render

El proyecto incluye `render.yaml` para desplegarlo como Web Service.

1. Sube el código a un repositorio privado de GitHub, sin `.streamlit/secrets.toml`.
2. En [Render](https://render.com), selecciona **New > Blueprint** y conecta el repositorio.
3. Mantén el plan **Free** y configura estas variables de entorno en Render:

```text
SUPABASE_URL
SUPABASE_KEY
APP_PASSWORD
```

Render asignará una URL pública. En el plan gratuito, el servicio se suspende tras 15 minutos sin tráfico y tarda unos segundos en reactivarse.

## 3. Actualizar maestros

En la app abre **Migración / actualización de maestros**.

### Coordinadores
Carga `DATA COORDINADORES - NUEVA.xlsx` y pulsa **Actualizar coordinadores en Supabase**.

La operación reemplaza el mapa anterior para evitar que queden sitios obsoletos y actualiza el maestro detallado.

### FME / celulares
Carga `CUADRILLAS ACT. NUEVO PROY. 2026.xlsx` cuando cambie la información de cuadrillas.

## 4. Uso diario

1. Abre **Importar WO**.
2. Carga el `WO Result Query` más reciente.
3. La app consulta coordinador y FME desde Supabase.
4. Guarda/actualiza actividades y ciclos.
5. Usa Mensajes FME, Feedback OWS, Closed / Exclusiones y Vista coordinadores.

## Regla de coordinadores

Para un `Site ID = X`:

```text
Libro3.CODIGO == X        -> Coordinador
si no:
Libro3.CodigoNetco == X   -> Coordinador
si no:
Hoja1.SITIO == X          -> COORDINADOR
si no:
Por validar
```

## Archivos del proyecto

- `app.py`: interfaz Streamlit.
- `core.py`: lectura de Excel, clasificación de ciclos y reglas de negocio.
- `database.py`: conexión y persistencia Supabase.
- `schema_v12.sql`: esquema/actualización recomendada.
- `schema.sql`: copia del esquema V12 para instalaciones nuevas.
- `requirements.txt`: dependencias.
- `.streamlit/secrets.example.toml`: ejemplo de configuración.
