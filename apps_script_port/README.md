# Mesa de Pedidos Judiciales — Port a Google Apps Script

Port funcional completo de la app de Google AI Studio (React + Express +
JSON local) a **Google Apps Script + Google Sheets como base de datos**.
Apps Script no ejecuta builds de Vite/React, así que la interfaz fue
reescrita en HTML/CSS/JS plano (con Tailwind vía CDN y `google.script.run`
en lugar de `fetch('/api/...')`).

## Qué incluye (paridad de las 7 pantallas)

1. **Registrar Pedido** — formulario completo + vista previa de WhatsApp
   (3 formatos: clásico, con emojis, compacto), plantillas rápidas de
   requerimiento, datos del solicitante ocultables.
2. **Consultar Estado** (vista pública) — buscador por expediente /
   juzgado / especialista.
3. **Bandeja Admin** — búsqueda, filtros (estado / tipo / juzgado),
   edición de estado y notas internas, eliminar, copiar texto, exportar
   CSV y JSON, imprimir.
4. **Directorio de Juzgados** — buscar/filtrar, seleccionar juzgado +
   especialista para el formulario, alta/edición/baja (admin), restablecer
   a datos originales.
5. **Configuración** (admin) — datos del grupo, tipos de trámite,
   materias frecuentes, cambio de contraseña de administrador.
6. **Importador rápido** — pega un mensaje de WhatsApp y autocompleta el
   formulario (mismo parser por expresiones regulares del original).
7. **Compartir** — enlace público, enlace de administrador
   (`?view=admin`) y mensaje de difusión sugerido.

## Diferencias respecto al original

- **Base de datos:** en vez de un archivo `data/pedidos.json` en un
  servidor Node, los datos viven en 3 hojas de una Google Sheet:
  `Pedidos`, `Juzgados` y `Config`. Esto además te permite abrir/editar
  los datos directamente en Sheets si lo necesitas (filtros, tablas
  dinámicas, etc.).
- **Contraseña de administrador:** se guarda en *Script Properties* (no
  en la hoja ni en el HTML), para que no quede visible si compartes el
  Sheet.
- **Autenticación:** igual que el original, el "modo administrador" es
  una bandera en el cliente (una vez ingresada la contraseña correcta),
  no una sesión con tokens server-side. Es suficiente para coordinar un
  grupo interno, pero no trates la contraseña como un control de
  seguridad fuerte — cualquiera con el enlace `?view=admin` puede
  intentar iniciar sesión.
- **Sin dependencias de Node/Express/Vite/React**: todo corre dentro de
  Apps Script + HTML Service.
- Los textos de las plantillas rápidas de "Requerimiento" se
  generalizaron (el original tenía algunas plantillas de ejemplo
  incompletas, tipo marcador de posición); edítalas a gusto en
  `JavaScript.html` → `App.templates()`.

## Despliegue paso a paso

1. **Crea una Google Sheet nueva** (sheets.new) y ponle un nombre, p.
   ej. "Mesa de Pedidos Judiciales — Base de Datos".
2. Ve a **Extensiones → Apps Script**. Se abrirá el editor con un
   `Code.gs` vacío.
3. Borra el contenido de `Code.gs` y pega el contenido de
   [`Code.gs`](Code.gs) de este proyecto.
4. En el editor, crea 3 archivos HTML nuevos (ícono **+** → HTML) con
   estos nombres exactos y pega el contenido correspondiente:
   - `Index` → contenido de [`Index.html`](Index.html)
   - `Stylesheet` → contenido de [`Stylesheet.html`](Stylesheet.html)
   - `JavaScript` → contenido de [`JavaScript.html`](JavaScript.html)
5. Abre el archivo de manifiesto: menú **Configuración del proyecto**
   (ícono de engranaje) → activa **"Mostrar el archivo de manifiesto
   'appsscript.json' en el editor"**. Luego abre `appsscript.json` y
   reemplaza su contenido por el de [`appsscript.json`](appsscript.json)
   de este proyecto (define que el Web App es accesible por "Cualquier
   usuario").
6. Guarda todo (Ctrl+S / ícono de disquete).
7. En la barra de funciones del editor, selecciona la función
   `inicializarHojas` y presiona **Ejecutar** (▶). La primera vez te
   pedirá autorizar permisos — acepta (es tu propio script, sobre tu
   propia hoja). Esto crea las hojas `Pedidos`, `Juzgados` y `Config`
   con los datos de ejemplo, y fija la contraseña de administrador
   inicial en `admin123`.
8. Vuelve a la hoja de cálculo y confirma que aparecieron las 3
   pestañas con datos.
9. En el editor de Apps Script, haz clic en **Implementar → Nueva
   implementación**. Tipo: **Aplicación web**. Configura:
   - Ejecutar como: **Yo (tu correo)**
   - Quién tiene acceso: **Cualquier usuario** (para que los 500
     miembros del grupo puedan usarla sin iniciar sesión de Google), o
     **Cualquier usuario de [tu organización]** si prefieres
     restringirlo.
   - Haz clic en **Implementar** y copia la **URL de la aplicación
     web**. Esa es la URL que compartes en el grupo (equivalente a la
     URL de AI Studio que tenías antes).
10. Para el enlace de administrador, usa la misma URL agregando
    `?view=admin` al final. Al abrirla te pedirá la contraseña
    (`admin123` por defecto — cámbiala de inmediato desde la pestaña
    **Configuración** una vez dentro).
11. Cada vez que edites el código en el editor y quieras que el cambio
    se refleje en la URL pública, debes hacer **Implementar → Administrar
    implementaciones → editar (ícono de lápiz) → Nueva versión →
    Implementar**. (Guardar el archivo NO actualiza automáticamente la
    versión publicada, es una particularidad de Apps Script.)

## Notas de mantenimiento

- Si algo falla al cargar ("Error cargando datos..."), casi siempre es
  porque falta ejecutar `inicializarHojas` (paso 7), o porque se borró
  alguna de las 3 hojas.
- El menú personalizado **"Mesa de Pedidos Judiciales"** en la hoja de
  cálculo (visible al abrir el Sheet) te permite reinicializar las
  hojas, restablecer el directorio de juzgados o restablecer la
  configuración sin tocar el editor de código.
- Los datos quedan en tu propia Google Sheet: puedes hacer copias de
  seguridad simplemente duplicando el archivo (Archivo → Crear una
  copia) o exportando desde la pestaña **Bandeja Admin** (botones CSV /
  JSON).
