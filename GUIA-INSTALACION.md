# App NÁCAR móvil — Guía de puesta en marcha

La aplicación ya está construida y probada. Está en esta carpeta (`NACAR-APP`) y
ahora mismo funciona en **modo demo** (con una muestra de datos reales incrustada).
Para conectarla a tus datos de verdad hay que hacer **dos gestiones, una sola vez**.
Ninguna de las dos requiere tocar MN Program ni las carpetas de documentos.

> Consejo: estas dos gestiones se pueden hacer en una sesión de Claude
> ("ayúdame a registrar la app de NACAR-APP en Entra y publicarla") y Claude
> puede ir guiándote pantalla a pantalla, o incluso manejar el navegador contigo.

---

## Paso 1 — Registrar la aplicación en Microsoft Entra (10 min)

Esto le dice a Microsoft "existe una app del despacho que puede leer (solo leer)
OneDrive y el calendario del usuario que inicie sesión". Hace falta hacerlo con
una cuenta administradora de vuestro Microsoft 365 (si no lo eres tú, lo es quien
os gestione el correo).

1. Entra en **https://entra.microsoft.com** con la cuenta del despacho.
2. Menú izquierdo: **Identidad → Aplicaciones → Registros de aplicaciones → Nuevo registro**.
3. Nombre: `NACAR App Movil`.
4. Tipos de cuenta: **Solo las cuentas de este directorio organizativo**.
5. URI de redirección: elige **Aplicación de página única (SPA)** y escribe la
   dirección donde vivirá la app (la del Paso 2, por ejemplo
   `https://TUUSUARIO.github.io/nacar-app/`). Se puede añadir o cambiar después.
6. Pulsa **Registrar**. En la pantalla que aparece, copia dos datos:
   - **Id. de aplicación (cliente)**  →  es el `clientId`
   - **Id. de directorio (inquilino)**  →  es el `tenantId`
7. Menú **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos delegados**
   y marca estos tres:
   - `User.Read`
   - `Files.Read.All`
   - `Calendars.Read`
8. Pulsa **Agregar permisos** y después **Conceder consentimiento de administrador**
   (botón con el tic ✓).

> Todos los permisos son **delegados y de lectura**: la app solo puede ver lo que
> ya puede ver el usuario que inicia sesión, y no puede modificar nada.

## Paso 2 — Publicar la app (10 min)

La app son 7 ficheros estáticos; vale cualquier sitio que sirva páginas web con
https. La opción más sencilla y gratuita es **GitHub Pages** (el código de la app
es público, pero NO contiene ningún dato del despacho — los datos solo se cargan
tras iniciar sesión y van directos de Microsoft a tu teléfono):

1. Crea una cuenta gratuita en **https://github.com** (si no tienes).
2. **New repository** → nombre `nacar-app` → público → **Create repository**.
3. **Add file → Upload files** y arrastra los 7 ficheros de esta carpeta
   (`index.html`, `app.js`, `config.js`, `demo-data.js`, `manifest.json`,
   `icon-180.png`, `icon-512.png`). **Commit changes**.
4. **Settings → Pages → Branch: main → Save**. En un minuto tu app estará en
   `https://TUUSUARIO.github.io/nacar-app/`.

Alternativa si el despacho tiene informático o suscripción de Azure: Azure Static
Web Apps (plan gratuito) dentro del propio tenant. El resultado es el mismo.

## Paso 3 — Conectar (2 min)

1. Abre `config.js` (con el Bloc de notas) y rellena:
   ```
   clientId: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
   tenantId: "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy",
   ```
   con los dos valores copiados en el Paso 1.6.
2. Vuelve a subir `config.js` a GitHub (Upload files → reemplazar).
3. Comprueba que en Entra (Paso 1.5) la URI de redirección es exactamente la
   dirección real de la app.

## Paso 4 — Instalarla en el iPhone (1 min)

1. Abre la dirección de la app en **Safari**.
2. Inicia sesión con tu cuenta `sotero@nacarabogados.com`.
3. Botón **Compartir** (cuadrado con flecha) → **Añadir a pantalla de inicio**.
4. Ya tienes el icono de NÁCAR como una app más. La sesión queda guardada y el
   propio bloqueo del iPhone (Face ID) protege el acceso.

---

## Uso semanal (2 minutos)

- **Documentos y agenda**: se actualizan solos, en tiempo real. No hay que hacer nada.
- **Clientes y expedientes**: una vez por semana, desde MN Program, exporta
  **Clientes** a Excel y **Expedientes** a Excel y guárdalos en la carpeta
  `Descargas` de OneDrive (como ya haces). La app detecta sola los ficheros más
  recientes — da igual cómo se llamen.
- Si pasan más de 7 días sin exportación nueva, la app te lo avisa en la pantalla Hoy.

## Qué hace y qué no hace

- Solo lectura: la app no puede crear, modificar ni borrar nada, ni en OneDrive,
  ni en el calendario, ni en MN Program.
- Los datos viajan directamente de Microsoft 365 a tu teléfono con tu sesión;
  no pasan por ningún servidor de terceros.
- Pestañas: **Hoy** (juicios, plazos y avisos), **Buscar** (todo: clientes, autos,
  NIG, contrarios, juzgados, documentos) y **Agenda** (calendario clasificado).
  El engranaje arriba a la derecha abre Ajustes (estado de los datos, calidad de
  datos, actualizar, cerrar sesión).
