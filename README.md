# vinext-starter

A clean full-stack starter running on
[vinext](https://github.com/cloudflare/vinext), with optional Cloudflare D1 and
Drizzle support.

## Prerequisites

- Node.js `>=22.13.0`

## Quick Start

```bash
npm install
npm run dev
npm run build
```

This starter does not use `wrangler.jsonc`.

## El documento de Finnegans

La app no devuelve sólo texto: genera el `.docx` de "BUGS / PROBLEMAS
REPORTADOS" que soporte usa hoy.

No lo dibujamos desde cero. `app/docx.ts` toma la plantilla real
(`public/plantilla-finnegans.docx`), reemplaza únicamente el cuerpo de
`word/document.xml` y vuelve a comprimir el paquete. Todo el resto —estilos,
numeración, cabecera con la marca, pie y las 16 tipografías embebidas— se copia
sin tocar, así que lo generado es indistinguible de un documento escrito a mano
sobre la plantilla.

La disposición del cuerpo sigue un caso real escrito a mano, no la plantilla en
blanco: las preguntas de "Problema" van como lista de dos niveles (pregunta
arriba, respuesta anidada), los párrafos de cuerpo van justificados, y
"¿Tiene alguna contingencia manual?" es la quinta pregunta de ese bloque y no
una línea suelta en "Situación actual".

Sobre esa base el generador agrega:

- las capturas que el usuario pega con Ctrl+V dentro de cada cuadro de texto,
  incrustadas en la sección a la que ese cuadro corresponde y reescaladas al
  ancho útil de la caja sin deformarlas;
- los enlaces de Drive como hipervínculos reales (`TargetMode="External"`);
- dos secciones que la plantilla no pide pero que a soporte le sirven: el título
  del caso y las pruebas ya intentadas.

Las secciones opcionales —Finni, pruebas intentadas, evidencias en Drive— no se
escriben si están vacías: un "pendiente" dentro de un documento que ya se envía
es sólo ruido.

Si Finnegans publica una versión nueva de la plantilla, reemplazá el archivo de
`public/` y corré `npm run test:app`. Ese test corre el código de producción
contra la plantilla real y verifica que sigan estando los estilos, el `sectPr`
que engancha cabecera y pie, y todas las relaciones que el documento referencia.

## Subida a Drive y mail a soporte

Además de descargar el `.docx`, la app puede subirlo a una carpeta de Drive
**convertido a Google Doc** y abrir el mail a soporte con el enlace en el cuerpo.

El archivo no se adjunta al mail y no es un descuido: `mailto:` no puede
adjuntar —lo prohíbe el RFC 6068— y Finnegans pide recibir enlaces de Drive, no
archivos. Mandar el enlace evita además tener que registrar una app en Entra ID
para hablar con Microsoft Graph: `mailto:` funciona en cualquier Outlook.

Todo pasa en el navegador. El token sale de Google Identity Services con el
scope `drive.file`, que sólo alcanza los archivos que crea esta app y está
clasificado como no sensible, así que no hay que pasar por la verificación de
apps de Google.

Se configura con las variables de `.env.example`. **Si falta el client ID o la
carpeta, el botón no aparece** y la app sigue funcionando con la descarga.

El documento hereda los permisos de la carpeta destino: compartila una vez con
quien tenga que leer los casos y la app no necesita otorgar accesos por su
cuenta. Por eso el scope alcanza y ningún caso queda accesible por URL pública.

### Alta en Google Cloud, paso a paso

No hay que "aprender Google Cloud": la consola es sólo el lugar donde se emite
la credencial. Google no deja que una página cualquiera escriba en el Drive de
alguien, así que la app necesita identificarse. El client ID es ese documento.
Todo lo que sigue es gratis y no pide tarjeta.

1. **Crear el proyecto.** En <https://console.cloud.google.com> abrí el selector
   de proyecto arriba a la izquierda y creá uno nuevo (ej. `casos-finnegans`).
   Un proyecto es sólo una carpeta que agrupa credenciales y permisos.

2. **Habilitar la Drive API.** Entrá a
   <https://console.cloud.google.com/apis/library/drive.googleapis.com>, revisá
   que arriba figure el proyecto recién creado y dale *Habilitar*. Esto le dice
   al proyecto que puede hablar con Drive.

3. **Pantalla de consentimiento.** En *Google Auth Platform* (antes "OAuth
   consent screen"), <https://console.cloud.google.com/auth/overview>:
   - **Branding**: nombre de la app y un mail de contacto.
   - **Audience**: si la organización tiene Google Workspace, elegí **Internal**.
     Queda limitada a la organización y te ahorra el límite de 100 usuarios de
     prueba y el trámite de publicación. Si no, *External* + *Publish app*: con
     `drive.file` no hay verificación de por medio porque es un scope no sensible.

4. **Crear el client ID.** En *Google Auth Platform > Clients > Create client*:
   - Tipo: **Aplicación web**.
   - **Authorized JavaScript origins**: `http://localhost:3000` y la URL de
     Vercel (`https://…vercel.app`), sin barra final.
   - **Authorized redirect URIs**: se deja vacío. Google Identity Services abre
     un popup contra el origen, no redirige. Cargar acá la URL en vez de en
     orígenes es el error más común y da `origin_mismatch`.

   Copiá el client ID (`…apps.googleusercontent.com`). Es público por diseño:
   sólo sirve desde los orígenes que declaraste.

5. **La carpeta de Drive.** Esto ya no es Google Cloud. Creá la carpeta en Drive,
   compartila con quien tenga que leer los casos y copiá el ID de la URL: es lo
   que va después de `/folders/`.

6. **Cargar las variables.** Local: copiá `.env.example` a `.env.local` y
   completalo. En Vercel: *Settings > Environment Variables*, las mismas cuatro,
   y volvé a desplegar (los `NEXT_PUBLIC_*` se embeben en el build, así que un
   deploy viejo no las toma).

`npm run test:app` cubre el armado del multipart de conversión, los parámetros
de la subida y la codificación del `mailto`. El ida y vuelta real con Google no
está cubierto: hace falta un client ID y una cuenta.

## Deploy targets

El código de `app/` es Next.js App Router estándar, así que el proyecto compila
para dos runtimes distintos. Ninguno pisa al otro.

| Target | Comandos | Salida |
| --- | --- | --- |
| Cloudflare Workers (vinext) | `npm run dev` / `build` / `start` | `dist/` |
| Vercel (Next.js) | `npm run dev:next` / `build:next` / `start:next` | `.next/` |

Vercel lee `vercel.json`, que fuerza `next build`. Sin eso correría el script
`build`, que produce un bundle de Workers que Vercel no sabe ejecutar.

Los dos builds escriben en `.next/`, así que encadenarlos en el mismo árbol deja
tipos generados inconsistentes. Si vas a alternar, borrá `.next/` en el medio. Además,
`next build` reescribe `next-env.d.ts` con sus propias referencias en lugar de
`vinext/types`; los dos builds siguen pasando así, pero no lo edites a mano.

`worker/`, `db/` y `examples/` son exclusivos de Cloudflare y están excluidos del
`tsconfig.json`: apuntan al runtime de Workers, cuyos globals chocan con los del
DOM que usa la app.

## Included Shape

- edit site code under `app/`
- `.openai/hosting.json` declares optional Sites D1 and R2 bindings
- `vite.config.ts` simulates declared bindings for local development
- `db/schema.ts` starts intentionally empty
- `examples/d1/` contains an optional D1 example surface
- `drizzle.config.ts` supports local migration generation when needed

## Workspace Auth Headers

Signed-in visitors receive both `oai-authenticated-user-id` and `oai-authenticated-user-email`. Private Sites require every visitor to sign in; public Sites may also have anonymous visitors, for whom neither header is present.

The user ID is stable for the same user on the same Site and different across Sites. Email and name are intended for display or contact purposes.

SIWC-authenticated workspace sites may also receive
`oai-authenticated-user-full-name` when the user's SIWC profile has a non-empty
`name` claim. The full-name value is percent-encoded UTF-8 and is accompanied by
`oai-authenticated-user-full-name-encoding: percent-encoded-utf-8`.

Treat the full name as optional and fall back to email when it is absent:

```tsx
import { headers } from "next/headers";

export default async function Home() {
  const requestHeaders = await headers();
  const userId = requestHeaders.get("oai-authenticated-user-id");
  const email = requestHeaders.get("oai-authenticated-user-email");
  const encodedFullName = requestHeaders.get("oai-authenticated-user-full-name");
  const fullName =
    encodedFullName &&
    requestHeaders.get("oai-authenticated-user-full-name-encoding") ===
      "percent-encoded-utf-8"
      ? decodeURIComponent(encodedFullName)
      : null;

  const displayName = fullName ?? email;
  // ...
}
```

## Optional Dispatch-Owned ChatGPT Sign-In

Import the ready-to-use helpers from `app/chatgpt-auth.ts` when the site needs
optional or required ChatGPT sign-in:

- Use `getChatGPTUser()` for optional signed-in UI.
- Use `requireChatGPTUser(returnTo)` for server-rendered pages that should send
  anonymous visitors through Sign in with ChatGPT.
- Use `chatGPTSignInPath(returnTo)` and `chatGPTSignOutPath(returnTo)` for
  browser links or actions.
- Pass a same-origin relative `returnTo` path for the destination after sign-in
  or sign-out. The helper validates and safely encodes it.
- Mark protected pages with `export const dynamic = "force-dynamic"` because
  they depend on per-request identity headers.

Dispatch owns `/signin-with-chatgpt`, `/signout-with-chatgpt`, `/callback`, the
OAuth cookies, and identity header injection. Do not implement app routes for
those reserved paths. Routes that do not import and call the helper remain
anonymous-compatible.

SIWC establishes identity only; it does not prove workspace membership. Use the
Sites hosting platform's access policy controls for workspace-wide restrictions,
or enforce explicit server-side membership or allowlist checks.

Use SIWC for account pages, user-specific dashboards, saved records, and write
actions tied to the current ChatGPT user. Leave public content anonymous.

## Useful Commands

- `npm run dev`: start local development
- `npm run build`: verify the vinext build output
- `npm test`: build the starter and verify its rendered loading skeleton
- `npm run test:app`: verificar el .docx contra la plantilla real y el armado de la subida a Drive
- `npm run db:generate`: generate Drizzle migrations after schema changes

## Learn More

- [vinext Documentation](https://github.com/cloudflare/vinext)
- [Drizzle D1 Guide](https://orm.drizzle.team/docs/get-started/d1-new)
