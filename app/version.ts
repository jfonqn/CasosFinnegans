/*
 * Identidad de la build, para poder pedirle a alguien "¿qué versión tenés?"
 * cuando reporta algo raro y saber contra qué código estamos mirando.
 */

/* Duplicado a propósito respecto de package.json: importarlo acá arrastraría al
   bundle del cliente la lista entera de dependencias y sus versiones.
   tests/version.test.mjs verifica que los dos números no se desincronicen. */
export const APP_VERSION = "1.0.0";

/* Vercel lo completa solo, pero sólo si en el proyecto está tildado "Enable
   access to System Environment Variables". Si no, o si la build no vino de un
   commit, llega vacío y mostramos únicamente la versión. */
export const BUILD_COMMIT = (process.env.NEXT_PUBLIC_VERCEL_GIT_COMMIT_SHA ?? "").slice(0, 7);

export const BUILD_LABEL = BUILD_COMMIT ? `v${APP_VERSION} · ${BUILD_COMMIT}` : `v${APP_VERSION}`;

/* Lo que sirve tener a mano cuando alguien reporta un problema: además de la
   versión, el navegador y el tamaño de pantalla, que es donde más se rompe. */
export function diagnostics(extra: Record<string, string> = {}) {
  const viewport =
    typeof window === "undefined" ? "-" : `${window.innerWidth}x${window.innerHeight} @${window.devicePixelRatio}x`;
  const lines = {
    Version: APP_VERSION,
    Commit: BUILD_COMMIT || "(no disponible)",
    Pantalla: viewport,
    Navegador: typeof navigator === "undefined" ? "-" : navigator.userAgent,
    Fecha: new Date().toISOString(),
    ...extra,
  };
  return Object.entries(lines)
    .map(([key, value]) => `${key}: ${value}`)
    .join("\n");
}
