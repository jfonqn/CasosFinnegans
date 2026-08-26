/*
 * app/version.ts repite el número de package.json a mano para no arrastrar el
 * package.json entero al bundle del cliente. Este test es lo que sostiene esa
 * decisión: si los dos se desincronizan, falla acá y no en producción.
 *
 *   node --test tests/version.test.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const pkg = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));
const { APP_VERSION, BUILD_LABEL, BUILD_COMMIT, diagnostics } = await import(
  pathToFileURL(path.join(ROOT, "app/version.ts")).href
);

test("la versión mostrada coincide con la de package.json", () => {
  assert.equal(APP_VERSION, pkg.version);
  assert.match(APP_VERSION, /^\d+\.\d+\.\d+$/);
});

test("sin commit muestra sólo la versión", () => {
  /* Vercel deja la variable vacía si no está habilitado el acceso a las
     variables de sistema, o si la build no vino de un commit. */
  assert.equal(BUILD_COMMIT, "");
  assert.equal(BUILD_LABEL, `v${pkg.version}`);
});

test("el diagnóstico trae lo que hace falta para triaje", () => {
  const texto = diagnostics({ Drive: "configurado" });
  for (const clave of ["Version:", "Commit:", "Pantalla:", "Navegador:", "Fecha:", "Drive: configurado"]) {
    assert.ok(texto.includes(clave), `falta ${clave}`);
  }
  assert.ok(texto.includes(pkg.version));
  /* Fuera del navegador no debe romper: se usa también en el render del server. */
  assert.ok(texto.includes("Commit: (no disponible)"));
});
