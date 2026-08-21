/*
 * Verifica el .docx que genera app/docx.ts contra la plantilla real.
 *
 * Corre el código de producción tal cual: sólo stubbea los dos globals del
 * navegador que Node no tiene (fetch al asset público y createImageBitmap).
 *
 *   node --test tests/docx.test.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import path from "node:path";
import { unzipSync, strFromU8 } from "fflate";

const ROOT = path.resolve(import.meta.dirname, "..");
const template = readFileSync(path.join(ROOT, "public/plantilla-finnegans.docx"));

globalThis.fetch = async (url) => {
  assert.equal(String(url), "/plantilla-finnegans.docx");
  return {
    ok: true,
    status: 200,
    arrayBuffer: async () => template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength),
  };
};

/* Dimensiones fijas por nombre: en el navegador esto lo resuelve createImageBitmap. */
const SIZES = { "ancha.png": [1600, 700], "chica.png": [420, 260] };
globalThis.createImageBitmap = async (file) => {
  const [width, height] = SIZES[file.name];
  return { width, height, close() {} };
};

const { buildCaseDocx, caseFileName } = await import(pathToFileURL(path.join(ROOT, "app/docx.ts")).href);

/* Un PNG 1x1 valido: alcanza porque las dimensiones las da el stub. */
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);
const shot = (name) => new File([PNG], name, { type: "image/png" });

const data = {
  caseNumber: "",
  date: "21/8/2026",
  domain: "INMIX",
  client: "Cooperativa Agrícola & Lechera <Sur>",
  author: "Bruno Merino",
  company: "Planta Osorno / Sucursal 02",
  dedicatedServer: "Sí",
  priority: "Alta",
  title: "El nuevo motor de retenciones no se aplica",
  finni: "Prompt: ¿por qué?\nRespuesta: revisar el padrón.",
  whatFails: "El cálculo usa procesos desactivados.",
  whatHappens: "Calcula con el motor viejo.",
  firstTime: "No",
  standardOperation: "Sí",
  impact: "No se pueden emitir comprobantes.\nAfecta a 14 clientes.",
  workaround: "Se carga una excepción manual.",
  steps: "1. Ingresar a Retenciones\n2. Actualizar el padrón\n3. Emitir la factura",
  reportFormat: "Partes por máquina",
  attempts: "Se probó recalcular desde el ABM.",
  expected: "Debería aplicar el nuevo motor.",
  driveLinks: ["https://drive.google.com/file/d/abc123/view"],
  accessConfirmed: true,
  images: { steps: [shot("ancha.png")], impact: [shot("chica.png")] },
};

const blob = await buildCaseDocx(data);
const parts = unzipSync(new Uint8Array(await blob.arrayBuffer()));
const doc = strFromU8(parts["word/document.xml"]);
const rels = strFromU8(parts["word/_rels/document.xml.rels"]);

const at = (text) => doc.indexOf(`>${text}</w:t>`);
/* Devuelve el <w:p> completo que contiene ese texto. */
const paragraphOf = (text) => {
  const i = at(text);
  return doc.slice(doc.lastIndexOf("<w:p>", i), doc.indexOf("</w:p>", i));
};

test("conserva las partes de la plantilla", () => {
  for (const part of [
    "[Content_Types].xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/header1.xml",
    "word/footer2.xml",
    "word/media/image1.png",
    "word/media/image2.png",
    "word/fonts/Poppins-regular.ttf",
  ]) {
    assert.ok(parts[part], `falta ${part}`);
  }
  /* El sectPr es el que engancha cabecera y pie: sin él se pierde la marca. */
  assert.match(doc, /<w:headerReference r:id="rId6" w:type="default"\/>/);
  assert.match(doc, /<w:footerReference r:id="rId8" w:type="default"\/>/);
});

test("escribe las secciones que pide la plantilla", () => {
  for (const heading of [
    "BUGS / PROBLEMAS REPORTADOS",
    "Pegar respuesta y promt de Finni",
    "Problema",
    "Situación actual",
    "Caso de uso",
    "Resultado esperado",
    "Evidencias en Google Drive",
  ]) {
    assert.ok(at(heading) >= 0, `falta la sección ${heading}`);
  }
});

test("plantea las preguntas como lista de dos niveles", () => {
  /* El caso escrito a mano pone la pregunta en el nivel 0 y la respuesta
     anidada en el nivel 1, no ambas en la misma viñeta. */
  const level = (text) => paragraphOf(text).match(/<w:ilvl w:val="(\d)"\/>/)?.[1] ?? null;
  assert.equal(level("¿Qué no funciona?"), "0");
  assert.equal(level("El cálculo usa procesos desactivados."), "1");
  assert.equal(level("¿Es una operación estándar?"), "0");
  assert.equal(level("Sí"), "1");
  /* La contingencia manual es la quinta pregunta del bloque Problema. */
  assert.ok(at("¿Tiene alguna contingencia manual?") > at("Problema"));
  assert.ok(at("¿Tiene alguna contingencia manual?") < at("Situación actual"));
  assert.equal(level("Se carga una excepción manual."), "1");
  assert.ok(at("Debe responder las siguientes preguntas") > at("Problema"));
});

test("justifica los párrafos de cuerpo", () => {
  for (const text of ["Debe responder las siguientes preguntas", "¿Qué no funciona?", "No se pueden emitir comprobantes."]) {
    assert.match(paragraphOf(text), /<w:jc w:val="both"\/>/, `sin justificar: ${text}`);
  }
});

test("respeta los saltos de línea de los pasos", () => {
  for (const step of ["1. Ingresar a Retenciones", "2. Actualizar el padrón", "3. Emitir la factura"]) {
    assert.ok(at(step) >= 0, `el paso "${step}" no quedó en su propio párrafo`);
  }
});

test("escapa el XML de los valores del usuario", () => {
  assert.ok(doc.includes("Cooperativa Agrícola &amp; Lechera &lt;Sur&gt;"));
  assert.ok(!doc.includes("Lechera <Sur>"));
});

test("incrusta cada captura en la sección donde fue pegada", () => {
  assert.ok(parts["word/media/evidencia1.png"]);
  assert.ok(parts["word/media/evidencia2.png"]);
  const shots = [...doc.matchAll(/<a:blip r:embed="(rId\d+)"\/>/g)].map((m) => ({ id: m[1], at: m.index }));
  assert.equal(shots.length, 2);
  const byId = Object.fromEntries(shots.map((s) => [s.id, s.at]));
  /* El slot "impact" se emite en Situación actual y "steps" en Caso de uso.
     rId1001 es el primero preparado y el orden de slots pone impact antes de steps. */
  const [first, second] = shots;
  assert.ok(first.at > at("Situación actual") && first.at < at("Caso de uso"), "la captura de impacto no quedó en Situación actual");
  assert.ok(second.at > at("Caso de uso"), "la captura de pasos no quedó en Caso de uso");
  assert.ok(second.at < at("Pruebas o soluciones intentadas"), "la captura de pasos se pasó de sección");
  assert.equal(Object.keys(byId).length, 2);
});

test("escala las capturas al ancho útil sin deformarlas", () => {
  const extents = [...doc.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => [+m[1], +m[2]]);
  assert.equal(extents.length, 2);
  /* chica.png (420x260) va primera por el orden de slots y no se estira. */
  assert.deepEqual(extents[0], [420 * 9525, 260 * 9525]);
  /* ancha.png (1600x700) se reduce al ancho de caja conservando proporción. */
  assert.equal(extents[1][0], 5847461);
  assert.equal(Math.round((extents[1][0] / extents[1][1]) * 100), Math.round((1600 / 700) * 100));
});

test("pisa el cuerpo de los títulos como hace la plantilla", () => {
  /* Los estilos traen Title=48pt: sin override el encabezado ocupa media página. */
  const sized = (text) => {
    const i = at(text);
    const open = doc.lastIndexOf("<w:r>", i);
    return doc.slice(open, i).match(/<w:sz w:val="(\d+)"\/>/)?.[1] ?? null;
  };
  assert.equal(sized("BUGS / PROBLEMAS REPORTADOS"), "34");
  assert.equal(sized("N° de caso"), "60");
  assert.equal(sized("Servidor propio / dedicado: SÍ"), "36");
  /* Subtítulos y Heading4 sí usan el cuerpo del estilo, como en la plantilla. */
  assert.equal(sized("Fecha de Ingreso: 21/8/2026"), null);
  assert.equal(sized("Situación actual"), null);
});

test("omite las secciones opcionales vacías", async () => {
  const pelado = await buildCaseDocx({
    ...data,
    finni: "",
    attempts: "",
    driveLinks: [],
    accessConfirmed: false,
    images: {},
  });
  const xml = strFromU8(unzipSync(new Uint8Array(await pelado.arrayBuffer()))["word/document.xml"]);
  assert.ok(!xml.includes("Pegar respuesta y promt de Finni"));
  assert.ok(!xml.includes("Pruebas o soluciones intentadas"));
  assert.ok(!xml.includes("Evidencias en Google Drive"));
  assert.ok(!xml.includes("Acceso verificado para soporte"));
  assert.ok(xml.includes(">Resultado esperado</w:t>"), "el resto del documento sigue intacto");
});

test("una sección opcional sobrevive si sólo tiene capturas", async () => {
  const soloImagen = await buildCaseDocx({ ...data, finni: "", images: { finni: [shot("chica.png")] } });
  const xml = strFromU8(unzipSync(new Uint8Array(await soloImagen.arrayBuffer()))["word/document.xml"]);
  assert.ok(xml.includes("Pegar respuesta y promt de Finni"), "no debe descartar la sección por tener sólo imagen");
});

test("declara toda relación que referencia", () => {
  const declared = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  const used = new Set([...doc.matchAll(/r:(?:id|embed)="([^"]+)"/g)].map((m) => m[1]));
  const dangling = [...used].filter((id) => !declared.has(id));
  assert.deepEqual(dangling, [], `r:id sin declarar: ${dangling}`);
  assert.match(rels, /Target="https:\/\/drive\.google\.com\/file\/d\/abc123\/view" TargetMode="External"/);
});

test("deja vacío el resultado esperado cuando no se sabe", async () => {
  const sinResultado = await buildCaseDocx({ ...data, expected: "   ", images: {} });
  const xml = strFromU8(unzipSync(new Uint8Array(await sinResultado.arrayBuffer()))["word/document.xml"]);
  const after = xml.slice(xml.indexOf(">Resultado esperado</w:t>"));
  const next = after.slice(0, after.indexOf("Evidencias en Google Drive"));
  assert.ok(!next.includes("Pendiente de completar"), "no debe forzar un placeholder");
});

test("arma un nombre de archivo usable", () => {
  assert.equal(caseFileName(data), "Caso - INMIX - El nuevo motor de retenciones no se aplica.docx");
  assert.equal(caseFileName({ domain: "", title: 'a/b:c*d?e"f<g>h|i' }), "Caso - abcdefghi.docx");
});
