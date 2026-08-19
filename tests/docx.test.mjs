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

const data = {
  caseNumber: "",
  date: "19/8/2026",
  domain: "COLUN",
  client: "Cooperativa Agrícola & Lechera <Sur>",
  author: "Bruno Merino",
  company: "Planta Osorno / Sucursal 02",
  dedicatedServer: "Sí",
  priority: "Alta",
  title: "El parte de trabajo no actualiza el horómetro",
  finni: "Prompt: ¿por qué?\nRespuesta: revisar el ABM.",
  whatFails: "El horómetro no se actualiza.",
  whatHappens: "Se guarda sin error pero el valor no cambia.",
  firstTime: "No",
  standardOperation: "Sí",
  impact: "Mantenimiento no puede programar los service.\nAfecta a 14 máquinas.",
  workaround: "Se lleva una planilla aparte.",
  steps: "1. Ingresar a Partes de trabajo\n2. Cargar 8 horas\n3. Guardar\n4. Abrir el ABM",
  reportFormat: "Partes por máquina",
  attempts: "Se probó recalcular desde el ABM.",
  expected: "Debería pasar de 1.240 a 1.248 horas.",
  driveLinks: ["https://drive.google.com/file/d/abc123/view"],
  accessConfirmed: true,
  images: [
    new File([PNG], "ancha.png", { type: "image/png" }),
    new File([PNG], "chica.png", { type: "image/png" }),
  ],
};

const blob = await buildCaseDocx(data);
const parts = unzipSync(new Uint8Array(await blob.arrayBuffer()));
const doc = strFromU8(parts["word/document.xml"]);
const rels = strFromU8(parts["word/_rels/document.xml.rels"]);

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
    assert.ok(doc.includes(`>${heading}</w:t>`), `falta la sección ${heading}`);
  }
});

test("respeta los saltos de línea de los pasos", () => {
  const steps = ["1. Ingresar a Partes de trabajo", "2. Cargar 8 horas", "3. Guardar", "4. Abrir el ABM"];
  for (const step of steps) assert.ok(doc.includes(`>${step}</w:t>`), `el paso "${step}" no quedó en su propio párrafo`);
});

test("escapa el XML de los valores del usuario", () => {
  assert.ok(doc.includes("Cooperativa Agrícola &amp; Lechera &lt;Sur&gt;"));
  assert.ok(!doc.includes("Lechera <Sur>"));
});

test("embebe las imágenes y las escala al ancho útil", () => {
  assert.ok(parts["word/media/evidencia1.png"]);
  assert.ok(parts["word/media/evidencia2.png"]);
  const extents = [...doc.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => [+m[1], +m[2]]);
  assert.equal(extents.length, 2);
  /* La ancha se reduce al ancho de caja conservando proporción; la chica no se estira. */
  assert.equal(extents[0][0], 5847461);
  assert.equal(Math.round((extents[0][0] / extents[0][1]) * 100), Math.round((1600 / 700) * 100));
  assert.deepEqual(extents[1], [420 * 9525, 260 * 9525]);
});

test("declara toda relación que referencia", () => {
  const declared = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  const used = new Set([...doc.matchAll(/r:(?:id|embed)="([^"]+)"/g)].map((m) => m[1]));
  const dangling = [...used].filter((id) => !declared.has(id));
  assert.deepEqual(dangling, [], `r:id sin declarar: ${dangling}`);
  assert.match(rels, /Target="https:\/\/drive\.google\.com\/file\/d\/abc123\/view" TargetMode="External"/);
});

test("deja vacío el resultado esperado cuando no se sabe", async () => {
  const sinResultado = await buildCaseDocx({ ...data, expected: "   ", images: [] });
  const xml = strFromU8(unzipSync(new Uint8Array(await sinResultado.arrayBuffer()))["word/document.xml"]);
  const after = xml.slice(xml.indexOf(">Resultado esperado</w:t>"));
  const next = after.slice(0, after.indexOf("Evidencias en Google Drive"));
  assert.ok(!next.includes("Pendiente de completar"), "no debe forzar un placeholder");
});

test("arma un nombre de archivo usable", () => {
  assert.equal(caseFileName(data), "Caso - COLUN - El parte de trabajo no actualiza el horómetro.docx");
  assert.equal(caseFileName({ domain: "", title: 'a/b:c*d?e"f<g>h|i' }), "Caso - abcdefghi.docx");
});
