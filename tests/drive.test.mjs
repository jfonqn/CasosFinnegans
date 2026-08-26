/*
 * Verifica el armado de la subida a Drive y del mail.
 *
 * El ida y vuelta real con Google no se puede probar acá —hace falta un client
 * ID y una cuenta—, así que se stubbea fetch y se comprueba exactamente qué
 * pide la app: el multipart de conversión, los headers y la URL del mailto.
 *
 *   node --test tests/drive.test.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";
import { pathToFileURL } from "node:url";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const drive = await import(pathToFileURL(path.join(ROOT, "app/drive.ts")).href);
const {
  driveMetadata,
  multipartBody,
  uploadAsGoogleDoc,
  caseMailSubject,
  caseMailBody,
  mailtoUrl,
  outlookWebUrl,
  isDriveConfigured,
} = drive;

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const GOOGLE_DOC_MIME = "application/vnd.google-apps.document";
const ZIP_SIGNATURE = "PK";

const caso = {
  caseNumber: "",
  domain: "INMIX",
  client: "FISTERRA",
  author: "Alejandro Fernández",
  priority: "Alta",
  title: "El nuevo motor de retenciones no se aplica",
};

test("pide la conversión a Google Doc, no un .docx suelto", () => {
  const metadata = driveMetadata("Caso - INMIX - Retenciones.docx", "carpeta123");
  /* El mimeType del metadato es el DESTINO de la conversión. */
  assert.equal(metadata.mimeType, GOOGLE_DOC_MIME);
  assert.deepEqual(metadata.parents, ["carpeta123"]);
  /* En Drive un Google Doc no lleva extensión. */
  assert.equal(metadata.name, "Caso - INMIX - Retenciones");
});

test("arma el multipart con el origen y el destino correctos", async () => {
  const file = new Blob([new Uint8Array([80, 75, 3, 4])], { type: DOCX_MIME });
  const body = multipartBody({ name: "x", mimeType: GOOGLE_DOC_MIME }, file, "limite");
  const text = await body.text();
  assert.equal(body.type, "multipart/related; boundary=limite");
  assert.ok(text.startsWith("--limite\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"));
  /* La parte binaria declara el formato de ORIGEN. */
  assert.ok(text.includes(`\r\n--limite\r\nContent-Type: ${DOCX_MIME}\r\n\r\n`));
  assert.ok(text.trimEnd().endsWith("--limite--"));
  assert.ok(text.includes(ZIP_SIGNATURE), "el .docx tiene que viajar en el cuerpo");
});

test("el boundary declarado siempre coincide con el del cuerpo", async () => {
  /* Blob baja su `type` a minúsculas, pero el boundary del cuerpo es sensible a
     mayúsculas: si divergen, Drive no encuentra las partes y rechaza la subida. */
  const body = multipartBody({ name: "x" }, new Blob(["x"]), "Limite-CON-Mayusculas");
  const declared = body.type.match(/boundary=(.+)$/)[1];
  const text = await body.text();
  assert.ok(text.startsWith(`--${declared}\r\n`), `el cuerpo no arranca con ${declared}`);
  assert.ok(text.trimEnd().endsWith(`--${declared}--`));
});

test("sube con el token y devuelve el enlace", async () => {
  let request = null;
  globalThis.fetch = async (url, init) => {
    request = { url: String(url), init };
    return {
      ok: true,
      status: 200,
      json: async () => ({ id: "doc1", name: "Caso", webViewLink: "https://docs.google.com/document/d/doc1/edit" }),
    };
  };
  const file = await uploadAsGoogleDoc({
    blob: new Blob(["x"], { type: DOCX_MIME }),
    fileName: "Caso - INMIX.docx",
    token: "token-abc",
    folderId: "carpeta123",
  });
  assert.equal(file.webViewLink, "https://docs.google.com/document/d/doc1/edit");
  assert.equal(request.init.headers.Authorization, "Bearer token-abc");
  assert.ok(request.url.includes("uploadType=multipart"));
  assert.ok(request.url.includes("fields=id,name,webViewLink"), "sin fields Drive no devuelve el enlace");
  /* Sin esto la carpeta destino no puede estar en una unidad compartida. */
  assert.ok(request.url.includes("supportsAllDrives=true"));
});

test("explica los errores de Drive en lugar de filtrar el status pelado", async () => {
  const subir = () => uploadAsGoogleDoc({ blob: new Blob(["x"]), fileName: "a.docx", token: "t", folderId: "f" });

  globalThis.fetch = async () => ({ ok: false, status: 403, text: async () => "{}" });
  await assert.rejects(subir, /permiso de escritura/);

  globalThis.fetch = async () => ({ ok: false, status: 404, text: async () => "{}" });
  await assert.rejects(subir, /no se encontró la carpeta/i);

  globalThis.fetch = async () => ({
    ok: false,
    status: 500,
    text: async () => JSON.stringify({ error: { message: "Backend Error" } }),
  });
  await assert.rejects(subir, /Backend Error/);
});

test("falla claro si Drive no devuelve el enlace", async () => {
  globalThis.fetch = async () => ({ ok: true, status: 200, json: async () => ({ id: "doc1" }) });
  await assert.rejects(
    () => uploadAsGoogleDoc({ blob: new Blob(["x"]), fileName: "a.docx", token: "t", folderId: "f" }),
    /no devolvió el enlace/,
  );
});

test("arma un asunto que se puede escanear en la bandeja", () => {
  assert.equal(caseMailSubject(caso), "[INMIX] El nuevo motor de retenciones no se aplica");
  assert.equal(
    caseMailSubject({ ...caso, caseNumber: "48210" }),
    "[INMIX · Caso 48210] El nuevo motor de retenciones no se aplica",
  );
  assert.equal(caseMailSubject({ ...caso, domain: "", title: "" }), "Caso nuevo");
});

test("el cuerpo lleva el enlace, no el caso entero", () => {
  const link = "https://docs.google.com/document/d/doc1/edit";
  const body = caseMailBody(caso, link);
  assert.ok(body.includes(link));
  assert.ok(body.includes("Dominio: INMIX"));
  assert.ok(body.includes("Alejandro Fernández"));
  /* La URL de mailto la truncan Windows y Outlook, así que el cuerpo es corto. */
  assert.ok(body.length < 600, `el cuerpo creció a ${body.length} caracteres`);
});

test("codifica el mailto como lo espera Outlook", () => {
  const url = mailtoUrl({
    to: " soporte@finnegans.com , casos@finnegans.com ",
    cc: "",
    subject: "[INMIX] Retenciones & percepciones",
    body: "Hola,\nsegunda línea",
  });
  /* Las direcciones van sin codificar y sin espacios sueltos. */
  assert.ok(url.startsWith("mailto:soporte@finnegans.com,casos@finnegans.com?"));
  assert.ok(!url.includes("cc="), "sin cc no debe mandar el parámetro vacío");
  assert.ok(url.includes("%20"), "los espacios van como %20, no como +");
  assert.ok(!url.includes("+"), "un + literal en el asunto se ve como signo más");
  assert.ok(url.includes("%0D%0A"), "los saltos necesitan CRLF o Outlook los pierde");
  assert.ok(url.includes("%26"), "el & del asunto tiene que ir escapado");

  /* Outlook separa con ";" al copiar una lista; el mailto necesita comas. */
  const puntoYComa = mailtoUrl({
    to: "mgalarza@ejemplo.com; soporte@ejemplo.com ; casos@ejemplo.com",
    subject: "s",
    body: "b",
  });
  assert.ok(puntoYComa.startsWith("mailto:mgalarza@ejemplo.com,soporte@ejemplo.com,casos@ejemplo.com?"));
  const conCopia = mailtoUrl({ to: "a@b.com", cc: "jefe@fisterra.com", subject: "s", body: "b" });
  assert.ok(conCopia.includes("cc=jefe%40fisterra.com"));
});

test("el flujo se apaga solo si falta configuración", () => {
  assert.equal(isDriveConfigured({ clientId: "", folderId: "", to: "", cc: "" }), false);
  assert.equal(isDriveConfigured({ clientId: "abc", folderId: "", to: "", cc: "" }), false);
  assert.equal(isDriveConfigured({ clientId: "abc", folderId: "xyz", to: "", cc: "" }), true);
});

test("ofrece Outlook Web como salida cuando mailto no abre nada", () => {
  /* mailto depende de que el equipo tenga un cliente asociado y de que el
     navegador permita el salto; este es un https y siempre abre. */
  const url = outlookWebUrl({
    to: "a@b.com; c@d.com",
    cc: "jefe@fisterra.com",
    subject: "[INMIX] Retenciones & percepciones",
    body: ["Hola,", "segunda línea"].join(String.fromCharCode(10)),
  });
  assert.ok(url.startsWith("https://outlook.office.com/mail/deeplink/compose?"));
  const params = new URL(url).searchParams;
  assert.equal(params.get("to"), "a@b.com,c@d.com");
  assert.equal(params.get("cc"), "jefe@fisterra.com");
  assert.equal(params.get("subject"), "[INMIX] Retenciones & percepciones");
  assert.ok(params.get("body").includes("segunda línea"));
  assert.ok(!url.includes("+"), "los espacios van como %20");
});
