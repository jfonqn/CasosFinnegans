/*
 * Genera el .docx de "BUGS / PROBLEMAS REPORTADOS" de Finnegans.
 *
 * No dibujamos el documento desde cero: usamos la plantilla original
 * (public/plantilla-finnegans.docx) como base y reemplazamos solo el cuerpo de
 * word/document.xml. Así se conservan intactos los estilos, la numeración, la
 * cabecera con la marca, el pie y las tipografías embebidas (Poppins y
 * Montserrat). Todo el resto del paquete OPC se copia tal cual.
 *
 * La disposición del cuerpo sigue un caso real escrito a mano: preguntas y
 * respuestas como lista de dos niveles, párrafos justificados, y las capturas
 * intercaladas en la sección a la que pertenecen.
 */
import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

const TEMPLATE_URL = "/plantilla-finnegans.docx";

/* Ancho útil de la caja de texto en EMU: pgSz 11906 twips menos los márgenes
   izquierdo (1440) y derecho (1257,4) que declara el sectPr de la plantilla.
   Coincide con las 6,4 pulgadas a las que el autor del caso real llevó cada
   captura. */
const CONTENT_WIDTH_EMU = Math.round((11906 - 1440 - 1257.4) * 635);
const EMU_PER_PX = 9525;

/* Los estilos de la plantilla traen cuerpos enormes (Title = 48pt) y el documento
   los pisa run por run. Sin replicar ese override los títulos tapan la página. */
const SIZE_DOC_TITLE = 34; // 17pt — "BUGS / PROBLEMAS REPORTADOS"
const SIZE_CASE_TITLE = 60; // 30pt — "N° de caso"
const SIZE_HEADER_LINE = 36; // 18pt — las líneas sueltas del encabezado

const MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
};

/* Campos de texto que aceptan capturas pegadas. El orden fija el nombre de las
   partes dentro del paquete, así que es estable entre generaciones. */
export const IMAGE_SLOTS = [
  "finni",
  "whatFails",
  "whatHappens",
  "workaround",
  "impact",
  "steps",
  "attempts",
  "expected",
] as const;

export type ImageSlot = (typeof IMAGE_SLOTS)[number];
export type CaseImages = Partial<Record<ImageSlot, File[]>>;

export function isSupportedImage(file: File) {
  return Object.values(MIME_BY_EXT).includes(file.type);
}

export type CaseDoc = {
  caseNumber: string;
  date: string;
  domain: string;
  client: string;
  author: string;
  company: string;
  dedicatedServer: string;
  priority: string;
  title: string;
  finni: string;
  whatFails: string;
  whatHappens: string;
  firstTime: string;
  standardOperation: string;
  impact: string;
  workaround: string;
  steps: string;
  reportFormat: string;
  attempts: string;
  expected: string;
  driveLinks: string[];
  accessConfirmed: boolean;
  images: CaseImages;
};

/* ---------- helpers de XML ---------- */

function esc(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

type RunOptions = { bold?: boolean; link?: boolean; size?: number };

/* La plantilla fija Poppins y el gris 434343 en cada run, no sólo en el estilo.
   Replicamos ese rPr para que lo generado sea indistinguible de lo escrito a mano.
   El orden de los hijos es el que exige el esquema CT_RPr. */
function rpr(options?: RunOptions) {
  let out = '<w:rPr><w:rFonts w:ascii="Poppins" w:cs="Poppins" w:eastAsia="Poppins" w:hAnsi="Poppins"/>';
  if (options?.bold) out += "<w:b/>";
  out += `<w:color w:val="${options?.link ? "1155CC" : "434343"}"/>`;
  if (options?.size) out += `<w:sz w:val="${options.size}"/><w:szCs w:val="${options.size}"/>`;
  if (options?.link) out += '<w:u w:val="single"/>';
  return out + '<w:rtl w:val="0"/></w:rPr>';
}

function run(text: string, options?: RunOptions) {
  return `<w:r>${rpr(options)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

type ParaOptions = { style?: string; level?: number; size?: number; justify?: boolean };

/* El orden de los hijos del pPr lo fija el esquema CT_PPrBase: pStyle, numPr,
   ind, jc, y recién al final el rPr de la marca de párrafo. */
function paraProps(options?: ParaOptions) {
  let props = "<w:pPr>";
  if (options?.style) props += `<w:pStyle w:val="${options.style}"/>`;
  if (options?.level !== undefined) {
    props += `<w:numPr><w:ilvl w:val="${options.level}"/><w:numId w:val="1"/></w:numPr>`;
    props += `<w:ind w:left="${720 * (options.level + 1)}" w:hanging="360"/>`;
  }
  if (options?.justify) props += '<w:jc w:val="both"/>';
  /* El rPr del pPr aplica a la marca de párrafo: si no lleva el mismo cuerpo,
     el interlineado queda calculado sobre el tamaño del estilo. */
  return props + rpr({ size: options?.size }) + "</w:pPr>";
}

function para(runs: string, options?: ParaOptions) {
  return `<w:p>${paraProps(options)}${runs}</w:p>`;
}

const title = (text: string, size: number) => para(run(text, { size }), { style: "Title", size });
const subtitle = (text: string) => para(run(text), { style: "Subtitle" });
const heading = (text: string) => para(run(text), { style: "Heading4" });
const normal = (text: string, size?: number) => para(text ? run(text, { size }) : "", { size });
const body = (text: string) => para(run(text), { justify: true });
const question = (text: string) => para(run(text), { level: 0, justify: true });
const answer = (text: string) => para(run(text), { level: 1, justify: true });
const labeled = (label: string, value: string) => para(run(label, { bold: true }) + run(value), { justify: true });

/* Un párrafo por línea: el .docx tiene que respetar los saltos que escribió el
   usuario, sobre todo en los pasos del caso de uso. */
function splitLines(text: string) {
  const lines = text.split("\n").map((line) => line.replace(/[^\S\n]+/g, " ").trim());
  while (lines.length && !lines[lines.length - 1]) lines.pop();
  return lines.some(Boolean) ? lines : null;
}

function block(text: string, fallback = "") {
  const lines = splitLines(text);
  if (!lines) return fallback ? body(fallback) : normal("");
  return lines.map((line) => (line ? body(line) : normal(""))).join("");
}

/* Igual que block() pero como respuesta anidada de la lista de preguntas. */
function answerBlock(text: string, fallback: string) {
  const lines = splitLines(text);
  if (!lines) return answer(fallback);
  return lines.filter(Boolean).map(answer).join("");
}

function hyperlink(relId: string, url: string) {
  return `<w:p>${paraProps({ level: 0 })}<w:hyperlink r:id="${relId}">${run(url, { link: true })}</w:hyperlink></w:p>`;
}

function picture(image: PreparedImage, id: number) {
  const { relId, name, widthEmu, heightEmu } = image;
  return (
    `<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:lineRule="auto"/></w:pPr><w:r>` +
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent b="0" l="0" r="0" t="0"/>` +
    `<wp:docPr id="${id}" name="${esc(name)}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
  );
}

/* ---------- imágenes ---------- */

type PreparedImage = {
  slot: ImageSlot;
  relId: string;
  part: string;
  /* fflate tipa las partes del zip sobre ArrayBuffer y no sobre ArrayBufferLike. */
  bytes: Uint8Array<ArrayBuffer>;
  ext: string;
  name: string;
  widthEmu: number;
  heightEmu: number;
};

async function measure(file: File) {
  try {
    const bitmap = await createImageBitmap(file);
    const size = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return size;
  } catch {
    /* Formato que createImageBitmap no acepta: caemos al <img> de toda la vida. */
    return new Promise<{ width: number; height: number }>((resolve, reject) => {
      const url = URL.createObjectURL(file);
      const image = new Image();
      image.onload = () => {
        URL.revokeObjectURL(url);
        resolve({ width: image.naturalWidth, height: image.naturalHeight });
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error(`No se pudo leer la imagen ${file.name}`));
      };
      image.src = url;
    });
  }
}

function extensionOf(file: File) {
  const byType = Object.entries(MIME_BY_EXT).find(([, mime]) => mime === file.type)?.[0];
  if (byType) return byType;
  const byName = (file.name.split(".").pop() ?? "").toLowerCase();
  return MIME_BY_EXT[byName] ? byName : null;
}

async function prepareImages(images: CaseImages, startId: number): Promise<PreparedImage[]> {
  const prepared: PreparedImage[] = [];
  for (const slot of IMAGE_SLOTS) {
    for (const file of images[slot] ?? []) {
      const ext = extensionOf(file);
      if (!ext) continue;
      const { width, height } = await measure(file);
      /* Escalamos sólo hacia abajo: una captura chica no se estira. */
      const scale = Math.min(1, CONTENT_WIDTH_EMU / (width * EMU_PER_PX));
      const index = prepared.length + 1;
      prepared.push({
        slot,
        relId: `rId${startId + prepared.length}`,
        part: `word/media/evidencia${index}.${ext}`,
        bytes: new Uint8Array(await file.arrayBuffer()),
        ext,
        name: file.name || `Captura ${index}`,
        widthEmu: Math.round(width * EMU_PER_PX * scale),
        heightEmu: Math.round(height * EMU_PER_PX * scale),
      });
    }
  }
  return prepared;
}

/* ---------- cuerpo del documento ---------- */

const PENDING = "Pendiente de completar";

function buildBody(data: CaseDoc, images: PreparedImage[], linkRels: string[]) {
  const parts: string[] = [];
  let pictureId = 1000;
  /* Las capturas se emiten dentro de la sección en la que fueron pegadas. */
  const shots = (slot: ImageSlot) =>
    images
      .filter((image) => image.slot === slot)
      .map((image) => picture(image, (pictureId += 1)))
      .join("");
  const hasShots = (slot: ImageSlot) => images.some((image) => image.slot === slot);

  parts.push(title("BUGS / PROBLEMAS REPORTADOS", SIZE_DOC_TITLE));
  parts.push(subtitle(""));
  /* Sin número, dejamos el rótulo solo, como viene en la plantilla: es Finnegans
     quien lo completa. Además evita que el título de 30pt se parta en tres líneas. */
  const caseNumber = data.caseNumber.trim();
  parts.push(title(caseNumber ? `N° de caso: ${caseNumber}` : "N° de caso", SIZE_CASE_TITLE));
  parts.push(subtitle(`Fecha de Ingreso: ${data.date}`));
  parts.push(subtitle(`Dominio: ${data.domain.trim() || PENDING}`));
  parts.push(subtitle(`Cliente: ${data.client.trim() || PENDING}`));
  parts.push(subtitle(`Redactor: ${data.author.trim() || PENDING}`));
  parts.push(normal(`Servidor propio / dedicado: ${data.dedicatedServer.toUpperCase()}`, SIZE_HEADER_LINE));
  parts.push(normal(`Prioridad del Caso: ${data.priority}`, SIZE_HEADER_LINE));
  if (data.company.trim()) parts.push(normal(`Empresa / sucursal: ${data.company.trim()}`, SIZE_HEADER_LINE));
  parts.push(normal(""));

  parts.push(heading("Título del caso"));
  parts.push(block(data.title, PENDING));

  /* Secciones opcionales: si no hay ni texto ni capturas, no se escribe el título. */
  if (data.finni.trim() || hasShots("finni")) {
    parts.push(heading("Pegar respuesta y promt de Finni"));
    parts.push(block(data.finni));
    parts.push(shots("finni"));
  }

  parts.push(heading("Problema"));
  parts.push(body("Debe responder las siguientes preguntas"));
  parts.push(question("¿Qué no funciona?"));
  parts.push(answerBlock(data.whatFails, PENDING));
  parts.push(shots("whatFails"));
  parts.push(question("¿Qué es lo que pasa?"));
  parts.push(answerBlock(data.whatHappens, PENDING));
  parts.push(shots("whatHappens"));
  parts.push(question("¿Es la primera vez que hace la transacción?"));
  parts.push(answer(data.firstTime));
  parts.push(question("¿Es una operación estándar?"));
  parts.push(answer(data.standardOperation));
  parts.push(question("¿Tiene alguna contingencia manual?"));
  parts.push(answerBlock(data.workaround, "No"));
  parts.push(shots("workaround"));

  parts.push(heading("Situación actual"));
  parts.push(block(data.impact, PENDING));
  parts.push(shots("impact"));

  parts.push(heading("Caso de uso"));
  parts.push(block(data.steps, PENDING));
  if (data.reportFormat.trim()) parts.push(labeled("Formato de grilla / informe: ", data.reportFormat.trim()));
  parts.push(shots("steps"));

  if (data.attempts.trim() || hasShots("attempts")) {
    parts.push(heading("Pruebas o soluciones intentadas"));
    parts.push(block(data.attempts));
    parts.push(shots("attempts"));
  }

  /* La plantilla aclara que si no se sabe qué debería dar, esta sección se deja
     vacía. Por eso acá no forzamos el "Pendiente de completar". */
  parts.push(heading("Resultado esperado"));
  parts.push(block(data.expected));
  parts.push(shots("expected"));

  /* Los enlaces de Drive son opcionales: si no hay ninguno, la sección no se
     escribe. Un "pendiente" en un documento que ya se envía es sólo ruido. */
  if (data.driveLinks.length) {
    parts.push(heading("Evidencias en Google Drive"));
    data.driveLinks.forEach((link, index) => parts.push(hyperlink(linkRels[index], link)));
    parts.push(normal(`Acceso verificado para soporte: ${data.accessConfirmed ? "Sí" : "Pendiente de confirmar"}`));
  }

  return parts.join("");
}

/* ---------- armado del paquete ---------- */

export async function buildCaseDocx(data: CaseDoc): Promise<Blob> {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) throw new Error(`No se pudo cargar la plantilla (${response.status})`);
  const files = unzipSync(new Uint8Array(await response.arrayBuffer()));

  const images = await prepareImages(data.images, 1001);
  const linkRels = data.driveLinks.map((_, index) => `rId${2001 + index}`);

  /* Relaciones nuevas: una por imagen embebida y una por enlace de Drive. */
  const relsPath = "word/_rels/document.xml.rels";
  const added =
    images
      .map(
        (image) =>
          `<Relationship Id="${image.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${image.part.replace("word/", "")}"/>`,
      )
      .join("") +
    data.driveLinks
      .map(
        (link, index) =>
          `<Relationship Id="${linkRels[index]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(link)}" TargetMode="External"/>`,
      )
      .join("");
  files[relsPath] = strToU8(strFromU8(files[relsPath]).replace("</Relationships>", `${added}</Relationships>`));

  /* La plantilla ya declara png; jpeg y gif hay que agregarlos si aparecen. */
  const typesPath = "[Content_Types].xml";
  let types = strFromU8(files[typesPath]);
  for (const ext of new Set(images.map((image) => image.ext))) {
    if (!types.includes(`Extension="${ext}"`)) {
      types = types.replace("<Override", `<Default ContentType="${MIME_BY_EXT[ext]}" Extension="${ext}"/><Override`);
    }
  }
  files[typesPath] = strToU8(types);

  for (const image of images) files[image.part] = image.bytes;

  /* Conservamos la raíz de <w:document> (declara los namespaces que usan las
     imágenes) y el <w:sectPr> final, que es el que engancha cabecera y pie. */
  const documentPath = "word/document.xml";
  const xml = strFromU8(files[documentPath]);
  const bodyStart = xml.indexOf("<w:body>");
  const sectStart = xml.lastIndexOf("<w:sectPr");
  if (bodyStart < 0 || sectStart < 0) throw new Error("La plantilla no tiene la estructura esperada");
  const bodyOpen = bodyStart + "<w:body>".length;
  files[documentPath] = strToU8(
    xml.slice(0, bodyOpen) + buildBody(data, images, linkRels) + xml.slice(sectStart),
  );

  return new Blob([zipSync(files) as unknown as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export function caseFileName(data: Pick<CaseDoc, "domain" | "title">) {
  const joined = ["Caso", data.domain.trim(), data.title.trim()].filter(Boolean).join(" - ");
  const safe = joined
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
  return `${safe || "Caso"}.docx`;
}
