/*
 * Genera el .docx de "BUGS / PROBLEMAS REPORTADOS" de Finnegans.
 *
 * No dibujamos el documento desde cero: usamos la plantilla original
 * (public/plantilla-finnegans.docx) como base y reemplazamos solo el cuerpo de
 * word/document.xml. Así se conservan intactos los estilos, la numeración, la
 * cabecera con la marca, el pie y las tipografías embebidas (Poppins y
 * Montserrat). Todo el resto del paquete OPC se copia tal cual.
 */
import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

const TEMPLATE_URL = "/plantilla-finnegans.docx";

/* Ancho útil de la caja de texto en EMU: pgSz 11906 twips menos los márgenes
   izquierdo (1440) y derecho (1257,4) que declara el sectPr de la plantilla. */
const CONTENT_WIDTH_EMU = Math.round((11906 - 1440 - 1257.4) * 635);
const EMU_PER_PX = 9525;

const MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
};

export const ACCEPTED_IMAGE_TYPES = ".png,.jpg,.jpeg,.gif";

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
  images: File[];
};

/* ---------- helpers de XML ---------- */

function esc(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* La plantilla fija Poppins y el gris 434343 en cada run, no sólo en el estilo.
   Replicamos ese rPr para que lo generado sea indistinguible de lo escrito a mano. */
function rpr(options?: { bold?: boolean; link?: boolean }) {
  let out = '<w:rPr><w:rFonts w:ascii="Poppins" w:cs="Poppins" w:eastAsia="Poppins" w:hAnsi="Poppins"/>';
  if (options?.bold) out += "<w:b/>";
  out += `<w:color w:val="${options?.link ? "1155CC" : "434343"}"/>`;
  if (options?.link) out += '<w:u w:val="single"/>';
  return out + '<w:rtl w:val="0"/></w:rPr>';
}

function run(text: string, options?: { bold?: boolean; link?: boolean }) {
  return `<w:r>${rpr(options)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function listProps() {
  return '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/>';
}

function para(runs: string, style?: string, numbered = false) {
  let props = "<w:pPr>";
  if (style) props += `<w:pStyle w:val="${style}"/>`;
  if (numbered) props += listProps();
  props += rpr() + "</w:pPr>";
  return `<w:p>${props}${runs}</w:p>`;
}

const title = (text: string) => para(run(text), "Title");
const subtitle = (text: string) => para(run(text), "Subtitle");
const heading = (text: string) => para(run(text), "Heading4");
const normal = (text: string) => para(text ? run(text) : "");
const bullet = (label: string, value: string) => para(run(label, { bold: true }) + run(value), undefined, true);
const labeled = (label: string, value: string) => para(run(label, { bold: true }) + run(value));

/* Un párrafo por línea: el .docx tiene que respetar los saltos que escribió el
   usuario, sobre todo en los pasos del caso de uso. */
function block(text: string, fallback = "") {
  const lines = text.split("\n").map((line) => line.replace(/[^\S\n]+/g, " ").trim());
  while (lines.length && !lines[lines.length - 1]) lines.pop();
  if (!lines.some(Boolean)) return normal(fallback);
  return lines.map((line) => normal(line)).join("");
}

function hyperlink(relId: string, url: string) {
  return `<w:p><w:pPr>${listProps()}${rpr()}</w:pPr><w:hyperlink r:id="${relId}">${run(url, { link: true })}</w:hyperlink></w:p>`;
}

function picture(relId: string, id: number, name: string, widthEmu: number, heightEmu: number) {
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

async function prepareImages(files: File[], startId: number): Promise<PreparedImage[]> {
  const prepared: PreparedImage[] = [];
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    const ext = (file.name.split(".").pop() ?? "png").toLowerCase();
    if (!MIME_BY_EXT[ext]) continue;
    const { width, height } = await measure(file);
    /* Escalamos sólo hacia abajo: una captura chica no se estira. */
    const scale = Math.min(1, CONTENT_WIDTH_EMU / (width * EMU_PER_PX));
    prepared.push({
      relId: `rId${startId + index}`,
      part: `word/media/evidencia${index + 1}.${ext}`,
      bytes: new Uint8Array(await file.arrayBuffer()),
      ext,
      name: file.name,
      widthEmu: Math.round(width * EMU_PER_PX * scale),
      heightEmu: Math.round(height * EMU_PER_PX * scale),
    });
  }
  return prepared;
}

/* ---------- cuerpo del documento ---------- */

const PENDING = "Pendiente de completar";

function buildBody(data: CaseDoc, images: PreparedImage[], linkRels: string[]) {
  const parts: string[] = [];

  parts.push(title("BUGS / PROBLEMAS REPORTADOS"));
  parts.push(subtitle(""));
  parts.push(title(`N° de caso: ${data.caseNumber.trim() || "A completar por Finnegans"}`));
  parts.push(subtitle(`Fecha de Ingreso: ${data.date}`));
  parts.push(subtitle(`Dominio: ${data.domain.trim() || PENDING}`));
  parts.push(subtitle(`Cliente: ${data.client.trim() || PENDING}`));
  parts.push(subtitle(`Redactor: ${data.author.trim() || PENDING}`));
  parts.push(normal(`Servidor propio / dedicado: ${data.dedicatedServer.toUpperCase()}`));
  parts.push(normal(`Prioridad del Caso: ${data.priority}`));
  if (data.company.trim()) parts.push(normal(`Empresa / sucursal: ${data.company.trim()}`));
  parts.push(normal(""));

  parts.push(heading("Título del caso"));
  parts.push(block(data.title, PENDING));

  parts.push(heading("Pegar respuesta y promt de Finni"));
  parts.push(block(data.finni));

  parts.push(heading("Problema"));
  parts.push(bullet("¿Qué no funciona? ", data.whatFails.trim() || PENDING));
  parts.push(bullet("¿Qué es lo que pasa? ", data.whatHappens.trim() || PENDING));
  parts.push(bullet("¿Es la primera vez que hace la transacción? ", data.firstTime));
  parts.push(bullet("¿Es una operación estándar? ", data.standardOperation));

  parts.push(heading("Situación actual"));
  parts.push(block(data.impact, PENDING));
  if (data.workaround.trim()) parts.push(labeled("Contingencia manual: ", data.workaround.trim()));

  parts.push(heading("Caso de uso"));
  parts.push(block(data.steps, PENDING));
  if (data.reportFormat.trim()) parts.push(labeled("Formato de grilla / informe: ", data.reportFormat.trim()));
  images.forEach((image, index) =>
    parts.push(picture(image.relId, 1000 + index, image.name, image.widthEmu, image.heightEmu)),
  );

  parts.push(heading("Pruebas o soluciones intentadas"));
  parts.push(block(data.attempts));

  /* La plantilla aclara que si no se sabe qué debería dar, esta sección se deja
     vacía. Por eso acá no forzamos el "Pendiente de completar". */
  parts.push(heading("Resultado esperado"));
  parts.push(block(data.expected));

  parts.push(heading("Evidencias en Google Drive"));
  if (data.driveLinks.length) {
    data.driveLinks.forEach((link, index) => parts.push(hyperlink(linkRels[index], link)));
  } else {
    parts.push(normal("Pendiente: cargar evidencia en Google Drive y agregar el enlace"));
  }
  parts.push(normal(`Acceso verificado para soporte: ${data.accessConfirmed ? "Sí" : "Pendiente de confirmar"}`));

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
