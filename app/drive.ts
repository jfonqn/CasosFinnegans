/*
 * Sube el caso a Google Drive convertido a Google Doc y arma el mail para
 * soporte con el enlace en el cuerpo.
 *
 * Todo pasa en el navegador: no hay backend ni secretos del lado del servidor.
 * El token sale de Google Identity Services y se pide el scope `drive.file`,
 * que sólo alcanza a los archivos que crea esta app. Es el scope no sensible,
 * así que no requiere pasar por la verificación de apps de Google.
 *
 * El archivo NO se adjunta al mail: `mailto:` no puede adjuntar (lo prohíbe el
 * RFC 6068). Va el enlace, que además es como Finnegans pide recibir evidencia.
 */

const SCOPE = "https://www.googleapis.com/auth/drive.file";
const GIS_SRC = "https://accounts.google.com/gsi/client";
const UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3/files";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const GOOGLE_DOC_MIME = "application/vnd.google-apps.document";

export type DriveConfig = {
  clientId: string;
  folderId: string;
  to: string;
  cc: string;
};

/* Se leen en build: Next reemplaza los NEXT_PUBLIC_* por su valor literal. */
export const DRIVE_CONFIG: DriveConfig = {
  clientId: process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "",
  folderId: process.env.NEXT_PUBLIC_DRIVE_FOLDER_ID ?? "",
  to: process.env.NEXT_PUBLIC_SUPPORT_TO ?? "",
  cc: process.env.NEXT_PUBLIC_SUPPORT_CC ?? "",
};

/* Sin client ID ni carpeta el flujo no existe y la app sigue funcionando con la
   descarga del .docx. */
export function isDriveConfigured(config: DriveConfig = DRIVE_CONFIG) {
  return Boolean(config.clientId && config.folderId);
}

/* ---------- Google Identity Services ---------- */

type TokenResponse = { access_token?: string; error?: string; error_description?: string };
type TokenClient = { requestAccessToken: (overrides?: { prompt?: string }) => void };
type GoogleIdentity = {
  accounts: {
    oauth2: {
      initTokenClient(config: {
        client_id: string;
        scope: string;
        callback: (response: TokenResponse) => void;
        error_callback?: (error: { type?: string; message?: string }) => void;
      }): TokenClient;
    };
  };
};

declare global {
  interface Window {
    google?: GoogleIdentity;
  }
}

function loadIdentityServices() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  return new Promise<void>((resolve, reject) => {
    const fail = () => reject(new Error("No se pudo cargar Google Identity Services."));
    const existing = document.querySelector<HTMLScriptElement>(`script[src="${GIS_SRC}"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve());
      existing.addEventListener("error", fail);
      return;
    }
    const script = document.createElement("script");
    script.src = GIS_SRC;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = fail;
    document.head.appendChild(script);
  });
}

export async function requestDriveToken(clientId: string) {
  await loadIdentityServices();
  const identity = window.google?.accounts?.oauth2;
  if (!identity) throw new Error("Google Identity Services no quedó disponible.");
  return new Promise<string>((resolve, reject) => {
    identity
      .initTokenClient({
        client_id: clientId,
        scope: SCOPE,
        callback: (response) => {
          if (response.access_token) resolve(response.access_token);
          else reject(new Error(response.error_description || "No se autorizó el acceso a Drive."));
        },
        error_callback: (error) => reject(new Error(error?.message || "Se canceló la autorización de Drive.")),
      })
      .requestAccessToken();
  });
}

/* ---------- subida ---------- */

/* El mimeType del metadato es el destino de la conversión y el Content-Type de
   la parte binaria es el origen: así Drive convierte el .docx a Google Doc. */
export function driveMetadata(fileName: string, folderId: string) {
  return {
    name: fileName.replace(/\.docx$/i, ""),
    mimeType: GOOGLE_DOC_MIME,
    parents: [folderId],
  };
}

export function multipartBody(metadata: object, file: Blob, rawBoundary: string) {
  /* Blob normaliza su `type` a minúsculas, pero el boundary del cuerpo es
     sensible a mayúsculas: si no lo bajamos acá, el header declararía un
     separador que no existe en el cuerpo y Drive rechazaría la subida. */
  const boundary = rawBoundary.toLowerCase();
  return new Blob(
    [
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      JSON.stringify(metadata),
      `\r\n--${boundary}\r\nContent-Type: ${DOCX_MIME}\r\n\r\n`,
      file,
      `\r\n--${boundary}--\r\n`,
    ],
    { type: `multipart/related; boundary=${boundary}` },
  );
}

export async function uploadAsGoogleDoc(options: {
  blob: Blob;
  fileName: string;
  token: string;
  folderId: string;
}) {
  const boundary = `casosfinnegans-${Math.random().toString(36).slice(2)}`;
  const metadata = driveMetadata(options.fileName, options.folderId);
  /* supportsAllDrives permite que la carpeta destino esté en una unidad compartida. */
  const url = `${UPLOAD_URL}?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink`;
  const response = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${options.token}` },
    body: multipartBody(metadata, options.blob, boundary),
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(driveErrorMessage(response.status, detail));
  }
  const file: { id?: string; name?: string; webViewLink?: string } = await response.json();
  if (!file.webViewLink) throw new Error("Drive no devolvió el enlace del documento.");
  return { id: file.id ?? "", name: file.name ?? "", webViewLink: file.webViewLink };
}

function driveErrorMessage(status: number, detail: string) {
  if (status === 401 || status === 403) {
    return "Drive rechazó la subida: revisá que la cuenta tenga permiso de escritura en la carpeta destino.";
  }
  if (status === 404) return "No se encontró la carpeta de Drive configurada.";
  const parsed = safeMessage(detail);
  return parsed ? `Drive respondió ${status}: ${parsed}` : `Drive respondió ${status}.`;
}

function safeMessage(detail: string) {
  try {
    return (JSON.parse(detail) as { error?: { message?: string } })?.error?.message ?? "";
  } catch {
    return "";
  }
}

/* ---------- mail ---------- */

type MailData = {
  caseNumber: string;
  domain: string;
  client: string;
  author: string;
  priority: string;
  title: string;
};

export function caseMailSubject(data: MailData) {
  const prefix = [data.domain.trim(), data.caseNumber.trim() && `Caso ${data.caseNumber.trim()}`]
    .filter(Boolean)
    .join(" · ");
  const title = data.title.trim() || "Caso nuevo";
  return prefix ? `[${prefix}] ${title}` : title;
}

/* Cuerpo corto a propósito: el detalle vive en el documento, y una URL de
   `mailto:` larga la truncan tanto Windows como el propio Outlook. */
export function caseMailBody(data: MailData, link: string) {
  const lines = [
    "Hola,",
    "",
    "Les compartimos un caso nuevo.",
    "",
    `Dominio: ${data.domain.trim() || "-"}`,
    `Cliente: ${data.client.trim() || "-"}`,
    `Prioridad: ${data.priority}`,
    "",
    data.title.trim(),
    "",
    "El detalle completo está en el documento:",
    link,
    "",
    "Gracias.",
  ];
  const author = data.author.trim();
  if (author) lines.push(author);
  return lines.join("\n");
}

/* Se aceptan comas y punto y coma porque Outlook separa con ";" al copiar una
   lista de destinatarios, pero el mailto tiene que emitir comas. */
const addressList = (value: string) =>
  value
    .split(/[,;]/)
    .map((address) => address.trim())
    .filter(Boolean)
    .join(",");

/* Alternativa para cuando `mailto:` no abre nada: o porque el sistema no tiene
   un cliente de mail asociado, o porque el navegador bloqueó el salto a un
   protocolo externo. Esto es un https común y corriente, así que siempre abre. */
export function outlookWebUrl(options: { to: string; cc?: string; subject: string; body: string }) {
  const params = new URLSearchParams();
  params.set("to", addressList(options.to));
  const cc = addressList(options.cc ?? "");
  if (cc) params.set("cc", cc);
  params.set("subject", options.subject);
  params.set("body", options.body);
  return `https://outlook.office.com/mail/deeplink/compose?${params.toString().replace(/\+/g, "%20")}`;
}

export function mailtoUrl(options: { to: string; cc?: string; subject: string; body: string }) {
  const params = new URLSearchParams();
  const cc = addressList(options.cc ?? "");
  if (cc) params.set("cc", cc);
  params.set("subject", options.subject);
  /* Outlook corta los saltos si vienen como %0A solo. */
  params.set("body", options.body.replace(/\r?\n/g, "\r\n"));
  /* URLSearchParams codifica el espacio como "+", que en un mailto se ve
     literal; el resto de los clientes espera %20. */
  const query = params.toString().replace(/\+/g, "%20");
  /* Las direcciones no se codifican: van tal cual, separadas por coma. */
  return `mailto:${addressList(options.to)}?${query}`;
}
