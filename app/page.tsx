"use client";

import { useEffect, useMemo, useState } from "react";
import { Lockup, PairedHeading, Panel, SectionPill } from "./fisterra";
import { buildCaseDocx, caseFileName, isSupportedImage, type CaseDoc, type CaseImages, type ImageSlot } from "./docx";
import {
  DRIVE_CONFIG,
  caseMailBody,
  caseMailSubject,
  isDriveConfigured,
  mailtoUrl,
  outlookWebUrl,
  requestDriveToken,
  uploadAsGoogleDoc,
} from "./drive";

type CaseData = {
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
  evidence: string;
  accessConfirmed: boolean;
};

const today = new Intl.DateTimeFormat("es-AR").format(new Date());

const initialData: CaseData = {
  caseNumber: "",
  date: today,
  domain: "",
  client: "",
  author: "",
  company: "",
  dedicatedServer: "No",
  priority: "Media",
  title: "",
  finni: "",
  whatFails: "",
  whatHappens: "",
  firstTime: "No",
  standardOperation: "Sí",
  impact: "",
  workaround: "",
  steps: "",
  reportFormat: "",
  attempts: "",
  expected: "",
  evidence: "",
  accessConfirmed: false,
};

/* Props de marca del documento de diseño (sección "Marca"). */
type CtaTone = "rojo" | "navy";

const CTA_BG: Record<CtaTone, string> = {
  rojo: "#F52125",
  navy: "linear-gradient(100deg,#1C4257 0%,#0A2F43 100%)",
};

const CTA_TONE: CtaTone = "rojo";
const CLAIM = "Brindamos soluciones estratégicas en contabilidad e impuestos.";

const ctaBg = CTA_BG[CTA_TONE];

const drivePattern = /^https?:\/\/(drive|docs)\.google\.com\//i;

/* Colapsa espacios horizontales pero conserva los saltos de línea: los pasos del
   caso de uso pierden todo su sentido si se aplanan en un párrafo. */
function tidy(text: string) {
  return text
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/g, " ").trim())
    .join("\n")
    .trim();
}

function clean(text: string) {
  const value = tidy(text);
  if (!value) return "Pendiente de completar";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

/* Igual que clean(), pero para campos que la plantilla permite dejar vacíos. */
function cleanOptional(text: string) {
  const value = tidy(text);
  if (!value) return "";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

/* El portapapeles expone la captura en items; files queda vacío en algunos
   navegadores, así que lo usamos sólo como respaldo. */
function imagesFrom(transfer: DataTransfer | null) {
  if (!transfer) return [];
  const fromItems = Array.from(transfer.items)
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter((file): file is File => file !== null);
  const files = fromItems.length ? fromItems : Array.from(transfer.files);
  return files.filter(isSupportedImage);
}

export default function Home() {
  const [data, setData] = useState(initialData);
  const [shots, setShots] = useState<CaseImages>({});
  const [copied, setCopied] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [building, setBuilding] = useState(false);
  const [docxError, setDocxError] = useState<string | null>(null);
  const [sending, setSending] = useState(false);
  const [driveLink, setDriveLink] = useState<string | null>(null);
  const [driveError, setDriveError] = useState<string | null>(null);
  const [linkCopied, setLinkCopied] = useState(false);

  const update = <K extends keyof CaseData>(key: K, value: CaseData[K]) =>
    setData((current) => ({ ...current, [key]: value }));

  const addShots = (slot: ImageSlot, files: File[]) =>
    setShots((current) => ({ ...current, [slot]: [...(current[slot] ?? []), ...files] }));

  const removeShot = (slot: ImageSlot, index: number) =>
    setShots((current) => ({ ...current, [slot]: (current[slot] ?? []).filter((_, i) => i !== index) }));

  const shotCount = Object.values(shots).reduce((total, files) => total + files.length, 0);

  const driveLinks = data.evidence
    .split(/\n|,/)
    .map((item) => item.trim())
    .filter(Boolean);
  const invalidLinks = driveLinks.filter((link) => !drivePattern.test(link));

  const required = [
    data.domain,
    data.client,
    data.author,
    data.title,
    data.whatFails,
    data.whatHappens,
    data.impact,
    data.steps,
    data.expected,
  ];
  const completed = required.filter((value) => value.trim()).length;
  const progress = Math.round((completed / required.length) * 100);

  const output = useMemo(() => {
    /* Sin enlaces no se escribe la sección: las evidencias en Drive son opcionales. */
    const evidenceBlock = driveLinks.length
      ? `\n\nEVIDENCIAS EN GOOGLE DRIVE\n${driveLinks
          .map((link, index) => `- Evidencia ${index + 1}: ${link}`)
          .join("\n")}\nAcceso verificado para soporte: ${data.accessConfirmed ? "Sí" : "Pendiente de confirmar"}`
      : "";
    const optional = (label: string, value: string) => {
      const text = cleanOptional(value);
      return text ? `\n\n${label}\n${text}` : "";
    };

    return `BUGS / PROBLEMAS REPORTADOS

N° de caso: ${data.caseNumber || "A completar por Finnegans"}
Fecha de Ingreso: ${data.date}
Dominio: ${clean(data.domain)}
Cliente: ${clean(data.client)}
Redactor: ${clean(data.author)}
Servidor propio / dedicado: ${data.dedicatedServer.toUpperCase()}
Prioridad del Caso: ${data.priority}${data.company.trim() ? `\nEmpresa / sucursal: ${tidy(data.company)}` : ""}

TÍTULO DEL CASO
${clean(data.title)}${optional("PEGAR RESPUESTA Y PROMT DE FINNI", data.finni)}

PROBLEMA
Debe responder las siguientes preguntas

- ¿Qué no funciona?
  ${clean(data.whatFails)}
- ¿Qué es lo que pasa?
  ${clean(data.whatHappens)}
- ¿Es la primera vez que hace la transacción?
  ${data.firstTime}
- ¿Es una operación estándar?
  ${data.standardOperation}
- ¿Tiene alguna contingencia manual?
  ${cleanOptional(data.workaround) || "No"}

SITUACIÓN ACTUAL
${clean(data.impact)}

CASO DE USO
${clean(data.steps)}${data.reportFormat.trim() ? `\nFormato de grilla / informe: ${tidy(data.reportFormat)}` : ""}${
      shotCount ? `\n${shotCount} captura(s) se incrustan en el .docx.` : ""
    }${optional("PRUEBAS O SOLUCIONES INTENTADAS", data.attempts)}${optional("RESULTADO ESPERADO", data.expected)}${evidenceBlock}`;
  }, [data, driveLinks, shotCount]);

  /* Las evidencias son opcionales: sólo frenan el caso si los enlaces están mal
     escritos, porque ahí soporte se va a encontrar con un link que no abre. */
  const linksOk = invalidLinks.length === 0;
  const hasEvidence = driveLinks.length > 0 || shotCount > 0;
  const ready = progress === 100 && linksOk;
  const readiness = !linksOk
    ? "Revisá los enlaces de Drive"
    : ready
      ? "Caso listo para enviar"
      : "Completá los datos obligatorios";

  async function copyCase() {
    try {
      await navigator.clipboard.writeText(output);
    } catch {
      /* El portapapeles puede estar bloqueado por permisos; el texto sigue visible. */
    }
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1800);
  }

  async function downloadDocx() {
    setBuilding(true);
    setDocxError(null);
    try {
      const payload: CaseDoc = { ...data, driveLinks, images: shots };
      const blob = await buildCaseDocx(payload);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = caseFileName(data);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (error) {
      setDocxError(error instanceof Error ? error.message : "No se pudo generar el documento.");
    } finally {
      setBuilding(false);
    }
  }

  const mailFor = (link: string) => ({
    to: DRIVE_CONFIG.to,
    cc: DRIVE_CONFIG.cc,
    subject: caseMailSubject(data),
    body: caseMailBody(data, link),
  });

  /* Chrome sólo deja saltar a un protocolo externo con una activación de usuario
     reciente, y acá venimos de un await largo (popup de Google + subida). Así que
     este intento puede no hacer nada, en silencio y sin error. Por eso el bloque
     de resultado deja el botón a mano: ese click sí es un gesto fresco. */
  function openMail(link: string) {
    window.location.href = mailtoUrl(mailFor(link));
  }

  async function copyDriveLink(link: string) {
    try {
      await navigator.clipboard.writeText(link);
    } catch {
      /* Sin permiso de portapapeles el enlace sigue visible y seleccionable. */
    }
    setLinkCopied(true);
    window.setTimeout(() => setLinkCopied(false), 1800);
  }

  /* Sube el caso a Drive convertido a Google Doc y abre el mail con el enlace.
     El .docx no se adjunta: `mailto:` no puede, y Finnegans pide enlaces. */
  async function sendViaDrive() {
    setSending(true);
    setDriveError(null);
    try {
      const blob = await buildCaseDocx({ ...data, driveLinks, images: shots });
      const token = await requestDriveToken(DRIVE_CONFIG.clientId);
      const file = await uploadAsGoogleDoc({
        blob,
        fileName: caseFileName(data),
        token,
        folderId: DRIVE_CONFIG.folderId,
      });
      setDriveLink(file.webViewLink);
      openMail(file.webViewLink);
    } catch (error) {
      setDriveError(error instanceof Error ? error.message : "No se pudo subir el caso a Drive.");
    } finally {
      setSending(false);
    }
  }

  /* Cada cuadro de texto acepta capturas pegadas, y van a parar a su propia
     sección del documento. */
  const pasteable = (slot: ImageSlot) => ({
    slot,
    files: shots[slot] ?? [],
    onAdd: addShots,
    onRemove: removeShot,
  });

  return (
    <main className="page">
      <header className="topbar">
        <Lockup size={34} />
        <div className="topbar-id">
          <strong>Casos Finnegans</strong>
          <span>Asistente de carga</span>
        </div>
        <div className="status-dot"><i /> Borrador guardado en este dispositivo</div>
      </header>

      <section className="hero">
        <SectionPill size="15px" bleed={false}>NUEVO CASO</SectionPill>
        <div className="hero-heading">
          <PairedHeading
            as="h1"
            line1="CONTANOS QUÉ PASÓ."
            line2="NOSOTROS LO ORDENAMOS."
            size="clamp(32px,4.4vw,52px)"
          />
        </div>
        <p>Completá la información esencial y descargá el documento de Finnegans ya armado, listo para enviar a soporte.</p>
        <div className="progress-wrap" aria-label={`Progreso ${progress}%`}>
          <div className="progress-copy"><span>Progreso del caso</span><b>{progress}%</b></div>
          <div className="progress"><span style={{ width: `${progress}%` }} /></div>
        </div>
      </section>

      <div className="workspace">
        <form className="case-form" onSubmit={(event) => { event.preventDefault(); setShowPreview(true); }}>
          <Section number="01" title="Datos generales" note="Identificá el entorno donde ocurrió el problema.">
            <div className="grid three">
              <Field label="Fecha de ingreso"><input value={data.date} onChange={(e) => update("date", e.target.value)} /></Field>
              <Field label="Dominio *"><input placeholder="Ej. INMIX" value={data.domain} onChange={(e) => update("domain", e.target.value)} /></Field>
              <Field label="Cliente *"><input placeholder="Razón social" value={data.client} onChange={(e) => update("client", e.target.value)} /></Field>
            </div>
            <div className="grid three">
              <Field label="Redactor *"><input placeholder="Nombre y apellido" value={data.author} onChange={(e) => update("author", e.target.value)} /></Field>
              <Field label="Servidor propio / dedicado"><Select value={data.dedicatedServer} onChange={(v) => update("dedicatedServer", v)} options={["No", "Sí"]} /></Field>
              <Field label="Prioridad"><Select value={data.priority} onChange={(v) => update("priority", v)} options={["Baja", "Media", "Alta", "Crítica"]} /></Field>
            </div>
            <Field label="Empresa / sucursal" hint="La plantilla pide identificarla cuando el problema depende del contexto.">
              <input placeholder="Empresa y sucursal donde se reproduce" value={data.company} onChange={(e) => update("company", e.target.value)} />
            </Field>
          </Section>

          <Section number="02" title="Consulta a Finni" note="Si ya consultaste a Finni, pegá acá el prompt y su respuesta.">
            <Field label="Prompt y respuesta de Finni" hint="Opcional. Evita que soporte repita un camino ya recorrido.">
              <Pasteable placeholder="Pegá la consulta que hiciste y lo que respondió." value={data.finni} onChange={(v) => update("finni", v)} {...pasteable("finni")} />
            </Field>
          </Section>

          <Section number="03" title="Descripción del problema" note="Usá lenguaje concreto. No es necesario conocer términos técnicos.">
            <Field label="Título breve del caso *" hint="Ejemplo: El nuevo motor de retenciones no se aplica en el cálculo">
              <input placeholder="Resumen del problema en una línea" value={data.title} onChange={(e) => update("title", e.target.value)} />
            </Field>
            <Field label="¿Qué no funciona? *">
              <Pasteable placeholder="Indicá la pantalla, proceso u operación afectada." value={data.whatFails} onChange={(v) => update("whatFails", v)} {...pasteable("whatFails")} />
            </Field>
            <Field label="¿Qué es lo que pasa? *">
              <Pasteable placeholder="Describí el mensaje, bloqueo o resultado obtenido." value={data.whatHappens} onChange={(v) => update("whatHappens", v)} {...pasteable("whatHappens")} />
            </Field>
            <div className="grid two">
              <Field label="¿Es la primera vez que hace la transacción?"><Select value={data.firstTime} onChange={(v) => update("firstTime", v)} options={["No", "Sí", "No sé"]} /></Field>
              <Field label="¿Es una operación estándar?"><Select value={data.standardOperation} onChange={(v) => update("standardOperation", v)} options={["Sí", "No", "No sé"]} /></Field>
            </div>
            <Field label="¿Tiene alguna contingencia manual?" hint="Opcional. Si hay un rodeo temporal, contalo acá. Si no, se envía como «No».">
              <Pasteable placeholder="Ej. se cargan los partes en una planilla aparte." value={data.workaround} onChange={(v) => update("workaround", v)} {...pasteable("workaround")} />
            </Field>
          </Section>

          <Section number="04" title="Situación actual" note="Qué impacto tiene hoy el cliente mientras el problema sigue abierto.">
            <Field label="Impacto en el cliente *">
              <Pasteable placeholder="Qué no puede hacer, desde cuándo y a cuánta gente u operación afecta." value={data.impact} onChange={(v) => update("impact", v)} {...pasteable("impact")} />
            </Field>
          </Section>

          <Section number="05" title="Caso de uso y resultado" note="Estos datos ayudan a soporte a repetir el error sin pedir aclaraciones.">
            <Field label="Pasos para reproducir el problema *" hint="Incluí empresa, menú, datos cargados y momento exacto del error.">
              <Pasteable className="large" placeholder={"1. Ingresar a…\n2. Seleccionar…\n3. Cargar…\n4. Guardar…"} value={data.steps} onChange={(v) => update("steps", v)} {...pasteable("steps")} />
            </Field>
            <Field label="Formato de grilla o informe" hint="Opcional. Solo si el problema aparece en una grilla o un informe.">
              <input placeholder="Nombre del formato que usa el cliente" value={data.reportFormat} onChange={(e) => update("reportFormat", e.target.value)} />
            </Field>
            <Field label="¿Qué soluciones o pruebas ya se intentaron?">
              <Pasteable placeholder="Indicá cambios de configuración, pruebas y sus resultados." value={data.attempts} onChange={(v) => update("attempts", v)} {...pasteable("attempts")} />
            </Field>
            <Field label="Resultado esperado *" hint="Si no sabés qué debería dar, la plantilla permite dejarlo vacío.">
              <Pasteable placeholder="Explicá qué debería permitir hacer el sistema." value={data.expected} onChange={(v) => update("expected", v)} {...pasteable("expected")} />
            </Field>
          </Section>

          <Section number="06" title="Evidencias en Drive" note="Opcional. Para videos y archivos que no entran como captura.">
            <div className="callout">
              <strong>Las capturas van arriba</strong>
              <p>Pegá cada captura con Ctrl+V dentro del cuadro de texto que corresponda y se incrusta en esa misma sección del documento. Acá van sólo los videos y archivos grandes.</p>
            </div>
            <Field label="Enlaces de Google Drive" hint="Opcional. Pegá un enlace por línea.">
              <textarea placeholder={"https://drive.google.com/…\nhttps://docs.google.com/…"} value={data.evidence} onChange={(e) => update("evidence", e.target.value)} />
            </Field>
            {invalidLinks.length > 0 && <p className="error">Revisá los enlaces: todos deben pertenecer a Google Drive o Google Docs.</p>}
            {/* Sin enlaces no hay permisos que confirmar. */}
            {driveLinks.length > 0 && (
              <label className="check"><input type="checkbox" checked={data.accessConfirmed} onChange={(e) => update("accessConfirmed", e.target.checked)} /><span>Confirmo que soporte puede abrir todos los enlaces</span></label>
            )}
          </Section>

          <button className="cta" type="submit" style={{ background: ctaBg }}>Revisar y generar caso</button>
        </form>

        <aside className="side">
          <div className="side-card">
            <span className="side-label">Control de calidad</span>
            <h3>Un buen caso acelera<br /><b>la solución.</b></h3>
            <ul>
              <li className={progress === 100 ? "done" : ""}><span>{progress === 100 ? "✓" : "1"}</span>Datos obligatorios completos</li>
              <li className={hasEvidence ? "done" : ""}><span>{hasEvidence ? "✓" : "2"}</span>Evidencias adjuntas <em>opcional</em></li>
              {driveLinks.length > 0 && (
                <li className={data.accessConfirmed ? "done" : ""}><span>{data.accessConfirmed ? "✓" : "3"}</span>Permisos de Drive confirmados</li>
              )}
            </ul>
            <div className={`readiness ${ready ? "ready" : ""}`}>{readiness}</div>
          </div>
          <div className="tip"><b>Consejo</b><p>Describí hechos observables: qué hiciste, qué mostró el sistema y qué esperabas que sucediera.</p></div>
          <p className="claim">{CLAIM}</p>
        </aside>
      </div>

      {showPreview && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="preview-title">
          <div className="modal">
            <button className="close" onClick={() => setShowPreview(false)} aria-label="Cerrar">×</button>
            <span className="side-label">Vista previa</span>
            <div className="modal-heading">
              <PairedHeading id="preview-title" line1="CASO LISTO" line2="PARA REVISAR" size="34px" />
            </div>
            {!ready && (
              <div className="warning">
                {linksOk
                  ? "El borrador fue generado, pero todavía hay datos obligatorios sin completar."
                  : "Hay enlaces que no son de Google Drive ni de Google Docs: soporte no va a poder abrirlos."}
              </div>
            )}
            <textarea className="output" value={output} onChange={() => {}} readOnly />
            {driveLink && (
              <div className="drive-result">
                <strong>Documento creado en Drive</strong>
                <a href={driveLink} target="_blank" rel="noreferrer">{driveLink}</a>
                <p>Si Outlook no se abrió solo, abrilo desde acá.</p>
                <div className="drive-actions">
                  <button type="button" className="cta compact" onClick={() => openMail(driveLink)} style={{ background: ctaBg }}>
                    Abrir en Outlook
                  </button>
                  {/* Un https común: abre aunque el equipo no tenga Outlook de escritorio. */}
                  <a className="secondary" href={outlookWebUrl(mailFor(driveLink))} target="_blank" rel="noreferrer">
                    Abrir en Outlook Web
                  </a>
                  <button type="button" className="secondary" onClick={() => copyDriveLink(driveLink)}>
                    {linkCopied ? "¡Copiado!" : "Copiar enlace"}
                  </button>
                </div>
              </div>
            )}
            {docxError && <p className="error">{docxError}</p>}
            {driveError && <p className="error">{driveError}</p>}
            <div className="modal-actions">
              <button className="secondary" onClick={() => setShowPreview(false)}>Volver a editar</button>
              <button className="secondary" onClick={copyCase}>{copied ? "¡Copiado!" : "Copiar texto"}</button>
              <button className="secondary" onClick={downloadDocx} disabled={building}>
                {building ? "Generando…" : "Descargar .docx"}
              </button>
              {isDriveConfigured() && (
                <button className="cta compact" onClick={sendViaDrive} disabled={sending} style={{ background: ctaBg }}>
                  {sending ? "Subiendo a Drive…" : "Subir a Drive y enviar"}
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

/* Cuadro de texto que acepta capturas pegadas o arrastradas. */
function Pasteable({
  slot,
  value,
  onChange,
  placeholder,
  className,
  files,
  onAdd,
  onRemove,
}: {
  slot: ImageSlot;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  className?: string;
  files: File[];
  onAdd: (slot: ImageSlot, files: File[]) => void;
  onRemove: (slot: ImageSlot, index: number) => void;
}) {
  const [over, setOver] = useState(false);
  return (
    <div className={`pasteable ${over ? "over" : ""}`}>
      <textarea
        className={className}
        placeholder={placeholder}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onPaste={(event) => {
          const images = imagesFrom(event.clipboardData);
          /* Sólo interceptamos si vino una imagen; el texto se pega normal. */
          if (!images.length) return;
          event.preventDefault();
          onAdd(slot, images);
        }}
        onDragOver={(event) => {
          if (!Array.from(event.dataTransfer.types).includes("Files")) return;
          event.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(event) => {
          const images = imagesFrom(event.dataTransfer);
          setOver(false);
          if (!images.length) return;
          event.preventDefault();
          onAdd(slot, images);
        }}
      />
      {files.length === 0 ? (
        <small className="paste-hint">Pegá una captura con Ctrl+V o arrastrala acá.</small>
      ) : (
        <Shots files={files} onRemove={(index) => onRemove(slot, index)} />
      )}
    </div>
  );
}

function Shots({ files, onRemove }: { files: File[]; onRemove: (index: number) => void }) {
  const urls = useMemo(() => files.map((file) => URL.createObjectURL(file)), [files]);
  useEffect(() => () => urls.forEach(URL.revokeObjectURL), [urls]);
  return (
    <ul className="shots">
      {urls.map((url, index) => (
        <li key={url}>
          {/* eslint-disable-next-line @next/next/no-img-element -- es un blob local, no un asset servido */}
          <img src={url} alt={files[index].name || `Captura ${index + 1}`} />
          <button type="button" onClick={() => onRemove(index)} aria-label={`Quitar captura ${index + 1}`}>×</button>
        </li>
      ))}
    </ul>
  );
}

function Section({ number, title, note, children }: { number: string; title: string; note: string; children: React.ReactNode }) {
  return (
    <Panel pad="0">
      <div className="section-head">
        <span>{number}</span>
        <div><h2>{title}</h2><p>{note}</p></div>
      </div>
      <div className="section-body">{children}</div>
    </Panel>
  );
}

function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return <label className="field"><span>{label}</span>{hint && <small>{hint}</small>}{children}</label>;
}

function Select({ value, options, onChange }: { value: string; options: string[]; onChange: (value: string) => void }) {
  return <select value={value} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select>;
}
