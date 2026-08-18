"use client";

import { useMemo, useState } from "react";

type CaseData = {
  caseNumber: string;
  date: string;
  domain: string;
  client: string;
  author: string;
  dedicatedServer: string;
  priority: string;
  title: string;
  whatFails: string;
  whatHappens: string;
  firstTime: string;
  standardOperation: string;
  steps: string;
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
  dedicatedServer: "No",
  priority: "Media",
  title: "",
  whatFails: "",
  whatHappens: "",
  firstTime: "No",
  standardOperation: "Sí",
  steps: "",
  attempts: "",
  expected: "",
  evidence: "",
  accessConfirmed: false,
};

const drivePattern = /^https?:\/\/(drive|docs)\.google\.com\//i;

function clean(text: string) {
  const value = text.trim();
  if (!value) return "Pendiente de completar";
  return value.charAt(0).toUpperCase() + value.slice(1).replace(/\s+/g, " ");
}

export default function Home() {
  const [data, setData] = useState(initialData);
  const [copied, setCopied] = useState(false);
  const [showPreview, setShowPreview] = useState(false);

  const update = <K extends keyof CaseData>(key: K, value: CaseData[K]) =>
    setData((current) => ({ ...current, [key]: value }));

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
    data.steps,
    data.expected,
  ];
  const completed = required.filter((value) => value.trim()).length;
  const progress = Math.round((completed / required.length) * 100);

  const output = useMemo(() => {
    const evidenceText = driveLinks.length
      ? driveLinks.map((link, index) => `- Evidencia ${index + 1}: ${link}`).join("\n")
      : "- Pendiente: cargar evidencia en Google Drive y agregar el enlace";

    return `BUG / PROBLEMA REPORTADO

N° de caso: ${data.caseNumber || "A completar por Finnegans"}
Fecha de ingreso: ${data.date}
Dominio: ${clean(data.domain)}
Cliente: ${clean(data.client)}
Redactor: ${clean(data.author)}
Servidor propio / dedicado: ${data.dedicatedServer}
Prioridad del caso: ${data.priority}

TÍTULO
${clean(data.title)}

PROBLEMA

¿Qué no funciona?
${clean(data.whatFails)}

¿Qué es lo que pasa?
${clean(data.whatHappens)}

¿Es la primera vez que hace la transacción?
${data.firstTime}

¿Es una operación estándar?
${data.standardOperation}

CASO DE USO / PASOS PARA REPRODUCIR
${clean(data.steps)}

PRUEBAS O SOLUCIONES INTENTADAS
${clean(data.attempts)}

RESULTADO ESPERADO
${clean(data.expected)}

EVIDENCIAS EN GOOGLE DRIVE
${evidenceText}
Acceso verificado para soporte: ${data.accessConfirmed ? "Sí" : "Pendiente de confirmar"}`;
  }, [data, driveLinks]);

  const ready = progress === 100 && driveLinks.length > 0 && invalidLinks.length === 0 && data.accessConfirmed;

  async function copyCase() {
    await navigator.clipboard.writeText(output);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1800);
  }

  return (
    <main>
      <header className="topbar">
        <div className="brand-mark">F</div>
        <div>
          <strong>Casos Finnegans</strong>
          <span>Asistente de carga</span>
        </div>
        <div className="status-dot"><i /> Borrador guardado en este dispositivo</div>
      </header>

      <section className="hero">
        <div className="eyebrow">NUEVO CASO</div>
        <h1>Contanos qué pasó.<br /><em>Nosotros lo ordenamos.</em></h1>
        <p>Completá la información esencial y obtené un caso claro, verificable y listo para enviar a soporte.</p>
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
              <Field label="Dominio *"><input placeholder="Ej. COLUN" value={data.domain} onChange={(e) => update("domain", e.target.value)} /></Field>
              <Field label="Cliente *"><input placeholder="Razón social" value={data.client} onChange={(e) => update("client", e.target.value)} /></Field>
            </div>
            <div className="grid three">
              <Field label="Redactor *"><input placeholder="Nombre y apellido" value={data.author} onChange={(e) => update("author", e.target.value)} /></Field>
              <Field label="Servidor propio / dedicado"><Select value={data.dedicatedServer} onChange={(v) => update("dedicatedServer", v)} options={["No", "Sí"]} /></Field>
              <Field label="Prioridad"><Select value={data.priority} onChange={(v) => update("priority", v)} options={["Baja", "Media", "Alta", "Crítica"]} /></Field>
            </div>
          </Section>

          <Section number="02" title="Descripción del problema" note="Usá lenguaje concreto. No es necesario conocer términos técnicos.">
            <Field label="Título breve del caso *" hint="Ejemplo: El plan de ventas no permite distribuir por centro de costo">
              <input placeholder="Resumen del problema en una línea" value={data.title} onChange={(e) => update("title", e.target.value)} />
            </Field>
            <Field label="¿Qué no funciona? *">
              <textarea placeholder="Indicá la pantalla, proceso u operación afectada." value={data.whatFails} onChange={(e) => update("whatFails", e.target.value)} />
            </Field>
            <Field label="¿Qué sucede cuando intentás hacerlo? *">
              <textarea placeholder="Describí el mensaje, bloqueo o resultado obtenido." value={data.whatHappens} onChange={(e) => update("whatHappens", e.target.value)} />
            </Field>
            <div className="grid two">
              <Field label="¿Es la primera vez que hacés la transacción?"><Select value={data.firstTime} onChange={(v) => update("firstTime", v)} options={["No", "Sí", "No sé"]} /></Field>
              <Field label="¿Es una operación estándar?"><Select value={data.standardOperation} onChange={(v) => update("standardOperation", v)} options={["Sí", "No", "No sé"]} /></Field>
            </div>
          </Section>

          <Section number="03" title="Reproducción y resultado" note="Estos datos ayudan a soporte a repetir el error sin pedir aclaraciones.">
            <Field label="Pasos para reproducir el problema *" hint="Incluí empresa, menú, datos cargados y momento exacto del error.">
              <textarea className="large" placeholder={'1. Ingresar a…\n2. Seleccionar…\n3. Cargar…\n4. Guardar…'} value={data.steps} onChange={(e) => update("steps", e.target.value)} />
            </Field>
            <Field label="¿Qué soluciones o pruebas ya se intentaron?">
              <textarea placeholder="Indicá cambios de configuración, pruebas y sus resultados." value={data.attempts} onChange={(e) => update("attempts", e.target.value)} />
            </Field>
            <Field label="Resultado esperado *">
              <textarea placeholder="Explicá qué debería permitir hacer el sistema." value={data.expected} onChange={(e) => update("expected", e.target.value)} />
            </Field>
          </Section>

          <Section number="04" title="Evidencias en Google Drive" note="Finnegans recibe enlaces; no adjuntes archivos físicos.">
            <div className="drive-callout">
              <span className="drive-icon">▲</span>
              <div><strong>Antes de continuar</strong><p>Subí capturas, videos o documentos a Drive y habilitá el acceso para las personas que recibirán el caso.</p></div>
            </div>
            <Field label="Enlaces de Google Drive" hint="Pegá un enlace por línea.">
              <textarea placeholder={'https://drive.google.com/…\nhttps://docs.google.com/…'} value={data.evidence} onChange={(e) => update("evidence", e.target.value)} />
            </Field>
            {invalidLinks.length > 0 && <p className="error">Revisá los enlaces: todos deben pertenecer a Google Drive o Google Docs.</p>}
            <label className="check"><input type="checkbox" checked={data.accessConfirmed} onChange={(e) => update("accessConfirmed", e.target.checked)} /><span>Confirmo que soporte puede abrir todos los enlaces</span></label>
          </Section>

          <button className="primary" type="submit">Revisar y generar caso <span>→</span></button>
        </form>

        <aside>
          <div className="side-card">
            <span className="side-label">CONTROL DE CALIDAD</span>
            <h3>Un buen caso acelera la solución.</h3>
            <ul>
              <li className={progress === 100 ? "done" : ""}><span>{progress === 100 ? "✓" : "1"}</span>Datos obligatorios completos</li>
              <li className={driveLinks.length > 0 && !invalidLinks.length ? "done" : ""}><span>{driveLinks.length > 0 && !invalidLinks.length ? "✓" : "2"}</span>Evidencias enlazadas</li>
              <li className={data.accessConfirmed ? "done" : ""}><span>{data.accessConfirmed ? "✓" : "3"}</span>Permisos confirmados</li>
            </ul>
            <div className={`readiness ${ready ? "ready" : ""}`}>{ready ? "Caso listo para enviar" : "Completá los puntos pendientes"}</div>
          </div>
          <div className="tip"><b>Consejo</b><p>Describí hechos observables: qué hiciste, qué mostró el sistema y qué esperabas que sucediera.</p></div>
        </aside>
      </div>

      {showPreview && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="preview-title">
          <div className="modal">
            <button className="close" onClick={() => setShowPreview(false)} aria-label="Cerrar">×</button>
            <span className="side-label">VISTA PREVIA</span>
            <h2 id="preview-title">Caso listo para revisar</h2>
            {!ready && <div className="warning">El borrador fue generado, pero todavía hay información o evidencias pendientes.</div>}
            <textarea className="output" value={output} onChange={() => {}} readOnly />
            <div className="modal-actions">
              <button className="secondary" onClick={() => setShowPreview(false)}>Volver a editar</button>
              <button className="primary compact" onClick={copyCase}>{copied ? "¡Copiado!" : "Copiar caso completo"}</button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

function Section({ number, title, note, children }: { number: string; title: string; note: string; children: React.ReactNode }) {
  return <section className="form-section"><div className="section-head"><span>{number}</span><div><h2>{title}</h2><p>{note}</p></div></div><div className="section-body">{children}</div></section>;
}

function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return <label className="field"><span>{label}</span>{hint && <small>{hint}</small>}{children}</label>;
}

function Select({ value, options, onChange }: { value: string; options: string[]; onChange: (value: string) => void }) {
  return <select value={value} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select>;
}
