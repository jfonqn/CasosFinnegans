/* Componentes del design system Fisterra, portados desde
   _ds/fisterra-design-system-ebd22bc4-acea-4bd4-9219-217e96906278/_ds_bundle.js.
   La geometría y las proporciones son verbatim del bundle: no ajustar a ojo. */
import type { CSSProperties, ReactNode } from "react";

type Tone = "color" | "navy" | "white" | "outline";

const FACES: Record<Tone, [string, string]> = {
  color: ["var(--fs-red)", "var(--fs-red-deep)"],
  navy: ["var(--fs-navy)", "var(--fs-navy)"],
  white: ["var(--fs-white)", "var(--fs-white)"],
  outline: ["none", "none"],
};

/* Isotipo Fisterra. La base cóncava (Q…95.2) es lo que lo distingue de un
   triángulo — no aplanarla. */
export function Isotipo({ size = 64, tone = "color", shadow = true, style }: {
  size?: number; tone?: Tone; shadow?: boolean; style?: CSSProperties;
}) {
  const faces = FACES[tone] ?? FACES.color;
  const outline = tone === "outline";
  return (
    <svg
      viewBox="0 0 100 100" width={size} height={size} role="img" aria-label="Fisterra"
      style={{ display: "block", flex: "none", filter: shadow && !outline ? "drop-shadow(var(--fs-shadow-logo))" : undefined, ...style }}
    >
      <path d="M50 0 L0 100 Q25 95.2 50 95.2 Z" fill={faces[0]} stroke={outline ? "var(--fs-red)" : undefined} strokeWidth={outline ? 1.4 : undefined} />
      <path d="M50 0 L50 95.2 Q75 95.2 100 100 Z" fill={faces[1]} stroke={outline ? "var(--fs-red-deep)" : undefined} strokeWidth={outline ? 1.4 : undefined} />
    </svg>
  );
}

/* Lockup horizontal: isotipo + wordmark FISTERRA en Montserrat Regular.
   Wordmark ≈ 68% de la altura del isotipo; separación ≈ 40% de su ancho. */
export function Lockup({ size = 58, tone = "color", shadow = true, style }: {
  size?: number; tone?: Tone; shadow?: boolean; style?: CSSProperties;
}) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: size * 0.4, ...style }}>
      <Isotipo size={size} tone={tone === "white" ? "white" : tone} shadow={shadow} />
      <span style={{
        fontFamily: "var(--fs-font)", fontSize: size * 0.68, fontWeight: "var(--fs-weight-regular)",
        letterSpacing: "var(--fs-tracking-wordmark)", lineHeight: 1, color: tone === "white" ? "var(--fs-white)" : "var(--fs-navy)",
      }}>FISTERRA</span>
    </span>
  );
}

/* Pill de sección — degradé navy, radio 999px, texto blanco Medium.
   Con bleed sangra fuera del margen izquierdo hasta el borde del lienzo. */
export function SectionPill({ children, size, bleed = true, style }: {
  children: ReactNode; size?: string; bleed?: boolean; style?: CSSProperties;
}) {
  return (
    <span style={{
      display: "inline-block", background: "var(--fs-gradient-navy)", color: "var(--fs-white)",
      borderRadius: "var(--fs-radius-pill)", fontFamily: "var(--fs-font)", fontSize: size ?? "var(--fs-text-pill)",
      fontWeight: "var(--fs-weight-medium)", lineHeight: 1.25, padding: "0.5em 1.6em",
      paddingLeft: bleed ? "2.2em" : "1.6em", marginLeft: bleed ? "calc(var(--fs-margin-canvas) * -1)" : 0,
      ...style,
    }}>{children}</span>
  );
}

/* Titular pareado — el recurso central del sistema. Dos líneas, mismo cuerpo,
   contraste de peso y color: línea 1 Bold rojo, línea 2 Regular tinta.
   Interlínea 1.05 para que lean como una unidad. */
export function PairedHeading({ line1, line2, size, align = "left", invert = false, as: Tag = "h2", id, style }: {
  line1: string; line2?: string; size?: string; align?: CSSProperties["textAlign"];
  /* El bundle fija h2; `as` sólo cambia el tag para no romper el esquema de
     encabezados de la página. No altera nada visual. */
  invert?: boolean; as?: "h1" | "h2" | "h3"; id?: string; style?: CSSProperties;
}) {
  return (
    <Tag id={id} style={{
      textAlign: align, fontFamily: "var(--fs-font)", fontSize: size ?? "var(--fs-text-display)",
      lineHeight: "var(--fs-leading-tight)", margin: 0, ...style,
    }}>
      <span style={{ display: "block", fontWeight: invert ? 400 : 700, color: invert ? "var(--fs-ink)" : "var(--fs-red)" }}>{line1}</span>
      {line2 ? <span style={{ display: "block", fontWeight: invert ? 700 : 400, color: invert ? "var(--fs-red)" : "var(--fs-ink)" }}>{line2}</span> : null}
    </Tag>
  );
}

/* Panel — el contenedor principal de toda pieza: radio 28px, filo blanco 1.5px,
   fondo apenas más claro que la página. */
export function Panel({ children, withTabs = false, pad = "3.4%", className, style }: {
  children: ReactNode; withTabs?: boolean; pad?: string; className?: string; style?: CSSProperties;
}) {
  return (
    <div className={className} style={{
      position: "relative", background: "var(--fs-panel-fill)", border: "var(--fs-hairline)",
      borderRadius: withTabs ? "0 var(--fs-radius-panel) var(--fs-radius-panel) var(--fs-radius-panel)" : "var(--fs-radius-panel)",
      boxShadow: "var(--fs-shadow-panel)", padding: pad, ...style,
    }}>{children}</div>
  );
}
