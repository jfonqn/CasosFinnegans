import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Asistente de Casos Finnegans",
  description: "Creá casos claros y completos para soporte de Finnegans.",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="es"><body>{children}</body></html>;
}
