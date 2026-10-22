import { createPortal } from "react-dom";

interface PrintBoletaPortalProps {
  text: string | string[];
}

export function PrintBoletaPortal({ text }: PrintBoletaPortalProps) {
  const pages = (Array.isArray(text) ? text : [text]).filter(Boolean);
  if (!pages.length) return null;

  return createPortal(
    <div id="boleta-print-root">
      {pages.map((page, index) => (
        <div
          className="boleta-print-page"
          key={index}
          style={{ pageBreakAfter: index < pages.length - 1 ? "always" : "auto" }}
        >
          <pre className="boleta-print-mono">{page}</pre>
          <img className="boleta-print-seal" src="/SELLO.png" alt="Sello y firma de planillas" />
        </div>
      ))}
    </div>,
    document.body
  );
}
