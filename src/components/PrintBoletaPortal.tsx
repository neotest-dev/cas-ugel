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
        <pre
          className="boleta-print-mono"
          key={index}
          style={{ pageBreakAfter: index < pages.length - 1 ? "always" : "auto" }}
        >
          {page}
        </pre>
      ))}
    </div>,
    document.body
  );
}
