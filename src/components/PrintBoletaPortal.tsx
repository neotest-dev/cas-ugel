import { createPortal } from "react-dom";

interface PrintBoletaPortalProps {
  text: string;
}

export function PrintBoletaPortal({ text }: PrintBoletaPortalProps) {
  if (!text) return null;

  return createPortal(
    <div id="boleta-print-root">
      <pre className="boleta-print-mono">{text}</pre>
    </div>,
    document.body
  );
}
