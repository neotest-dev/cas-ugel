import jsPDF from "jspdf";
import { Worker } from "./boleta";

export function exportBoletaToPDF(worker: Worker, boletaText: string, filenamePrefix = "Boleta"): void {
  if (!worker || !boletaText) return;

  const pdf = new jsPDF({ unit: "mm", format: "a4" });
  addBoletaPage(pdf, boletaText);

  const cleanName = `${filenamePrefix}_${worker.n || "1"}_${worker.apPaterno}_${worker.nombres}`
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "_");

  pdf.save(`${cleanName}.pdf`);
}

export function exportBoletasToPDF(
  boletas: Array<{ worker: Worker; text: string }>,
  filename = "Boletas_seleccionadas"
): void {
  const validBoletas = boletas.filter(({ worker, text }) => worker && text);
  if (!validBoletas.length) return;

  const pdf = new jsPDF({ unit: "mm", format: "a4" });
  validBoletas.forEach(({ text }, index) => {
    if (index > 0) pdf.addPage();
    addBoletaPage(pdf, text);
  });

  const cleanName = filename.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_");
  pdf.save(`${cleanName}.pdf`);
}

function addBoletaPage(pdf: jsPDF, text: string): void {
  pdf.setFont("courier", "normal");
  pdf.setFontSize(9);

  let y = 15;
  for (const line of text.split("\n")) {
    if (y > 285) {
      pdf.addPage();
      y = 15;
    }
    pdf.text(line, 15, y);
    y += 4.2;
  }
}
