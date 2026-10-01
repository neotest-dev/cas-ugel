import jsPDF from "jspdf";
import { Worker } from "./boleta";

export function exportBoletaToPDF(worker: Worker, boletaText: string, filenamePrefix = "Boleta"): void {
  if (!worker || !boletaText) return;

  const pdf = new jsPDF({ unit: "mm", format: "a4" });
  pdf.setFont("courier", "normal");
  pdf.setFontSize(9);

  let y = 15;
  for (const line of boletaText.split("\n")) {
    if (y > 285) {
      pdf.addPage();
      y = 15;
    }
    pdf.text(line, 15, y);
    y += 4.2;
  }

  const cleanName = `${filenamePrefix}_${worker.n || "1"}_${worker.apPaterno}_${worker.nombres}`
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "_");

  pdf.save(`${cleanName}.pdf`);
}
