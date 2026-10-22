import jsPDF from "jspdf";
import { Worker } from "./boleta";

export async function exportBoletaToPDF(worker: Worker, boletaText: string, filenamePrefix = "Boleta"): Promise<void> {
  if (!worker || !boletaText) return;

  const pdf = new jsPDF({ unit: "mm", format: "a4" });
  await addBoletaPage(pdf, boletaText);

  const cleanName = `${filenamePrefix}_${worker.nombres}`
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "_");

  pdf.save(`${cleanName}.pdf`);
}

export async function exportBoletasToPDF(
  boletas: Array<{ worker: Worker; text: string }>,
  filename = "Boletas_seleccionadas"
): Promise<void> {
  const validBoletas = boletas.filter(({ worker, text }) => worker && text);
  if (!validBoletas.length) return;

  const pdf = new jsPDF({ unit: "mm", format: "a4" });
  for (const [{ text }, index] of validBoletas.map((item, index) => [item, index] as const)) {
    if (index > 0) pdf.addPage();
    await addBoletaPage(pdf, text);
  }

  const cleanName = filename.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_");
  pdf.save(`${cleanName}.pdf`);
}

async function addBoletaPage(pdf: jsPDF, text: string): Promise<void> {
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

  const response = await fetch("/SELLO.png");
  if (response.ok) {
    const seal = await response.blob();
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = reject;
      reader.readAsDataURL(seal);
    });
    pdf.addImage(dataUrl, "PNG", 108, 120, 57, 22.8);
  }
}
