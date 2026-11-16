import html2canvas from "html2canvas-pro";
import { jsPDF } from "jspdf";

/**
 * Renders a DOM element to a multi-page A4 PDF.
 * Elements marked with `data-pdf-expand` lose their scroll clipping so wide tables are captured whole.
 */
export async function createPdfBlob(el: HTMLElement, fileName: string): Promise<{ blob: Blob; fileName: string }> {
  const expanded = Array.from(el.querySelectorAll<HTMLElement>("[data-pdf-expand]"));
  const prev = expanded.map((n) => n.style.overflow);
  expanded.forEach((n) => (n.style.overflow = "visible"));
  let canvas: HTMLCanvasElement;
  try {
    const bg = getComputedStyle(document.body).backgroundColor || "#ffffff";
    canvas = await html2canvas(el, { scale: 2, backgroundColor: bg, useCORS: true, windowWidth: el.scrollWidth });
  } catch {
    throw new Error("PDF로 만들 화면을 그리지 못했습니다. 다른 형식(Word·HTML)으로 내려받아 인쇄하세요.");
  } finally {
    expanded.forEach((n, i) => (n.style.overflow = prev[i]));
  }

  const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
  const margin = 10;
  const pageW = pdf.internal.pageSize.getWidth() - margin * 2;
  const pageH = pdf.internal.pageSize.getHeight() - margin * 2;
  const pxPerMm = canvas.width / pageW;
  const slicePx = Math.floor(pageH * pxPerMm);

  for (let y = 0, page = 0; y < canvas.height; y += slicePx, page++) {
    const h = Math.min(slicePx, canvas.height - y);
    const part = document.createElement("canvas");
    part.width = canvas.width;
    part.height = h;
    const ctx = part.getContext("2d");
    if (!ctx) throw new Error("PDF 페이지를 만들지 못했습니다.");
    ctx.drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);
    if (page > 0) pdf.addPage();
    pdf.addImage(part.toDataURL("image/jpeg", 0.92), "JPEG", margin, margin, pageW, h / pxPerMm);
  }
  return { blob: pdf.output("blob"), fileName };
}
