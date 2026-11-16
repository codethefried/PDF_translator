import * as XLSX from "xlsx";
import { DOC_KIND_LABEL, type Seg, type StructuredDoc } from "@/lib/pdf-parse";

export type ExportFormat = "pdf" | "xlsx" | "doc" | "txt" | "html";
export type ShowMode = "translated" | "both" | "original";

export const EXPORT_FORMATS: { id: ExportFormat; label: string; ext: string; note: string }[] = [
  { id: "pdf", label: "PDF", ext: ".pdf", note: "화면에 보이는 양식 그대로" },
  { id: "xlsx", label: "Excel", ext: ".xlsx", note: "항목·표·문항을 시트별로" },
  { id: "doc", label: "Word", ext: ".doc", note: "Word에서 열고 편집" },
  { id: "txt", label: "텍스트", ext: ".txt", note: "서식 없는 텍스트" },
  { id: "html", label: "HTML", ext: ".html", note: "브라우저로 보는 문서" },
];

type Tr = Record<string, string>;

const pick = (s: Seg, tr: Tr, mode: ShowMode) => (mode === "original" ? s.src : tr[s.id] ?? s.src);
const pair = (s: Seg, tr: Tr, mode: ShowMode) => {
  const main = pick(s, tr, mode);
  return mode === "both" && tr[s.id] !== undefined && tr[s.id] !== s.src ? `${main}\n  (${s.src})` : main;
};

/* ---------- text ---------- */

export function toText(doc: StructuredDoc, tr: Tr, mode: ShowMode): string {
  const out: string[] = [];
  if (doc.title) out.push(pair(doc.title, tr, mode), "");
  if (doc.fields.length) {
    for (const f of doc.fields) out.push(`${pick(f.label, tr, mode)}: ${pair(f.value, tr, mode)}`);
    out.push("");
  }
  doc.tables.forEach((t, i) => {
    out.push(`[표 ${i + 1}]`, t.header.map((c) => pick(c, tr, mode)).join("\t"));
    t.rows.forEach((r) => out.push(r.map((c) => pick(c, tr, mode)).join("\t")));
    out.push("");
  });
  for (const c of doc.clauses) {
    out.push(pair(c.title, tr, mode));
    c.body.forEach((b) => out.push(pair(b, tr, mode)));
    out.push("");
  }
  for (const q of doc.qa) {
    out.push(`${q.no} ${q.question.map((s) => pair(s, tr, mode)).join(" ")}`);
    q.choices.forEach((c) => out.push(`  ${c.label}. ${pair(c.seg, tr, mode)}`));
    if (q.answer) out.push(`  정답: ${q.answer}`);
    q.explanation.forEach((e) => out.push(`  해설: ${pair(e, tr, mode)}`));
    out.push("");
  }
  for (const b of doc.blocks) out.push(pair(b.seg, tr, mode), "");
  return out.join("\n").trim() + "\n";
}

/* ---------- html / word ---------- */

const esc = (s: string) => s.replace(/[&<>"]/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[m] ?? m);
function cell(s: Seg, tr: Tr, mode: ShowMode) {
  const main = esc(pick(s, tr, mode));
  return mode === "both" && tr[s.id] !== undefined && tr[s.id] !== s.src ? `${main}<span class="src">${esc(s.src)}</span>` : main;
}

export function toHtml(doc: StructuredDoc, tr: Tr, mode: ShowMode, fileTitle: string): string {
  const parts: string[] = [];
  parts.push(`<p class="kind">${esc(DOC_KIND_LABEL[doc.kind])}</p>`);
  if (doc.title) parts.push(`<h1>${cell(doc.title, tr, mode)}</h1>`);
  if (doc.fields.length) {
    parts.push(`<table class="fields">${doc.fields.map((f) => `<tr><th>${cell(f.label, tr, mode)}</th><td>${cell(f.value, tr, mode)}</td></tr>`).join("")}</table>`);
  }
  for (const t of doc.tables) {
    parts.push(
      `<table class="grid"><thead><tr>${t.header.map((c) => `<th>${cell(c, tr, mode)}</th>`).join("")}</tr></thead><tbody>${t.rows
        .map((r) => `<tr>${r.map((c) => `<td>${cell(c, tr, mode)}</td>`).join("")}</tr>`)
        .join("")}</tbody></table>`,
    );
  }
  for (const c of doc.clauses) {
    parts.push(`<section class="clause"><h2>${cell(c.title, tr, mode)}</h2>${c.body.map((b) => `<p>${cell(b, tr, mode)}</p>`).join("")}</section>`);
  }
  for (const q of doc.qa) {
    parts.push(
      `<section class="qa"><h2>${esc(q.no)}</h2>${q.question.map((s) => `<p>${cell(s, tr, mode)}</p>`).join("")}<ol class="choices">${q.choices
        .map((c) => `<li${q.answer.toUpperCase().includes(c.label) ? ' class="ans"' : ""}><b>${c.label}.</b> ${cell(c.seg, tr, mode)}</li>`)
        .join("")}</ol>${q.answer ? `<p class="answer">정답: ${esc(q.answer)}</p>` : ""}${q.explanation
        .map((e) => `<p class="ex"><b>해설</b> ${cell(e, tr, mode)}</p>`)
        .join("")}</section>`,
    );
  }
  for (const b of doc.blocks) parts.push(b.kind === "heading" ? `<h2>${cell(b.seg, tr, mode)}</h2>` : `<p>${cell(b.seg, tr, mode)}</p>`);

  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${esc(fileTitle)}</title><style>
body{font-family:"Malgun Gothic","Apple SD Gothic Neo",Arial,sans-serif;color:#1f2933;max-width:900px;margin:32px auto;padding:0 16px;line-height:1.6}
h1{font-size:22px;margin:4px 0 16px}h2{font-size:16px;margin:20px 0 6px}.kind{color:#52606d;font-size:12px;margin:0}
table{border-collapse:collapse;width:100%;margin:12px 0;font-size:13px}th,td{border:1px solid #cbd2d9;padding:6px 8px;text-align:left;vertical-align:top}
.fields th{width:28%;background:#f5f7fa}.grid thead th{background:#e4e7eb}.src{display:block;color:#616e7c;font-size:11px}
.qa{border:1px solid #cbd2d9;border-radius:6px;padding:10px 14px;margin:12px 0}.choices{list-style:none;padding:0}.choices li{margin:4px 0}
.ans{font-weight:600}.answer{font-weight:600}.ex{background:#f5f7fa;padding:6px 8px}
</style></head><body>${parts.join("\n")}</body></html>`;
}

/* ---------- excel ---------- */

export function toXlsx(doc: StructuredDoc, tr: Tr): Blob {
  const wb = XLSX.utils.book_new();
  const t = (s: Seg) => tr[s.id] ?? "";
  const add = (name: string, rows: (string | number)[][]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);

  if (doc.fields.length || doc.title) {
    const rows: string[][] = [["항목(원문)", "항목(번역)", "값(원문)", "값(번역)"]];
    if (doc.title) rows.push(["제목", "", doc.title.src, t(doc.title)]);
    doc.fields.forEach((f) => rows.push([f.label.src, t(f.label), f.value.src, t(f.value)]));
    add("기본 정보", rows);
  }
  doc.tables.forEach((tb, i) => {
    add(`표${i + 1} 번역`, [tb.header.map((c) => tr[c.id] ?? c.src), ...tb.rows.map((r) => r.map((c) => tr[c.id] ?? c.src))]);
    add(`표${i + 1} 원문`, [tb.header.map((c) => c.src), ...tb.rows.map((r) => r.map((c) => c.src))]);
  });
  if (doc.clauses.length) {
    const rows: string[][] = [["조항(원문)", "조항(번역)", "내용(원문)", "내용(번역)"]];
    doc.clauses.forEach((c) => {
      if (!c.body.length) rows.push([c.title.src, t(c.title), "", ""]);
      c.body.forEach((b, i) => rows.push(i === 0 ? [c.title.src, t(c.title), b.src, t(b)] : ["", "", b.src, t(b)]));
    });
    add("조항", rows);
  }
  if (doc.qa.length) {
    const maxCh = Math.max(...doc.qa.map((q) => q.choices.length), 0);
    const labels = Array.from({ length: maxCh }, (_, i) => String.fromCharCode(65 + i));
    const rows: string[][] = [["번호", "문제(원문)", "문제(번역)", ...labels.flatMap((l) => [`${l}(원문)`, `${l}(번역)`]), "정답", "해설(원문)", "해설(번역)"]];
    doc.qa.forEach((q) => {
      const ch = labels.flatMap((l) => {
        const c = q.choices.find((x) => x.label === l);
        return c ? [c.seg.src, t(c.seg)] : ["", ""];
      });
      rows.push([
        q.no,
        q.question.map((s) => s.src).join(" "),
        q.question.map(t).join(" "),
        ...ch,
        q.answer,
        q.explanation.map((s) => s.src).join(" "),
        q.explanation.map(t).join(" "),
      ]);
    });
    add("문항", rows);
  }
  if (doc.blocks.length) add("본문", [["구분", "원문", "번역"], ...doc.blocks.map((b) => [b.kind === "heading" ? "제목" : "문단", b.seg.src, t(b.seg)])]);
  if (!wb.SheetNames.length) add("본문", [["내용 없음"]]);
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

/* ---------- download ---------- */

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
