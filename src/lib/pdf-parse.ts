import { extractTextItems } from "unpdf";

/* ---------- types ---------- */

export type DocKind = "qa" | "order" | "estimate" | "contract" | "general";

export const DOC_KIND_LABEL: Record<DocKind, string> = {
  qa: "문제집 / Dump",
  order: "발주서",
  estimate: "견적서",
  contract: "계약서",
  general: "일반 문서",
};

export interface Line {
  page: number;
  y: number;
  size: number;
  text: string;
  cells: string[];
}

export interface ParsedPdf {
  pageCount: number;
  /** Cleaned lines, page order, top to bottom. */
  lines: Line[];
}

export interface Detection {
  kind: DocKind;
  reasons: string[];
  scores: Record<DocKind, number>;
}

/** A translatable unit. `id` keys the translation map. */
export interface Seg {
  id: string;
  src: string;
}

export interface FieldRow { label: Seg; value: Seg }
export interface TableBlock { header: Seg[]; rows: Seg[][] }
export interface Clause { title: Seg; body: Seg[] }
export interface QaItem { no: string; question: Seg[]; choices: { label: string; seg: Seg }[]; answer: string; explanation: Seg[] }
export interface Block { kind: "heading" | "paragraph"; seg: Seg }

export interface StructuredDoc {
  kind: DocKind;
  title: Seg | null;
  fields: FieldRow[];
  tables: TableBlock[];
  clauses: Clause[];
  qa: QaItem[];
  blocks: Block[];
  segs: Seg[];
}

/* ---------- extraction ---------- */

const PAGE_NO = /^\s*(?:page\s*)?\d{1,4}(?:\s*(?:\/|of)\s*\d{1,4})?\s*$|^\s*-\s*\d{1,4}\s*-\s*$/i;

export async function parsePdf(bytes: Uint8Array): Promise<ParsedPdf> {
  const { totalPages, items } = await extractTextItems(bytes);
  const pages: Line[][] = items.map((pageItems, pi) => {
    const clean = pageItems
      .filter((it) => it.str.trim())
      .map((it) => ({ str: it.str, x: it.x, y: it.y, w: it.width || it.str.length * (it.fontSize || 10) * 0.5, size: it.fontSize || it.height || 10 }));
    clean.sort((a, b) => b.y - a.y || a.x - b.x);
    const rows: (typeof clean)[] = [];
    for (const it of clean) {
      const row = rows.find((r) => Math.abs(r[0].y - it.y) <= Math.max(2, Math.min(r[0].size, it.size) * 0.45));
      if (row) row.push(it);
      else rows.push([it]);
    }
    return rows
      .map((row) => {
        row.sort((a, b) => a.x - b.x);
        const cells: string[] = [];
        let cur = "";
        let prevEnd = -Infinity;
        for (const it of row) {
          const gap = it.x - prevEnd;
          if (cur && gap > Math.max(it.size * 1.4, 12)) {
            cells.push(cur.trim());
            cur = "";
          }
          cur += (cur && gap > it.size * 0.15 && !cur.endsWith(" ") ? " " : "") + it.str;
          prevEnd = it.x + it.w;
        }
        if (cur.trim()) cells.push(cur.trim());
        const size = Math.max(...row.map((r) => r.size));
        return { page: pi + 1, y: row[0].y, size, cells, text: cells.join(" ").replace(/\s+/g, " ").trim() };
      })
      .filter((l) => l.text);
  });

  // repeated headers / footers: same text (digits ignored) in top/bottom two lines of most pages
  const norm = (t: string) => t.replace(/\d+/g, "#").toLowerCase();
  const edgeCount = new Map<string, number>();
  for (const pg of pages) {
    const edges = new Set([...pg.slice(0, 2), ...pg.slice(-2)].map((l) => norm(l.text)));
    edges.forEach((k) => edgeCount.set(k, (edgeCount.get(k) ?? 0) + 1));
  }
  const minRepeat = Math.max(3, Math.ceil(pages.length * 0.5));
  const repeated = new Set([...edgeCount].filter(([, n]) => n >= minRepeat).map(([k]) => k));

  const lines = pages.flatMap((pg) =>
    pg.filter((l, i) => {
      const edge = i < 2 || i >= pg.length - 2;
      if (edge && PAGE_NO.test(l.text)) return false;
      if (edge && repeated.has(norm(l.text))) return false;
      return true;
    }),
  );
  return { pageCount: totalPages, lines };
}

/* ---------- page range ---------- */

export function parseRange(input: string, max: number): number[] {
  const t = input.trim();
  if (!t || t === "전체") return Array.from({ length: max }, (_, i) => i + 1);
  const set = new Set<number>();
  for (const part of t.split(",")) {
    const p = part.trim();
    if (!p) continue;
    const m = p.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) {
      let [a, b] = [Number(m[1]), Number(m[2])];
      if (a > b) [a, b] = [b, a];
      for (let n = a; n <= b; n++) if (n >= 1 && n <= max) set.add(n);
    } else if (/^\d+$/.test(p)) {
      const n = Number(p);
      if (n >= 1 && n <= max) set.add(n);
    } else {
      throw new Error("페이지 범위는 1-5, 8 처럼 입력하세요.");
    }
  }
  if (!set.size) throw new Error(`페이지는 1~${max} 범위로 입력하세요.`);
  return [...set].sort((a, b) => a - b);
}

/* ---------- classification ---------- */

const Q_START = /^(?:NO\.?\s*\d+|QUESTION\s*(?:NO\.?\s*)?\d+|Q\s*\d+\s*[.:)]|문제\s*\d+|\d+\s*번\s*문제|Topic\s*\d+\s*,?\s*Question\s*#?\s*\d+)/i;
const CHOICE = /^\(?([A-Ha-h])[.)．]\s*(.*)$/;
const ANSWER = /^(?:Correct\s+Answer|Answer|정답|답)\s*[:：]?\s*(.*)$/i;
const EXPLAIN = /^(?:Explanation|해설|설명|Reference)\s*[:：]?\s*(.*)$/i;
const CLAUSE = /^(?:제\s*\d+\s*조(?:\s*[(（].{1,40}[)）])?.*|Article\s+\d+[.:]?.*|ARTICLE\s+[IVXLC\d]+.*|Section\s+\d+[.:].*)$/;
const FIELD = /^([^:：]{1,30}?)\s*[:：]\s*(.+)$/;

export function detectKind(lines: Line[]): Detection {
  const sample = lines.slice(0, 400);
  const all = sample.map((l) => l.text).join("\n");
  const low = all.toLowerCase();
  const count = (re: RegExp) => (all.match(re) ?? []).length;
  const lineCount = (re: RegExp) => sample.filter((l) => re.test(l.text)).length;

  const qStarts = lineCount(Q_START);
  const choices = lineCount(CHOICE);
  const answers = lineCount(ANSWER);
  const clauses = lineCount(CLAUSE);
  const multiCol = sample.filter((l) => l.cells.length >= 3).length;
  const amounts = count(/\d{1,3}(?:,\d{3})+(?:\.\d+)?/g);

  const scores: Record<DocKind, number> = { qa: 0, order: 0, estimate: 0, contract: 0, general: 1 };
  const reasons: Record<DocKind, string[]> = { qa: [], order: [], estimate: [], contract: [], general: ["뚜렷한 양식 패턴 없음"] };
  const add = (k: DocKind, n: number, why: string) => {
    if (n <= 0) return;
    scores[k] += n;
    reasons[k].push(why);
  };

  add("qa", Math.min(qStarts, 10) * 2, `문제 번호 ${qStarts}개`);
  if (choices >= 3) add("qa", Math.min(choices, 20) * 0.6, `A~D 형식 보기 ${choices}줄`);
  add("qa", Math.min(answers, 10) * 1.5, `정답/Answer 표기 ${answers}개`);

  const contractKw = count(/계약서|계약기간|계약 기간|계약금액|계약금|위약|갑\s*[과와은는]|을\s*[은는]|agreement|contract|whereas|hereby|hereinafter|terminat/gi);
  add("contract", Math.min(contractKw, 10) * 1.5, `계약 관련 표현 ${contractKw}회`);
  add("contract", Math.min(clauses, 12) * 1.5, `조항 제목(제n조·Article) ${clauses}개`);

  const orderKw = count(/발주서|발주번호|발주일|purchase\s*order|p\.?\s*o\.?\s*(?:no|number|#)|납기|납품|delivery\s*date|ship\s*to|배송지/gi);
  add("order", Math.min(orderKw, 10) * 2, `발주·납기 표현 ${orderKw}회`);
  const estKw = count(/견적서|견적|quotation|quote\s*no|estimate|공급가액|유효기간|valid\s*(?:until|for)/gi);
  add("estimate", Math.min(estKw, 10) * 2, `견적 관련 표현 ${estKw}회`);
  if (multiCol >= 3) {
    add("order", Math.min(multiCol, 15) * 0.4, `3열 이상 표 행 ${multiCol}줄`);
    add("estimate", Math.min(multiCol, 15) * 0.4, `3열 이상 표 행 ${multiCol}줄`);
  }
  if (amounts >= 3) {
    add("order", 1, `금액 표기 ${amounts}개`);
    add("estimate", 1, `금액 표기 ${amounts}개`);
  }
  if (/견적서|quotation/.test(low) && !/발주서|purchase\s*order/.test(low)) scores.estimate += 3;
  if (/발주서|purchase\s*order/.test(low)) scores.order += 3;

  const kind = (Object.keys(scores) as DocKind[]).reduce((a, b) => (scores[b] > scores[a] ? b : a), "general");
  return { kind: scores[kind] >= 4 ? kind : "general", reasons: scores[kind] >= 4 ? reasons[kind] : reasons.general, scores };
}

/* ---------- structuring ---------- */

export function structure(lines: Line[], kind: DocKind): StructuredDoc {
  const segs: Seg[] = [];
  let n = 0;
  const seg = (src: string): Seg => {
    const s = { id: `s${++n}`, src: src.trim() };
    segs.push(s);
    return s;
  };
  const doc: StructuredDoc = { kind, title: null, fields: [], tables: [], clauses: [], qa: [], blocks: [], segs };
  if (!lines.length) return doc;

  const bodySize = median(lines.map((l) => l.size));
  const first = lines[0];
  let rest = lines;
  if (first.size > bodySize * 1.2 || (kind !== "qa" && first.text.length < 60)) {
    doc.title = seg(first.text);
    rest = lines.slice(1);
  }

  if (kind === "qa") buildQa(rest, doc, seg);
  else if (kind === "order" || kind === "estimate") buildForm(rest, doc, seg);
  else if (kind === "contract") buildContract(rest, doc, seg);
  else buildGeneral(rest, doc, seg, bodySize);
  return doc;
}

type SegFn = (s: string) => Seg;

function buildQa(lines: Line[], doc: StructuredDoc, seg: SegFn) {
  let cur: { no: string; q: string[]; ch: { label: string; text: string }[]; ans: string; ex: string[] } | null = null;
  let mode: "q" | "ch" | "ex" = "q";
  const flush = () => {
    if (!cur) return;
    doc.qa.push({
      no: cur.no,
      question: cur.q.length ? [seg(cur.q.join(" "))] : [],
      choices: cur.ch.map((c) => ({ label: c.label, seg: seg(c.text) })),
      answer: cur.ans,
      explanation: cur.ex.length ? [seg(cur.ex.join(" "))] : [],
    });
  };
  for (const l of lines) {
    const t = l.text;
    const qm = t.match(Q_START);
    if (qm) {
      flush();
      const restText = t.slice(qm[0].length).replace(/^[\s.:)\-–]+/, "");
      cur = { no: qm[0].replace(/\s+/g, " ").trim(), q: restText ? [restText] : [], ch: [], ans: "", ex: [] };
      mode = "q";
      continue;
    }
    if (!cur) {
      doc.blocks.push({ kind: "paragraph", seg: seg(t) });
      continue;
    }
    const am = t.match(ANSWER);
    if (am) {
      cur.ans = am[1].trim();
      mode = "ex";
      continue;
    }
    const em = t.match(EXPLAIN);
    if (em) {
      if (em[1].trim()) cur.ex.push(em[1].trim());
      mode = "ex";
      continue;
    }
    const cm = mode !== "ex" ? t.match(CHOICE) : null;
    if (cm) {
      cur.ch.push({ label: cm[1].toUpperCase(), text: cm[2] });
      mode = "ch";
      continue;
    }
    if (mode === "q") cur.q.push(t);
    else if (mode === "ch") cur.ch[cur.ch.length - 1].text += " " + t;
    else cur.ex.push(t);
  }
  flush();
}

function buildForm(lines: Line[], doc: StructuredDoc, seg: SegFn) {
  let table: { header: string[]; rows: string[][] } | null = null;
  const closeTable = () => {
    if (table && table.rows.length) {
      const width = Math.max(table.header.length, ...table.rows.map((r) => r.length));
      const pad = (r: string[]) => [...r, ...Array(width - r.length).fill("")];
      doc.tables.push({ header: pad(table.header).map(seg), rows: table.rows.map((r) => pad(r).map(seg)) });
    } else if (table) {
      doc.blocks.push({ kind: "paragraph", seg: seg(table.header.join("  ")) });
    }
    table = null;
  };
  for (const l of lines) {
    if (l.cells.length >= 3) {
      if (!table) table = { header: l.cells, rows: [] };
      else table.rows.push(l.cells);
      continue;
    }
    closeTable();
    pushFields(l, doc, seg);
  }
  closeTable();
}

function pushFields(l: Line, doc: StructuredDoc, seg: SegFn) {
  if (l.cells.length === 2 && l.cells[0].length <= 30) {
    const label = l.cells[0].replace(/[:：]\s*$/, "");
    doc.fields.push({ label: seg(label), value: seg(l.cells[1]) });
    return;
  }
  const fm = l.text.match(FIELD);
  if (fm && !/^https?$/i.test(fm[1])) {
    doc.fields.push({ label: seg(fm[1]), value: seg(fm[2]) });
    return;
  }
  doc.blocks.push({ kind: "paragraph", seg: seg(l.text) });
}

function buildContract(lines: Line[], doc: StructuredDoc, seg: SegFn) {
  let cur: { title: string; body: string[] } | null = null;
  const flush = () => {
    if (cur) doc.clauses.push({ title: seg(cur.title), body: mergeLines(cur.body).map(seg) });
  };
  for (const l of lines) {
    if (CLAUSE.test(l.text)) {
      flush();
      cur = { title: l.text, body: [] };
      continue;
    }
    if (!cur) {
      pushFields(l, doc, seg);
      continue;
    }
    cur.body.push(l.text);
  }
  flush();
}

function buildGeneral(lines: Line[], doc: StructuredDoc, seg: SegFn, bodySize: number) {
  let para: string[] = [];
  let prev: Line | null = null;
  const flush = () => {
    if (para.length) doc.blocks.push({ kind: "paragraph", seg: seg(para.join(" ")) });
    para = [];
  };
  for (const l of lines) {
    const heading = l.size > bodySize * 1.15 && l.text.length < 90;
    const gap = prev && prev.page === l.page ? prev.y - l.y : 0;
    if (heading) {
      flush();
      doc.blocks.push({ kind: "heading", seg: seg(l.text) });
    } else {
      if (gap > l.size * 1.9 || (prev && /[.!?。다]$/.test(prev.text) && /^[-•·▪\d]/.test(l.text))) flush();
      para.push(l.text);
    }
    prev = l;
  }
  flush();
}

/** Joins wrapped lines into sentences-ending paragraphs. */
function mergeLines(texts: string[]): string[] {
  const out: string[] = [];
  let cur = "";
  for (const t of texts) {
    const startsItem = /^(?:[①-⑳]|\(\d+\)|\d+\.|[가-하]\.|[a-z]\)|-|•)/.test(t);
    if (cur && (startsItem || /[.。다:]$/.test(cur))) {
      out.push(cur);
      cur = t;
    } else cur = cur ? `${cur} ${t}` : t;
  }
  if (cur) out.push(cur);
  return out;
}

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)] || 10;
}
