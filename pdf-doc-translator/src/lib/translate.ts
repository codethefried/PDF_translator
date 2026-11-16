import type { Seg } from "@/lib/pdf-parse";

export type Engine = "copilot" | "browser";

/** The Copilot reply could not be read as a translation list. */
export class AgentReplyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AgentReplyError";
  }
}

/** A problem the user can fix in the app (settings, language, browser support). */
export class TranslateInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "TranslateInputError";
  }
}

export const LANGS = [
  { code: "ko", name: "한국어", en: "Korean" },
  { code: "en", name: "영어", en: "English" },
  { code: "ja", name: "일본어", en: "Japanese" },
  { code: "zh", name: "중국어(간체)", en: "Simplified Chinese" },
  { code: "es", name: "스페인어", en: "Spanish" },
  { code: "vi", name: "베트남어", en: "Vietnamese" },
  { code: "de", name: "독일어", en: "German" },
  { code: "fr", name: "프랑스어", en: "French" },
] as const;

export const langName = (code: string) => LANGS.find((l) => l.code === code)?.name ?? code;
const langEn = (code: string) => LANGS.find((l) => l.code === code)?.en ?? code;

export interface TranslateSettings {
  engine: Engine;
  /** "" = detect automatically */
  from: string;
  to: string;
}

export interface Progress {
  done: number;
  total: number;
}

export const hasWords = (s: string) => /\p{L}{2,}/u.test(s);

/* ---------- Copilot chat (copy & paste, no plugin) ---------- */

export function buildPrompt(items: { id: string; t: string }[], from: string, to: string) {
  const source = from ? `from ${langEn(from)} ` : "";
  return [
    `Translate the "t" value of every item ${source}into ${langEn(to)}.`,
    "Rules: keep numbers, amounts, dates, codes, model numbers, e-mail addresses and URLs unchanged; keep the tone of a business document; do not merge, split, skip or add items; do not add explanations.",
    'Reply with ONLY a JSON array in the form [{"id":"...","t":"..."}] using exactly the same ids.',
    "Items:",
    JSON.stringify(items),
  ].join("\n");
}

export function parseReply(text: string, ids: Set<string>): Record<string, string> {
  const body = text.replace(/```(?:json)?/gi, "");
  const start = body.indexOf("[");
  const end = body.lastIndexOf("]");
  if (start < 0 || end <= start) throw new AgentReplyError("Copilot 답변에서 번역 목록(JSON)을 찾지 못했습니다.");
  let parsed: unknown;
  try {
    parsed = JSON.parse(body.slice(start, end + 1));
  } catch {
    throw new AgentReplyError("Copilot 답변의 번역 목록(JSON) 형식이 올바르지 않습니다.");
  }
  if (!Array.isArray(parsed)) throw new AgentReplyError("Copilot 답변이 목록 형식이 아닙니다.");
  const out: Record<string, string> = {};
  for (const row of parsed) {
    if (typeof row !== "object" || row === null) continue;
    const { id, t } = row as Record<string, unknown>;
    if (typeof id === "string" && typeof t === "string" && ids.has(id)) out[id] = t;
  }
  if (!Object.keys(out).length) throw new AgentReplyError("Copilot 답변에 요청한 문장의 번역이 없습니다.");
  return out;
}

export function batches(segs: Seg[], maxChars = 6000, maxItems = 60): Seg[][] {
  const out: Seg[][] = [];
  let cur: Seg[] = [];
  let size = 0;
  for (const s of segs) {
    if (cur.length && (size + s.src.length > maxChars || cur.length >= maxItems)) {
      out.push(cur);
      cur = [];
      size = 0;
    }
    cur.push(s);
    size += s.src.length;
  }
  if (cur.length) out.push(cur);
  return out;
}

/* ---------- browser built-in translator ---------- */

interface BrowserTranslator {
  translate(text: string): Promise<string>;
  destroy?: () => void;
}
interface TranslatorStatic {
  availability(o: { sourceLanguage: string; targetLanguage: string }): Promise<string>;
  create(o: { sourceLanguage: string; targetLanguage: string }): Promise<BrowserTranslator>;
}
interface DetectorStatic {
  create(): Promise<{ detect(text: string): Promise<{ detectedLanguage: string; confidence: number }[]> }>;
}
const browserTranslator = () => (globalThis as unknown as { Translator?: TranslatorStatic }).Translator ?? null;
const browserDetector = () => (globalThis as unknown as { LanguageDetector?: DetectorStatic }).LanguageDetector ?? null;

export type BrowserStatus = "unsupported" | "unavailable" | "downloadable" | "downloading" | "available" | "need-source";

export async function browserStatus(from: string, to: string): Promise<BrowserStatus> {
  const api = browserTranslator();
  if (!api) return "unsupported";
  if (!from && !browserDetector()) return "need-source";
  if (!from) return "available";
  try {
    const s = await api.availability({ sourceLanguage: from, targetLanguage: to });
    return s === "available" || s === "downloadable" || s === "downloading" ? s : "unavailable";
  } catch {
    return "unavailable";
  }
}

async function detectSource(sample: string): Promise<string> {
  const det = browserDetector();
  if (!det) throw new TranslateInputError("원문 언어를 직접 선택하세요.");
  let best: { detectedLanguage: string; confidence: number } | undefined;
  try {
    const d = await det.create();
    [best] = await d.detect(sample.slice(0, 2000));
  } catch {
    best = undefined;
  }
  if (!best || best.confidence < 0.4) throw new TranslateInputError("원문 언어를 알아내지 못했습니다. 원문 언어를 직접 선택하세요.");
  return best.detectedLanguage.split("-")[0];
}

/* ---------- entry point ---------- */

/** Browser built-in translator (Chrome/Edge). Runs fully on the device. */
export async function translateSegs(
  segs: Seg[],
  settings: TranslateSettings,
  onProgress: (p: Progress) => void,
  isCancelled: () => boolean,
): Promise<{ map: Record<string, string>; missing: number }> {
  if (settings.from && settings.from === settings.to) throw new TranslateInputError("원문 언어와 번역 언어가 같습니다.");
  const map: Record<string, string> = {};
  const work: Seg[] = [];
  for (const s of segs) {
    if (hasWords(s.src)) work.push(s);
    else map[s.id] = s.src;
  }
  const total = work.length;
  let done = 0;
  onProgress({ done, total });

  {
    const api = browserTranslator();
    if (!api) throw new TranslateInputError("이 브라우저는 내장 번역을 지원하지 않습니다.");
    const from = settings.from || (await detectSource(work.map((s) => s.src).join("\n")));
    if (from === settings.to) throw new TranslateInputError("원문 언어와 번역 언어가 같습니다.");
    let tr: BrowserTranslator;
    try {
      tr = await api.create({ sourceLanguage: from, targetLanguage: settings.to });
    } catch {
      throw new TranslateInputError("브라우저 내장 번역을 시작하지 못했습니다. 이 언어 조합을 지원하는지 확인하거나 'Copilot 채팅 번역'을 선택하세요.");
    }
    try {
      for (const s of work) {
        if (isCancelled()) break;
        map[s.id] = await tr.translate(s.src);
        done += 1;
        if (done % 5 === 0 || done === total) onProgress({ done, total });
      }
    } catch {
      throw new TranslateInputError("브라우저 내장 번역 중 오류가 났습니다. 다시 시도하세요.");
    } finally {
      tr.destroy?.();
    }
  }
  const missing = work.filter((s) => map[s.id] === undefined).length;
  return { map, missing };
}
