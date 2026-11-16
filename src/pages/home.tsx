import { useCallback, useRef, useState } from "react";
import { useMutation, useQuery } from "@tanstack/react-query";
import { ClipboardCopy, Download, ExternalLink, FileText, Languages, Square, Trash2, Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Progress } from "@/components/ui/progress";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { Table, TableBody, TableCaption, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState } from "@/components/states";
import { Textarea } from "@/components/ui/textarea";
import { LiveStatus } from "@/components/live-status";
import { cn } from "@/lib/utils";
import { DOC_KIND_LABEL, detectKind, parsePdf, parseRange, structure, type Detection, type DocKind, type ParsedPdf, type Seg, type StructuredDoc } from "@/lib/pdf-parse";
import { AgentReplyError, LANGS, batches, browserStatus, buildPrompt, hasWords, langName, parseReply, translateSegs, type Engine, type Progress as TrProgress } from "@/lib/translate";
import { EXPORT_FORMATS, downloadBlob, toHtml, toText, toXlsx, type ExportFormat, type ShowMode } from "@/lib/exporters";
import { createPdfBlob } from "@/lib/pdf-export";

/* ---------- settings (browser) ---------- */

const SETTINGS_KEY = "pdf-doc-translator:settings";
interface Settings { engine: Engine; from: string; to: string }
const DEFAULT_SETTINGS: Settings = { engine: "browser", from: "", to: "ko" };
const COPILOT_URL = "https://m365.cloud.microsoft/chat";

const confirm = async (o: { title: string; description?: string; confirmLabel?: string; destructive?: boolean }) => window.confirm(o.description ? `${o.title}\n\n${o.description}` : o.title);

interface ManualSession { fileId: string; groups: Seg[][]; idx: number; reply: string; error: string | null; copied: boolean }
function loadSettings(): Settings {
  try {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    const v = JSON.parse(raw) as Partial<Settings>;
    return {
      engine: v.engine === "copilot" ? "copilot" : "browser",
      from: typeof v.from === "string" ? v.from : "",
      to: typeof v.to === "string" && v.to ? v.to : "ko",
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}
function saveSettings(s: Settings) {
  try {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  } catch {
    /* settings stay in memory */
  }
}

/* ---------- document state ---------- */

interface DocFile {
  id: string;
  name: string;
  parsed: ParsedPdf;
  range: string;
  detection: Detection;
  kind: DocKind;
  doc: StructuredDoc;
  tr: Record<string, string>;
  missing: number;
}

function build(parsed: ParsedPdf, range: string, kind?: DocKind) {
  const pages = new Set(parseRange(range, parsed.pageCount));
  const lines = parsed.lines.filter((l) => pages.has(l.page));
  const detection = detectKind(lines);
  const k = kind ?? detection.kind;
  return { detection, kind: k, doc: structure(lines, k) };
}

const baseName = (n: string) => n.replace(/\.pdf$/i, "");
const KINDS: DocKind[] = ["qa", "order", "estimate", "contract", "general"];

export function HomePage() {
  const [settings, setSettingsState] = useState<Settings>(loadSettings);
  const setSettings = (p: Partial<Settings>) =>
    setSettingsState((s) => {
      const next = { ...s, ...p };
      saveSettings(next);
      return next;
    });
  const [files, setFiles] = useState<DocFile[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [loading, setLoading] = useState<string | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [rangeDraft, setRangeDraft] = useState("전체");
  const [rangeError, setRangeError] = useState<string | null>(null);
  const [show, setShow] = useState<ShowMode>("both");
  const [format, setFormat] = useState<ExportFormat>("pdf");
  const [exportBoth, setExportBoth] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [progress, setProgress] = useState<TrProgress | null>(null);
  const [status, setStatus] = useState({ text: "", seq: 0 });
  const [manual, setManual] = useState<ManualSession | null>(null);
  const cancelRef = useRef(false);
  const resultRef = useRef<HTMLDivElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);

  const announce = (text: string) => setStatus((s) => ({ text, seq: s.seq + 1 }));
  const active = files.find((f) => f.id === activeId) ?? null;
  const patchActive = useCallback((fn: (f: DocFile) => DocFile) => setFiles((fs) => fs.map((f) => (f.id === activeId ? fn(f) : f))), [activeId]);

  const browser = useQuery({
    queryKey: ["browser-translator", settings.from, settings.to],
    queryFn: () => browserStatus(settings.from, settings.to),
    staleTime: 60_000,
    retry: false,
  });

  /* ----- upload ----- */
  const onFiles = async (list: FileList | null) => {
    if (!list?.length) return;
    setLoadError(null);
    const added: DocFile[] = [];
    for (const f of Array.from(list)) {
      if (f.type !== "application/pdf" && !/\.pdf$/i.test(f.name)) {
        setLoadError(`${f.name}: PDF 파일만 올릴 수 있습니다.`);
        continue;
      }
      if (f.size > 25 * 1024 * 1024) {
        setLoadError(`${f.name}: 25MB 이하 파일만 처리할 수 있습니다.`);
        continue;
      }
      setLoading(f.name);
      try {
        const parsed = await parsePdf(new Uint8Array(await f.arrayBuffer()));
        if (!parsed.lines.length) {
          setLoadError(`${f.name}: 글자를 읽을 수 없는 PDF입니다. 스캔 이미지로 된 문서는 번역할 수 없습니다.`);
          continue;
        }
        const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
        added.push({ id, name: f.name, parsed, range: "전체", ...build(parsed, "전체"), tr: {}, missing: 0 });
      } catch {
        setLoadError(`${f.name}: PDF를 읽지 못했습니다. 암호가 걸려 있거나 손상된 파일인지 확인하세요.`);
      } finally {
        setLoading(null);
      }
    }
    if (added.length) {
      setFiles((fs) => [...fs, ...added]);
      const last = added[added.length - 1];
      setActiveId(last.id);
      setRangeDraft("전체");
      setRangeError(null);
      announce(`${added.length}개 문서를 읽었습니다. ${last.name}: ${DOC_KIND_LABEL[last.kind]}로 판별.`);
    }
  };

  const selectFile = (id: string) => {
    const f = files.find((x) => x.id === id);
    setActiveId(id);
    setRangeDraft(f?.range ?? "전체");
    setRangeError(null);
  };

  const removeFile = async (f: DocFile) => {
    if (Object.keys(f.tr).length && !(await confirm({ title: `${f.name}을(를) 목록에서 뺄까요?`, description: "번역 결과도 함께 사라집니다.", confirmLabel: "빼기", destructive: true }))) return;
    if (manual?.fileId === f.id) setManual(null);
    const rest = files.filter((x) => x.id !== f.id);
    setFiles(rest);
    if (activeId === f.id) {
      setActiveId(rest[0]?.id ?? null);
      setRangeDraft(rest[0]?.range ?? "전체");
    }
    announce(`${f.name}을(를) 뺐습니다. 남은 문서 ${rest.length}개.`);
  };

  const rebuild = async (opts: { range?: string; kind?: DocKind }) => {
    if (!active) return;
    if (Object.keys(active.tr).length && !(await confirm({ title: "번역 결과를 지우고 다시 구성할까요?", description: "페이지 범위나 문서 유형을 바꾸면 양식이 새로 만들어져 다시 번역해야 합니다.", confirmLabel: "다시 구성" }))) return false;
    if (manual?.fileId === active.id) setManual(null);
    const range = opts.range ?? active.range;
    try {
      const next = build(active.parsed, range, opts.kind ?? (opts.range ? undefined : active.kind));
      patchActive((f) => ({ ...f, range, ...next, tr: {}, missing: 0 }));
      setRangeError(null);
      announce(`${DOC_KIND_LABEL[next.kind]} 양식으로 구성했습니다.`);
      return true;
    } catch (e) {
      setRangeError(e instanceof Error ? e.message : "페이지 범위를 확인하세요.");
      return false;
    }
  };

  /* ----- translate ----- */
  const translate = useMutation({
    retry: false,
    mutationFn: async (target: DocFile) => {
      cancelRef.current = false;
      return { id: target.id, ...(await translateSegs(target.doc.segs, {
        engine: settings.engine,
        from: settings.from,
        to: settings.to,
      }, setProgress, () => cancelRef.current)) };
    },
    onSuccess: ({ id, map, missing }) => {
      setFiles((fs) => fs.map((f) => (f.id === id ? { ...f, tr: { ...f.tr, ...map }, missing } : f)));
      const n = Object.keys(map).length;
      announce(cancelRef.current ? `번역을 중지했습니다. ${n}개 문장 번역됨.` : `번역을 마쳤습니다. ${n}개 문장${missing ? `, 번역되지 않은 문장 ${missing}개` : ""}.`);
    },
  });
  const trError = translate.error;

  /* ----- Copilot chat (copy & paste) ----- */
  const manualFile = manual ? files.find((f) => f.id === manual.fileId) ?? null : null;
  const manualPrompt = manual && manual.idx < manual.groups.length
    ? buildPrompt(manual.groups[manual.idx].map((s) => ({ id: s.id, t: s.src })), settings.from, settings.to)
    : "";

  const startManual = (target: DocFile) => {
    if (settings.from && settings.from === settings.to) {
      setManual(null);
      setLoadError("원문 언어와 번역 언어가 같습니다.");
      return;
    }
    const keep: Record<string, string> = {};
    const work: Seg[] = [];
    for (const s of target.doc.segs) (hasWords(s.src) ? work.push(s) : (keep[s.id] = s.src));
    setFiles((fs) => fs.map((f) => (f.id === target.id ? { ...f, tr: keep, missing: work.length } : f)));
    const groups = batches(work);
    setManual({ fileId: target.id, groups, idx: 0, reply: "", error: null, copied: false });
    announce(`Copilot 채팅 번역을 시작합니다. 묶음 ${groups.length}개.`);
  };

  const copyPrompt = async () => {
    if (!manual) return;
    try {
      await navigator.clipboard.writeText(manualPrompt);
      setManual({ ...manual, copied: true, error: null });
      announce(`묶음 ${manual.idx + 1} 요청문을 복사했습니다. Copilot 채팅에 붙여 넣으세요.`);
    } catch {
      setManual({ ...manual, error: "클립보드에 복사하지 못했습니다. 아래 요청문을 직접 선택해 복사하세요." });
    }
  };

  const applyReply = () => {
    if (!manual) return;
    const group = manual.groups[manual.idx];
    try {
      const got = parseReply(manual.reply, new Set(group.map((s) => s.id)));
      const next = manual.idx + 1;
      setFiles((fs) => fs.map((f) => {
        if (f.id !== manual.fileId) return f;
        const tr = { ...f.tr, ...got };
        return { ...f, tr, missing: f.doc.segs.filter((s) => tr[s.id] === undefined).length };
      }));
      const lack = group.length - Object.keys(got).length;
      if (next >= manual.groups.length) {
        setManual(null);
        announce(`Copilot 채팅 번역을 마쳤습니다.${lack ? ` 마지막 묶음에서 ${lack}개 문장이 빠졌습니다.` : ""}`);
      } else {
        setManual({ ...manual, idx: next, reply: "", error: null, copied: false });
        announce(`묶음 ${manual.idx + 1} 적용 완료${lack ? `(${lack}개 누락)` : ""}. 다음 묶음 ${next + 1}/${manual.groups.length}.`);
      }
    } catch (e) {
      setManual({ ...manual, error: e instanceof AgentReplyError ? e.message : "답변을 읽지 못했습니다." });
    }
  };

  const engineReady = settings.engine === "copilot" ? true : browser.data === "available" || browser.data === "downloadable" || browser.data === "downloading";
  const translatedCount = active ? Object.keys(active.tr).length : 0;

  /* ----- export ----- */
  const doExport = async () => {
    if (!active) return;
    setExportError(null);
    setExporting(true);
    const mode: ShowMode = exportBoth ? "both" : "translated";
    const stem = `${baseName(active.name)}_${langName(settings.to)}`;
    try {
      if (format === "pdf") {
        if (!resultRef.current) throw new Error("결과 화면을 찾지 못했습니다.");
        const prev = show;
        if (prev !== mode) {
          setShow(mode);
          await new Promise((r) => setTimeout(r, 120));
        }
        try {
          const res = await createPdfBlob(resultRef.current, `${stem}.pdf`);
          downloadBlob(res.blob, res.fileName);
        } finally {
          if (prev !== mode) setShow(prev);
        }
      } else if (format === "xlsx") {
        downloadBlob(toXlsx(active.doc, active.tr), `${stem}.xlsx`);
      } else if (format === "doc") {
        downloadBlob(new Blob(["\ufeff" + toHtml(active.doc, active.tr, mode, stem)], { type: "application/msword" }), `${stem}.doc`);
      } else if (format === "txt") {
        downloadBlob(new Blob(["\ufeff" + toText(active.doc, active.tr, mode)], { type: "text/plain;charset=utf-8" }), `${stem}.txt`);
      } else {
        downloadBlob(new Blob([toHtml(active.doc, active.tr, mode, stem)], { type: "text/html;charset=utf-8" }), `${stem}.html`);
      }
      announce(`${EXPORT_FORMATS.find((f) => f.id === format)?.label} 파일 다운로드를 시작했습니다.`);
    } catch (e) {
      setExportError(e instanceof Error ? e.message : "파일을 만들지 못했습니다.");
    } finally {
      setExporting(false);
    }
  };

  const pct = progress && progress.total ? Math.round((progress.done / progress.total) * 100) : 0;

  return (
    <div className="mx-auto grid w-full max-w-[1500px] gap-6 px-4 py-5 lg:grid-cols-[340px_minmax(0,1fr)]">
      <LiveStatus>{status.seq > 0 ? `${status.text} (${status.seq})` : ""}</LiveStatus>

      <aside aria-label="문서와 번역 설정" className="flex flex-col gap-6">
        {/* files */}
        <section aria-labelledby="files-h" className="flex flex-col gap-2">
          <h2 id="files-h" className="text-sm font-semibold">PDF 문서</h2>
          <input ref={fileInput} id="pdf-input" type="file" accept="application/pdf,.pdf" multiple className="sr-only" tabIndex={-1}
            onChange={(e) => { void onFiles(e.target.files); e.target.value = ""; }} />
          <Button variant="outline" className="h-auto flex-col gap-1 border-dashed py-5" onClick={() => fileInput.current?.click()} disabled={!!loading}>
            <Upload aria-hidden="true" />
            <span>{loading ? `${loading} 읽는 중…` : "PDF 올리기"}</span>
          </Button>
          {loadError && <p role="alert" className="text-sm text-red-700 dark:text-red-400">{loadError}</p>}
          {files.length > 0 && (
            <ul className="flex flex-col gap-1">
              {files.map((f) => (
                <li key={f.id} className={cn("flex items-center gap-1 rounded-md border", f.id === activeId ? "border-primary bg-muted" : "border-transparent")}>
                  <button type="button" onClick={() => selectFile(f.id)} aria-current={f.id === activeId ? "true" : undefined}
                    className="flex min-w-0 flex-1 flex-col px-3 py-2 text-left focus-visible:outline-2 focus-visible:outline-ring">
                    <span className="truncate text-sm font-medium" title={f.name}>{f.name}</span>
                    <span className="text-xs text-muted-foreground">
                      {f.parsed.pageCount}쪽 · {DOC_KIND_LABEL[f.kind]}{Object.keys(f.tr).length ? " · 번역됨" : ""}
                    </span>
                  </button>
                  <Button variant="ghost" size="icon-sm" aria-label={`${f.name} 목록에서 빼기`} onClick={() => void removeFile(f)}>
                    <Trash2 aria-hidden="true" />
                  </Button>
                </li>
              ))}
            </ul>
          )}
        </section>

        {active && (
          <section aria-labelledby="kind-h" className="flex flex-col gap-2">
            <h2 id="kind-h" className="text-sm font-semibold">문서 유형</h2>
            <Select value={active.kind} onValueChange={(v) => void rebuild({ kind: v as DocKind })}>
              <SelectTrigger aria-label="문서 유형" className="w-full"><SelectValue /></SelectTrigger>
              <SelectContent>
                {KINDS.map((k) => (
                  <SelectItem key={k} value={k}>{DOC_KIND_LABEL[k]}{k === active.detection.kind ? " (자동 판별)" : ""}</SelectItem>
                ))}
              </SelectContent>
            </Select>
            <div className="text-xs text-muted-foreground">
              <p className="font-medium text-foreground">판별 근거 · 규칙 기반</p>
              <ul className="mt-1 list-disc pl-4">
                {active.detection.reasons.slice(0, 4).map((r) => <li key={r}>{r}</li>)}
              </ul>
            </div>
            <div className="mt-1 grid gap-1.5">
              <Label htmlFor="range">페이지 범위 <span className="font-normal text-muted-foreground">(전체 {active.parsed.pageCount}쪽)</span></Label>
              <div className="flex gap-2">
                <Input id="range" value={rangeDraft} onChange={(e) => setRangeDraft(e.target.value)} placeholder="예: 1-5, 8"
                  aria-invalid={!!rangeError} aria-describedby={rangeError ? "range-err" : undefined} />
                <Button variant="outline" onClick={() => void rebuild({ range: rangeDraft })} disabled={rangeDraft.trim() === active.range}>적용</Button>
              </div>
              {rangeError && <p id="range-err" className="text-xs text-red-700 dark:text-red-400">{rangeError}</p>}
            </div>
          </section>
        )}

        {/* translate settings */}
        <section aria-labelledby="tr-h" className="flex flex-col gap-3">
          <h2 id="tr-h" className="text-sm font-semibold">번역</h2>
          <RadioGroup value={settings.engine} onValueChange={(v) => setSettings({ engine: v as Engine })} aria-label="번역 방식" className="gap-2">
            <label htmlFor="eng-copilot" className={cn("flex cursor-pointer gap-3 rounded-md border p-3", settings.engine === "copilot" && "border-primary")}>
              <RadioGroupItem id="eng-copilot" value="copilot" className="mt-0.5" />
              <span className="flex flex-col"><span className="text-sm font-medium">Copilot 채팅 번역</span><span className="text-xs text-muted-foreground">요청문을 복사해 Microsoft 365 Copilot에 붙여 넣고 답변을 다시 붙여 넣기 · 플러그인 불필요</span></span>
            </label>
            <label htmlFor="eng-browser" className={cn("flex cursor-pointer gap-3 rounded-md border p-3", settings.engine === "browser" && "border-primary")}>
              <RadioGroupItem id="eng-browser" value="browser" className="mt-0.5" />
              <span className="flex flex-col"><span className="text-sm font-medium">브라우저 내장 번역</span><span className="text-xs text-muted-foreground">Chrome·Edge 기기 내 번역, 설정 불필요</span></span>
            </label>
          </RadioGroup>

          {settings.engine === "copilot" ? (
            <p className="text-xs text-muted-foreground">
              문서를 묶음으로 나눠 번역 요청문을 만듭니다. 회사 계정으로 로그인한 Copilot 채팅에서 번역하므로 별도 키나 플러그인이 필요 없습니다.
            </p>
          ) : (
            <p className="text-xs" role="status">
              {browser.isPending ? "브라우저 지원 여부 확인 중…"
                : browser.data === "available" ? "이 브라우저에서 바로 번역할 수 있습니다."
                : browser.data === "downloadable" || browser.data === "downloading" ? "처음 한 번 언어 모델을 내려받은 뒤 번역합니다."
                : browser.data === "need-source" ? "원문 언어를 직접 선택하세요."
                : browser.data === "unavailable" ? "이 언어 조합은 브라우저 내장 번역에서 지원되지 않습니다."
                : "이 브라우저는 내장 번역을 지원하지 않습니다. 최신 Chrome 또는 Edge를 쓰거나 Copilot 채팅 번역을 선택하세요."}
            </p>
          )}

          <div className="grid grid-cols-2 gap-2">
            <div className="grid gap-1.5">
              <Label htmlFor="from">원문 언어</Label>
              <Select value={settings.from || "auto"} onValueChange={(v) => setSettings({ from: v === "auto" ? "" : v })}>
                <SelectTrigger id="from" className="w-full"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="auto">자동 감지</SelectItem>
                  {LANGS.map((l) => <SelectItem key={l.code} value={l.code}>{l.name}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
            <div className="grid gap-1.5">
              <Label htmlFor="to">번역 언어</Label>
              <Select value={settings.to} onValueChange={(v) => setSettings({ to: v })}>
                <SelectTrigger id="to" className="w-full"><SelectValue /></SelectTrigger>
                <SelectContent>
                  {LANGS.map((l) => <SelectItem key={l.code} value={l.code}>{l.name}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
          </div>

          <div className="flex gap-2">
            <Button className="flex-1" disabled={!active || translate.isPending || !!manual || !engineReady} aria-describedby={!engineReady ? "tr-why" : undefined}
              onClick={() => active && (settings.engine === "copilot" ? void startManual(active) : translate.mutate(active))}>
              <Languages aria-hidden="true" /> {translate.isPending ? "번역 중…" : translatedCount ? "다시 번역" : "번역 시작"}
            </Button>
            {translate.isPending && (
              <Button variant="outline" onClick={() => { cancelRef.current = true; }}>
                <Square aria-hidden="true" /> 중지
              </Button>
            )}
          </div>
          {!engineReady && <p id="tr-why" className="text-xs text-muted-foreground">이 브라우저에서는 내장 번역을 쓸 수 없습니다. Copilot 채팅 번역을 선택하세요.</p>}
          {manual && (
            <div className="flex flex-col gap-2 rounded-md border border-primary p-3" aria-labelledby="manual-h">
              <p id="manual-h" className="text-sm font-semibold">
                Copilot 채팅 번역 · 묶음 {manual.idx + 1} / {manual.groups.length}
                {manualFile ? <span className="block text-xs font-normal text-muted-foreground">{manualFile.name} · 이 묶음 {manual.groups[manual.idx].length}문장</span> : null}
              </p>
              <Progress value={Math.round((manual.idx / manual.groups.length) * 100)} aria-label={`묶음 진행 ${manual.idx} / ${manual.groups.length}`} />
              <ol className="list-decimal pl-4 text-xs text-muted-foreground">
                <li>요청문 복사</li>
                <li>Copilot 채팅을 열어 붙여 넣고 보내기</li>
                <li>Copilot 답변 전체를 복사해 아래에 붙여 넣고 적용</li>
              </ol>
              <div className="flex gap-2">
                <Button size="sm" variant="outline" className="flex-1" onClick={() => void copyPrompt()}>
                  <ClipboardCopy aria-hidden="true" /> {manual.copied ? "다시 복사" : "요청문 복사"}
                </Button>
                <Button size="sm" variant="outline" className="flex-1" asChild>
                  <a href={COPILOT_URL} target="_blank" rel="noopener noreferrer">
                    <ExternalLink aria-hidden="true" /> Copilot 열기<span className="sr-only"> (새 창)</span>
                  </a>
                </Button>
              </div>
              <details className="text-xs">
                <summary className="cursor-pointer text-muted-foreground">요청문 보기</summary>
                <Textarea readOnly value={manualPrompt} aria-label="번역 요청문" className="mt-1 h-28 font-mono text-xs" onFocus={(e) => e.currentTarget.select()} />
              </details>
              <Label htmlFor="reply">Copilot 답변</Label>
              <Textarea id="reply" value={manual.reply} onChange={(e) => setManual({ ...manual, reply: e.target.value, error: null })}
                placeholder='[{"id":"...","t":"..."}] 형태의 답변을 붙여 넣으세요' className="h-28 font-mono text-xs"
                aria-invalid={!!manual.error} aria-describedby={manual.error ? "reply-err" : undefined} />
              {manual.error && <p id="reply-err" role="alert" className="text-xs text-red-700 dark:text-red-400">{manual.error}</p>}
              <div className="flex gap-2">
                <Button size="sm" className="flex-1" disabled={!manual.reply.trim()} onClick={applyReply}>답변 적용</Button>
                <Button size="sm" variant="ghost" onClick={() => { setManual(null); announce("Copilot 채팅 번역을 멈췄습니다. 적용한 묶음은 남아 있습니다."); }}>
                  <Square aria-hidden="true" /> 멈추기
                </Button>
              </div>
            </div>
          )}
          {translate.isPending && progress && (
            <div className="grid gap-1" aria-live="polite">
              <Progress value={pct} aria-label={`번역 진행 ${progress.done} / ${progress.total}문장`} />
              <span className="text-xs text-muted-foreground">{progress.done} / {progress.total}문장</span>
            </div>
          )}
          {trError && !translate.isPending && (
            <p role="alert" className="text-sm text-red-700 dark:text-red-400">{trError instanceof Error ? trError.message : "번역하지 못했습니다."}</p>
          )}
        </section>

        {/* export */}
        <section aria-labelledby="ex-h" className="flex flex-col gap-3">
          <h2 id="ex-h" className="text-sm font-semibold">내보내기</h2>
          <RadioGroup value={format} onValueChange={(v) => setFormat(v as ExportFormat)} aria-label="파일 형식" className="grid grid-cols-2 gap-2">
            {EXPORT_FORMATS.map((f) => (
              <label key={f.id} htmlFor={`fmt-${f.id}`} className={cn("flex cursor-pointer gap-2 rounded-md border p-2.5", format === f.id && "border-primary")}>
                <RadioGroupItem id={`fmt-${f.id}`} value={f.id} className="mt-0.5" />
                <span className="flex flex-col"><span className="text-sm font-medium">{f.label} <span className="font-normal text-muted-foreground">{f.ext}</span></span><span className="text-xs text-muted-foreground">{f.note}</span></span>
              </label>
            ))}
          </RadioGroup>
          <ToggleGroup type="single" variant="outline" value={exportBoth ? "both" : "translated"} onValueChange={(v) => v && setExportBoth(v === "both")} aria-label="파일에 담을 내용" className="w-full">
            <ToggleGroupItem value="translated" className="flex-1">번역만</ToggleGroupItem>
            <ToggleGroupItem value="both" className="flex-1">번역 + 원문</ToggleGroupItem>
          </ToggleGroup>
          <Button onClick={() => void doExport()} disabled={!active || !translatedCount || exporting} aria-describedby={!translatedCount ? "ex-why" : undefined}>
            <Download aria-hidden="true" /> {exporting ? "파일 만드는 중…" : "다운로드"}
          </Button>
          {!translatedCount && <p id="ex-why" className="text-xs text-muted-foreground">번역을 마치면 내려받을 수 있습니다.</p>}
          {exportError && <p role="alert" className="text-sm text-red-700 dark:text-red-400">{exportError}</p>}
        </section>
      </aside>

      {/* result */}
      <section id="result" tabIndex={-1} aria-labelledby="result-h" className="min-w-0 outline-none">
        {!active ? (
          <EmptyState icon={FileText} title="PDF를 올리면 문서 유형에 맞는 양식으로 보여 줍니다" description="문제집, 발주서, 견적서, 계약서, 일반 문서를 구분합니다." />
        ) : (
          <div className="flex flex-col gap-4">
            <div className="flex flex-wrap items-center gap-2">
              <h1 id="result-h" className="min-w-0 truncate text-lg font-semibold" title={active.name}>{baseName(active.name)}</h1>
              <Badge variant="secondary">{DOC_KIND_LABEL[active.kind]}</Badge>
              <span className="text-xs text-muted-foreground">
                {translatedCount ? `${langName(settings.to)} 번역 ${translatedCount}/${active.doc.segs.length}` : `번역 전 · 문장 ${active.doc.segs.length}개`}
                {active.range !== "전체" ? ` · ${active.range}쪽` : ""}
              </span>
              <ToggleGroup type="single" variant="outline" size="sm" value={show} onValueChange={(v) => v && setShow(v as ShowMode)} aria-label="보기" className="ml-auto">
                <ToggleGroupItem value="translated">번역</ToggleGroupItem>
                <ToggleGroupItem value="both">번역 + 원문</ToggleGroupItem>
                <ToggleGroupItem value="original">원문</ToggleGroupItem>
              </ToggleGroup>
            </div>
            {active.missing > 0 && (
              <p role="status" className="rounded-md border border-amber-700 px-3 py-2 text-sm dark:border-amber-400">
                {active.missing}개 문장은 번역되지 않아 원문으로 표시됩니다. 다시 번역하면 채울 수 있습니다.
              </p>
            )}
            <div ref={resultRef} className="rounded-lg border bg-card p-6">
              <DocView doc={active.doc} tr={active.tr} show={show} />
            </div>
          </div>
        )}
      </section>
    </div>
  );
}

/* ---------- document forms ---------- */

function T({ s, tr, show, className }: { s: Seg; tr: Record<string, string>; show: ShowMode; className?: string }) {
  const t = tr[s.id];
  const main = show === "original" || t === undefined ? s.src : t;
  const showSrc = show === "both" && t !== undefined && t !== s.src;
  return (
    <span className={className}>
      {main}
      {showSrc && <span className="mt-0.5 block text-xs text-muted-foreground">{s.src}</span>}
    </span>
  );
}

function DocView({ doc, tr, show }: { doc: StructuredDoc; tr: Record<string, string>; show: ShowMode }) {
  const p = { tr, show };
  return (
    <article className="flex flex-col gap-5 text-sm leading-relaxed">
      <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{DOC_KIND_LABEL[doc.kind]}</p>
      {doc.title && <h2 className="-mt-3 text-xl font-semibold"><T s={doc.title} {...p} /></h2>}

      {doc.fields.length > 0 && (
        <Table>
          <TableCaption className="sr-only">기본 정보</TableCaption>
          <TableHeader><TableRow><TableHead scope="col" className="w-[32%]">항목</TableHead><TableHead scope="col">내용</TableHead></TableRow></TableHeader>
          <TableBody>
            {doc.fields.map((f) => (
              <TableRow key={f.label.id}>
                <TableHead scope="row" className="whitespace-normal bg-muted/40 align-top font-medium"><T s={f.label} {...p} /></TableHead>
                <TableCell className="whitespace-normal align-top"><T s={f.value} {...p} /></TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}

      {doc.tables.map((t, i) => (
        <div key={i} className="overflow-x-auto" data-pdf-expand>
          <Table>
            <TableCaption className="text-left">{doc.kind === "order" || doc.kind === "estimate" ? "품목" : "표"} {i + 1} · {t.rows.length}행</TableCaption>
            <TableHeader>
              <TableRow>{t.header.map((c) => <TableHead key={c.id} scope="col" className="whitespace-normal align-top"><T s={c} {...p} /></TableHead>)}</TableRow>
            </TableHeader>
            <TableBody>
              {t.rows.map((r, ri) => (
                <TableRow key={ri}>{r.map((c) => <TableCell key={c.id} className="whitespace-normal align-top"><T s={c} {...p} /></TableCell>)}</TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      ))}

      {doc.clauses.length > 0 && (
        <ol className="flex flex-col gap-4">
          {doc.clauses.map((c) => (
            <li key={c.title.id} className="border-l-2 border-primary pl-4">
              <h3 className="font-semibold"><T s={c.title} {...p} /></h3>
              {c.body.map((b) => <p key={b.id} className="mt-1.5"><T s={b} {...p} /></p>)}
            </li>
          ))}
        </ol>
      )}

      {doc.qa.length > 0 && (
        <ol className="flex flex-col gap-4">
          {doc.qa.map((q, qi) => (
            <li key={qi} className="rounded-md border p-4">
              <p className="mb-1 text-xs font-semibold text-muted-foreground">{q.no}</p>
              {q.question.map((s) => <p key={s.id} className="font-medium"><T s={s} {...p} /></p>)}
              {q.choices.length > 0 && (
                <ul className="mt-3 flex flex-col gap-1.5">
                  {q.choices.map((c) => {
                    const correct = !!q.answer && q.answer.toUpperCase().replace(/[^A-H]/g, "").includes(c.label);
                    return (
                      <li key={c.seg.id} className={cn("flex gap-2 rounded-md border px-3 py-2", correct && "border-green-700 dark:border-green-400")}>
                        <span className="font-semibold">{c.label}.</span>
                        <T s={c.seg} {...p} className="min-w-0 flex-1" />
                        {correct && <Badge className="h-fit shrink-0 bg-green-100 text-green-900 dark:bg-green-950 dark:text-green-100">정답</Badge>}
                      </li>
                    );
                  })}
                </ul>
              )}
              {q.answer && <p className="mt-3 text-sm font-semibold">정답: {q.answer}</p>}
              {q.explanation.map((e) => (
                <div key={e.id} className="mt-2 rounded-md bg-muted p-3">
                  <p className="mb-1 text-xs font-semibold">해설</p>
                  <T s={e} {...p} />
                </div>
              ))}
            </li>
          ))}
        </ol>
      )}

      {doc.blocks.map((b) =>
        b.kind === "heading"
          ? <h3 key={b.seg.id} className="text-base font-semibold"><T s={b.seg} {...p} /></h3>
          : <p key={b.seg.id}><T s={b.seg} {...p} /></p>,
      )}
    </article>
  );
}
