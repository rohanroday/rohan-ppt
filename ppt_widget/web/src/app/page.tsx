"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { AccountSearch } from "@/components/AccountSearch";
import { SlidePreview } from "@/components/SlidePreview";
import { WarningsPanel } from "@/components/WarningsPanel";
import { buildPptx } from "@/lib/htmlToPptx";
import { resolveImages } from "@/lib/imageResolver";
import { buildDeck, deckFileName, groupTemplates, relatedModules } from "@/lib/templateEngine";
import type { CrmRecord, GeneratedSlide, GenerationWarning, Template } from "@/lib/types";
import { connect, describeError, type AccountSummary, type DataSource } from "@/lib/zoho";

type Notice = { kind: "error" | "success" | "info"; text: string } | null;

interface Deck {
  slides: GeneratedSlide[];
  warnings: GenerationWarning[];
  account: AccountSummary;
  templateName: string;
}

export default function Page() {
  const [source, setSource] = useState<DataSource | null>(null);
  const [phase, setPhase] = useState<"loading" | "ready" | "failed">("loading");
  const [userName, setUserName] = useState("");
  const [templates, setTemplates] = useState<Template[]>([]);
  const [account, setAccount] = useState<AccountSummary | null>(null);
  const [templateName, setTemplateName] = useState("");
  const [busy, setBusy] = useState<"" | "generate" | "download" | "attach">("");
  const [deck, setDeck] = useState<Deck | null>(null);
  const [index, setIndex] = useState(0);
  const [notice, setNotice] = useState<Notice>(null);
  const [exportWarnings, setExportWarnings] = useState<GenerationWarning[]>([]);

  // FR-01/FR-02: initialise the SDK, then load templates.
  useEffect(() => {
    (async () => {
      let src: DataSource;
      try {
        const conn = await connect();
        src = conn.source;
        setSource(src);
        setUserName(conn.userName);
        if (conn.fallbackReason) setNotice({ kind: "info", text: conn.fallbackReason });
      } catch (e) {
        setPhase("failed");
        setNotice({ kind: "error", text: `${describeError(e, "connect to Zoho CRM")} Reload the tab to retry.` });
        return;
      }
      try {
        const rows = await src.listTemplates();
        const grouped = groupTemplates(rows);
        setTemplates(grouped);
        if (grouped.length === 1) setTemplateName(grouped[0].name);
      } catch (e) {
        setNotice({ kind: "error", text: describeError(e, "load Presentation Templates") });
      }
      setPhase("ready");
    })();
  }, []);

  const template = useMemo(() => templates.find((t) => t.name === templateName) ?? null, [templates, templateName]);
  const canGenerate = phase === "ready" && !!account && !!template && !busy; // FR-09

  const generate = useCallback(async () => {
    if (!source || !account || !template) return;
    setBusy("generate");
    setNotice(null);
    setExportWarnings([]);
    try {
      const modules = relatedModules(template);
      const failures: string[] = [];
      const [record, format, ...relatedLists] = await Promise.all([
        source.getAccount(account.id),
        source.getFormatContext(modules),
        ...modules.map((m) =>
          source.getRelated(account.id, m).then(
            (rows) => [m, rows] as const,
            (e) => {
              failures.push(describeError(e, `load ${m.replace(/_/g, " ")}`));
              return null;
            },
          ),
        ),
      ]);
      const related: Record<string, CrmRecord[]> = {};
      for (const entry of relatedLists) if (entry) related[entry[0]] = entry[1];
      const built = buildDeck({ template, account: record, related, format });
      const pictures = await resolveImages(built.slides, source.getImage); // FR-25
      const out = { slides: pictures.slides, warnings: [...built.warnings, ...pictures.warnings] };
      setDeck({ ...out, account: { ...account, name: String(record.Account_Name ?? account.name) }, templateName: template.name });
      setIndex(0);
      if (failures.length) setNotice({ kind: "error", text: `${failures.join(" ")} The other slides were generated.` });
    } catch (e) {
      setNotice({ kind: "error", text: describeError(e, "generate the slides") }); // FR-15
    } finally {
      setBusy("");
    }
  }, [source, account, template]);

  const exportPptx = useCallback(async () => {
    if (!deck) throw new Error("Nothing to export");
    const { pptx, warnings } = await buildPptx(deck.slides, { title: `${deck.account.name} - ${deck.templateName}`, author: userName || undefined });
    setExportWarnings(warnings);
    return { pptx, fileName: deckFileName(deck.account.name, deck.templateName) };
  }, [deck, userName]);

  const download = useCallback(async () => {
    setBusy("download");
    setNotice(null);
    try {
      const { pptx, fileName } = await exportPptx();
      await pptx.writeFile({ fileName, compression: true });
      setNotice({ kind: "info", text: `Downloading “${fileName}”. If nothing downloads, CRM may be blocking downloads in widgets; use Save to Account instead.` });
    } catch (e) {
      setNotice({ kind: "error", text: `${describeError(e, "create the PowerPoint file")} Try Save to Account instead.` }); // FR-24
    } finally {
      setBusy("");
    }
  }, [exportPptx]);

  const attach = useCallback(async () => {
    if (!source || !deck) return;
    setBusy("attach");
    setNotice(null);
    try {
      const { pptx, fileName } = await exportPptx();
      const blob = (await pptx.write({ outputType: "blob", compression: true })) as Blob;
      await source.attachFile(deck.account.id, fileName, blob);
      setNotice({
        kind: "success",
        text: source.mode === "mock"
          ? `Mock mode: “${fileName}” would now be attached to ${deck.account.name}.`
          : `Saved “${fileName}” to ${deck.account.name}’s Attachments.`,
      });
    } catch (e) {
      setNotice({ kind: "error", text: describeError(e, "attach the file to the Account") });
    } finally {
      setBusy("");
    }
  }, [source, deck, exportPptx]);

  const noTemplates = phase === "ready" && templates.length === 0;
  const stale = deck && (deck.account.id !== account?.id || deck.templateName !== templateName);
  const allWarnings = deck ? [...deck.warnings, ...exportWarnings] : [];

  return (
    <div className="mx-auto flex w-full max-w-6xl flex-1 flex-col gap-4 p-4 md:p-6">
      {/* Header */}
      <header className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-slate-900">Presentation Builder</h1>
          <p className="text-sm text-slate-500">Pick an Account and a template, preview the slides, then download an editable PowerPoint.</p>
        </div>
        <div className="text-right text-xs text-slate-500">
          {source?.mode === "mock" && <span className="mr-2 rounded bg-amber-100 px-1.5 py-0.5 font-medium text-amber-800">Mock data</span>}
          {userName}
        </div>
      </header>

      {/* Selection bar */}
      <section className="grid grid-cols-1 items-end gap-3 rounded-lg border border-slate-200 bg-white p-4 shadow-sm md:grid-cols-[1fr_220px_auto]">
        <div>
          <label className="mb-1 block text-xs font-medium uppercase tracking-wide text-slate-500">Account</label>
          {source && <AccountSearch source={source} value={account} onChange={setAccount} disabled={phase !== "ready" || busy === "generate"} />}
        </div>
        <div>
          <label htmlFor="template" className="mb-1 block text-xs font-medium uppercase tracking-wide text-slate-500">Template</label>
          <select
            id="template"
            value={templateName}
            onChange={(e) => setTemplateName(e.target.value)}
            disabled={phase !== "ready" || noTemplates || busy === "generate"}
            className="h-9 w-full rounded-md border border-slate-300 bg-white px-2 text-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100 disabled:bg-slate-100"
          >
            <option value="">{phase === "loading" ? "Loading…" : "Choose a template"}</option>
            {templates.map((t) => (
              <option key={t.name} value={t.name}>
                {t.name} ({t.slides.length} slide{t.slides.length === 1 ? "" : "s"})
              </option>
            ))}
          </select>
        </div>
        <button
          type="button"
          onClick={generate}
          disabled={!canGenerate}
          className="h-9 rounded-md bg-blue-600 px-5 text-sm font-medium text-white shadow-sm hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-slate-300"
        >
          {busy === "generate" ? "Generating…" : deck && !stale ? "Regenerate" : "Generate"}
        </button>
      </section>

      {/* Status and warnings */}
      {phase === "loading" && <Status kind="info" text="Connecting to Zoho CRM…" />}
      {noTemplates && <Status kind="info" text="No templates found — add records to Presentation Templates." />}
      {notice && <Status kind={notice.kind} text={notice.text} onClose={() => setNotice(null)} />}
      {stale && <Status kind="info" text="The selection changed. Click Generate to update the preview." />}
      {deck && <WarningsPanel warnings={allWarnings} />}

      {/* Preview */}
      {deck ? (
        <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm">
          <div className="mb-3 flex items-baseline justify-between">
            <h2 className="text-sm font-medium text-slate-700">
              {deck.account.name} · {deck.templateName} · {deck.slides.length} slide{deck.slides.length === 1 ? "" : "s"}
            </h2>
            <span className="text-xs text-slate-400">← → keys change slides</span>
          </div>
          <SlidePreview slides={deck.slides} index={index} onIndex={setIndex} />

          {/* Action bar */}
          <div className="mt-2 flex flex-wrap items-center justify-end gap-2 border-t border-slate-100 pt-3">
            <button type="button" onClick={attach} disabled={!!busy}
              className="h-9 rounded-md border border-slate-300 bg-white px-4 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:opacity-50">
              {busy === "attach" ? "Saving…" : "Save to Account"}
            </button>
            <button type="button" onClick={download} disabled={!!busy}
              className="h-9 rounded-md bg-blue-600 px-4 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50">
              {busy === "download" ? "Preparing…" : "Download .pptx"}
            </button>
          </div>
        </section>
      ) : (
        phase === "ready" && !noTemplates && (
          <section className="flex flex-1 items-center justify-center rounded-lg border border-dashed border-slate-300 bg-white/60 p-10 text-center text-sm text-slate-500">
            {busy === "generate" ? "Fetching CRM data and building slides…" : "Choose an Account and a template, then click Generate to preview the slides."}
          </section>
        )
      )}
    </div>
  );
}

function Status({ kind, text, onClose }: { kind: "error" | "success" | "info"; text: string; onClose?: () => void }) {
  const styles = {
    error: "border-red-300 bg-red-50 text-red-800",
    success: "border-green-300 bg-green-50 text-green-800",
    info: "border-slate-200 bg-slate-50 text-slate-700",
  }[kind];
  return (
    <div role={kind === "error" ? "alert" : "status"} className={`flex items-start justify-between gap-3 rounded-md border px-3 py-2 text-sm ${styles}`}>
      <span>{text}</span>
      {onClose && (
        <button type="button" onClick={onClose} className="shrink-0 opacity-60 hover:opacity-100" aria-label="Dismiss">✕</button>
      )}
    </div>
  );
}
