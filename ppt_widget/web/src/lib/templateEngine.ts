// Template engine (SRS §4, rules T1–T10). Pure TypeScript: no Zoho SDK, no DOM.
//   #Field$            value from the slide's data source (T1)
//   #Lookup.field$     sub-field of a lookup/user field, e.g. #Owner.name$ (T2)
//   #Account.Field$    always the selected Account (T6)
//   #RowCountIterator … $RowCountIterator   repeat once per related record (T3)
//   #Sno$              1-based row number, continuing across overflow slides (T4)

import type {
  CrmRecord,
  DeckInput,
  DeckOutput,
  FormatContext,
  GeneratedSlide,
  GenerationWarning,
  Template,
  TemplateRow,
} from "./types";
import { fileRef } from "./imageResolver";

const PLACEHOLDER = /#([A-Za-z0-9_.]+)\$/g; // T10: anything else is left as-is
const ITERATOR = /#RowCountIterator([\s\S]*?)\$RowCountIterator/g;
const ACCOUNTS = "Accounts";

// ---------------------------------------------------------------- templates

/** Groups template rows by Template Name; slides sorted by Slide_Order, then creation time (FR-02–FR-04). */
export function groupTemplates(rows: TemplateRow[]): Template[] {
  const byName = new Map<string, TemplateRow[]>();
  for (const row of rows) {
    const name = row.Name?.trim();
    if (!name) continue;
    const list = byName.get(name) ?? [];
    list.push(row);
    byName.set(name, list);
  }
  const order = (r: TemplateRow) => (typeof r.Slide_Order === "number" ? r.Slide_Order : Number.POSITIVE_INFINITY);
  return [...byName.entries()]
    .map(([name, slides]) => ({
      name,
      slides: [...slides].sort(
        (a, b) =>
          order(a) - order(b) ||
          (a.Created_Time ?? "").localeCompare(b.Created_Time ?? "") ||
          a.id.localeCompare(b.id),
      ),
    }))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
}

/** Modules (other than Accounts) whose related records a template needs. */
export function relatedModules(template: Template): string[] {
  const set = new Set<string>();
  for (const s of template.slides) {
    if (s.Module_Name && s.Module_Name !== ACCOUNTS && s.Module_Name !== "-None-") set.add(s.Module_Name);
  }
  return [...set];
}

// ---------------------------------------------------------------- formatting

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function formatDate(y: string, m: string, d: string, pattern?: string): string {
  if (!pattern) return `${y}-${m}-${d}`;
  return pattern.replace(/yyyy|yy|MMM|MM|M|dd|d/g, (t) => {
    switch (t) {
      case "yyyy": return y;
      case "yy": return y.slice(2);
      case "MMM": return MONTHS[Number(m) - 1] ?? m;
      case "MM": return m;
      case "M": return String(Number(m));
      case "dd": return d;
      default: return String(Number(d));
    }
  });
}

/** T9: turns a CRM value into display text (not yet escaped). */
export function formatValue(value: unknown, fieldType: string | undefined, ctx: FormatContext = {}): string {
  if (value === null || value === undefined) return "";
  if (fieldType === "imageupload" || fieldType === "fileupload") return fileRef(value); // T12
  if (typeof value === "boolean") return value ? "Yes" : "No";
  if (Array.isArray(value)) {
    return value.map((v) => formatValue(v, undefined, ctx)).filter(Boolean).join(", ");
  }
  if (typeof value === "object") {
    const o = value as Record<string, unknown>;
    if ("name" in o) return formatValue(o.name, undefined, ctx);
    if ("display_value" in o) return formatValue(o.display_value, undefined, ctx);
    return "";
  }
  if (typeof value === "number") {
    if (fieldType === "currency") {
      const n = value.toLocaleString(ctx.locale, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
      return ctx.currencySymbol ? `${ctx.currencySymbol}${n}` : n;
    }
    if (fieldType === "bigint" || fieldType === "autonumber") return String(value);
    return value.toLocaleString(ctx.locale);
  }
  const s = String(value);
  const dt = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(s);
  if (dt && (fieldType === "date" || fieldType === "datetime" || s.length === 10 || /T\d{2}:\d{2}:\d{2}/.test(s))) {
    const date = formatDate(dt[1], dt[2], dt[3], ctx.dateFormat);
    return dt[4] ? `${date} ${dt[4]}:${dt[5]}` : date;
  }
  return s;
}

// ---------------------------------------------------------------- placeholders

type Lookup = { found: boolean; value?: unknown };

function lookupPath(record: CrmRecord | undefined, path: string[]): Lookup {
  let cur: unknown = record;
  for (const seg of path) {
    if (cur === null) return { found: true, value: null }; // empty lookup, e.g. no Owner
    if (typeof cur !== "object" || cur === undefined || !(seg in (cur as object))) return { found: false };
    cur = (cur as Record<string, unknown>)[seg];
  }
  return { found: true, value: cur };
}

interface FillScope {
  module: string;
  record: CrmRecord | undefined;
  account: CrmRecord;
  sno?: number;
  format: FormatContext;
  warn: (message: string) => void;
}

function resolve(name: string, scope: FillScope): string {
  if (name === "Sno") {
    if (scope.sno !== undefined) return String(scope.sno);
    scope.warn("#Sno$ is used outside a #RowCountIterator block, so it is left blank.");
    return "";
  }
  const path = name.split(".");
  const types = scope.format.fieldTypes ?? {};

  // T6: #Account.Field$ reads the selected Account.
  if (path[0] === "Account" && path.length > 1) {
    const hit = lookupPath(scope.account, path.slice(1));
    if (hit.found) return formatValue(hit.value, path.length === 2 ? types[ACCOUNTS]?.[path[1]] : undefined, scope.format);
  }

  // A related slide with no records has nothing to resolve against: blank, no warning.
  if (scope.record === undefined) return "";

  const hit = lookupPath(scope.record, path);
  const fieldType = path.length === 1 ? types[scope.module]?.[path[0]] : undefined;
  if (hit.found) return formatValue(hit.value, fieldType, scope.format);
  // Known in field metadata but absent on the record: just empty.
  if (types[scope.module]?.[path[0]]) return "";
  scope.warn(`Field "${name}" was not found in ${scope.module}, so it is left blank.`);
  return "";
}

/** Replaces every placeholder. `escape` is true for HTML, false for plain text such as titles. */
export function fillPlaceholders(text: string, scope: FillScope, escape = true): string {
  return text.replace(PLACEHOLDER, (_, name: string) => {
    const value = resolve(name, scope);
    return escape ? escapeHtml(value) : value;
  });
}

function noRecordsMarkup(block: string): string {
  if (/<tr[\s>]/i.test(block)) {
    const cells = (block.match(/<t[dh][\s>]/gi) ?? []).length || 1;
    return `<tr><td colspan="${cells}">No records</td></tr>`;
  }
  if (/<li[\s>]/i.test(block)) return "<li>No records</li>";
  return "<p>No records</p>";
}

// ---------------------------------------------------------------- slides

function renderSlide(row: TemplateRow, input: DeckInput, warnings: GenerationWarning[]): GeneratedSlide[] {
  const format = input.format ?? {};
  const sourceModule = !row.Module_Name || row.Module_Name === "-None-" ? ACCOUNTS : row.Module_Name;
  const isAccount = sourceModule === ACCOUNTS;
  const baseTitle = row.Slide_Title || row.Name;
  const seen = new Set<string>();
  const warn = (message: string) => {
    if (seen.has(message)) return;
    seen.add(message);
    warnings.push({ slide: baseTitle, message });
  };

  if (!isAccount && !(sourceModule in input.related)) warn(`No related records were loaded for ${sourceModule}.`);
  const records: CrmRecord[] = isAccount ? [input.account] : input.related[sourceModule] ?? [];
  const scope = (record: CrmRecord | undefined, sno?: number): FillScope => ({
    module: sourceModule,
    record,
    account: input.account,
    sno,
    format,
    warn,
  });

  const html = row.Slide_HTML ?? "";
  const blocks = [...html.matchAll(ITERATOR)];
  const title = (n: number, total: number) => {
    const t = fillPlaceholders(baseTitle, scope(records[0]), false);
    return total > 1 ? `${t} (${n}/${total})` : t;
  };

  if (blocks.length === 0) {
    if (/#RowCountIterator|\$RowCountIterator/.test(html)) {
      warn("A #RowCountIterator marker has no matching end marker, so it is ignored.");
    }
    return [{ title: title(1, 1), html: fillPlaceholders(html, scope(records[0])), templateRowId: row.id }];
  }

  if (blocks.length > 1) warn("Only the first #RowCountIterator block is repeated; the others are filled once (T7).");
  const first = blocks[0];
  const start = first.index ?? 0;
  const before = html.slice(0, start);
  const after = html.slice(start + first[0].length).replace(ITERATOR, (_, inner: string) => inner);
  const inner = first[1];

  const limit = row.Limit_of_Rows && row.Limit_of_Rows > 0 ? Math.floor(row.Limit_of_Rows) : records.length || 1;
  const chunks: CrmRecord[][] = [];
  for (let i = 0; i < records.length; i += limit) chunks.push(records.slice(i, i + limit));
  if (chunks.length === 0) chunks.push([]);

  // T5: static parts resolve against the first related record (or the Account).
  const head = fillPlaceholders(before, scope(records[0]));
  const tail = fillPlaceholders(after, scope(records[0]));

  return chunks.map((chunk, i) => {
    const body =
      chunk.length === 0
        ? noRecordsMarkup(inner)
        : chunk.map((rec, j) => fillPlaceholders(inner, scope(rec, i * limit + j + 1))).join("");
    return { title: title(i + 1, chunks.length), html: head + body + tail, templateRowId: row.id };
  });
}

/** Builds every slide of a template for one Account (FR-10–FR-14). */
export function buildDeck(input: DeckInput): DeckOutput {
  const warnings: GenerationWarning[] = [];
  const slides = input.template.slides.flatMap((row) => renderSlide(row, input, warnings));
  return { slides, warnings };
}

/** FR-22: "<Account Name> - <Template Name> - <YYYY-MM-DD>.pptx" with invalid characters removed. */
export function deckFileName(accountName: string, templateName: string, date = new Date()): string {
  const clean = (s: string) =>
    s.replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/\s+/g, " ").trim() || "Untitled";
  const d = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  return `${clean(accountName)} - ${clean(templateName)} - ${d}.pptx`;
}
