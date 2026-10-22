// Converts generated slide HTML into native, editable PowerPoint objects (FR-19–FR-21).
//   h1–h6, p, div text, blockquote  -> text box
//   ul / ol                          -> text box with bullets / numbers
//   table                            -> native table (th = bold, colspan/rowspan kept)
//   img                              -> image
// Inline styles kept: bold, italic, underline, strike, colour, background (highlight / cell fill),
// text-align, font-size, links. Anything else becomes plain text (FR-21).
// Blocks are stacked top to bottom; the preview (slidePreview.ts) uses the same sizes.

import type PptxGenJS from "pptxgenjs";
import type { GeneratedSlide, GenerationWarning } from "./types";

// 16:9 "LAYOUT_WIDE" is 13.333 x 7.5 in. 1 in = 96 preview px.
export const SLIDE_W = 13.333;
export const SLIDE_H = 7.5;
export const MARGIN = 0.5;
export const FOOTER_H = 0.4;
export const CONTENT_W = SLIDE_W - 2 * MARGIN;
export const FONT_FACE = "Arial";
export const BASE_PT = 16;
export const HEADING_PT: Record<string, number> = { h1: 32, h2: 26, h3: 22, h4: 18, h5: 16, h6: 14 };
const GAP = 0.12;
const BOTTOM = SLIDE_H - MARGIN - FOOTER_H + 0.2;

/** Table font size so `rows` rows fit in `availableIn` inches (same rule as the preview). */
export function tableFontPt(rows: number, availableIn: number): number {
  const pt = Math.floor((availableIn * 72) / (Math.max(rows, 1) * 1.75));
  return Math.max(8, Math.min(14, pt));
}

// ---------------------------------------------------------------- styles

type TextOpts = PptxGenJS.TextPropsOptions;
type Run = { text: string; options: TextOpts };

interface InlineStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  fontSize?: number;
  link?: string;
  superscript?: boolean;
  subscript?: boolean;
}

const NAMED: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  orange: "FFA500", purple: "800080", gray: "808080", grey: "808080", navy: "000080", teal: "008080",
  maroon: "800000", silver: "C0C0C0", lime: "00FF00", aqua: "00FFFF", fuchsia: "FF00FF", olive: "808000",
};

/** CSS colour -> "RRGGBB", or undefined. */
export function parseColor(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();
  if (v === "transparent" || v === "inherit" || v === "initial") return undefined;
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return m[1].split("").map((c) => c + c).join("").toUpperCase();
  m = /^#([0-9a-f]{6})/.exec(v);
  if (m) return m[1].toUpperCase();
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(v);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return undefined;
    return [m[1], m[2], m[3]].map((n) => Math.min(255, Number(n)).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  return NAMED[v];
}

function parseFontSize(value: string, basePt: number): number | undefined {
  const m = /^([\d.]+)(px|pt|em|rem|%)?$/.exec(value.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case "pt": return n;
    case "em": case "rem": return n * basePt;
    case "%": return (n / 100) * basePt;
    default: return n * 0.75; // px
  }
}

function backgroundOf(el: Element): string | undefined {
  const s = (el as HTMLElement).style;
  return parseColor(s?.backgroundColor) ?? parseColor(el.getAttribute("bgcolor")) ??
    parseColor(/(?:^|;)\s*background\s*:\s*([^;]+)/i.exec(el.getAttribute("style") ?? "")?.[1]?.split(/\s+/)[0]);
}

function alignOf(el: Element): PptxGenJS.HAlign | undefined {
  const a = ((el as HTMLElement).style?.textAlign || el.getAttribute("align") || "").toLowerCase();
  if (a === "center" || a === "right" || a === "left" || a === "justify") return a;
  return undefined;
}

function inlineStyleOf(el: Element, parent: InlineStyle, basePt: number): InlineStyle {
  const tag = el.tagName.toLowerCase();
  const s: InlineStyle = { ...parent };
  if (tag === "b" || tag === "strong" || tag === "th") s.bold = true;
  if (tag === "i" || tag === "em") s.italic = true;
  if (tag === "u" || tag === "ins") s.underline = true;
  if (tag === "s" || tag === "strike" || tag === "del") s.strike = true;
  if (tag === "sup") s.superscript = true;
  if (tag === "sub") s.subscript = true;
  if (tag === "a" && el.getAttribute("href")) s.link = el.getAttribute("href") ?? undefined;
  if (tag === "font") s.color = parseColor(el.getAttribute("color")) ?? s.color;
  const st = (el as HTMLElement).style;
  if (st) {
    if (st.fontWeight === "bold" || Number(st.fontWeight) >= 600) s.bold = true;
    if (st.fontWeight === "normal" || (Number(st.fontWeight) > 0 && Number(st.fontWeight) < 600)) s.bold = false;
    if (st.fontStyle === "italic") s.italic = true;
    if (st.textDecoration?.includes("underline")) s.underline = true;
    if (st.textDecoration?.includes("line-through")) s.strike = true;
    s.color = parseColor(st.color) ?? s.color;
    if (st.fontSize) s.fontSize = parseFontSize(st.fontSize, s.fontSize ?? basePt) ?? s.fontSize;
    const bg = tag === "td" || tag === "th" || tag === "tr" ? undefined : parseColor(st.backgroundColor);
    if (bg) s.highlight = bg;
  }
  return s;
}

function runOptions(s: InlineStyle, basePt: number): TextOpts {
  const o: TextOpts = { fontFace: FONT_FACE, fontSize: Math.round(s.fontSize ?? basePt) };
  if (s.bold) o.bold = true;
  if (s.italic) o.italic = true;
  if (s.underline) o.underline = { style: "sng" };
  if (s.strike) o.strike = "sngStrike";
  if (s.color) o.color = s.color;
  if (s.highlight) o.highlight = s.highlight;
  if (s.superscript) o.superscript = true;
  if (s.subscript) o.subscript = true;
  if (s.link) o.hyperlink = { url: s.link };
  return o;
}

// ---------------------------------------------------------------- inline runs

const BLOCK_TAGS = new Set([
  "address", "article", "aside", "blockquote", "div", "dl", "dt", "dd", "fieldset", "figure", "figcaption", "footer",
  "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav", "ol", "p", "pre", "section",
  "table", "ul", "img", "center",
]);
const SKIP_TAGS = new Set(["script", "style", "head", "title", "meta", "link", "noscript", "template", "iframe", "object", "svg"]);

/** Collects text runs; nested blocks become line breaks. */
function collectRuns(node: Node, style: InlineStyle, basePt: number, out: Run[]) {
  if (node.nodeType === 3) {
    const text = (node.textContent ?? "").replace(/\s+/g, " ");
    if (text) out.push({ text, options: runOptions(style, basePt) });
    return;
  }
  if (node.nodeType !== 1) return;
  const el = node as Element;
  const tag = el.tagName.toLowerCase();
  if (SKIP_TAGS.has(tag)) return;
  if (tag === "br") {
    lineBreak(out, style, basePt);
    return;
  }
  if (tag === "img") {
    const alt = el.getAttribute("alt");
    if (alt) out.push({ text: alt, options: runOptions(style, basePt) });
    return;
  }
  const isBlock = BLOCK_TAGS.has(tag);
  if (isBlock) lineBreak(out, style, basePt, true);
  const s = inlineStyleOf(el, style, basePt);
  const prefix = tag === "li" ? "• " : "";
  if (prefix) out.push({ text: prefix, options: runOptions(s, basePt) });
  el.childNodes.forEach((c) => collectRuns(c, s, basePt, out));
  if (isBlock) lineBreak(out, style, basePt, true);
}

function lineBreak(out: Run[], style: InlineStyle, basePt: number, soft = false) {
  const last = out[out.length - 1];
  if (soft && (!last || last.options.breakLine)) return; // no empty lines between blocks
  if (last) last.options = { ...last.options, breakLine: true };
  else out.push({ text: "", options: { ...runOptions(style, basePt), breakLine: true } });
}

/** Trims spaces around line breaks and drops a trailing break. */
function tidy(runs: Run[]): Run[] {
  const res = runs
    .filter((r) => r.text !== "" || r.options.breakLine)
    .map((r) => ({ text: r.text, options: { ...r.options } }));
  let lineStart = true;
  for (let i = 0; i < res.length; i++) {
    const r = res[i];
    if (lineStart) r.text = r.text.replace(/^ +/, "");
    if (r.options.breakLine || i === res.length - 1) r.text = r.text.replace(/ +$/, "");
    lineStart = !!r.options.breakLine || (lineStart && r.text === "");
  }
  while (res.length && res[res.length - 1].text === "" && res[res.length - 1].options.breakLine) res.pop();
  if (res.length) {
    const last = res[res.length - 1];
    last.options = { ...last.options, breakLine: false };
  }
  return res;
}

function runsText(runs: Run[]): string[] {
  const lines: string[] = [""];
  for (const r of runs) {
    lines[lines.length - 1] += r.text;
    if (r.options.breakLine) lines.push("");
  }
  return lines;
}

function textHeight(runs: Run[], widthIn: number, pt: number): number {
  const charsPerLine = Math.max(10, (widthIn * 72) / (pt * 0.52));
  const lines = runsText(runs).reduce((n, l) => n + Math.max(1, Math.ceil(l.length / charsPerLine)), 0);
  return (lines * pt * 1.25) / 72 + 0.12;
}

// ---------------------------------------------------------------- slide builder

interface Ctx {
  pptx: PptxGenJS;
  slide: PptxGenJS.Slide;
  y: number;
  warn: (m: string) => void;
  inherited: InlineStyle;
  align?: PptxGenJS.HAlign;
}

function addTextBlock(ctx: Ctx, runs: Run[], pt: number, opts: Partial<TextOpts> = {}) {
  const clean = tidy(runs);
  if (clean.length === 0 || clean.every((r) => !r.text.trim())) return;
  const h = textHeight(clean, CONTENT_W, pt);
  ctx.slide.addText(clean, { x: MARGIN, y: ctx.y, w: CONTENT_W, h, margin: [0, 0, 0, 0], valign: "top", align: ctx.align, ...opts });
  ctx.y += h + GAP;
}

function addParagraph(ctx: Ctx, el: Element, pt: number, extra: InlineStyle = {}) {
  const style = inlineStyleOf(el, { ...ctx.inherited, ...extra, fontSize: extra.fontSize ?? ctx.inherited.fontSize }, pt);
  const runs: Run[] = [];
  el.childNodes.forEach((c) => collectRuns(c, style, pt, runs));
  const fill = backgroundOf(el);
  addTextBlock(ctx, runs, style.fontSize ?? pt, { align: alignOf(el) ?? ctx.align, ...(fill ? { fill: { color: fill } } : {}) });
}

function addList(ctx: Ctx, el: Element) {
  const ordered = el.tagName.toLowerCase() === "ol";
  const style = inlineStyleOf(el, ctx.inherited, BASE_PT);
  const pt = style.fontSize ?? BASE_PT;
  const runs: Run[] = [];
  const items = [...el.children].filter((c) => c.tagName.toLowerCase() === "li");
  items.forEach((li) => {
    const itemRuns: Run[] = [];
    const s = inlineStyleOf(li, style, pt);
    li.childNodes.forEach((c) => collectRuns(c, s, pt, itemRuns));
    const clean = tidy(itemRuns);
    if (clean.length === 0) clean.push({ text: " ", options: runOptions(s, pt) });
    // Only the first run of an item carries the bullet: PowerPoint starts a new paragraph at each bulleted run.
    clean[0].options = { ...clean[0].options, bullet: ordered ? { type: "number" } : true };
    clean[clean.length - 1].options.breakLine = true;
    runs.push(...clean);
  });
  if (runs.length === 0) return;
  runs[runs.length - 1].options.breakLine = false;
  const h = textHeight(runs, CONTENT_W - 0.4, pt) + items.length * 0.02;
  ctx.slide.addText(runs, { x: MARGIN, y: ctx.y, w: CONTENT_W, h, margin: [0, 0, 0, 0], valign: "top", align: alignOf(el) ?? ctx.align });
  ctx.y += h + GAP;
}

function addTable(ctx: Ctx, table: HTMLTableElement) {
  const trs = [...table.rows];
  if (trs.length === 0) return;
  const pt = tableFontPt(trs.length, BOTTOM - ctx.y);
  const bordered = true;
  const border: PptxGenJS.BorderProps = { type: "solid", pt: 0.75, color: "BFC5CC" };
  const tableBg = backgroundOf(table);
  const tableStyle = inlineStyleOf(table, ctx.inherited, pt);
  let cols = 0;
  const rows: PptxGenJS.TableRow[] = trs.map((tr) => {
    const rowBg = backgroundOf(tr);
    const rowStyle = inlineStyleOf(tr, tableStyle, pt);
    let span = 0;
    const cells = [...tr.cells].map((cell) => {
      const s = inlineStyleOf(cell, rowStyle, pt);
      const runs: Run[] = [];
      cell.childNodes.forEach((c) => collectRuns(c, s, pt, runs));
      const clean = tidy(runs);
      const fill = backgroundOf(cell) ?? rowBg ?? tableBg;
      const colspan = Math.max(1, Number(cell.getAttribute("colspan")) || 1);
      const rowspan = Math.max(1, Number(cell.getAttribute("rowspan")) || 1);
      span += colspan;
      const options: PptxGenJS.TableCellProps = {
        fontFace: FONT_FACE,
        fontSize: pt,
        valign: "middle",
        align: alignOf(cell) ?? alignOf(tr) ?? (cell.tagName.toLowerCase() === "th" ? "left" : undefined),
        bold: s.bold,
        color: s.color,
        margin: [2, 4, 2, 4],
        ...(fill ? { fill: { color: fill } } : {}),
        ...(bordered ? { border: [border, border, border, border] } : {}),
        ...(colspan > 1 ? { colspan } : {}),
        ...(rowspan > 1 ? { rowspan } : {}),
      };
      return { text: clean.length ? (clean as PptxGenJS.TableCell[]) : "", options } as PptxGenJS.TableCell;
    });
    cols = Math.max(cols, span);
    return cells;
  });
  const rowH = (pt * 1.75) / 72;
  ctx.slide.addTable(rows, { x: MARGIN, y: ctx.y, w: CONTENT_W, colW: columnWidths(trs, cols || 1), rowH, autoPage: false });
  ctx.y += rows.length * rowH + GAP;
}

/** Column widths weighted by the longest text in each column, like a browser's auto table layout. */
function columnWidths(trs: HTMLTableRowElement[], cols: number): number[] {
  const weight = Array(cols).fill(4);
  for (const tr of trs) {
    let c = 0;
    for (const cell of [...tr.cells]) {
      const span = Math.max(1, Number(cell.getAttribute("colspan")) || 1);
      if (span === 1 && c < cols) {
        const len = (cell.textContent ?? "").replace(/\s+/g, " ").trim().length;
        weight[c] = Math.max(weight[c], Math.min(len, 40));
      }
      c += span;
    }
  }
  const total = weight.reduce((a, b) => a + b, 0);
  return weight.map((w) => (CONTENT_W * w) / total);
}

async function toDataUrl(src: string): Promise<string> {
  if (src.startsWith("data:")) return src;
  const res = await fetch(src);
  if (!res.ok) throw new Error(String(res.status));
  const blob = await res.blob();
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

function imageSize(dataUrl: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 400, h: img.naturalHeight || 300 });
    img.onerror = () => resolve({ w: 400, h: 300 });
    img.src = dataUrl;
  });
}

async function addImage(ctx: Ctx, img: HTMLImageElement) {
  const src = img.getAttribute("src");
  if (!src) return;
  try {
    const data = await toDataUrl(src);
    const natural = await imageSize(data);
    let w = (Number(img.getAttribute("width")) || natural.w) / 96;
    let h = (Number(img.getAttribute("height")) || natural.h * (w * 96) / natural.w) / 96;
    const maxH = Math.max(0.5, BOTTOM - ctx.y);
    const scale = Math.min(1, CONTENT_W / w, maxH / h);
    w *= scale;
    h *= scale;
    const align = alignOf(img.parentElement ?? img) ?? ctx.align;
    const x = align === "center" ? MARGIN + (CONTENT_W - w) / 2 : align === "right" ? MARGIN + CONTENT_W - w : MARGIN;
    ctx.slide.addImage({ data, x, y: ctx.y, w, h, altText: img.getAttribute("alt") ?? undefined });
    ctx.y += h + GAP;
  } catch {
    ctx.warn(`An image could not be loaded (${src.slice(0, 60)}), so it was left out of the PowerPoint.`);
  }
}

const HEADINGS = new Set(["h1", "h2", "h3", "h4", "h5", "h6"]);
const TEXT_BLOCKS = new Set(["p", "blockquote", "pre", "address", "figcaption", "dt", "dd", "center"]);

async function addBlocks(ctx: Ctx, parent: Element) {
  let pending: Node[] = [];
  const flush = () => {
    if (pending.length === 0) return;
    const runs: Run[] = [];
    pending.forEach((n) => collectRuns(n, ctx.inherited, BASE_PT, runs));
    addTextBlock(ctx, runs, ctx.inherited.fontSize ?? BASE_PT);
    pending = [];
  };

  for (const node of [...parent.childNodes]) {
    if (node.nodeType !== 1) {
      if (node.nodeType === 3) pending.push(node);
      continue;
    }
    const el = node as Element;
    const tag = el.tagName.toLowerCase();
    if (SKIP_TAGS.has(tag)) continue;
    if (!BLOCK_TAGS.has(tag)) {
      // Inline element: if it hides block content, treat it as a container.
      if (el.querySelector("table,ul,ol,p,div,h1,h2,h3,h4,h5,h6,img")) {
        flush();
        await addContainer(ctx, el);
      } else pending.push(el);
      continue;
    }
    flush();
    if (HEADINGS.has(tag)) {
      const pt = HEADING_PT[tag];
      addParagraph(ctx, el, pt, { bold: true, fontSize: pt });
    } else if (TEXT_BLOCKS.has(tag)) {
      addParagraph(ctx, el, ctx.inherited.fontSize ?? BASE_PT, tag === "blockquote" ? { italic: true } : {});
    } else if (tag === "ul" || tag === "ol") {
      addList(ctx, el);
    } else if (tag === "table") {
      addTable(ctx, el as HTMLTableElement);
    } else if (tag === "img") {
      await addImage(ctx, el as HTMLImageElement);
    } else if (tag === "hr") {
      ctx.slide.addShape(ctx.pptx.ShapeType.line, { x: MARGIN, y: ctx.y, w: CONTENT_W, h: 0, line: { color: "BFC5CC", width: 1 } });
      ctx.y += GAP;
    } else if (tag === "li") {
      addParagraph(ctx, el, ctx.inherited.fontSize ?? BASE_PT);
    } else {
      // div, section, … : recurse if it has blocks, else it is a paragraph.
      const hasBlocks = [...el.children].some((c) => BLOCK_TAGS.has(c.tagName.toLowerCase()));
      if (hasBlocks) await addContainer(ctx, el);
      else addParagraph(ctx, el, ctx.inherited.fontSize ?? BASE_PT);
    }
  }
  flush();
}

async function addContainer(ctx: Ctx, el: Element) {
  const saved = { inherited: ctx.inherited, align: ctx.align };
  ctx.inherited = inlineStyleOf(el, ctx.inherited, BASE_PT);
  ctx.align = alignOf(el) ?? ctx.align;
  await addBlocks(ctx, el);
  ctx.inherited = saved.inherited;
  ctx.align = saved.align;
}

/** Adds one generated slide to the presentation. */
export async function addHtmlSlide(
  pptx: PptxGenJS,
  slideData: GeneratedSlide,
  index: number,
  total: number,
  warnings: GenerationWarning[],
) {
  const slide = pptx.addSlide();
  const doc = new DOMParser().parseFromString(`<!doctype html><html><body>${slideData.html}</body></html>`, "text/html");
  const body = doc.body;

  // A single wrapping element with a background colours the whole slide.
  const only = body.children.length === 1 ? body.children[0] : null;
  const bg = only ? backgroundOf(only) : undefined;
  if (bg) slide.background = { color: bg };

  const seen = new Set<string>();
  const ctx: Ctx = {
    pptx,
    slide,
    y: MARGIN,
    inherited: {},
    warn: (message) => {
      if (seen.has(message)) return;
      seen.add(message);
      warnings.push({ slide: slideData.title, message });
    },
  };
  await addBlocks(ctx, body);
  if (ctx.y > BOTTOM + 0.3) ctx.warn("The content is taller than the slide; some of it may be cut off. Lower Limit of Rows or shorten the text.");

  const footer: TextOpts = { fontFace: FONT_FACE, fontSize: 10, color: "7B8794", margin: [0, 0, 0, 0], y: SLIDE_H - MARGIN - 0.1, h: 0.3, valign: "bottom" };
  slide.addText(slideData.title, { ...footer, x: MARGIN, w: CONTENT_W * 0.8 });
  slide.addText(`${index + 1} / ${total}`, { ...footer, x: MARGIN + CONTENT_W * 0.8, w: CONTENT_W * 0.2, align: "right" });
}

/** Builds the whole .pptx. PptxGenJS is loaded only when needed. */
export async function buildPptx(slides: GeneratedSlide[], meta: { title: string; author?: string }) {
  const { default: PptxGen } = await import("pptxgenjs");
  const pptx = new PptxGen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = meta.title;
  if (meta.author) pptx.author = meta.author;
  pptx.company = "Zoho CRM Presentation Builder";
  const warnings: GenerationWarning[] = [];
  for (let i = 0; i < slides.length; i++) await addHtmlSlide(pptx, slides[i], i, slides.length, warnings);
  return { pptx, warnings };
}
