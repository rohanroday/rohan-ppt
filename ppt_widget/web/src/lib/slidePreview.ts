// Builds the srcdoc for one preview iframe. Sizes mirror htmlToPptx.ts
// (1280x720 px = 13.333x7.5 in at 96 px/in; 1 pt = 4/3 px) so the preview matches the export.

import { BASE_PT, FONT_FACE, FOOTER_H, HEADING_PT, MARGIN, SLIDE_H, tableFontPt } from "./htmlToPptx";
import type { GeneratedSlide } from "./types";

export const PREVIEW_W = 1280;
export const PREVIEW_H = 720;
const px = (pt: number) => `${((pt * 4) / 3).toFixed(2)}px`;
const IN = 96;

function maxTableRows(html: string): number {
  const tables = html.split(/<table[\s>]/i).slice(1);
  return tables.reduce((m, t) => Math.max(m, (t.split(/<\/table>/i)[0].match(/<tr[\s>]/gi) ?? []).length), 0);
}

export function previewDoc(slide: GeneratedSlide, index: number, total: number): string {
  const rows = maxTableRows(slide.html);
  // Tables usually start below a heading, about 1.3 in from the top.
  const tablePt = rows ? tableFontPt(rows, SLIDE_H - MARGIN - FOOTER_H + 0.2 - 1.3) : 12;
  const headings = Object.entries(HEADING_PT)
    .map(([tag, pt]) => `${tag}{font-size:${px(pt)};margin:0 0 ${0.12 * IN}px;font-weight:bold;line-height:1.2}`)
    .join("");
  const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  return `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{margin:0;width:${PREVIEW_W}px;height:${PREVIEW_H}px;overflow:hidden;background:#fff}
body{box-sizing:border-box;position:relative;padding:${MARGIN * IN}px;font-family:${FONT_FACE},Helvetica,sans-serif;font-size:${px(BASE_PT)};line-height:1.25;color:#1f2933}
${headings}
p,ul,ol,blockquote,pre{margin:0 0 ${0.12 * IN}px}
ul,ol{padding-left:${0.4 * IN}px}
table{border-collapse:collapse;width:100%;font-size:${px(tablePt)};margin:0 0 ${0.12 * IN}px}
td,th{border:1px solid #bfc5cc;padding:2px 6px;text-align:left;vertical-align:middle;height:${((tablePt * 1.75) / 72) * IN - 6}px}
img{max-width:100%}
.pb-footer{position:absolute;left:${MARGIN * IN}px;right:${MARGIN * IN}px;bottom:${(MARGIN - 0.1) * IN}px;display:flex;justify-content:space-between;font-size:${px(10)};color:#7b8794;font-weight:normal}
</style></head><body>${slide.html}<div class="pb-footer"><span>${escape(slide.title)}</span><span>${index + 1} / ${total}</span></div></body></html>`;
}
