// A data view of a filled .pptx: every text value and table on each slide, in reading order.
// A browser cannot draw a real PowerPoint faithfully, so this lets a user check the data before exporting.

import JSZip from "jszip";

export type OutlineBlock =
  | { kind: "text"; lines: string[] }
  | { kind: "table"; rows: string[][] };

export interface SlideOutline {
  index: number;
  title: string;
  blocks: OutlineBlock[];
}

const unescapeXml = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function paragraphs(xml: string): string[] {
  return (xml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? [])
    .map((p) => unescapeXml((p.match(/<a:t>[\s\S]*?<\/a:t>/g) ?? []).map((t) => t.replace(/<\/?a:t>/g, "")).join("")).replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

/** Text boxes and tables of one slide, in document order. Pictures and decoration are not listed. */
export function outlineSlideXml(xml: string, index: number): SlideOutline {
  const blocks: OutlineBlock[] = [];
  for (const m of xml.matchAll(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>|<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const shape = m[0];
    if (shape.startsWith("<p:graphicFrame>")) {
      const rows = (shape.match(/<a:tr\b[\s\S]*?<\/a:tr>/g) ?? [])
        .map((tr) => (tr.match(/<a:tc\b[\s\S]*?<\/a:tc>/g) ?? []).map((tc) => paragraphs(tc).join(" / ")))
        .filter((r) => r.some(Boolean));
      if (rows.length) blocks.push({ kind: "table", rows });
    } else {
      const lines = paragraphs(shape);
      if (lines.length) blocks.push({ kind: "text", lines });
    }
  }
  const first = blocks.find((b): b is Extract<OutlineBlock, { kind: "text" }> => b.kind === "text");
  return { index, title: first?.lines[0] ?? `Slide ${index}`, blocks };
}

export async function outlineSlides(data: Uint8Array): Promise<SlideOutline[]> {
  const zip = await JSZip.loadAsync(data);
  const names = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]));
  return Promise.all(names.map(async (n, i) => outlineSlideXml(await zip.file(n)!.async("string"), i + 1)));
}
