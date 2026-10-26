// Fills a real .pptx template without touching its design (pure client side, no network).
//   - {{token}} text runs are replaced with escaped values
//   - the first table on a slide is rebuilt from rows (cloning the template's data rows)
//   - named pictures are swapped for CRM pictures, fitted inside the original frame
//   - slides are cloned per spec, so one template slide can appear many times, in any order
// The template (e.g. Anarock's deck) is never stored in this repo: it is downloaded from CRM at run time.

import JSZip from "jszip";

export interface SlideSpec {
  /** Which template slide to clone: one of the `roles` passed to fillTemplate. */
  role: string;
  /** Values for {{tokens}} on the slide. Missing tokens become empty text. */
  tokens?: Record<string, string>;
  /** Rows for the first table (columns after the numbering column). The row number is added. */
  table?: string[][];
  /** Picture name on the template slide -> data URL (png/jpeg). */
  images?: Record<string, string>;
}

export interface FillResult {
  data: Uint8Array;
  warnings: string[];
}

const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

// ---------------------------------------------------------------- text

/** Replaces `{{name}}` tokens. Values are XML-escaped; a missing value becomes empty text. */
export function fillTokens(xml: string, values: Record<string, string>): string {
  return xml.replace(/\{\{([A-Za-z0-9_]+)\}\}/g, (_, name: string) => escapeXml((values[name] ?? "").replace(/\s*\n\s*/g, " ")));
}

function setCellText(cell: string, text: string): string {
  const safe = escapeXml(text.replace(/\s*\n\s*/g, " "));
  const para = /<a:p>[\s\S]*?<\/a:p>/.exec(cell);
  if (!para) return cell;
  const p = para[0];
  const pPr = /<a:pPr[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/.exec(p)?.[0] ?? "";
  const end = /<a:endParaRPr([^>]*?)(\/>|>([\s\S]*?)<\/a:endParaRPr>)/.exec(p);
  const runPr = /<a:r>\s*(<a:rPr[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>))/.exec(p)?.[1] ?? (end ? (end[3] ? `<a:rPr${end[1]}>${end[3]}</a:rPr>` : `<a:rPr${end[1]}/>`) : "");
  const run = text === "" ? "" : `<a:r>${runPr}<a:t>${safe}</a:t></a:r>`;
  return cell.replace(p, () => `<a:p>${pPr}${run}${end?.[0] ?? ""}</a:p>`);
}

function fillRow(row: string, values: string[]): string {
  const cells = row.match(/<a:tc\b[\s\S]*?<\/a:tc>/g) ?? [];
  let out = row;
  cells.forEach((cell, i) => {
    if (i < values.length) out = out.replace(cell, () => setCellText(cell, values[i]));
  });
  return out;
}

/** Rebuilds the first table: header row kept, then one row per entry of `rows`. */
export function fillFirstTable(xml: string, rows: string[][]): string {
  const tbl = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(xml);
  if (!tbl) return xml;
  const trs = tbl[0].match(/<a:tr\b[\s\S]*?<\/a:tr>/g) ?? [];
  if (trs.length < 2) return xml;
  // The first data row sits under the header's heavy rule; later rows share the lighter style.
  const firstStyle = trs[1];
  const nextStyle = trs[2] ?? trs[1];
  const built = rows.map((vals, i) => fillRow(i === 0 ? firstStyle : nextStyle, [String(i + 1), ...vals]));
  const newTbl = tbl[0].replace(/<a:tr\b[\s\S]*<\/a:tr>/, () => trs[0] + built.join(""));
  return xml.replace(tbl[0], () => newTbl);
}

// ---------------------------------------------------------------- pictures

export function parseDataUrl(dataUrl: string): { bytes: Uint8Array; ext: "png" | "jpeg" } | null {
  const m = /^data:image\/(png|jpeg|jpg);base64,([A-Za-z0-9+/=\s]+)$/.exec(dataUrl);
  if (!m) return null;
  const bin = atob(m[2].replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { bytes, ext: m[1] === "png" ? "png" : "jpeg" };
}

export function imageSize(bytes: Uint8Array, ext: "png" | "jpeg"): { w: number; h: number } | null {
  if (ext === "png") {
    if (bytes.length < 24) return null;
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { w: dv.getUint32(16), h: dv.getUint32(20) };
  }
  let i = 2;
  while (i + 9 < bytes.length) {
    if (bytes[i] !== 0xff) { i++; continue; }
    const marker = bytes[i + 1];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { h: (bytes[i + 5] << 8) | bytes[i + 6], w: (bytes[i + 7] << 8) | bytes[i + 8] };
    }
    i += 2 + ((bytes[i + 2] << 8) | bytes[i + 3]);
  }
  return null;
}

/** Largest rectangle of the picture's proportions that fits inside the frame, centred. */
export function fitInside(frame: { x: number; y: number; cx: number; cy: number }, img: { w: number; h: number }) {
  const ratio = img.w / img.h;
  let cx = frame.cx;
  let cy = Math.round(cx / ratio);
  if (cy > frame.cy) {
    cy = frame.cy;
    cx = Math.round(cy * ratio);
  }
  return { x: frame.x + Math.round((frame.cx - cx) / 2), y: frame.y + Math.round((frame.cy - cy) / 2), cx, cy };
}

interface MediaWrite {
  path: string;
  bytes: Uint8Array;
}

function swapPicture(
  xml: string,
  relsXml: string,
  picName: string,
  dataUrl: string,
  mediaPath: string,
  relId: string,
  media: MediaWrite[],
  warnings: string[],
): { xml: string; rels: string } {
  const img = parseDataUrl(dataUrl);
  // Match the picture by its own name: a group or shape may carry the same name elsewhere on the slide.
  const named = new RegExp(`<p:cNvPr\\b[^>]*\\bname="${picName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`);
  const hit = [...xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)].find((m) => named.test(m[0]));
  if (!img || !hit) {
    warnings.push(img ? `Picture "${picName}" was not found on the template slide.` : `Picture "${picName}" is not a PNG or JPEG, so it was left as the template placeholder.`);
    return { xml, rels: relsXml };
  }
  const st = hit.index;
  const en = st + hit[0].length;
  let pic = hit[0];
  const size = imageSize(img.bytes, img.ext);
  const frame = /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(pic);
  if (size && frame) {
    const f = fitInside({ x: +frame[1], y: +frame[2], cx: +frame[3], cy: +frame[4] }, size);
    pic = pic.replace(frame[0], `<a:off x="${f.x}" y="${f.y}"/><a:ext cx="${f.cx}" cy="${f.cy}"/>`);
  }
  pic = pic.replace(/r:embed="[^"]*"/, `r:embed="${relId}"`);
  media.push({ path: mediaPath, bytes: img.bytes });
  const rel = `<Relationship Id="${relId}" Type="${REL_IMAGE}" Target="../media/${mediaPath.split("/").pop()}"/>`;
  return { xml: xml.slice(0, st) + pic + xml.slice(en), rels: relsXml.replace("</Relationships>", `${rel}</Relationships>`) };
}

// ---------------------------------------------------------------- package

function slidePathsInOrder(presXml: string, presRels: string): string[] {
  const rel = new Map<string, string>();
  for (const m of presRels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]*)"/.exec(m[0])?.[1];
    const target = /Target="([^"]*)"/.exec(m[0])?.[1];
    const type = /Type="([^"]*)"/.exec(m[0])?.[1];
    if (id && target && type === REL_SLIDE) rel.set(id, "ppt/" + target.replace(/^\/?(ppt\/)?/, ""));
  }
  const ids = [...presXml.matchAll(/<p:sldId\b[^>]*r:id="([^"]*)"/g)].map((m) => m[1]);
  return ids.map((id) => rel.get(id)).filter((p): p is string => !!p);
}

/**
 * Builds a new deck from `template`. `roles[i]` names template slide i+1; each spec clones the
 * slide of its role, fills it, and the specs become the slides of the result, in order.
 */
export async function fillTemplate(template: Uint8Array | ArrayBuffer, roles: string[], specs: SlideSpec[]): Promise<FillResult> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(template);
  const read = async (name: string) => {
    const f = zip.file(name);
    if (!f) throw new Error(`The template is missing ${name}`);
    return f.async("string");
  };

  const presXml = await read("ppt/presentation.xml");
  const presRels = await read("ppt/_rels/presentation.xml.rels");
  const order = slidePathsInOrder(presXml, presRels);
  if (order.length !== roles.length) throw new Error(`The template has ${order.length} slides, but ${roles.length} were expected.`);

  const sources = new Map<string, { xml: string; rels: string }>();
  for (let i = 0; i < roles.length; i++) {
    const path = order[i];
    const relsPath = path.replace("slides/", "slides/_rels/") + ".rels";
    sources.set(roles[i], { xml: await read(path), rels: await read(relsPath) });
  }

  const media: MediaWrite[] = [];
  const out: { xml: string; rels: string }[] = [];
  specs.forEach((spec, k) => {
    const src = sources.get(spec.role);
    if (!src) throw new Error(`Unknown slide role "${spec.role}".`);
    let xml = src.xml.replace(/(<p:sld\b[^>]*?) show="0"/, "$1"); // the template hides its slides
    let rels = src.rels;
    if (spec.table) xml = fillFirstTable(xml, spec.table);
    let n = 0;
    for (const [name, dataUrl] of Object.entries(spec.images ?? {})) {
      const ext = dataUrl.startsWith("data:image/png") ? "png" : "jpeg";
      const r = swapPicture(xml, rels, name, dataUrl, `ppt/media/fill-${k + 1}-${++n}.${ext}`, `rIdFill${n}`, media, warnings);
      xml = r.xml;
      rels = r.rels;
    }
    xml = fillTokens(xml, spec.tokens ?? {});
    out.push({ xml, rels });
  });

  // Replace the template's slides with the new ones.
  for (const path of order) {
    zip.remove(path);
    zip.remove(path.replace("slides/", "slides/_rels/") + ".rels");
  }
  out.forEach((s, k) => {
    zip.file(`ppt/slides/slide${k + 1}.xml`, s.xml);
    zip.file(`ppt/slides/_rels/slide${k + 1}.xml.rels`, s.rels);
  });
  for (const m of media) zip.file(m.path, m.bytes);

  const newRels = presRels.replace(/<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*\/>/g, "")
    .replace("</Relationships>", out.map((_, k) => `<Relationship Id="rIdOut${k + 1}" Type="${REL_SLIDE}" Target="slides/slide${k + 1}.xml"/>`).join("") + "</Relationships>");
  const ids = out.map((_, k) => 256 + k);
  let newPres = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids.map((id, k) => `<p:sldId id="${id}" r:id="rIdOut${k + 1}"/>`).join("")}</p:sldIdLst>`);
  newPres = newPres.replace(/<p14:sldIdLst>[\s\S]*?<\/p14:sldIdLst>/g, `<p14:sldIdLst>${ids.map((id) => `<p14:sldId id="${id}"/>`).join("")}</p14:sldIdLst>`);
  zip.file("ppt/presentation.xml", newPres);
  zip.file("ppt/_rels/presentation.xml.rels", newRels);

  let ct = await read("[Content_Types].xml");
  ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "");
  for (const ext of new Set(media.map((m) => (m.path.endsWith(".png") ? "png" : "jpeg")))) {
    if (!new RegExp(`<Default\\b[^>]*Extension="${ext}"`).test(ct)) {
      ct = ct.replace("<Override", `<Default Extension="${ext}" ContentType="image/${ext}"/><Override`);
    }
  }
  ct = ct.replace("</Types>", out.map((_, k) => `<Override PartName="/ppt/slides/slide${k + 1}.xml" ContentType="${CT_SLIDE}"/>`).join("") + "</Types>");
  zip.file("[Content_Types].xml", ct);

  const app = zip.file("docProps/app.xml");
  if (app) {
    const a = (await app.async("string"))
      .replace(/<Slides>\d+<\/Slides>/, `<Slides>${out.length}</Slides>`)
      .replace(/<HiddenSlides>\d+<\/HiddenSlides>/, "<HiddenSlides>0</HiddenSlides>")
      .replace(/<Notes>\d+<\/Notes>/, "<Notes>0</Notes>");
    zip.file("docProps/app.xml", a);
  }

  return { data: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), warnings };
}
