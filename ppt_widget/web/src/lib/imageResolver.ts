// Image placeholders (SRS T11–T13, FR-25–FR-28). Pure TypeScript: the file download is injected.
// The engine turns an Image Upload / File Upload value into `crmfile:<file_Id>`; this step downloads
// each unique file once and rewrites the <img src> to a data URL, so preview and export need no changes.

import type { GeneratedSlide, GenerationWarning } from "./types";

export const FILE_REF_PREFIX = "crmfile:";
export const MAX_IMAGES = 50; // FR-26
const CONCURRENCY = 4; // FR-26

const IMG_TAG = /<img\b[^>]*>/gi;
const SRC_ATTR = /\bsrc\s*=\s*(["'])(.*?)\1/i;
const FILE_REF = /crmfile:([A-Za-z0-9_-]+)/;
const STRAY_REF = /crmfile:[A-Za-z0-9_-]+/g;

/** `crmfile:<id>` for the first file of an Image Upload / File Upload value, or "" when there is none. */
export function fileRef(value: unknown): string {
  const first = Array.isArray(value) ? value[0] : value;
  if (!first || typeof first !== "object") return "";
  const o = first as Record<string, unknown>;
  const id = o.file_Id ?? o.File_Id ?? o.fileId ?? o.id;
  return typeof id === "string" || typeof id === "number" ? `${FILE_REF_PREFIX}${id}` : "";
}

export interface ResolveResult {
  slides: GeneratedSlide[];
  warnings: GenerationWarning[];
}

const reasonOf = (e: unknown) => {
  const msg = e instanceof Error ? e.message : typeof e === "string" ? e : "";
  return msg ? ` (${msg.slice(0, 120)})` : "";
};

/**
 * Replaces every `crmfile:` picture with a data URL from `load`. A picture that is empty, over the
 * limit or fails to load is removed from its slide and reported (T13); generation never stops (FR-28).
 */
export async function resolveImages(
  slides: GeneratedSlide[],
  load: (fileId: string) => Promise<string>,
): Promise<ResolveResult> {
  const ids: string[] = [];
  for (const s of slides) {
    for (const tag of s.html.match(IMG_TAG) ?? []) {
      const id = FILE_REF.exec(SRC_ATTR.exec(tag)?.[2] ?? "")?.[1];
      if (id && !ids.includes(id)) ids.push(id);
    }
  }

  const loaded = new Map<string, string>();
  const failed = new Map<string, string>();
  const skipped = new Set(ids.slice(MAX_IMAGES));
  const queue = ids.slice(0, MAX_IMAGES);
  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, queue.length) }, async () => {
      for (let id = queue.shift(); id !== undefined; id = queue.shift()) {
        try {
          loaded.set(id, await load(id));
        } catch (e) {
          failed.set(id, reasonOf(e));
        }
      }
    }),
  );

  const warnings: GenerationWarning[] = [];
  const out = slides.map((slide) => {
    const seen = new Set<string>();
    const warn = (message: string) => {
      if (seen.has(message)) return;
      seen.add(message);
      warnings.push({ slide: slide.title, message });
    };
    let html = slide.html.replace(IMG_TAG, (tag) => {
      const src = SRC_ATTR.exec(tag)?.[2]?.trim() ?? "";
      if (!src) {
        warn("A picture was left out because its field is empty.");
        return "";
      }
      const id = FILE_REF.exec(src)?.[1];
      if (!id) return tag; // an ordinary web address or data URL: left for the exporter (C2)
      const data = loaded.get(id);
      if (data) return tag.replace(FILE_REF, () => data);
      warn(
        skipped.has(id)
          ? `Only the first ${MAX_IMAGES} pictures are loaded, so one picture was left out.`
          : `A picture could not be loaded, so it was left out${failed.get(id) ?? ""}.`,
      );
      return "";
    });
    if (html.includes(FILE_REF_PREFIX)) {
      html = html.replace(STRAY_REF, "");
      warn("An image field was used outside an <img src>, so its file reference was removed.");
    }
    return { ...slide, html };
  });

  return { slides: out, warnings };
}
