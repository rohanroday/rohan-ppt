"use client";

import { useEffect, useRef, useState } from "react";
import type { PptxViewer } from "@aiden0z/pptx-renderer";
import type { SlideOutline } from "@/lib/pptxFill/outline";

/**
 * Slide preview of a filled template file: the real slides, drawn in the browser, one at a time.
 * If the renderer cannot draw the file, a text view of every value is shown instead.
 */
export function FileDeckPreview({ data, outline }: { data: Uint8Array; outline: SlideOutline[] }) {
  const host = useRef<HTMLDivElement>(null);
  const viewer = useRef<PptxViewer | null>(null);
  const [index, setIndex] = useState(0);
  const [ready, setReady] = useState(false);
  const [failed, setFailed] = useState(false);
  const last = outline.length - 1;

  useEffect(() => {
    let cancelled = false;
    const el = host.current;
    (async () => {
      try {
        const { PptxViewer, RECOMMENDED_ZIP_LIMITS } = await import("@aiden0z/pptx-renderer");
        if (!el || cancelled) return;
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
        const v = await PptxViewer.open(buffer, el, { renderMode: "slide", fitMode: "contain", zipLimits: RECOMMENDED_ZIP_LIMITS });
        if (cancelled) {
          v.destroy();
          return;
        }
        viewer.current = v;
        setReady(true);
      } catch {
        if (!cancelled) setFailed(true);
      }
    })();
    return () => {
      cancelled = true;
      viewer.current?.destroy();
      viewer.current = null;
    };
  }, [data]);

  useEffect(() => {
    if (!ready) return;
    viewer.current?.renderSlide(index).catch(() => setFailed(true));
  }, [index, ready]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight") setIndex((i) => Math.min(i + 1, last));
      if (e.key === "ArrowLeft") setIndex((i) => Math.max(i - 1, 0));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [last]);

  const slide = outline[Math.min(index, last)];
  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center gap-1" role="tablist" aria-label="Slides">
        <button type="button" onClick={() => setIndex((i) => Math.max(i - 1, 0))} disabled={index === 0}
          className="mr-1 h-7 rounded border border-slate-300 bg-white px-2 text-xs text-slate-600 hover:bg-slate-50 disabled:opacity-40">← Previous</button>
        {outline.map((s, i) => (
          <button
            key={s.index}
            type="button"
            role="tab"
            aria-selected={i === index}
            onClick={() => setIndex(i)}
            className={`h-7 min-w-7 rounded border px-2 text-xs ${i === index ? "border-blue-600 bg-blue-50 font-semibold text-blue-700" : "border-slate-300 bg-white text-slate-600 hover:bg-slate-50"}`}
          >
            {s.index}
          </button>
        ))}
        <button type="button" onClick={() => setIndex((i) => Math.min(i + 1, last))} disabled={index === last}
          className="ml-1 h-7 rounded border border-slate-300 bg-white px-2 text-xs text-slate-600 hover:bg-slate-50 disabled:opacity-40">Next →</button>
        <span className="ml-auto text-xs text-slate-400">
          Slide {index + 1} / {outline.length} · ← → keys
        </span>
      </div>

      {!failed && (
        <div className="overflow-hidden rounded-md border border-slate-200 bg-slate-100">
          <div ref={host} className="w-full" aria-label={`Slide ${index + 1}: ${slide?.title ?? ""}`} />
          {!ready && <p className="p-6 text-center text-sm text-slate-500">Drawing the slides…</p>}
        </div>
      )}

      {failed && slide && (
        <div className="rounded-md border border-slate-200 bg-slate-50 p-4">
          <p className="mb-3 text-xs text-slate-500">The slide drawing is not available here, so the values are listed instead. The downloaded file has the full design.</p>
          <h3 className="mb-2 text-sm font-semibold text-slate-900">{slide.title}</h3>
          <div className="flex flex-col gap-3">
            {slide.blocks.map((b, i) =>
              b.kind === "table" ? (
                <div key={i} className="overflow-x-auto rounded border border-slate-200 bg-white">
                  <table className="w-full text-left text-xs">
                    <tbody>
                      {b.rows.map((row, r) => (
                        <tr key={r} className={r === 0 ? "bg-slate-100 font-medium text-slate-700" : "border-t border-slate-100 text-slate-800"}>
                          {row.map((cell, c) => (
                            <td key={c} className="px-2 py-1 align-top">{cell}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div key={i} className="rounded border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800">
                  {b.lines.map((l, k) => (
                    <p key={k}>{l}</p>
                  ))}
                </div>
              ),
            )}
          </div>
        </div>
      )}
    </div>
  );
}
