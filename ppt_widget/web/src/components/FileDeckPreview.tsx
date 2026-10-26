"use client";

import { useEffect, useState } from "react";
import type { SlideOutline } from "@/lib/pptxFill/outline";

/** Slide-by-slide data view of a filled template file, so the user can check the values before exporting. */
export function FileDeckPreview({ outline }: { outline: SlideOutline[] }) {
  const [index, setIndex] = useState(0);
  const last = outline.length - 1;
  const slide = outline[Math.min(index, last)];

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight") setIndex((i) => Math.min(i + 1, last));
      if (e.key === "ArrowLeft") setIndex((i) => Math.max(i - 1, 0));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [last]);

  if (!slide) return null;
  return (
    <div>
      <p className="mb-2 text-xs text-slate-500">
        Data preview: check the values here. The downloaded file keeps the customer&apos;s full design, pictures and icons.
      </p>
      <div className="mb-3 flex flex-wrap gap-1" role="tablist" aria-label="Slides">
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
      </div>

      <div className="rounded-md border border-slate-200 bg-slate-50 p-4">
        <div className="mb-3 flex items-baseline justify-between">
          <h3 className="text-sm font-semibold text-slate-900">
            Slide {slide.index} / {outline.length} · {slide.title}
          </h3>
          <span className="text-xs text-slate-400">← → keys change slides</span>
        </div>
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
                  <p key={k} className={k === 0 && b.lines.length > 1 ? "font-medium text-slate-600" : undefined}>{l}</p>
                ))}
              </div>
            ),
          )}
        </div>
      </div>
    </div>
  );
}
