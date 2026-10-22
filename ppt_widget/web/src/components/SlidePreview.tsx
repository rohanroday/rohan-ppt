"use client";

import { useEffect, useRef, useState } from "react";
import { PREVIEW_H, PREVIEW_W, previewDoc } from "@/lib/slidePreview";
import type { GeneratedSlide } from "@/lib/types";

/** A 16:9 slide in a sandboxed iframe (no scripts, no access to the widget; FR-16, NFR-S2). */
function SlideFrame({ slide, index, total, width, title }: { slide: GeneratedSlide; index: number; total: number; width: number; title: string }) {
  const scale = width / PREVIEW_W;
  return (
    <div className="relative overflow-hidden bg-white" style={{ width, height: PREVIEW_H * scale }}>
      <iframe
        title={title}
        sandbox=""
        srcDoc={previewDoc(slide, index, total)}
        loading="lazy"
        tabIndex={-1}
        className="pointer-events-none absolute left-0 top-0 origin-top-left border-0"
        style={{ width: PREVIEW_W, height: PREVIEW_H, transform: `scale(${scale})` }}
      />
    </div>
  );
}

interface Props {
  slides: GeneratedSlide[];
  index: number;
  onIndex: (i: number) => void;
}

/** FR-16/FR-17: main slide, previous/next, counter and thumbnail strip. */
export function SlidePreview({ slides, index, onIndex }: Props) {
  const mainRef = useRef<HTMLDivElement>(null);
  const stripRef = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(800);
  const total = slides.length;
  const go = (i: number) => onIndex(Math.max(0, Math.min(total - 1, i)));

  useEffect(() => {
    const el = mainRef.current;
    if (!el) return;
    const ro = new ResizeObserver(([entry]) => setWidth(Math.floor(entry.contentRect.width)));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null;
      if (t && (t.tagName === "INPUT" || t.tagName === "SELECT" || t.tagName === "TEXTAREA")) return;
      if (e.key === "ArrowRight") onIndex(Math.min(total - 1, index + 1));
      if (e.key === "ArrowLeft") onIndex(Math.max(0, index - 1));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [index, total, onIndex]);

  useEffect(() => {
    stripRef.current?.querySelector<HTMLElement>(`[data-thumb="${index}"]`)?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [index]);

  const slide = slides[index];
  return (
    <div className="flex flex-col gap-3">
      <div ref={mainRef} className="w-full">
        {slide && (
          <div className="overflow-hidden rounded-md border border-slate-300 shadow-sm">
            <SlideFrame slide={slide} index={index} total={total} width={width} title={`Slide ${index + 1}: ${slide.title}`} />
          </div>
        )}
      </div>

      <div className="flex items-center justify-between gap-3">
        <button type="button" onClick={() => go(index - 1)} disabled={index === 0}
          className="rounded-md border border-slate-300 bg-white px-3 py-1.5 text-sm hover:bg-slate-50 disabled:opacity-40">
          ← Previous
        </button>
        <div className="min-w-0 text-center text-sm">
          <span className="font-medium tabular-nums">{index + 1} / {total}</span>
          <span className="ml-2 truncate text-slate-500">{slide?.title}</span>
        </div>
        <button type="button" onClick={() => go(index + 1)} disabled={index >= total - 1}
          className="rounded-md border border-slate-300 bg-white px-3 py-1.5 text-sm hover:bg-slate-50 disabled:opacity-40">
          Next →
        </button>
      </div>

      <div ref={stripRef} className="flex gap-2 overflow-x-auto pb-2" role="tablist" aria-label="Slides">
        {slides.map((s, i) => (
          <button
            key={`${s.templateRowId}-${i}`}
            type="button"
            data-thumb={i}
            role="tab"
            aria-selected={i === index}
            onClick={() => onIndex(i)}
            title={s.title}
            className={`shrink-0 overflow-hidden rounded border-2 text-left ${i === index ? "border-blue-500" : "border-transparent hover:border-slate-300"}`}
          >
            <SlideFrame slide={s} index={i} total={total} width={144} title={`Thumbnail ${i + 1}`} />
            <div className="w-[144px] truncate bg-slate-50 px-1.5 py-0.5 text-[11px] text-slate-600">{i + 1}. {s.title}</div>
          </button>
        ))}
      </div>
    </div>
  );
}
