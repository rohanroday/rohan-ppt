"use client";

import { useState } from "react";
import type { GenerationWarning } from "@/lib/types";

/** FR-14: warnings shown with the preview; they never stop generation. */
export function WarningsPanel({ warnings }: { warnings: GenerationWarning[] }) {
  const [open, setOpen] = useState(false);
  if (warnings.length === 0) return null;
  return (
    <div className="rounded-md border border-amber-300 bg-amber-50 text-sm text-amber-900">
      <button type="button" onClick={() => setOpen((o) => !o)} className="flex w-full items-center justify-between px-3 py-2 text-left" aria-expanded={open}>
        <span>
          <strong>{warnings.length} warning{warnings.length > 1 ? "s" : ""}</strong> — the slides were still generated.
        </span>
        <span className="text-xs">{open ? "Hide" : "Show"}</span>
      </button>
      {open && (
        <ul className="space-y-1 border-t border-amber-200 px-3 py-2">
          {warnings.map((w, i) => (
            <li key={i}>
              <span className="font-medium">{w.slide}:</span> {w.message}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
