"use client";

import { useEffect, useRef, useState } from "react";
import { describeError, type AccountSummary, type DataSource } from "@/lib/zoho";

interface Props {
  source: DataSource;
  value: AccountSummary | null;
  onChange: (account: AccountSummary | null) => void;
  disabled?: boolean;
}

/**
 * FR-06/FR-07: debounced Account search (2+ characters, 300 ms, up to 20 matches).
 * Focusing or clicking the empty box lists the most recently modified Accounts.
 */
export function AccountSearch({ source, value, onChange, disabled }: Props) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<AccountSummary[]>([]);
  const [suggestions, setSuggestions] = useState<AccountSummary[] | null>(null);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadingSuggestions, setLoadingSuggestions] = useState(false);
  const [error, setError] = useState("");
  const [active, setActive] = useState(0);
  const requestId = useRef(0);
  const boxRef = useRef<HTMLDivElement>(null);

  const term = query.trim();
  useEffect(() => {
    if (term.length < 2) return;
    const id = ++requestId.current;
    const timer = setTimeout(async () => {
      setLoading(true);
      setError("");
      try {
        const found = await source.searchAccounts(term);
        if (id !== requestId.current) return;
        setResults(found);
        setActive(0);
        setOpen(true);
      } catch (e) {
        if (id === requestId.current) setError(describeError(e, "search Accounts"));
      } finally {
        if (id === requestId.current) setLoading(false);
      }
    }, 300);
    return () => clearTimeout(timer);
  }, [term, source]);

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, []);

  const openSuggestions = async () => {
    setActive(0);
    if (suggestions) {
      setOpen(true);
      return;
    }
    if (loadingSuggestions) return;
    setLoadingSuggestions(true);
    setError("");
    try {
      setSuggestions(await source.suggestAccounts());
      setOpen(true);
    } catch (e) {
      setError(describeError(e, "load Accounts"));
    } finally {
      setLoadingSuggestions(false);
    }
  };

  const pick = (a: AccountSummary) => {
    onChange(a);
    setOpen(false);
    setQuery("");
    setResults([]);
  };

  if (value) {
    return (
      <div className="flex h-9 items-center justify-between gap-2 rounded-md border border-blue-300 bg-blue-50 px-3 text-sm">
        <span className="truncate">
          <span className="font-medium text-slate-900">{value.name}</span>
          {value.secondary && <span className="ml-2 text-slate-500">{value.secondary}</span>}
        </span>
        <button
          type="button"
          onClick={() => onChange(null)}
          disabled={disabled}
          className="shrink-0 rounded px-1.5 text-slate-500 hover:bg-blue-100 hover:text-slate-800 disabled:opacity-50"
          aria-label="Clear selected Account"
          title="Clear"
        >
          ✕
        </button>
      </div>
    );
  }

  const searching = term.length >= 2;
  const items = searching ? results : (suggestions ?? []);
  const showList = open && (searching || suggestions !== null);
  return (
    <div ref={boxRef} className="relative">
      <input
        type="search"
        value={query}
        disabled={disabled}
        onChange={(e) => {
          setQuery(e.target.value);
          if (e.target.value.trim().length < 2) {
            requestId.current++;
            setResults([]);
            setLoading(false);
            setError("");
            void openSuggestions();
          }
        }}
        onFocus={() => (searching ? results.length > 0 && setOpen(true) : void openSuggestions())}
        onClick={() => !searching && void openSuggestions()}
        onKeyDown={(e) => {
          if (e.key === "Escape") { setOpen(false); return; }
          if (!showList || items.length === 0) return;
          if (e.key === "ArrowDown") { e.preventDefault(); setActive((a) => Math.min(a + 1, items.length - 1)); }
          else if (e.key === "ArrowUp") { e.preventDefault(); setActive((a) => Math.max(a - 1, 0)); }
          else if (e.key === "Enter") { e.preventDefault(); pick(items[active]); }
        }}
        placeholder="Search Accounts by name…"
        className="h-9 w-full rounded-md border border-slate-300 bg-white px-3 text-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100 disabled:bg-slate-100"
        aria-label="Search Accounts"
        role="combobox"
        aria-expanded={showList}
        aria-controls="account-results"
      />
      {(loading || loadingSuggestions) && (
        <span className="absolute right-3 top-2 text-xs text-slate-400">{loading ? "Searching…" : "Loading…"}</span>
      )}
      {error && <p className="mt-1 text-xs text-red-600">{error}</p>}
      {showList && !loading && (
        <ul id="account-results" role="listbox" className="absolute z-20 mt-1 max-h-72 w-full overflow-auto rounded-md border border-slate-200 bg-white py-1 shadow-lg">
          {!searching && items.length > 0 && (
            <li role="presentation" className="px-3 pb-1 pt-2 text-xs font-semibold uppercase tracking-wide text-slate-400">
              Recent Accounts
            </li>
          )}
          {items.length === 0 && (
            <li className="px-3 py-2 text-sm text-slate-500">{searching ? `No Accounts match “${term}”.` : "No Accounts found."}</li>
          )}
          {items.map((a, i) => (
            <li
              key={a.id}
              role="option"
              aria-selected={i === active}
              onMouseDown={(e) => { e.preventDefault(); pick(a); }}
              onMouseEnter={() => setActive(i)}
              className={`cursor-pointer px-3 py-2 text-sm ${i === active ? "bg-blue-50" : ""}`}
            >
              <div className="font-medium text-slate-900">{a.name}</div>
              {a.secondary && <div className="text-xs text-slate-500">{a.secondary}</div>}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
