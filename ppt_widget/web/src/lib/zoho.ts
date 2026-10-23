// CRM data access. Two implementations behind one interface:
// - crmSource: the Zoho CRM widget JS SDK (inside CRM)
// - mockSource: sample data (plain `next dev`, or any URL with ?mock=1)
// The vendored SDK always defines window.ZOHO, so "inside CRM" is detected by
// running in an iframe, not by the SDK being present.

import { mockSource } from "./mock/sampleData";
import type { CrmRecord, FieldTypes, FormatContext, TemplateRow } from "./types";

export interface AccountSummary {
  id: string;
  name: string;
  secondary?: string;
}

export interface DataSource {
  mode: "crm" | "mock";
  init(): Promise<{ userName: string }>;
  listTemplates(): Promise<TemplateRow[]>;
  searchAccounts(query: string): Promise<AccountSummary[]>;
  /** Most recently modified Accounts, shown when the search box is focused before typing. */
  suggestAccounts(): Promise<AccountSummary[]>;
  getAccount(id: string): Promise<CrmRecord>;
  getRelated(accountId: string, module: string): Promise<CrmRecord[]>;
  getFormatContext(modules: string[]): Promise<FormatContext>;
  attachFile(accountId: string, fileName: string, file: Blob): Promise<void>;
}

/** Related-list API names on Accounts, by source module (checked in the Demo org; OI-3). */
export const RELATED_LISTS: Record<string, string> = {
  Property_Options: "Property_Options",
  Contacts: "Contacts",
};

const PER_PAGE = 200;

// ---------------------------------------------------------------- SDK typing

type ApiResponse = { data?: unknown; info?: { more_records?: boolean }; status?: number | string } | "" | undefined;

interface ZohoSdk {
  embeddedApp: { on(event: string, cb: (data: unknown) => void): void; init(): Promise<unknown> };
  CRM: {
    API: {
      getAllRecords(p: Record<string, unknown>): Promise<ApiResponse>;
      getRecord(p: Record<string, unknown>): Promise<ApiResponse>;
      searchRecord(p: Record<string, unknown>): Promise<ApiResponse>;
      getRelatedRecords(p: Record<string, unknown>): Promise<ApiResponse>;
      attachFile(p: Record<string, unknown>): Promise<ApiResponse>;
    };
    CONFIG: {
      getCurrentUser(): Promise<{ users?: { full_name?: string; date_format?: string; locale?: string }[] }>;
      getOrgInfo(): Promise<{ org?: { currency_symbol?: string }[] }>;
    };
    META: { getFields(p: { Entity: string }): Promise<{ fields?: { api_name: string; data_type: string }[] }> };
  };
}

function sdk(): ZohoSdk {
  const z = (window as unknown as { ZOHO?: ZohoSdk }).ZOHO;
  if (!z) throw new Error("The Zoho CRM SDK did not load. Reload the tab; if it keeps happening, re-upload the widget.");
  return z;
}

/** Records from an SDK response; throws on API-level errors (e.g. NO_PERMISSION). */
function records(resp: ApiResponse): CrmRecord[] {
  if (!resp || typeof resp !== "object" || !Array.isArray(resp.data)) return []; // 204 = no records
  const rows = resp.data as CrmRecord[];
  const err = rows.find((r) => r && (r.status === "error" || (typeof r.code === "string" && r.code !== "SUCCESS" && !r.id)));
  if (err) throw err;
  return rows;
}

/** A readable message for any SDK or JS error (NFR-U2). */
export function describeError(e: unknown, action: string): string {
  const o = (e && typeof e === "object" ? e : {}) as Record<string, unknown>;
  const raw = (o.data && Array.isArray(o.data) ? (o.data[0] as Record<string, unknown>) : o) ?? {};
  const code = String(raw.code ?? o.code ?? "");
  const message = String(raw.message ?? o.message ?? (typeof e === "string" ? e : ""));
  if (/NO_PERMISSION|OAUTH_SCOPE|permission/i.test(code + message)) {
    return `Couldn't ${action}: your CRM profile doesn't have access. Ask your CRM admin for read access, then try again.`;
  }
  if (/LIMIT|TOO_MANY/i.test(code)) return `Couldn't ${action}: the org's API limit was reached. Try again later.`;
  if (/INVALID_MODULE|INVALID_DATA|INVALID_URL/i.test(code)) {
    return `Couldn't ${action}: CRM rejected the request (${code}${message ? `: ${message}` : ""}). Check the module and field setup.`;
  }
  return `Couldn't ${action}${message ? `: ${message}` : ""}. Check your connection and try again.`;
}

// ---------------------------------------------------------------- CRM implementation

const fieldTypeCache = new Map<string, Record<string, string>>();

const crmSource: DataSource = {
  mode: "crm",

  async init() {
    const z = sdk();
    // PageLoad carries record context for record-page widgets; a Web Tab doesn't need it,
    // so don't wait for it. init() resolving is enough to make API calls.
    z.embeddedApp.on("PageLoad", () => {});
    await withTimeout(z.embeddedApp.init(), INIT_TIMEOUT_MS);
    const me = await withTimeout(z.CRM.CONFIG.getCurrentUser(), INIT_TIMEOUT_MS);
    return { userName: me.users?.[0]?.full_name ?? "" };
  },

  async listTemplates() {
    const z = sdk();
    const all: CrmRecord[] = [];
    for (let page = 1; ; page++) {
      const resp = await z.CRM.API.getAllRecords({ Entity: "Presentation_Templates", page, per_page: PER_PAGE, sort_by: "Created_Time", sort_order: "asc" });
      all.push(...records(resp));
      if (!(resp && typeof resp === "object" && resp.info?.more_records)) break;
    }
    // List responses can cut long text; fetch the full record when CRM says there's more.
    const full = await Promise.all(
      all.map(async (r) => {
        const more = (r.$has_more as Record<string, boolean> | undefined)?.Slide_HTML;
        if (!more) return r;
        const one = records(await z.CRM.API.getRecord({ Entity: "Presentation_Templates", RecordID: r.id }));
        return one[0] ?? r;
      }),
    );
    return full.map(toTemplateRow);
  },

  async searchAccounts(query) {
    const q = query.replace(/[()\\,]/g, " ").trim();
    if (q.length < 2) return [];
    const resp = await sdk().CRM.API.searchRecord({ Entity: "Accounts", Type: "word", Query: q, page: 1, per_page: 20 });
    return records(resp).slice(0, 20).map(toAccountSummary);
  },

  async suggestAccounts() {
    const resp = await sdk().CRM.API.getAllRecords({ Entity: "Accounts", page: 1, per_page: 20, sort_by: "Modified_Time", sort_order: "desc" });
    return records(resp).slice(0, 20).map(toAccountSummary);
  },

  async getAccount(id) {
    const rec = records(await sdk().CRM.API.getRecord({ Entity: "Accounts", RecordID: id }))[0];
    if (!rec) throw new Error("The selected Account no longer exists.");
    return rec;
  },

  async getRelated(accountId, module) {
    const list = RELATED_LISTS[module] ?? module;
    const all: CrmRecord[] = [];
    for (let page = 1; ; page++) {
      const resp = await sdk().CRM.API.getRelatedRecords({ Entity: "Accounts", RecordID: accountId, RelatedList: list, page, per_page: PER_PAGE });
      all.push(...records(resp));
      if (!(resp && typeof resp === "object" && resp.info?.more_records)) break;
    }
    return all;
  },

  async getFormatContext(modules) {
    const z = sdk();
    const ctx: FormatContext = {};
    try {
      const me = (await z.CRM.CONFIG.getCurrentUser()).users?.[0];
      ctx.dateFormat = me?.date_format;
      ctx.locale = me?.locale?.replace("_", "-");
    } catch { /* optional */ }
    try {
      ctx.currencySymbol = (await z.CRM.CONFIG.getOrgInfo()).org?.[0]?.currency_symbol;
    } catch { /* optional */ }
    const fieldTypes: FieldTypes = {};
    await Promise.all(
      ["Accounts", ...modules].map(async (m) => {
        if (!fieldTypeCache.has(m)) {
          try {
            const res = await z.CRM.META.getFields({ Entity: m });
            fieldTypeCache.set(m, Object.fromEntries((res.fields ?? []).map((f) => [f.api_name, f.data_type])));
          } catch {
            fieldTypeCache.set(m, {});
          }
        }
        fieldTypes[m] = fieldTypeCache.get(m)!;
      }),
    );
    ctx.fieldTypes = fieldTypes;
    return ctx;
  },

  async attachFile(accountId, fileName, file) {
    const resp = await sdk().CRM.API.attachFile({ Entity: "Accounts", RecordID: accountId, File: { Name: fileName, Content: file } });
    records(resp); // throws on API error
  },
};

function toAccountSummary(r: CrmRecord): AccountSummary {
  return {
    id: String(r.id),
    name: String(r.Account_Name ?? "(no name)"),
    secondary: [r.Billing_City, (r.Owner as { name?: string } | null)?.name].filter(Boolean).join(" · ") || undefined,
  };
}

function toTemplateRow(r: CrmRecord): TemplateRow {
  const num = (v: unknown) => (typeof v === "number" ? v : v === null || v === undefined || v === "" ? null : Number(v));
  return {
    id: String(r.id),
    Name: String(r.Name ?? ""),
    Slide_Title: String(r.Slide_Title ?? ""),
    Slide_HTML: String(r.Slide_HTML ?? ""),
    Limit_of_Rows: num(r.Limit_of_Rows),
    Module_Name: String(r.Module_Name ?? "Accounts"),
    Slide_Order: num(r.Slide_Order),
    Created_Time: typeof r.Created_Time === "string" ? r.Created_Time : undefined,
  };
}

// ---------------------------------------------------------------- choose

const INIT_TIMEOUT_MS = 8000;

function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error("Zoho CRM did not respond in time")), ms);
    p.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
  });
}

/** True when the page that embeds this widget is a Zoho site. */
function embeddedInZoho(): boolean {
  if (window.parent === window) return false;
  const ancestors = Array.from(window.location.ancestorOrigins ?? []);
  return [...ancestors, document.referrer].some((o) => /zoho|zappsusercontent/i.test(o));
}

export interface Connection {
  source: DataSource;
  userName: string;
  /** Set when CRM didn't answer outside Zoho and the widget switched to mock data. */
  fallbackReason?: string;
}

/**
 * Picks the data source and connects.
 * - Not in a frame, or ?mock=1: mock data.
 * - In a frame: try the Zoho SDK. If it doesn't answer and the frame isn't on a Zoho site
 *   (e.g. an editor's preview pane), fall back to mock data instead of hanging.
 */
export async function connect(): Promise<Connection> {
  const forceMock = new URLSearchParams(window.location.search).has("mock");
  if (forceMock || window.parent === window) {
    return { source: mockSource, userName: (await mockSource.init()).userName };
  }
  try {
    return { source: crmSource, userName: (await crmSource.init()).userName };
  } catch (e) {
    if (embeddedInZoho()) throw e;
    return {
      source: mockSource,
      userName: (await mockSource.init()).userName,
      fallbackReason: "Zoho CRM didn't respond (this page isn't inside CRM), so mock data is shown.",
    };
  }
}
