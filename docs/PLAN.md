# Zoho CRM "Presentation Builder" Widget — Plan

## Context
Internship task: inside Zoho CRM, a user opens a **Web Tab widget**, picks an **Account**, picks a **template** (Temp1, Temp2…), sees a **preview**, and **downloads a .pptx**. Slides come from template rows stored in a new CRM custom module, where each row holds HTML with placeholders that get filled with live CRM data. Claude Code builds it, using Zoho's MCP servers to set up CRM and Zoho's official `zoho-crm` skill plus the ZET CLI to build the widget.

Decisions made: official Zoho CRM MCP servers on the **Demo** org (project-scoped, uniquely named) · **Next.js + Tailwind CSS**, built as a **static export** and hosted internally in Zoho · **editable native PPT** (PptxGenJS, not screenshots) · main record = **Accounts** · rows over the limit **continue on new slides**.

Project folder `D:\internship\Rohan -Widgit-ppt` is empty, so everything is new.

What research turned up:
- **Current MCPs aren't enough.** "Cosmo CRM MCP" only has 8 client-script tools. It can't create modules or read records.
- **Zoho hosts 4 CRM MCP servers:** Data Insights (read-only/COQL), Data Operations (record CRUD), Module Customization (modules, fields, layouts), and Workflow. OAuth runs on first use.
- **Zoho publishes a Claude Code skill** (`zoho/crm-skills`) that scaffolds widgets with ZET, including Web Tab widgets.
- **Widget JS SDK calls we'll use:** `ZOHO.embeddedApp.init`, `ZOHO.CRM.API.searchRecord`, `getRecord`, `getRelatedRecords`, `getAllRecords`, `attachFile`.

---

## Phase 0 — Tooling setup (one-time, mostly by you)
1. **Node.js LTS** (check with `node -v`), then the ZET CLI: `npm install -g zoho-extension-toolkit`.
2. **Zoho CRM skill** (project scope): `npx skills add zoho/crm-skills --agent claude-code`. This adds `.claude/skills/zoho-crm`.
3. **Zoho CRM MCP servers for the Demo org — project-scoped, uniquely named.** Add them in a `.mcp.json` at the project root, so they only load in this project and never sit beside your claude.ai connectors (books, Cosmo CRM MCP, Zoho CRM - VedSatavik, Catalyst-Rose, NTT catylst, Poshaya- POS, workindom-books, Workindom-creator).

   | Server key (= tool prefix `mcp__<key>__…`) | Zoho server | Purpose |
   |---|---|---|
   | `zoho-crm-demo-insights` | Data Insights | read-only COQL, modules, field schemas |
   | `zoho-crm-demo-dataops` | Data Operations | record CRUD, related records |
   | `zoho-crm-demo-modules` | Module Customization | create modules, fields, layouts |

   ```json
   {
     "mcpServers": {
       "zoho-crm-demo-insights": { "type": "http", "url": "<Data Insights URL>" },
       "zoho-crm-demo-dataops":  { "type": "http", "url": "<Data Operations URL>" },
       "zoho-crm-demo-modules":  { "type": "http", "url": "<Module Customization URL>" }
     }
   }
   ```
   - URLs come from Zoho's setup page (https://www.zoho.com/crm/developer/docs/mcp/setup/claude.html), or generate your own at Zoho MCP (mcp.zoho.in / .com, matching the Demo org's data centre).
   - `.mcp.json` goes in `.gitignore`, because a self-generated Zoho MCP URL carries a key.
   - First load: approve the 3 project servers, then sign in to the **Demo org** for each (OAuth). Check the org name on the consent screen.
   - **Wrong-org guard:** for this project's sessions, turn off the claude.ai connectors "Cosmo CRM MCP", "Zoho CRM - VedSatavik" and "books", so CRM writes can only go to the Demo org. Before the first write, Claude reads the org name back through `zoho-crm-demo-insights` and confirms it with you.
4. **CRM permissions.** Your profile needs Developer/Customize permissions, and **Developer Mode** turned on under Setup → Developer Hub → Widgets.
5. Run `git init` in the project folder.

## Phase 1 — CRM data model (Claude does this through MCP; confirms each create with you)
1. **Inspect the org.** List modules and check Accounts and Contacts fields. Check whether a "Property Options" module already exists.
2. **Property Options module.** If it's missing, create it with a **Lookup → Accounts** (so it shows as a related list on the Account) plus sample fields: Property Name, Location, Price, Area, Type.
3. **Create custom module `Presentation_Templates`:**

   | Field | Type | Example |
   |---|---|---|
   | Template Name (record Name field) | Single Line | Temp1 |
   | Slide Title | Single Line | Property Listings |
   | Slide HTML | Multi-line **Rich Text** (as specified) | `<table>…</table>` |
   | Limit of Rows | Number | 20 |
   | Module Name | Picklist: Accounts / Property_Options / Contacts | Property_Options |
   | **Slide Order** (proposed addition) | Number | 2 |

   Slide Order is needed so the slides come out in a fixed order. Without it, record order isn't guaranteed.
   **Risk:** Zoho's rich-text field may sanitize or strip parts of the HTML. Right after creating the field, Claude saves a sample template and reads it back. If tags get mangled, switch to **Multi-line (Large)** plain text.
4. **Seed data:** Temp1 (3 slides: Introduction / Property Listings / Contact Details), a small Temp2, one test Account with about 45 Property Options (to exercise pagination) and 2 Contacts.

## Phase 2 — Template language (fixed spec, documented in README)
- **`#Field_API_Name$`** → field value from that slide's Module Name source, e.g. `#Account_Name$`.
- **`#Owner.name$`** → dot notation for lookup and user fields.
- **`#RowCountIterator` … `$RowCountIterator`** → this block repeats once per related record.
- **`#Sno$`** → serial number (1, 2, 3…), continuing across overflow slides.
- **Module Name decides the data source:**
  - `Accounts` → the selected Account record.
  - `Property_Options` / `Contacts` → related records of that Account (`getRelatedRecords`). Outside an iterator, the first record is used.
- **Limit of Rows = N:** the related records are split into chunks of N, and each chunk becomes its own slide with the static parts repeated. Title becomes "Property Listings (2/3)".
- **Escaping:** every CRM value is HTML-escaped before it's substituted. Missing fields become an empty string, with a warning shown in the preview.

## Phase 3 — Widget: Next.js + Tailwind, static export (Web Tab)

**Step 3.0 — Hosting spike first (half a day, before any real UI).** Build a hello-world Next.js page that calls `ZOHO.embeddedApp.init()` and shows the current user's name. Export it, pack it with ZET, upload it as an internal widget, and open it in a Web Tab. This proves three things:
- relative asset paths load inside Zoho's hosting
- `zet validate` accepts the `_next/` folder
- the SDK initialises

Fallback if something fails: a post-build script rewrites `/_next/` to `./static/`, renames the folder, and fixes the references.

**Layout.** The ZET project wraps the Next.js app. `app/` is build output only and is never edited by hand.
```
ppt-widget/                         ZET project root (zet init via the zoho-crm skill)
  plugin-manifest.json              Web Tab widget → url "/app/index.html"
  app/                              ← generated: copy of web/out after build
  scripts/build-widget.mjs          next build → clean app/ → copy web/out → app/
  web/                              Next.js (App Router, TypeScript, Tailwind CSS)
    next.config.mjs                 output:'export', assetPrefix:'./', images:{unoptimized:true}, trailingSlash:true
    public/vendor/ZohoEmbededAppSDK.min.js   vendored Zoho JS SDK (no CDN at runtime)
    src/app/layout.tsx              loads the SDK via next/script (beforeInteractive), Tailwind globals
    src/app/page.tsx                'use client' — the widget screen
    src/components/                 AccountSearch, TemplateSelect, GenerateButton, SlidePreview (sandboxed iframe, 16:9),
                                    ThumbnailStrip, WarningsPanel, ActionBar (Download / Save to Account)
    src/lib/zoho.ts                 typed SDK wrapper: init(), listTemplates(), searchAccounts(q), getAccount(id),
                                    getRelated(id, relatedList), attachFile(id, blob)
    src/lib/templateEngine.ts       pure TS (no SDK/DOM): groupTemplates, fillPlaceholders, expandIterator, paginate → Slide[]
    src/lib/htmlToPptx.ts           DOMParser walker → PptxGenJS (npm, bundled): h1-h3/p/li → addText,
                                    table → addTable, img → addImage; bold/italic/color/align/bg/font-size
    src/lib/__tests__/templateEngine.test.ts   Vitest
    src/lib/mock/sampleData.ts      fake Account / Property Options / Contacts / templates for local dev outside CRM
```
- **Running outside CRM.** When `window.ZOHO` is missing (plain `next dev`), `zoho.ts` switches to `mock/sampleData.ts`. That makes UI work fast, and Zoho is only needed for integration checks.
- **Client-only.** Every SDK and PptxGenJS call runs in client components. Nothing needs a server, which is required for static export.
**User flow:**
1. The Web Tab loads and `ZOHO.embeddedApp.init()` runs.
2. Templates are fetched and grouped by Template Name for the dropdown.
3. The user types an Account name; `searchRecord` returns matches and they pick one.
4. The user picks a template and clicks Generate. The widget fetches the Account and its related lists, then the engine builds the slide list.
5. **Preview:** slides render as 16:9 cards in sandboxed iframes, with prev/next and thumbnails.
6. **Download:** `pptx.writeFile("<Account>-<Template>.pptx")`. A **"Save to Account"** option uses `ZOHO.CRM.API.attachFile`. This is also the fallback if the CRM iframe blocks downloads.

Later, optionally: the same widget as a Button on the Account detail page, reading the record ID from `PageLoad` so there's no search step.

## Phase 4 — Package & deploy
1. **Develop:**
   - UI work happens with `npm run dev` in `web/` (mock data).
   - CRM integration: run `node scripts/build-widget.mjs`, then `zet run`. In CRM, create the widget as a Web Tab with External hosting pointing at the localhost URL (Developer Mode).
2. **Package:** `node scripts/build-widget.mjs`, then `zet validate`, then `zet pack`. Upload `dist/*.zip` as **Internal hosting**.
3. **Place it:** Setup → Customization → Web Tabs → New Web Tab → Widget → choose profiles.

## Verification
- **Hosting spike (Step 3.0):** the hello-world page shows the CRM user's name inside the Web Tab, and the browser console shows no 404s for `_next` assets.
- **Lint and build:** `npm run lint` and `npm run build` in `web/` both pass, and `zet validate` passes on the packed output.
- **Unit tests** (`npx vitest run` in `web/`):
  - placeholder fill, plus HTML escaping
  - iterator expansion
  - 45 rows with limit 20 → 3 slides (20/20/5), with Sno 1–45 continuous
  - missing-field warnings
- **MCP setup check:** the connector list shows `zoho-crm-demo-insights`, `-dataops` and `-modules` as connected, with Cosmo/VedSatavik off. A read through `zoho-crm-demo-insights` returns the Demo org's modules.
- **MCP round-trip:** read the seeded template back through `zoho-crm-demo-insights` and confirm the HTML is intact.
- **In CRM (`zet run` + external hosting):**
  1. Open the Web Tab.
  2. Search the test Account and pick Temp1.
  3. The preview should show Intro + 3 Property slides + Contact slide.
  4. Download the .pptx, open it in PowerPoint/LibreOffice, and confirm the text and tables are editable and the data matches CRM.
  5. Switch to Temp2 and confirm the output differs.
- **After internal hosting:** repeat the flow once to confirm the vendored library loads and the download/attach works.

## Docs to update when building starts
- `docs/PLAN.md`: replace it with this plan.
- `docs/SRS.md`:
  - add a "Technology stack" item to §2.4/2.5 (Next.js static export + Tailwind CSS, TypeScript, PptxGenJS, Vitest)
  - C1 becomes "Next.js static export packed with ZET, internally hosted"
  - §6.3 names the Demo MCP servers (`zoho-crm-demo-*`)

## Open risks to check early
- Next.js static export inside Zoho internal hosting: relative asset paths and the `_next/` folder name (Step 3.0 spike).
- Rich-text sanitization (Phase 1.3 test).
- Whether the download works inside the CRM iframe (fallback: attachFile).
- `getRelatedRecords` uses the related-list API name, not the module name. Read it from the Account's related-list metadata.
- API credit usage. Fetch related records with per_page = 200 and paginate.

---

## Status log

### 2026-10-22 — Hosting spike, local part done
Built: `ppt_widget/` (ZET project, created with `zet init --zoho-service CRM --project-name ppt_widget`) containing `web/` (Next.js 16 + Tailwind 4, static export) and `scripts/build-widget.mjs`. `zet validate` passes and `zet pack` produces `dist/ppt_widget.zip`. The page loads from an `/app/` path with relative assets and no console errors.

Findings that change the plan:
- **Entry page is `app/widget.html`**, not `index.html`. The build script renames it.
- **ZET folder names allow only `a-zA-Z0-9_-$`.** Next's `__next.*` payload folders and `_not-found` are removed by the build script (not needed for a single screen).
- **Allowed file types:** `.ico` is rejected, so no favicon. Fonts must be woff/woff2/ttf/eot.
- **Next 16 scaffold enables `cacheComponents`, which needs the Node runtime.** It is switched off for static export. The scaffold's own `AGENTS.md` says to read `node_modules/next/dist/docs/` before writing Next code.
- **Zoho SDK is vendored** at `web/public/vendor/ZohoEmbededAppSDK.min.js`, but it loads one more script from a URL CRM supplies, so "no external calls" became "no third-party calls" (SRS C2).
- **Outside-CRM detection:** the vendored SDK always defines `window.ZOHO`, so "ZOHO missing" can't be the mock-mode switch. Use `window.parent === window` (or a `?mock=1` flag) instead.

Still to do for the spike: upload `dist/ppt_widget.zip` as an internal-hosting widget (index page `/widget.html`), add it to a Web Tab, and confirm it shows the CRM user's name.

Housekeeping for git: ignore `node_modules/`, `web/.next/`, `web/out/`, `app/` (generated), `dist/`, `cert.pem`, `key.pem`, and `.mcp.json`.

### 2026-10-22 — Phase 1 (CRM data model) done
Org confirmed as the Demo org by the user (org IDs start `7644593…`, sample Accounts like "Benton (Sample)"). All writes went through the `zoho-crm-demo-*` MCP servers.

Modules (both created earlier the same day):
- **Property_Options**: Name, Account (lookup → Accounts), Location, Price (currency), Area (decimal), **Type** (picklist: Apartment/Villa/Plot/Commercial). The picklist API name is `Type`, not `Property_Type`.
- **Presentation_Templates**: Name, Slide_Title, **Slide_HTML (Multi-line Large, 32,000 chars)**, Limit_of_Rows, Module_Name (Accounts/Property_Options/Contacts), Slide_Order.
- The related list on Accounts for Property Options has API name **`Property_Options`** (resolves OI-3).

**OI-1 resolved, rich text dropped:** HTML saved through the API into the rich-text `Slide_HTML` came back with every tag stripped. That field was deleted and recreated as Multi-line (Large) with the same API name `Slide_HTML`. The round-trip now returns the HTML unchanged, including inline styles and the iterator markers. Admins type raw HTML into a plain text box (fits SRS A2). Note: right after a field delete, creating one with the same label briefly fails with DUPLICATE_DATA; a retry a few seconds later works.

Seed data (source: `crm/seed-data.json`, generated by `node crm/make-seed.mjs`):
- Account **Test Realty Pvt Ltd**, id `7644593000000742001`
- 2 Contacts (Asha Menon, Vikram Rao) linked to it
- 45 Property Options linked to it; row 7 is named `Escape test <b>x</b>` for AC-4 (stored as typed)
- Templates: Temp1 = Introduction (Accounts, order 1), Property Listings (Property_Options, limit 20, order 2), Contact Details (Contacts, order 3); Temp2 = Account Summary (Accounts, 1), Shortlist (Property_Options, limit 10, 2)
- The 10 sample Accounts have no Property Options, which covers AC-9.

Expected widget output for Test Realty: Temp1 → 5 slides (Intro, Listings 1/3–3/3 with 20/20/5 rows, Contacts); Temp2 → 6 slides (Summary, Shortlist 1/5–5/5).

### 2026-10-22 — Phases 2–4 built (widget complete, ready to upload)
**Phase 2 (template engine):** `web/src/lib/templateEngine.ts`, pure TS covering T1–T10: placeholders, lookups, `#Account.X$`, iterator, `#Sno$`, overflow pagination with " (n/N)" titles, "No records" row, escaping, warnings, T9 formatting. Values known from field metadata but absent on a record don't warn.

**Phase 3 (widget):**
- `src/lib/zoho.ts`: SDK wrapper (`getAllRecords` with pagination, `searchRecord` word search, `getRecord`, `getRelatedRecords` per_page 200, `CONFIG.getCurrentUser/getOrgInfo`, `META.getFields`, `attachFile`). The list API can truncate long text (`$has_more.Slide_HTML`), so those template rows are re-fetched with `getRecord`. Mock mode applies when not in an iframe or with `?mock=1`.
- `src/lib/htmlToPptx.ts`: DOMParser walker to PptxGenJS 4.0.1 (lazy-loaded). Text boxes, bullets, native tables (content-weighted column widths, adaptive font so ~20 rows fit), images, inline styles. The slide title and "n / N" go in a footer.
- `src/lib/slidePreview.ts` + `components/SlidePreview.tsx`: sandboxed `srcdoc` iframes at 1280×720 (= 13.333×7.5 in at 96 px/in, same sizes as the export), prev/next, ←/→ keys, thumbnails.
- `page.tsx`: search → template → Generate; related-list failures (e.g. no permission) show an error while the other slides still generate (AC-8); Download and Save to Account.
- Seed generator also writes `web/src/lib/mock/seed-data.json`, so the mock and CRM data are identical.

**Phase 4:** `node scripts/build-widget.mjs` → `zet validate` passes → `zet pack` → `dist/ppt_widget.zip` (~415 KB).

**Verification done:**
- `npm test`: 22 Vitest tests (engine and real .pptx export via jsdom + JSZip), including AC-3 (20/20/5, S.No 1–45), AC-4, AC-5, AC-9, FR-21. `tsc`, `npm run lint` and `next build` are clean.
- Built `app/` driven in headless Edge (mock mode): search "test" → Temp1 → 5 slides, no warnings, no console errors; Download produced the .pptx; "King (Sample)" gives a "No records" row.
- The downloaded .pptx opened in PowerPoint (via COM) without a repair prompt: 5 slides, native tables. The first render exposed split bullets and equal-width columns; both are fixed and the deck re-checked.

**Not yet verified (needs the CRM UI):** upload `dist/ppt_widget.zip` as an internal Web Tab widget and run AC-1…AC-9 there, especially:
- whether downloads work inside the CRM iframe (OI-2); Save to Account is the fallback
- `searchRecord` "word" search for "Test"
- `attachFile` with a Blob

Housekeeping: `next build` also writes a small `ppt_widget/.next/build` folder (config compile, because of the two lockfiles). It's harmless and should be git-ignored.

### 2026-10-22 — Fix: widget stuck on "Connecting to Zoho CRM…"
- `next dev` never hydrated because of `assetPrefix: "./"`. It now applies only to production builds (static export still uses relative paths; re-verified the built widget).
- Connecting no longer waits for `PageLoad` (a Web Tab doesn't need it) and times out after 8 s. In a frame that isn't on a Zoho site (e.g. an editor preview pane), it falls back to mock data with a notice; inside Zoho a timeout shows an error.
- Repacked `dist/ppt_widget.zip`.
