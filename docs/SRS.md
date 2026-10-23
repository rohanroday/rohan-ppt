# Software Requirements Specification
## Zoho CRM Presentation Builder Widget

| | |
|---|---|
| Version | 1.1 (draft) |
| Date | 2026-10-23 |
| Author | Rohan |
| Status | For review |
| Related | [PLAN.md](PLAN.md) |

**Changes in 1.1** (after review feedback from the mentor):
- Templates can now include **images pulled from CRM** (T11–T13, FR-25–FR-29, AC-10, AC-11).
- The design will be checked against a **real customer's template and a deck generated from it** (§2.7, AC-12, OI-7). Those files are to be requested from Parth and have not been received, so anything that depends on them is marked as open.

---

## 1. Introduction

### 1.1 Purpose
This document specifies what the Presentation Builder must do. The Presentation Builder is a Zoho CRM widget that generates PowerPoint presentations from CRM data using templates stored in CRM. It is written for the developer, the reviewer/mentor, and anyone who tests or maintains the widget.

### 1.2 Scope
The system has two parts:
- **CRM data model**:
  - a custom module **Presentation Templates** that holds slide templates as HTML with placeholders
  - a custom module **Property Options**, related to Accounts, which supplies data for the listing slides
- **Web Tab widget**: an HTML/JS app embedded in Zoho CRM. In the widget a user:
  1. selects an Account record
  2. selects a template
  3. generates slides filled with that Account's data, both text and images
  4. previews the slides
  5. downloads them as an editable `.pptx` file

Out of scope for v1:
- editing slides inside the widget
- cropping, editing or uploading images in the widget (pictures come from CRM as they are, only scaled down if very large)
- PDF export
- template versioning
- multi-language templates
- scheduled or bulk generation

### 1.3 Definitions

| Term | Meaning |
|---|---|
| Template | A named set of slide definitions, e.g. "Temp1". It consists of every Presentation Template record that has the same Template Name. |
| Slide definition | One Presentation Template record: title, HTML, row limit, source module, order. |
| Placeholder | A token in the slide HTML, `#Field_API_Name$`, replaced with a CRM field value. |
| Iterator block | Markup between `#RowCountIterator` and `$RowCountIterator`. It repeats once per related record. |
| Row limit | The maximum number of iterator rows on one slide. Extra rows continue on new slides. |
| Overflow slide | An extra slide created when the rows exceed the row limit. |
| Image placeholder | An `<img>` tag in the slide HTML whose `src` is a placeholder, e.g. `<img src="#Photo$">`. It is replaced by the picture stored in that CRM field. |
| Image field | A CRM field that holds a picture or file: Image Upload or File Upload. The value is a file ID, not a web address. |
| Reference deck | A real customer's PowerPoint template together with one finished deck generated from it. Used to check that the widget's output matches what customers actually use (§2.7). |
| Web Tab | A Zoho CRM top-level navigation tab that hosts a widget. |
| ZET | Zoho Extension Toolkit, the command-line tool for scaffolding, running, validating and packing widgets. |
| MCP | Model Context Protocol. Zoho-hosted MCP servers let Claude Code read and modify CRM metadata and records during development. |
| JS SDK | Zoho CRM widget JavaScript SDK (`ZOHO.embeddedApp`, `ZOHO.CRM.API`). |

### 1.4 References
- Zoho CRM widgets: https://www.zoho.com/crm/developer/docs/widgets/create-widget.html
- Zoho CRM MCP servers: https://www.zoho.com/crm/developer/docs/mcp/overview.html
- Zoho CRM skill for AI agents: https://github.com/zoho/crm-skills
- PptxGenJS: https://github.com/gitbrent/PptxGenJS
- Original task brief (from the mentor), restated in Section 8.
- Customer reference deck (template plus generated output), to be requested from Parth: see §2.7 and OI-7.

---

## 2. Overall Description

### 2.1 Product perspective
The widget runs inside Zoho CRM's iframe sandbox. It has no backend of its own: all data access goes through the Zoho CRM JS SDK under the logged-in user's session and permissions. The PowerPoint file is generated entirely in the browser.

```
┌───────────── Zoho CRM (browser) ─────────────┐
│  Web Tab                                      │
│  ┌──────────── Widget (iframe) ────────────┐  │
│  │ UI ─► Template Engine ─► Preview         │  │
│  │  │                     └─► PptxGenJS ─► .pptx download
│  │  └─► JS SDK ◄──► CRM: Accounts,          │  │
│  │                 Property Options,        │  │
│  │                 Contacts,                │  │
│  │                 Presentation Templates   │  │
│  └──────────────────────────────────────────┘  │
└───────────────────────────────────────────────┘
Development time only: Claude Code ─► Zoho CRM MCP servers ─► CRM metadata and records
```

### 2.2 Product functions (summary)
1. List available templates.
2. Search for and select an Account.
3. Fetch the Account, its Property Options and its Contacts.
4. Render each slide definition, including iterator expansion and overflow pagination.
5. Download the pictures that image placeholders point to and put them into the slides.
6. Preview the slides.
7. Export an editable `.pptx` file and download it, or attach it to the Account.

### 2.3 User classes

| User | Needs | Permissions |
|---|---|---|
| Sales user | Generate and download decks for Accounts. | Read Accounts, Contacts, Property Options and Presentation Templates. Access to the Web Tab. |
| Template admin | Create and edit slide templates in the Presentation Templates module. | Create and edit Presentation Templates. |
| CRM admin/developer | Install, configure and deploy the widget and modules. | Administrator or Developer profile. |

### 2.4 Operating environment
- Zoho CRM, Enterprise edition or higher, or any edition that supports widgets, Web Tabs and custom modules.
- A current desktop browser: Chrome, Edge, Firefox or Safari, latest two versions.
- Generated files open in Microsoft PowerPoint 2016 or later, Google Slides and LibreOffice Impress.

### 2.4.1 Technology stack
| Layer | Choice |
|---|---|
| UI framework | Next.js (App Router), React, TypeScript |
| Styling | Tailwind CSS |
| PowerPoint generation | PptxGenJS (bundled via npm) |
| CRM access | Zoho CRM widget JS SDK (vendored) |
| Build and packaging | `next build` with static export, then Zoho Extension Toolkit (`zet validate`, `zet pack`) |
| Testing | Vitest (unit tests for the template engine) |
| Development tooling | Claude Code with the Zoho CRM MCP servers and the `zoho-crm` skill |

### 2.5 Design and implementation constraints
- **C1:** The widget is a Next.js static export (no server-side rendering or API routes), packed with ZET and hosted internally in CRM as an uploaded ZIP.
- **C2:** The widget makes no third-party network calls at runtime. All libraries, including PptxGenJS, are bundled in the package. The only outside traffic is Zoho's own: the vendored SDK loads a further script from a URL CRM supplies, and the SDK calls CRM APIs. Pictures from CRM fields are downloaded through the SDK, so they stay inside this rule. A template may also point `<img>` at an ordinary web address. That is best effort only: the browser may block it (CORS), and then the picture is left out with a warning (T13).
- **C6:** The package must pass `zet validate`: folder names use only letters, digits, `_`, `-` and `$`; only allowed file types (e.g. html, js, css, png, svg, woff2; not `.ico`) are included; the entry page is `app/widget.html`.
- **C3:** All CRM access uses the JS SDK, so it respects the user's role and profile.
- **C4:** CRM API calls consume the org's API credits.
- **C5:** Slide HTML is stored in a Multi-line (Large) field, so one slide's HTML is limited to 32,000 characters.

### 2.6 Assumptions and dependencies
- **A1:** Property Options has a lookup to Accounts, and Contacts are linked to Accounts through the standard Account lookup.
- **A2:** Template admins can write basic HTML: headings, paragraphs, lists, tables and images.
- **A3:** The org's data centre (.in, .com, etc.) supports the Zoho MCP servers used during development.
- **A4:** The modules that supply pictures (e.g. Property Options, Accounts) have an Image Upload or File Upload field, and the user's profile can read those files.
- **A5:** The customer's template and generated deck (from Parth, once received) show the real target layout. Until they arrive, the sample templates (Temp1, Temp2) stand in for it.

### 2.7 Customer reference deck
The mentor asked that the solution be based on a real customer's deck, not only on our own sample templates. Parth is to be asked for one customer's PowerPoint template and one deck generated from it (OI-7). The two files are used like this:
1. **Find the structure.** Count the slides, see which parts are fixed text, which come from CRM, and which repeat per record.
2. **Write the matching templates.** Turn each slide into one Presentation Template record, with placeholders (T1–T13), row limits and a Module Name.
3. **Find the images.** Note which pictures are in the deck and where they come from in CRM (OI-8).
4. **Compare the result.** Generate the deck for the same Account and compare it with the sample output (AC-12). Text, tables, data, order and pictures must match. Fonts, colours and exact positions may differ slightly (NFR-U3).

If the sample needs something the widget cannot do yet (for example a slide master, speaker notes or charts), it is added to OI-7 and discussed with the mentor, not silently dropped.

---

## 3. Data Requirements

### 3.1 Module: Presentation Templates (`Presentation_Templates`)

| # | Field label | API name | Type | Required | Notes |
|---|---|---|---|---|---|
| D1 | Template Name | `Name` | Single Line | Yes | Groups slide definitions into a template, e.g. "Temp1". |
| D2 | Slide Title | `Slide_Title` | Single Line | Yes | e.g. "Introduction". |
| D3 | Slide HTML | `Slide_HTML` | Multi-line (Large), 32,000 chars | Yes | Static HTML plus placeholders, typed as raw HTML. Rich Text was tried and dropped: it strips tags (OI-1). |
| D4 | Limit of Rows | `Limit_of_Rows` | Number | No | Rows per slide for an iterator block. Blank means no limit. |
| D5 | Module Name | `Module_Name` | Picklist: Accounts, Property_Options, Contacts | Yes | The data source for this slide. |
| D6 | Slide Order | `Slide_Order` | Number | Yes | Position of the slide in the deck, ascending. Added to the original brief. |

Example records:

| Name | Slide_Title | Limit_of_Rows | Module_Name | Slide_Order |
|---|---|---|---|---|
| Temp1 | Introduction | — | Accounts | 1 |
| Temp1 | Property Listings | 20 | Property_Options | 2 |
| Temp1 | Contact Details | — | Contacts | 3 |

### 3.2 Module: Property Options (`Property_Options`)
Create this module only if it does not already exist.

| Field | Type | Notes |
|---|---|---|
| Name (Property Name) | Single Line | Required |
| Account | Lookup → Accounts | Required. Creates the related list on Account. |
| Location | Single Line | |
| Price | Currency | |
| Area | Number | e.g. sq. ft. |
| Type (`Type`) | Picklist | Apartment, Villa, Plot, Commercial |

### 3.3 Existing modules used
- **Accounts:** read the selected record.
- **Contacts:** read through the Account's Contacts related list.

### 3.4 Image fields (new in 1.1)
Pictures reach a slide through an Image Upload or File Upload field on the module that supplies the slide's data.

| # | Where | Field | Type | Notes |
|---|---|---|---|---|
| D7 | Property Options | Photo (`Photo`) | Image Upload | One picture per property. To be created in the Demo org before the feature is tested there. |
| D8 | Accounts | Logo (`Logo`) | Image Upload | Optional. Shown on the introduction slide. |

- The value CRM returns for these fields is a file ID. The widget downloads the file by that ID (FR-25).
- Zoho's built-in *Record Image* (profile image) field uses a different download path. Support for it is not promised in v1 and is part of OI-8.
- The field names above are proposals. The real names come from the customer's data (OI-8).

---

## 4. Template Language Specification

| ID | Rule |
|---|---|
| T1 | `#Field_API_Name$` is replaced with the value of that field from the slide's data source. Field API names are case-sensitive. |
| T2 | `#Lookup.field$` reads a sub-field of a lookup or user field, e.g. `#Owner.name$`, `#Account_Name.name$`. |
| T3 | Text between `#RowCountIterator` and `$RowCountIterator` is repeated once per related record. Inside the block, placeholders resolve against the current related record. |
| T4 | `#Sno$` inside an iterator gives the 1-based row number. Numbering continues across overflow slides: slide 2 starts at Limit + 1. |
| T5 | Outside an iterator, placeholders on a related-module slide (Property_Options, Contacts) resolve against the first related record. Placeholders on an Accounts slide resolve against the selected Account. |
| T6 | `#Account.Field$` always resolves against the selected Account, whatever the slide's Module Name. This lets a listing slide show the Account name. |
| T7 | A slide has at most one iterator block. If more are found, only the first is expanded and a warning is raised. |
| T8 | Values are HTML-escaped before substitution. |
| T9 | Formatting rules: null or missing values render as an empty string; dates use the user's CRM date format if available, otherwise ISO; currency uses the org currency symbol if available; picklists render their display value; booleans render "Yes" or "No". |
| T10 | Text that looks like a placeholder but does not match the pattern `#[A-Za-z0-9_.]+\$` is left unchanged. |
| T11 | **Image placeholder.** `<img src="#Field$">` shows the picture stored in that field. It works on Accounts slides and inside an iterator, where each row gets its own picture. `width` and `height` attributes on the `<img>` are honoured; if they are missing the picture keeps its own proportions and is fitted to the slide. In v1 the `<img>` must sit on its own, not inside a paragraph, heading, list item or table cell: pictures placed there are not exported (OI-10). |
| T12 | If the field is an Image Upload or File Upload field, the placeholder resolves to the first file in it. If the field holds ordinary text, such as a web address, that text is used as the picture's address (see C2). |
| T13 | If a picture cannot be shown (empty field, no permission, file missing, not a picture, download failed), the `<img>` is removed from that slide and a warning names the slide and the field. The rest of the slide is produced as normal. |

Example:
```html
<h1>Property Listings for #Account.Account_Name$</h1>
<table>
  <tr><th>S.No</th><th>Property</th><th>Location</th><th>Price</th></tr>
  #RowCountIterator
  <tr><td>#Sno$</td><td>#Name$</td><td>#Location$</td><td>#Price$</td></tr>
  $RowCountIterator
</table>
```

Example with pictures (each property shows its photo above its details; the Account logo sits at the top). Each `<img>` is on its own line:
```html
<img src="#Account.Logo$" width="120">
<h1>Properties</h1>
#RowCountIterator
<img src="#Photo$" width="120">
<p><b>#Name$</b>, #Location$, #Price$</p>
$RowCountIterator
```

---

## 5. Functional Requirements

Priority: **M** = Must, **S** = Should, **C** = Could.

### 5.1 Initialisation and templates
| ID | Requirement | Pri |
|---|---|---|
| FR-01 | On load, the widget initialises the Zoho JS SDK and shows a loading state until it is ready. | M |
| FR-02 | The widget fetches all Presentation Template records, following pagination, and groups them by Template Name. | M |
| FR-03 | A dropdown lists the distinct template names, sorted alphabetically. | M |
| FR-04 | Slide definitions within a template are ordered by Slide_Order ascending. Ties are broken by record creation time. | M |
| FR-05 | If no templates exist, the widget shows "No templates found — add records to Presentation Templates" and disables Generate. | M |

### 5.2 Record selection
| ID | Requirement | Pri |
|---|---|---|
| FR-06 | A search box finds Accounts by name once at least 2 characters are typed, with a 300 ms debounce. It shows up to 20 matches with Account Name and one secondary field (e.g. City or Owner). | M |
| FR-07 | The user can select one Account. The selection is shown clearly and can be cleared. | M |
| FR-08 | When the widget is opened from an Account (button or detail-page placement), the Account is pre-selected from the page context. | C |

### 5.3 Generation
| ID | Requirement | Pri |
|---|---|---|
| FR-09 | Generate is enabled only when both an Account and a template are selected. | M |
| FR-10 | On Generate, the widget fetches the Account record and, for each related module used by the template, all of its related records (paginated, up to 200 per call). | M |
| FR-11 | Each slide definition is rendered according to Section 4 (T1–T10). | M |
| FR-12 | If an iterator yields more rows than Limit of Rows, the slide is split into ⌈rows ÷ limit⌉ slides. Each one repeats the static content, and the title gets the suffix " (n/N)". | M |
| FR-13 | If a related-module slide has zero records, the slide is still produced, with the iterator block replaced by a "No records" row. | S |
| FR-14 | Missing fields, unknown placeholders and extra iterators are collected as warnings and shown with the preview. They do not stop generation. | S |
| FR-15 | A fetch error (permission, network or API limit) shows a readable message and leaves the UI usable. | M |

### 5.3a Images (new in 1.1)
| ID | Requirement | Pri |
|---|---|---|
| FR-25 | After the slides are generated, every image placeholder (T11) that points to a CRM file is replaced by the real picture, downloaded through the JS SDK. | M |
| FR-26 | Each unique file is downloaded once per generation, at most 4 at a time, and at most 50 files per generation. Files beyond the limit are skipped with a warning. | M |
| FR-27 | A picture whose longest side is over 1280 px is scaled down before it goes into the preview and the file, so the `.pptx` stays small. | S |
| FR-28 | A picture that cannot be loaded is left out and reported as a warning (T13). Generation never stops because of an image. | M |
| FR-29 | The preview and the exported `.pptx` show the same pictures. In the `.pptx` each picture is a native picture object that can be moved, resized or replaced. | M |

### 5.4 Preview
| ID | Requirement | Pri |
|---|---|---|
| FR-16 | Generated slides are shown as 16:9 previews in sandboxed containers, so template CSS and scripts cannot affect the widget UI. | M |
| FR-17 | The preview has previous/next navigation, a slide counter ("3 / 7") and a thumbnail strip. | M |
| FR-18 | The preview reflects the same content that is exported, though PowerPoint rendering may differ slightly (see NFR-U3). | M |

### 5.5 Export
| ID | Requirement | Pri |
|---|---|---|
| FR-19 | Download produces a `.pptx` file in 16:9 layout with one PowerPoint slide per generated slide. | M |
| FR-20 | Export converts HTML into native, editable PowerPoint objects: `h1`–`h3`, `p` and `li` become text boxes, `table` becomes a native table, `img` becomes an image (including pictures from CRM fields, FR-25). Supported styles are bold, italic, underline, font colour, background colour, text alignment and font size. | M |
| FR-21 | HTML elements that are not supported are exported as plain text, without failing. | M |
| FR-22 | The file is named `<Account Name> - <Template Name> - <YYYY-MM-DD>.pptx`, with characters that are invalid in file names removed. | S |
| FR-23 | "Save to Account" attaches the generated `.pptx` to the selected Account's Attachments. | S |
| FR-24 | If the browser blocks the download inside the CRM iframe, the widget offers Save to Account or opening the file in a new tab. | S |

---

## 6. External Interface Requirements

### 6.1 User interface
Single page, top to bottom:
1. Header
2. Selection bar: Account search, Template dropdown, Generate button
3. Status and warning area
4. Preview area: main slide, navigation, thumbnails
5. Action bar: Download .pptx, Save to Account

The UI must be usable at widths of 1024 px and above. It follows Zoho CRM's look: neutral colours, system font.

### 6.2 Software interfaces (Zoho CRM JS SDK)
| Use | API |
|---|---|
| Initialise and page context | `ZOHO.embeddedApp.on("PageLoad")`, `ZOHO.embeddedApp.init()` |
| Templates | `ZOHO.CRM.API.getAllRecords({Entity:"Presentation_Templates", page, per_page})` |
| Account search | `ZOHO.CRM.API.searchRecord({Entity:"Accounts", Type:"word", Query})` |
| Account details | `ZOHO.CRM.API.getRecord({Entity:"Accounts", RecordID})` |
| Related data | `ZOHO.CRM.API.getRelatedRecords({Entity:"Accounts", RecordID, RelatedList:<related list API name>, page, per_page})` |
| Attach file | `ZOHO.CRM.API.attachFile({Entity:"Accounts", RecordID, File:{Name, Content}})` |
| Download a picture or file | `ZOHO.CRM.API.getFile({id: <file ID from the field>})`. The name, parameters and return type are to be confirmed against the SDK reference and tested inside CRM (OI-8). |

The exact parameter names must be checked against the current SDK reference during development.

### 6.3 Development interfaces (not used at runtime)
- **Zoho CRM MCP servers** (Data Insights, Data Operations, Module Customization), used from Claude Code to create the modules and fields, seed test data, and check that template HTML is stored intact. They are registered per project under unique names (`zoho-crm-demo-insights`, `zoho-crm-demo-dataops`, `zoho-crm-demo-modules`), so they cannot conflict with connectors for other organisations, and they point only at the Demo org.
- **ZET CLI:** `zet init`, `zet run`, `zet validate`, `zet pack`.

---

## 7. Non-Functional Requirements

| ID | Category | Requirement |
|---|---|---|
| NFR-P1 | Performance | The widget is interactive within 3 s of opening the Web Tab on a normal connection. |
| NFR-P2 | Performance | For a template of 3 slide definitions and 50 related records, generation plus preview finishes within 5 s. Export finishes within 5 s. With 20 pictures of up to 1 MB each, generation plus preview finishes within 10 s. |
| NFR-P3 | Performance | API calls per generation are at most 1 + 1 per related module + extra pages. No per-row calls, except one download per unique picture (FR-26, at most 50). |
| NFR-S1 | Security | All CRM values are HTML-escaped before rendering (T8). |
| NFR-S2 | Security | Preview containers block scripts and isolate template HTML from the widget. |
| NFR-S3 | Security | The widget only reads what the user's CRM profile allows, and stores no data outside CRM. |
| NFR-S4 | Security | No third-party network requests at runtime (C2). |
| NFR-U1 | Usability | A first-time user can generate and download a deck in 4 steps or fewer: search, select, choose template, Generate/Download. |
| NFR-U2 | Usability | Every error message says what went wrong and what to do. |
| NFR-U3 | Usability | The exported PPT keeps text, tables, data and order exactly. Visual styling can be approximate. |
| NFR-R1 | Reliability | Bad template HTML or missing fields never crash generation (FR-14, FR-21). |
| NFR-M1 | Maintainability | The template engine is pure JavaScript with no SDK or DOM dependency, and has automated unit tests. |
| NFR-M2 | Maintainability | Code is separated by concern: CRM data, template engine, HTML→PPTX conversion, UI. |
| NFR-M3 | Maintainability | The template language is documented for template admins (README). |
| NFR-C1 | Compatibility | Exported files open without repair prompts in PowerPoint, Google Slides and LibreOffice Impress. |

---

## 8. Traceability to the Original Brief

| Brief item | Covered by |
|---|---|
| 0. Use the Zoho CRM MCP in Claude Code to build widgets | §6.3, PLAN Phase 0–1 |
| 1. Widgets are added to Web Tabs in CRM | §2.1, C1, PLAN Phase 4 |
| 1. Custom module "Presentation Template" with Template Name, single line, rich text HTML, Limit of rows, Module Name | §3.1 (D1–D5, plus D6 Slide Order) |
| 1. Static + dynamic content via parameters and `#RowCountIterator` | §4 (T1–T10) |
| 2. User picks the record; Generate PPT button; template choice (Temp1, Temp2) | FR-03, FR-06–FR-09 |
| 3. Pull the record's data and generate slides from the template and business rules | FR-10–FR-14 |
| 4. Preview and download | FR-16–FR-24 |
| Mentor feedback: pull images from CRM and place them in the template's placeholders | T11–T13, FR-25–FR-29, D7–D8, AC-10, AC-11 |
| Mentor feedback: base the solution on a real customer's template and generated deck (from Parth) | §2.7, A5, AC-12, OI-7, OI-8 |

---

## 9. Acceptance Criteria

| ID | Scenario | Expected result |
|---|---|---|
| AC-1 | Open the Web Tab with Temp1 and Temp2 defined. | The dropdown shows Temp1 and Temp2. |
| AC-2 | Search "Test" and select the test Account. | The Account shows as selected and Generate is enabled once a template is chosen. |
| AC-3 | Generate Temp1 for an Account with 45 Property Options and 2 Contacts (limit 20). | 5 slides: Introduction; Property Listings (1/3), (2/3), (3/3) with 20/20/5 rows and S.No 1–45; Contact Details. |
| AC-4 | A Property Option name contains `<b>x</b>`. | The text appears literally and is not rendered as bold. |
| AC-5 | A template references `#Non_Existent$`. | The slide renders with a blank value and a warning lists the field. |
| AC-6 | Download. | The file opens in PowerPoint; tables are native and editable; values match CRM. |
| AC-7 | Save to Account. | The `.pptx` appears in the Account's Attachments. |
| AC-8 | The user lacks read access to Property Options. | A clear permission error appears and the rest of the UI still works. |
| AC-9 | Account with zero Property Options. | The Property Listings slide shows a "No records" row. |
| AC-10 | A template slide has `<img src="#Photo$">` in a Property Options iterator, and 3 of the properties have a photo. | Those 3 rows show their own picture in the preview. In the downloaded `.pptx` they are native pictures that can be selected and resized. |
| AC-11 | One property's photo is missing, or the user cannot read the file. | That picture is left out, a warning names the slide and field, and the rest of the deck is generated. |
| AC-12 | Generate the reference deck's template (§2.7) for the same Account as the sample output from Parth. | Same number of slides and order. Text, table data and pictures match the sample. Fonts, colours and exact positions may differ slightly. |

---

## 10. Open Issues

| ID | Issue | Plan |
|---|---|---|
| OI-1 | Zoho's Rich Text field may sanitise or strip HTML. | **Resolved 2026-10-22:** it strips all tags. `Slide_HTML` is now Multi-line (Large); round-trip verified. |
| OI-2 | Downloads may be blocked inside the CRM widget iframe. | Test early; fall back to FR-23/FR-24. |
| OI-3 | The related list API name for Property Options on Accounts may differ from the module API name. | **Resolved:** it is `Property_Options`. |
| OI-4 | The Limit of Rows semantics are confirmed as "continue on new slides". Confirm with the mentor whether overflow slides should repeat the table header. | Default: yes, repeat. |
| OI-5 | Should the main record ever be something other than an Account (e.g. Deals)? | v1 supports Accounts only. The design allows extension. |
| OI-6 | Which CRM org and data centre is the target? | Confirm before Phase 1 when connecting the MCP servers. |
| OI-7 | Customer reference deck: ask Parth for one customer's PowerPoint template and one generated deck. Request drafted 2026-10-23; not yet sent or received. | Rohan sends the request, then waits on Parth. When the files arrive, follow §2.7 and update templates, requirements and AC-12. Slide layout fidelity decisions (slide master, fonts, image positions) are blocked until then. |
| OI-8 | Where do the customer's pictures live: CRM Image Upload fields, Attachments, the Record Image field, or web addresses? Does `getFile` return a Blob inside a widget? | Ask Parth with the files. Meanwhile v1 supports Image Upload and File Upload fields. Test `getFile` in CRM with a real field before relying on it. Other sources are added only if the customer needs them. |
| OI-10 | The exporter only places pictures that stand on their own. A picture inside a paragraph, heading, list item or table cell (for example a photo column in a table, or a logo beside a title) is dropped from the `.pptx`; the preview shows it. | Found while implementing 1.1. Decide after seeing the reference deck (OI-7) whether to support side-by-side layouts and pictures in table cells. For now templates keep each `<img>` on its own line (T11). |
| OI-9 | Picture size and speed: is a 1280 px limit and 50 pictures per generation right for the customer's decks? | Defaults as in FR-26 and FR-27. Adjust after measuring with the reference deck. |
