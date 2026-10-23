# Presentation Builder: Zoho CRM widget

A Zoho CRM Web Tab widget that turns CRM data into an editable PowerPoint deck. The user picks an Account and a template, previews the slides, then clicks **Download .pptx** or **Save to Account**.

- Requirements: [docs/SRS.md](docs/SRS.md)
- Plan and status log: [docs/PLAN.md](docs/PLAN.md)

## Writing templates (for template admins)

A **template** is every record in **Presentation Templates** that has the same *Template Name* (for example `Temp1`). Each record is one slide definition:

| Field | What to put in it |
|---|---|
| Template Name | The template this slide belongs to, e.g. `Temp1` |
| Slide Title | Shown in the slide footer and the widget's slide list. Overflow slides get " (2/3)" added. |
| Slide HTML | The slide content as raw HTML with placeholders (plain text box, up to 32,000 characters) |
| Limit of Rows | Maximum repeated rows per slide. Extra rows continue on new slides. Leave blank for no limit. |
| Module Name | Where the data comes from: `Accounts`, `Property_Options` or `Contacts` |
| Slide Order | Position in the deck (1, 2, 3 …) |

### Placeholders

| Write | Gets replaced with |
|---|---|
| `#Field_API_Name$` | That field from the slide's Module Name source, e.g. `#Account_Name$`, `#Price$` |
| `#Owner.name$` | A sub-field of a lookup or user field |
| `#Account.Account_Name$` | Always the selected Account, on any slide |
| `#RowCountIterator` … `$RowCountIterator` | The text between the markers repeats once per related record |
| `#Sno$` | Row number 1, 2, 3 … inside the repeated block (continues across overflow slides) |
| `<img src="#Photo$">` | A picture from a CRM Image Upload / File Upload field (first file). Works on Accounts slides and per row inside the repeated block. Put each `<img>` on its own line, not inside a `<p>`, heading or table cell, or it won't be exported to the .pptx. A missing or unreadable picture is left out with a warning. |

Notes:
- Use **API names**, which are case-sensitive (Setup → Developer Hub → APIs → API names).
- On a `Property_Options` or `Contacts` slide, placeholders **outside** the repeated block use the first record.
- Values are inserted as text, so a value like `<b>x</b>` shows literally and never becomes HTML.
- An unknown field is left blank, and the widget lists it under **warnings**.
- If an Account has no related records, the repeated block becomes a single "No records" row.
- Dates use the user's CRM date format and currency uses the org's currency symbol.

### Example (a listing slide)

```html
<h2 style="color:#1f4e79">Property Listings for #Account.Account_Name$</h2>
<table>
<tr style="background-color:#1f4e79;color:#ffffff"><th>S.No</th><th>Property</th><th>Price</th></tr>
#RowCountIterator
<tr><td>#Sno$</td><td>#Name$</td><td>#Price$</td></tr>
$RowCountIterator
</table>
```

### HTML that converts to PowerPoint

| HTML | Becomes in PowerPoint |
|---|---|
| `h1`–`h6`, `p`, `div` text, `blockquote` | Text box |
| `ul` / `ol` + `li` | Bulleted / numbered text box |
| `table`, `tr`, `th`, `td` (incl. colspan/rowspan) | Native, editable table |
| `img src="…"` | Picture (it must be loadable from the browser) |
| `b`/`strong`, `i`/`em`, `u`, `s`, `a href` | Bold, italic, underline, strike-through, link |
| `style="color: … ; background-color: … ; text-align: … ; font-size: …"` | Text colour, cell fill or highlight, alignment, size |

Blocks are stacked top to bottom on a 16:9 slide. Any other tag is exported as plain text. Keep tables to about 20 rows per slide, and use Limit of Rows for longer lists.

## Project layout

```
crm/make-seed.mjs            builds the CRM seed data (also feeds the widget's mock mode)
ppt_widget/                  ZET project (zet validate / zet pack run here)
  plugin-manifest.json
  scripts/build-widget.mjs   next build -> copies web/out into app/ (entry app/widget.html)
  app/                       GENERATED: never edit
  dist/ppt_widget.zip        GENERATED: upload this to CRM
  web/                       Next.js + Tailwind (static export)
    src/app/page.tsx         the widget screen
    src/components/          AccountSearch, SlidePreview, WarningsPanel
    src/lib/templateEngine.ts  placeholders, iterator, pagination (pure TS, unit-tested)
    src/lib/htmlToPptx.ts      HTML -> native PowerPoint objects (PptxGenJS)
    src/lib/slidePreview.ts    preview iframe document (same sizes as the export)
    src/lib/zoho.ts            Zoho SDK wrapper + mock switch
    src/lib/mock/              sample data for local development
```

## Develop

```bash
cd ppt_widget/web
npm install
npm run dev      # http://localhost:3000, runs on mock data outside CRM
npm test         # Vitest: template engine + PowerPoint export
npm run lint
```

The widget uses **mock data** whenever it isn't inside a CRM iframe, or when the URL has `?mock=1`. A yellow "Mock data" badge shows when it is on.

## Package and deploy

```bash
cd ppt_widget
node scripts/build-widget.mjs   # build the static export into app/
zet validate
zet pack                        # -> dist/ppt_widget.zip
```

In Zoho CRM:
1. **Setup → Developer Hub → Widgets → Create New Widget**: Type **Web Tab**, Hosting **Zoho**, upload `dist/ppt_widget.zip`, Index page `/widget.html`. To update an existing widget, edit it and upload the new zip.
2. **Setup → Customization → Web Tabs → New Web Tab**: choose **Widget**, pick this widget, then choose the profiles that can see it.
3. Users need read access to Accounts, Contacts, Property Options and Presentation Templates. **Save to Account** also needs permission to add attachments to Accounts.

To test against CRM without re-uploading each time: run `zet run` in `ppt_widget/`, open https://127.0.0.1:5000 once in the browser and accept the certificate, then create the widget with **External** hosting pointing at `https://127.0.0.1:5000/app/widget.html`.
