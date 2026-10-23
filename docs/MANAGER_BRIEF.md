# Presentation Builder: Manager Brief

Talking points and a demo script for a 5–10 minute project update.

| | |
|---|---|
| Date | 2026-10-22 |
| Author | Rohan |
| Related | [PLAN.md](PLAN.md), [SRS.md](SRS.md), [README](../README.md) |

---

## 1. One-line summary
> "I built a Zoho CRM widget that turns an Account's data into an editable PowerPoint deck. You pick an Account and a template, preview the slides, and download the .pptx or save it to the Account."

## 2. The problem it solves
- Sales teams build client decks by hand: copying Account details, property listings and contacts into PowerPoint.
- This does it in one click, using live CRM data, with a consistent format every time.

## 3. How it works
1. **Templates live in CRM.** A new module, *Presentation Templates*, holds each slide as HTML with placeholders like `#Account_Name$`. Admins change templates in CRM with no code change.
2. **Data comes from CRM.** The widget reads the Account, its Property Options and its Contacts.
3. **Long lists split automatically.** With a limit of 20 rows, 45 properties become 3 slides (1/3, 2/3, 3/3) with continuous numbering.
4. **The output is a real PowerPoint.** Tables and text stay editable; they aren't screenshots.

## 4. Live demo (2 minutes)
1. Open the widget. Search **test** and pick **Test Realty Pvt Ltd**.
2. Pick **Temp1** and click **Generate**. Show the 5 slides and click through the 3 listing pages.
3. Switch to **Temp2** and show that the same data produces a different deck (6 slides).
4. Click **Download .pptx**, open it in PowerPoint, and **edit a table cell** to show it's native.
5. *(Optional)* Pick **King (Sample)** and show the "No records" slide, which handles empty data cleanly.

## 5. Status
| Done | Pending |
|---|---|
| CRM modules, fields and test data (set up through Zoho's MCP tools) | Upload to CRM as a Web Tab and test it there |
| Template engine with 22 automated tests | Confirm downloads work inside CRM (*Save to Account* is the fallback) |
| Widget UI: search, preview, download, save to Account | |
| Exported file checked in PowerPoint: opens cleanly, tables editable | |

## 6. Decisions worth mentioning
- **The Rich Text field didn't work.** Zoho stripped all the HTML tags, which I caught with a save-and-read-back test. I switched to a plain Multi-line field and confirmed the HTML now saves intact.
- **Added a "Slide Order" field** so slides always come out in a fixed order. The original spec didn't have one.
- **Values are escaped,** so CRM data can't inject HTML into slides.
- **It runs entirely inside CRM** with the user's own permissions. There's no separate server and no extra cost.

## 7. Questions for the manager
- On overflow slides, should the table header repeat? (It does now.)
- Should it work for other records later, such as Deals, or only Accounts?
- Which users or profiles should get the Web Tab?
- Images: where do customers keep the pictures that go into decks (CRM image fields, attachments, web links)? I'm asking Parth for a real customer template and a generated deck to check this. *(Request not sent yet.)*

---

**Tip:** lead with the demo, not the architecture. Managers remember "click, preview, download" more than the details.
