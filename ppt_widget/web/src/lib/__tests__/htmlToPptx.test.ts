// @vitest-environment jsdom
import { writeFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildPptx, parseColor, tableFontPt } from "../htmlToPptx";
import seed from "../mock/seed-data.json";
import { buildDeck, groupTemplates } from "../templateEngine";
import type { CrmRecord, TemplateRow } from "../types";

const templates = groupTemplates(
  seed.templates.map((t, i) => ({
    id: String(i),
    Name: t.Name,
    Slide_Title: t.Slide_Title,
    Slide_HTML: t.Slide_HTML,
    Limit_of_Rows: "Limit_of_Rows" in t ? (t.Limit_of_Rows as number) : null,
    Module_Name: t.Module_Name,
    Slide_Order: t.Slide_Order,
  })) as TemplateRow[],
);
const account: CrmRecord = { id: "a1", Owner: { name: "Demo User" }, ...seed.account };
const related = {
  Property_Options: seed.properties as CrmRecord[],
  Contacts: seed.contacts.map((c) => ({ ...c, Full_Name: `${c.First_Name} ${c.Last_Name}` })),
};
const format = { currencySymbol: "₹", locale: "en-IN", fieldTypes: { Property_Options: { Price: "currency" } } };

async function exportDeck(name: string) {
  const template = templates.find((t) => t.name === name)!;
  const deck = buildDeck({ template, account, related, format });
  const { pptx, warnings } = await buildPptx(deck.slides, { title: name });
  const buf = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  if (process.env.PPTX_OUT) writeFileSync(`${process.env.PPTX_OUT}/${name}.pptx`, buf);
  const zip = await JSZip.loadAsync(buf);
  const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]));
  const xml = await Promise.all(slideFiles.map((f) => zip.file(f)!.async("string")));
  return { deck, warnings, xml };
}

describe("htmlToPptx", () => {
  it("exports Temp1 as 5 native slides: 3 listing tables of 20/20/5 rows (AC-3, AC-6)", async () => {
    const { deck, warnings, xml } = await exportDeck("Temp1");
    expect(deck.warnings).toEqual([]);
    expect(warnings).toEqual([]);
    expect(xml).toHaveLength(5);
    // Native tables, header row + data rows.
    const tableRows = xml.map((x) => (x.includes("<a:tbl>") ? (x.match(/<a:tr /g) ?? []).length : 0));
    expect(tableRows).toEqual([0, 21, 21, 6, 3]);
    // Intro: heading text and bullets.
    expect(xml[0]).toContain("Test Realty Pvt Ltd");
    // One bullet per <li> (4), even though each item mixes bold and plain runs.
    expect((xml[0].match(/<a:buChar/g) ?? []).length).toBe(4);
    // Values match CRM, with currency formatting.
    expect(xml[1]).toContain("Whitefield Apartment 01");
    expect(xml[1]).toContain("₹32,37,000");
    // AC-4: markup in data stays literal text (escaped in XML), never bold markup.
    expect(xml[1]).toContain("Escape test &lt;b&gt;x&lt;/b&gt;");
    // Footer carries the overflow title.
    expect(xml[2]).toContain("Property Listings (2/3)");
    expect(xml[4]).toContain("Asha Menon");
  });

  it("exports Temp2 as 6 slides (summary + 5 shortlist pages)", async () => {
    const { xml } = await exportDeck("Temp2");
    expect(xml).toHaveLength(6);
    expect(xml[5]).toContain("Shortlist (5/5)");
  });

  it("keeps inline styles: header fill and white text", async () => {
    const { xml } = await exportDeck("Temp1");
    expect(xml[1]).toContain('<a:srgbClr val="1F4E79"/>');
    expect(xml[1]).toContain('<a:srgbClr val="FFFFFF"/>');
  });

  it("parses colours and sizes table fonts", () => {
    expect(parseColor("#abc")).toBe("AABBCC");
    expect(parseColor("rgb(31, 78, 121)")).toBe("1F4E79");
    expect(parseColor("navy")).toBe("000080");
    expect(parseColor("transparent")).toBeUndefined();
    expect(tableFontPt(21, 5.3)).toBeGreaterThanOrEqual(8);
    expect(tableFontPt(3, 5.3)).toBe(14);
  });

  it("turns unsupported elements into plain text without failing (FR-21)", async () => {
    const { pptx, warnings } = await buildPptx(
      [{ title: "Odd", templateRowId: "x", html: "<marquee>Hello</marquee><custom-tag>World</custom-tag><script>alert(1)</script><svg><text>skip</text></svg>" }],
      { title: "t" },
    );
    const buf = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
    const xml = await (await JSZip.loadAsync(buf)).file("ppt/slides/slide1.xml")!.async("string");
    expect(xml).toContain("Hello");
    expect(xml).toContain("World");
    expect(xml).not.toContain("alert");
    expect(warnings).toEqual([]);
  });
});
