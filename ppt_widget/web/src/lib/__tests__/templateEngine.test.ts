import { describe, expect, it } from "vitest";
import { buildDeck, deckFileName, escapeHtml, formatValue, groupTemplates, relatedModules } from "../templateEngine";
import type { CrmRecord, Template, TemplateRow } from "../types";

const row = (over: Partial<TemplateRow>): TemplateRow => ({
  id: over.id ?? Math.random().toString(36).slice(2),
  Name: "Temp1",
  Slide_Title: "Slide",
  Slide_HTML: "",
  Limit_of_Rows: null,
  Module_Name: "Accounts",
  Slide_Order: 1,
  ...over,
});

const account: CrmRecord = { id: "a1", Account_Name: "Test Realty", Owner: { name: "Asha", id: "u1" }, Industry: null };
const props = (n: number): CrmRecord[] =>
  Array.from({ length: n }, (_, i) => ({ id: `p${i + 1}`, Name: `Prop ${i + 1}`, Price: 1000 * (i + 1) }));

const listing = `<h2>Listings for #Account.Account_Name$</h2><table><tr><th>No</th><th>Name</th></tr>#RowCountIterator<tr><td>#Sno$</td><td>#Name$</td></tr>$RowCountIterator</table>`;

describe("groupTemplates", () => {
  it("groups by name, sorts names and orders slides by Slide_Order then created time", () => {
    const rows = [
      row({ id: "3", Name: "Temp2", Slide_Order: 1 }),
      row({ id: "2", Name: "Temp1", Slide_Order: 2 }),
      row({ id: "1", Name: "Temp1", Slide_Order: 1, Created_Time: "2026-01-02" }),
      row({ id: "0", Name: "Temp1", Slide_Order: 1, Created_Time: "2026-01-01" }),
      row({ id: "4", Name: "Temp10", Slide_Order: null }),
      row({ id: "5", Name: "  " }),
    ];
    const t = groupTemplates(rows);
    expect(t.map((x) => x.name)).toEqual(["Temp1", "Temp2", "Temp10"]);
    expect(t[0].slides.map((s) => s.id)).toEqual(["0", "1", "2"]);
  });

  it("lists the related modules a template needs", () => {
    const t: Template = { name: "T", slides: [row({}), row({ Module_Name: "Contacts" }), row({ Module_Name: "Property_Options" }), row({ Module_Name: "Contacts" })] };
    expect(relatedModules(t).sort()).toEqual(["Contacts", "Property_Options"]);
  });
});

describe("placeholders", () => {
  it("fills Account fields and lookup sub-fields (T1, T2)", () => {
    const t: Template = { name: "T", slides: [row({ Slide_HTML: "<h1>#Account_Name$</h1><p>by #Owner.name$</p>" })] };
    const { slides, warnings } = buildDeck({ template: t, account, related: {} });
    expect(slides[0].html).toBe("<h1>Test Realty</h1><p>by Asha</p>");
    expect(warnings).toEqual([]);
  });

  it("escapes CRM values (T8, AC-4)", () => {
    const t: Template = { name: "T", slides: [row({ Slide_HTML: "<p>#Account_Name$</p>" })] };
    const { slides } = buildDeck({ template: t, account: { ...account, Account_Name: "Escape <b>x</b> & 'y'" }, related: {} });
    expect(slides[0].html).toBe("<p>Escape &lt;b&gt;x&lt;/b&gt; &amp; &#39;y&#39;</p>");
  });

  it("leaves missing fields blank and warns (AC-5)", () => {
    const t: Template = { name: "T", slides: [row({ Slide_Title: "Intro", Slide_HTML: "<p>[#Non_Existent$] [#Industry$]</p>" })] };
    const { slides, warnings } = buildDeck({ template: t, account, related: {} });
    expect(slides[0].html).toBe("<p>[] []</p>"); // Industry exists but is null: no warning
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ slide: "Intro" });
    expect(warnings[0].message).toContain("Non_Existent");
  });

  it("does not warn for fields known from metadata but absent on the record", () => {
    const t: Template = { name: "T", slides: [row({ Slide_HTML: "#Website$" })] };
    const { warnings } = buildDeck({ template: t, account, related: {}, format: { fieldTypes: { Accounts: { Website: "website" } } } });
    expect(warnings).toEqual([]);
  });

  it("leaves text that isn't a placeholder alone (T10)", () => {
    const html = `<p style="color:#1f4e79">50% off # nothing $ here #a b$</p>`;
    const { slides } = buildDeck({ template: { name: "T", slides: [row({ Slide_HTML: html })] }, account, related: {} });
    expect(slides[0].html).toBe(html);
  });
});

describe("iterator and pagination", () => {
  it("45 rows with limit 20 give 3 slides (20/20/5) with continuous S.No (AC-3)", () => {
    const t: Template = { name: "T", slides: [row({ Slide_Title: "Property Listings", Module_Name: "Property_Options", Limit_of_Rows: 20, Slide_HTML: listing })] };
    const { slides, warnings } = buildDeck({ template: t, account, related: { Property_Options: props(45) } });
    expect(warnings).toEqual([]);
    expect(slides.map((s) => s.title)).toEqual(["Property Listings (1/3)", "Property Listings (2/3)", "Property Listings (3/3)"]);
    const rows = slides.map((s) => (s.html.match(/<tr><td>/g) ?? []).length);
    expect(rows).toEqual([20, 20, 5]);
    expect(slides[1].html).toContain("<td>21</td><td>Prop 21</td>");
    expect(slides[2].html).toContain("<td>45</td><td>Prop 45</td>");
    // Static parts (heading, table header) repeat on every slide.
    for (const s of slides) {
      expect(s.html.startsWith("<h2>Listings for Test Realty</h2><table><tr><th>No</th>")).toBe(true);
      expect(s.html.endsWith("</table>")).toBe(true);
    }
  });

  it("no limit puts every row on one slide without a suffix", () => {
    const t: Template = { name: "T", slides: [row({ Slide_Title: "All", Module_Name: "Property_Options", Slide_HTML: listing })] };
    const { slides } = buildDeck({ template: t, account, related: { Property_Options: props(30) } });
    expect(slides).toHaveLength(1);
    expect(slides[0].title).toBe("All");
    expect((slides[0].html.match(/<tr><td>/g) ?? []).length).toBe(30);
  });

  it("zero related records gives one slide with a No records row (FR-13, AC-9)", () => {
    const t: Template = { name: "T", slides: [row({ Module_Name: "Property_Options", Limit_of_Rows: 20, Slide_HTML: listing })] };
    const { slides } = buildDeck({ template: t, account, related: { Property_Options: [] } });
    expect(slides).toHaveLength(1);
    expect(slides[0].html).toContain(`<tr><td colspan="2">No records</td></tr>`);
  });

  it("outside an iterator, related slides use the first record (T5)", () => {
    const t: Template = { name: "T", slides: [row({ Module_Name: "Contacts", Slide_HTML: "<p>#Full_Name$</p>" })] };
    const { slides } = buildDeck({ template: t, account, related: { Contacts: [{ Full_Name: "Asha" }, { Full_Name: "Vikram" }] } });
    expect(slides[0].html).toBe("<p>Asha</p>");
  });

  it("expands only the first iterator and warns (T7)", () => {
    const html = "#RowCountIterator<i>#Name$</i>$RowCountIterator|#RowCountIterator<b>#Name$</b>$RowCountIterator";
    const t: Template = { name: "T", slides: [row({ Module_Name: "Property_Options", Slide_HTML: html })] };
    const { slides, warnings } = buildDeck({ template: t, account, related: { Property_Options: props(2) } });
    expect(slides[0].html).toBe("<i>Prop 1</i><i>Prop 2</i>|<b>Prop 1</b>");
    expect(warnings.some((w) => w.message.includes("first #RowCountIterator"))).toBe(true);
  });

  it("warns about #Sno$ outside an iterator and about unloaded modules", () => {
    const t: Template = { name: "T", slides: [row({ Module_Name: "Contacts", Slide_HTML: "#Sno$" })] };
    const { warnings } = buildDeck({ template: t, account, related: {} });
    expect(warnings.map((w) => w.message).join(" ")).toMatch(/Contacts[\s\S]*Sno|Sno[\s\S]*Contacts/);
  });

  it("builds a full deck in template order", () => {
    const t: Template = {
      name: "Temp1",
      slides: [
        row({ Slide_Title: "Introduction", Slide_HTML: "<h1>#Account_Name$</h1>" }),
        row({ Slide_Title: "Property Listings", Module_Name: "Property_Options", Limit_of_Rows: 20, Slide_HTML: listing }),
        row({ Slide_Title: "Contact Details", Module_Name: "Contacts", Slide_HTML: "#RowCountIterator<p>#Full_Name$</p>$RowCountIterator" }),
      ],
    };
    const { slides } = buildDeck({ template: t, account, related: { Property_Options: props(45), Contacts: [{ Full_Name: "A" }, { Full_Name: "B" }] } });
    expect(slides.map((s) => s.title)).toEqual([
      "Introduction",
      "Property Listings (1/3)",
      "Property Listings (2/3)",
      "Property Listings (3/3)",
      "Contact Details",
    ]);
  });
});

describe("formatValue (T9)", () => {
  it("formats empties, booleans, lookups, lists, currency and dates", () => {
    expect(formatValue(null, undefined)).toBe("");
    expect(formatValue(true, undefined)).toBe("Yes");
    expect(formatValue(false, undefined)).toBe("No");
    expect(formatValue({ name: "Asha", id: "1" }, undefined)).toBe("Asha");
    expect(formatValue(["A", "B"], undefined)).toBe("A, B");
    expect(formatValue(3237000, "currency", { currencySymbol: "₹", locale: "en-IN" })).toBe("₹32,37,000");
    expect(formatValue(1441, "double", { locale: "en-US" })).toBe("1,441");
    expect(formatValue("2026-10-22", "date", { dateFormat: "dd/MM/yyyy" })).toBe("22/10/2026");
    expect(formatValue("2026-10-22", "date")).toBe("2026-10-22");
    expect(formatValue("2026-10-22T16:41:11+05:30", "datetime", { dateFormat: "MMM d, yyyy" })).toBe("Oct 22, 2026 16:41");
    expect(formatValue("+91 98450 11111", "phone")).toBe("+91 98450 11111");
  });
});

describe("helpers", () => {
  it("escapeHtml", () => expect(escapeHtml(`<a href="x">&</a>`)).toBe("&lt;a href=&quot;x&quot;&gt;&amp;&lt;/a&gt;"));
  it("deckFileName removes invalid characters (FR-22)", () => {
    expect(deckFileName('Test: Realty/"Pvt"', "Temp1", new Date(2026, 9, 22))).toBe("Test RealtyPvt - Temp1 - 2026-10-22.pptx");
  });
});
