import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  ANAROCK_ROLES, SUMMARY_ROWS_PER_SLIDE, buildAnarockSpecs, formatCoverDate, type AnarockInput, type AnarockOption,
} from "../pptxFill/anarockDeck";
import { outlineSlideXml } from "../pptxFill/outline";
import { fillFirstTable, fillTemplate, fillTokens, fitInside, imageSize, parseDataUrl } from "../pptxFill/fillTemplate";

// A 1x1 PNG and a 2x1 PNG header-only size check.
const PNG_1X1 =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

const cell = (text: string) =>
  `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/>${text ? `<a:r><a:rPr lang="en-US" sz="900"/><a:t>${text}</a:t></a:r>` : ""}<a:endParaRPr lang="en-IN" sz="900"/></a:p></a:txBody><a:tcPr/></a:tc>`;
const row = (...cells: string[]) => `<a:tr h="100">${cells.map(cell).join("")}</a:tr>`;
const table = (rows: string[]) => `<a:tbl><a:tblGrid/>${rows.join("")}</a:tbl>`;
const texts = (xml: string) => [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]);

describe("fillTokens", () => {
  it("escapes values and blanks missing ones", () => {
    expect(fillTokens("<a:t>{{a}}</a:t><a:t>{{b}}</a:t>", { a: "<b>x</b> & y" })).toBe("<a:t>&lt;b&gt;x&lt;/b&gt; &amp; y</a:t><a:t></a:t>");
  });
  it("flattens line breaks", () => {
    expect(fillTokens("{{a}}", { a: "one\ntwo" })).toBe("one two");
  });
});

describe("fillFirstTable", () => {
  const xml = `<p:x>${table([row("#", "Name"), row("1", ""), row("2", ""), row("3", "")])}</p:x>`;
  it("keeps the header and builds one row per entry, numbered", () => {
    const out = fillFirstTable(xml, [["A"], ["B & C"]]);
    expect(texts(out)).toEqual(["#", "Name", "1", "A", "2", "B &amp; C"]);
    expect((out.match(/<a:tr\b/g) ?? []).length).toBe(3);
  });
  it("can add more rows than the template has", () => {
    const out = fillFirstTable(xml, Array.from({ length: 6 }, (_, i) => [`R${i}`]));
    expect((out.match(/<a:tr\b/g) ?? []).length).toBe(7);
    expect(texts(out).slice(-2)).toEqual(["6", "R5"]);
  });
  it("leaves only the header when there are no rows", () => {
    expect((fillFirstTable(xml, []).match(/<a:tr\b/g) ?? []).length).toBe(1);
  });
});

describe("pictures", () => {
  it("reads the size of a PNG", () => {
    const p = parseDataUrl(PNG_1X1)!;
    expect(p.ext).toBe("png");
    expect(imageSize(p.bytes, "png")).toEqual({ w: 1, h: 1 });
  });
  it("rejects other formats", () => {
    expect(parseDataUrl("data:image/svg+xml;base64,AAAA")).toBeNull();
  });
  it("fits a wide picture inside a frame, centred", () => {
    expect(fitInside({ x: 0, y: 0, cx: 1000, cy: 1000 }, { w: 200, h: 100 })).toEqual({ x: 0, y: 250, cx: 1000, cy: 500 });
    expect(fitInside({ x: 10, y: 10, cx: 1000, cy: 400 }, { w: 100, h: 200 })).toEqual({ x: 410, y: 10, cx: 200, cy: 400 });
  });
});

describe("formatCoverDate", () => {
  it("adds ordinals like the template", () => {
    expect(formatCoverDate(new Date(2026, 5, 2))).toBe("2nd June 2026");
    expect(formatCoverDate(new Date(2026, 9, 26))).toBe("26th October 2026");
    expect(formatCoverDate(new Date(2026, 0, 11))).toBe("11th January 2026");
    expect(formatCoverDate(new Date(2026, 0, 23))).toBe("23rd January 2026");
  });
});

// ---- a minimal two-slide template, built in memory
async function miniTemplate(): Promise<Uint8Array> {
  const z = new JSZip();
  z.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="x"/><Override PartName="/ppt/slides/slide2.xml" ContentType="x"/></Types>`);
  z.file("ppt/presentation.xml", `<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="2147481558" r:id="rId2"/><p:sldId id="2147481552" r:id="rId3"/></p:sldIdLst><p:extLst><p14:sectionLst><p14:section><p14:sldIdLst><p14:sldId id="2147481558"/><p14:sldId id="2147481552"/></p14:sldIdLst></p14:section></p14:sectionLst></p:extLst></p:presentation>`);
  z.file("ppt/_rels/presentation.xml.rels", `<Relationships><Relationship Id="rId1" Type="http://x/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`);
  z.file("ppt/slides/slide1.xml", `<p:sld xmlns:p="p" show="0"><a:t>Hello {{name}}</a:t><p:pic><p:nvPicPr><p:cNvPr id="3" name="Logo"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr></p:pic></p:sld>`);
  z.file("ppt/slides/_rels/slide1.xml.rels", `<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`);
  z.file("ppt/slides/slide2.xml", `<p:sld xmlns:p="p" show="0"><a:t>Second</a:t>${table([row("#", "Name"), row("1", ""), row("2", "")])}</p:sld>`);
  z.file("ppt/slides/_rels/slide2.xml.rels", `<Relationships/>`);
  z.file("docProps/app.xml", `<Properties><Slides>2</Slides><Notes>2</Notes><HiddenSlides>2</HiddenSlides></Properties>`);
  return z.generateAsync({ type: "uint8array" });
}

describe("fillTemplate", () => {
  it("clones, orders and fills slides, and fixes ids, sections and hidden flags", async () => {
    const { data, warnings } = await fillTemplate(await miniTemplate(), ["one", "two"], [
      { role: "one", tokens: { name: "A & B" }, images: { Logo: PNG_1X1 } },
      { role: "two", table: [["x"]] },
      { role: "one", tokens: { name: "C" } },
    ]);
    expect(warnings).toEqual([]);
    const z = await JSZip.loadAsync(data);
    const slides = Object.keys(z.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort();
    expect(slides).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml", "ppt/slides/slide3.xml"]);
    const s1 = await z.file("ppt/slides/slide1.xml")!.async("string");
    const s3 = await z.file("ppt/slides/slide3.xml")!.async("string");
    expect(s1).toContain("Hello A &amp; B");
    expect(s3).toContain("Hello C");
    expect(s1).not.toContain('show="0"');
    expect(s1).toContain('r:embed="rIdFill1"');
    expect(Object.keys(z.files).some((f) => f.startsWith("ppt/media/fill-1-1"))).toBe(true);
    expect(await z.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).toContain("fill-1-1.png");
    const pres = await z.file("ppt/presentation.xml")!.async("string");
    expect(pres.match(/<p:sldId id="(\d+)"/g)).toEqual(['<p:sldId id="256"', '<p:sldId id="257"', '<p:sldId id="258"']);
    expect(pres).toContain('<p14:sldId id="256"/><p14:sldId id="257"/><p14:sldId id="258"/>');
    expect(await z.file("[Content_Types].xml")!.async("string")).toContain("/ppt/slides/slide3.xml");
    expect(await z.file("docProps/app.xml")!.async("string")).toContain("<Slides>3</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides>");
  });

  it("swaps the picture with that name, not an earlier icon or a group that shares the name", async () => {
    const z = new JSZip();
    const icon = `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Icon"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic>`;
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Plan"/></p:nvGrpSpPr></p:grpSp>`;
    const target = `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Plan"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr></p:pic>`;
    z.file("[Content_Types].xml", `<Types><Default Extension="png" ContentType="image/png"/></Types>`);
    z.file("ppt/presentation.xml", `<p:presentation xmlns:r="r"><p:sldIdLst><p:sldId id="300" r:id="rId2"/></p:sldIdLst></p:presentation>`);
    z.file("ppt/_rels/presentation.xml.rels", `<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`);
    z.file("ppt/slides/slide1.xml", `<p:sld>${icon}${group}${target}</p:sld>`);
    z.file("ppt/slides/_rels/slide1.xml.rels", `<Relationships/>`);
    const { data, warnings } = await fillTemplate(await z.generateAsync({ type: "uint8array" }), ["s"], [{ role: "s", images: { Plan: PNG_1X1 } }]);
    expect(warnings).toEqual([]);
    const out = await (await JSZip.loadAsync(data)).file("ppt/slides/slide1.xml")!.async("string");
    expect(out).toContain('name="Icon"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/>'); // icon untouched
    expect(out).toContain('name="Plan"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFill1"/>');
  });

  it("warns when a picture is missing and keeps going", async () => {
    const { warnings } = await fillTemplate(await miniTemplate(), ["one", "two"], [{ role: "one", images: { Nope: PNG_1X1 } }]);
    expect(warnings[0]).toMatch(/Nope/);
  });

  it("rejects an unexpected template", async () => {
    await expect(fillTemplate(await miniTemplate(), ["only-one"], [])).rejects.toThrow(/2 slides/);
    await expect(fillTemplate(await miniTemplate(), ["one", "two"], [{ role: "zzz" }])).rejects.toThrow(/zzz/);
  });
});

describe("outlineSlideXml", () => {
  it("lists text boxes and tables in order, unescaped, skipping empty shapes", () => {
    const xml = `<p:sld><p:sp><a:p><a:r><a:t>Option 1: A &amp; B</a:t></a:r></a:p></p:sp><p:sp><a:p/></p:sp>` +
      `<p:graphicFrame>${table([row("#", "Name"), row("1", "315 LARKINS"), row("", "")])}</p:graphicFrame></p:sld>`;
    const o = outlineSlideXml(xml, 4);
    expect(o.title).toBe("Option 1: A & B");
    expect(o.blocks).toEqual([
      { kind: "text", lines: ["Option 1: A & B"] },
      { kind: "table", rows: [["#", "Name"], ["1", "315 LARKINS"]] },
    ]);
  });
});

describe("buildAnarockSpecs", () => {
  const mk = (kind: AnarockOption["kind"], i: number): AnarockOption => ({ kind, name: `Bldg ${i}`, operator: "Op", microMarket: "MM", rent: "325", cam: "11.5", area: "27,003", highlights: ["a", "b"] });
  const input: AnarockInput = {
    client: "Test Realty", requirementType: "Conventional", date: new Date(2026, 9, 26),
    user: { name: "U", email: "u@x", phone: "1" },
    options: [...Array.from({ length: 12 }, (_, i) => mk("conventional", i + 1)), mk("coworking", 1)],
  };
  const specs = buildAnarockSpecs(input);

  it("orders the deck and splits summary rows into groups of 10", () => {
    expect(SUMMARY_ROWS_PER_SLIDE).toBe(10);
    expect(specs.map((s) => s.role)).toEqual([
      "cover", "mapConv", "sumConv", "sumConv", ...Array(12).fill("detConv"), "mapCow", "sumCow", "detCow", "closing",
    ]);
    expect(specs.filter((s) => s.role === "sumConv").map((s) => s.table!.length)).toEqual([10, 2]);
  });
  it("formats units and titles like the template", () => {
    const d = specs.find((s) => s.role === "detConv")!;
    expect(d.tokens).toMatchObject({ title: "Option 1: Bldg 1", rent: "INR 325/SF/Month", maintenance: "INR 11.5 /SF/Month", area: "27,003 SF", hl1: "a", hl2: "b", hl3: "" });
    expect(specs.find((s) => s.role === "detCow")!.tokens).toMatchObject({ title: "Option 1: Op, Bldg 1" });
    expect(specs[0].tokens).toEqual({ date: "26th October 2026", client: "Test Realty", type: "Conventional" });
  });
  it("skips a type that has no options", () => {
    const only = buildAnarockSpecs({ ...input, options: [mk("coworking", 1)] }).map((s) => s.role);
    expect(only).toEqual(["cover", "mapCow", "sumCow", "detCow", "closing"]);
  });
});

// Real template (Anarock's own file, git-ignored): runs only where it exists. PPTX_OUT=<dir> writes the deck for viewing.
const REAL = resolve(__dirname, "../../../../templates/anarock-template.pptx");
describe.skipIf(!existsSync(REAL))("real Anarock template", () => {
  it("fills every slide role and produces a well-formed package", async () => {
    const conv = (i: number): AnarockOption => ({
      kind: "conventional", name: i === 1 ? "315 LARKINS" : `Prestige Tech Park ${i}`, city: "Thane", microMarket: "Thane", structure: "2B + GF + 8F", floor: "8",
      floorPlate: "780,000", area: "27,003", rent: "325", cam: "11.5", handover: "Warm Shell", availability: "Immediate", status: "Ready",
      ownership: "Institutional", developer: "Medanta", projectSize: "33,800,000", efficiency: "69.5%", hvac: "Chiller / PVRs", year: "2030",
      power: "90% Back-up", cert: "USGBC WEED Platinum", elevators: "7 Passenger + 4 Service Lifts", metro: "1.5 km from Dadar East",
      airport: "25.4 km from Chhatrapati Shivaji International Airport", leaseTerm: "3 Years", lockin: "2 Years (Warm Shell)",
      opHours: "16 Hours X 3.5 days", escalation: "7.5% YoY on last paid rent", parkingRatio: "2: 1000", parkingCharges: "9,300",
      addnParking: "INR 6,300 - Subject to Availability", deposit: "8 Months",
      highlights: ["Shuttle service till the nearest metro station", "Amenities like Daycare", "Gym", "Auditorium"],
      tenants: ["Mr. Arjun Sharma", "Mr. Sachin Tendulkar", "Ms. Priya Patel", "Mr. Rahul Dravid", "Ms. Anita Desai", "Mr. Virat Kohli"],
      buildingImage: PNG_1X1, floorImage: PNG_1X1,
    });
    const cow: AnarockOption = {
      kind: "coworking", operator: "WeWork", name: "Galaxy Tower", city: "Rajendra Place", microMarket: "Rajendra Place", floor: "6F", seats: "175",
      price: "33,754", wsSize: "8 ft x 2 ft", lockin: "16 Months", escalation: "8.7% per annum", centerArea: "69,690 sq. ft", capacity: "876 Workstations",
      density: "55 SF", opHours: "9 AM-8 PM (Mon-Fri)\n9 AM-2 PM (Sat)", configuration: "125 W.S + 3 Cabins + 2 Meeting Room",
      arrangement: "Scattered / Contiguous", availability: "Immediate", metro: "200m from Rajiv Chowk Metro", airport: "18 km from IGI Airport",
      developer: "Anarock", year: "2029", grade: "A", cert: "USGBC WEED Platinum", elevators: "8 Passenger + 4 Service Lifts",
      notice: "2 Month included in Lock-in", deposit: "6 Months", setupCost: "None", parkingSlots: "4", parkingCharges: "10,517",
      modifications: "Depends on Requirement", credits: "2-Meeting room / 22-Printing (per seat)", meetingRoomCharges: "TBD",
      highlights: ["High-speed Wi-Fi", "Unlimited Coffee", "Meeting Room Credits", "Printing Services"], projectHighlights: ["Prime CBD Location", "Grade A Building", "24/7 Access"],
    };
    const specs = buildAnarockSpecs({
      client: "Test Realty Pvt Ltd", requirementType: "Conventional", date: new Date(2026, 9, 26),
      user: { name: "Rohan", email: "rohan@example.com", phone: "+91 90000 00000" }, mapImage: PNG_1X1,
      options: [conv(1), conv(2), conv(3), cow],
    });
    const { data, warnings } = await fillTemplate(readFileSync(REAL), ANAROCK_ROLES, specs);
    expect(warnings).toEqual([]);
    if (process.env.PPTX_OUT) writeFileSync(`${process.env.PPTX_OUT}/anarock-filled.pptx`, data);
    const z = await JSZip.loadAsync(data);
    const slideXml = await Promise.all(
      Object.keys(z.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).map((f) => z.file(f)!.async("string")),
    );
    expect(slideXml).toHaveLength(specs.length);
    expect(slideXml.join("")).not.toMatch(/\{\{[A-Za-z0-9_]+\}\}/); // every token filled
    expect(slideXml.join("")).not.toContain('show="0"');
  });
});
