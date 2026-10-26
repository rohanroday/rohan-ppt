// Turns Anarock.pptm into a macro-free, tokenised template (ppt_widget/templates/anarock-template.pptx).
//   node scripts/make-anarock-template.mjs "C:/path/Anarock.pptm" [out.pptx]
// - removes the macro project and every macro click action (nothing is ever executed)
// - removes the macro buttons and the speaker notes (tool bookkeeping, not deck content)
// - replaces each sample value on the detail/cover/closing slides with a {{token}} (one run per token)
// The output contains Anarock's logo, photos and addresses: it is git-ignored and uploaded to CRM by hand.

import { createRequire } from "node:module";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const JSZip = createRequire(resolve(here, "../web/package.json"))("jszip");

const src = process.argv[2];
const out = process.argv[3] ?? resolve(here, "../templates/anarock-template.pptx");
if (!src) {
  console.error('usage: node scripts/make-anarock-template.mjs "<Anarock.pptm>" [out.pptx]');
  process.exit(1);
}

const MAIN_MACRO = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
const MAIN_PLAIN = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

// ---- sample value -> token, per slide (matched on the whole paragraph text, whitespace-normalised)
const norm = (s) => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/\s+/g, " ").trim();
const dict = (pairs) => new Map(pairs.map(([text, token]) => [norm(text), token]));

const TOKENS = {
  1: dict([["2nd June 2026", "date"], ["Deal Name", "client"], ["Requirement Type", "type"]]),
  6: dict([
    ["Option 1: 315 LARKINS", "title"], ["Thane", "city"],
    ["Shuttle service till the nearest metro station", "hl1"], ["Amenities like Daycare", "hl2"], ["Gym", "hl3"], ["Auditorium", "hl4"],
    ["Institutional", "ownership"], ["Medanta", "developer"], ["33,800,000", "project_size"], ["780,000", "floor_plate"],
    ["2B + GF + 8F", "structure"], ["69.5%", "efficiency"],
    ["8", "floor"], ["27,003 SF", "area"], ["Warm Shell", "handover"], ["Immediate", "availability"],
    ["1.5 km from Dadar East", "metro"], ["25. 4 km from Chhatrapati Shivaji International Airport", "airport"],
    ["Chiller / PVRs", "hvac"], ["2030", "year"], ["90% Back-up", "power"], ["USGBC WEED Platinum", "cert"],
    ["7 Passenger + 4 Service Lifts", "elevators"],
    ["3 Years", "lease_term"], ["2 Years (Warm Shell)", "lockin"], ["INR 325/SF/Month", "rent"], ["INR 11.5 /SF/Month", "maintenance"],
    ["16 Hours X 3. 5 days", "op_hours"], ["7.5% YoY on last paid rent", "escalation"], ["2: 1000", "parking_ratio"],
    ["INR 9,300/Slot/Month", "parking_charges"], ["INR 6,300 - Subject to Availability", "addn_parking"], ["8 Months", "deposit"],
    ["Mr. Arjun Sharma", "tp1"], ["Mr. Sachin Tendulkar", "tp2"], ["Ms. Priya Patel", "tp3"],
    ["Mr. Rahul Dravid", "tp4"], ["Ms. Anita Desai", "tp5"], ["Mr. Virat Kohli", "tp6"],
  ]),
  7: dict([
    ["Option 1: WeWork, Galaxy Tower", "title"], ["Rajendra Place", "city"],
    ["High-speed Wi-Fi", "ch1"], ["Unlimited Coffee", "ch2"], ["Meeting Room Credits", "ch3"], ["Printing Services", "ch4"],
    ["Prime CBD Location", "ph1"], ["Grade A Building", "ph2"], ["24/7 Access", "ph3"],
    ["69,690 sq. ft", "center_area"], ["876 Workstations", "capacity"], ["6F", "floor"], ["8 ft x 2 ft", "ws_size"],
    ["55 SF", "density"], ["9 AM-8 PM (Mon-Fri)", "op_hours1"], ["9 AM-2 PM (Sat)", "op_hours2"],
    ["175", "seats"], ["125 W.S + 3 Cabins + 2 Meeting Room", "configuration"], ["Scattered / Contiguous", "arrangement"],
    ["Immediate", "availability"],
    ["200m from Rajiv Chowk Metro", "metro"], ["18 km from IGI Airport", "airport"],
    ["Anarock", "developer"], ["2029", "year"], ["A", "grade"], ["USGBC WEED Platinum", "cert"], ["8 Passenger + 4 Service Lifts", "elevators"],
    ["16 Months", "lockin"], ["2 Month included in Lock-in", "notice"], ["6 Months", "deposit"], ["8.7% per annum", "escalation"],
    ["INR 33,754/Seat/Month", "price"], ["None", "setup_cost"], ["4", "parking_slots"], ["INR 10,517 /Slot/Month", "parking_charges"],
    ["Depends on Requirement", "modifications"], ["2-Meeting room / 22-Printing (per seat)", "credits"], ["TBD", "meeting_room_charges"],
  ]),
  8: dict([["User Full Name", "user_name"], ["User Email", "user_email"], ["User Phone Number", "user_phone"]]),
};

// ---- helpers
function rewriteParagraphs(xml, map, used) {
  return xml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (p) => {
    const text = norm((p.match(/<a:t>[\s\S]*?<\/a:t>/g) ?? []).map((t) => t.replace(/<\/?a:t>/g, "")).join(""));
    const token = map.get(text);
    if (!token) return p;
    used.add(token);
    const pPr = p.match(/<a:pPr[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/)?.[0] ?? "";
    const rPr = p.match(/<a:r>\s*(<a:rPr[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>))/)?.[1] ?? "";
    const end = p.match(/<a:endParaRPr[^>]*?(?:\/>|>[\s\S]*?<\/a:endParaRPr>)/)?.[0] ?? "";
    return `<a:p>${pPr}<a:r>${rPr}<a:t>{{${token}}}</a:t></a:r>${end}</a:p>`;
  });
}

function removeShapesByName(xml, pattern) {
  let removed = 0;
  for (;;) {
    const m = pattern.exec(xml);
    if (!m) return [xml, removed];
    const st = xml.lastIndexOf("<p:sp>", m.index);
    const en = xml.indexOf("</p:sp>", m.index) + "</p:sp>".length;
    xml = xml.slice(0, st) + xml.slice(en);
    removed++;
  }
}

// ---- build
const zin = await JSZip.loadAsync(readFileSync(src));
const zout = new JSZip();
const report = [];

for (const name of Object.keys(zin.files)) {
  const entry = zin.files[name];
  if (entry.dir) continue;
  if (name.includes("vbaProject")) continue; // macro project: dropped, never executed
  if (name.startsWith("ppt/notesSlides/")) continue; // tool bookkeeping notes
  let data = await entry.async("nodebuffer");

  if (name === "[Content_Types].xml") {
    let s = data.toString("utf8");
    s = s.replace(/<Default[^>]*vbaProject[^>]*\/>/g, "").replace(/<Override[^>]*vbaProject[^>]*\/>/g, "");
    s = s.replace(/<Override[^>]*notesSlides\/[^>]*\/>/g, "");
    s = s.replace(MAIN_MACRO, MAIN_PLAIN);
    data = Buffer.from(s, "utf8");
  } else if (name === "ppt/_rels/presentation.xml.rels") {
    data = Buffer.from(data.toString("utf8").replace(/<Relationship [^>]*vbaProject[^>]*\/>/g, ""), "utf8");
  } else if (/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(name)) {
    data = Buffer.from(data.toString("utf8").replace(/<Relationship [^>]*\/notesSlide"[^>]*\/>/g, ""), "utf8");
  } else if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) {
    const n = Number(name.match(/slide(\d+)\.xml/)[1]);
    let s = data.toString("utf8");
    s = s.replace(/<a:hlinkClick[^>]*ppaction:\/\/macro[^>]*?(?:\/>|>[\s\S]*?<\/a:hlinkClick>)/g, "");
    const [s2, removed] = removeShapesByName(s, /name="btn[A-Za-z]*"/);
    s = s2;
    if (removed) report.push(`slide${n}: removed ${removed} macro button shape(s)`);
    const map = TOKENS[n];
    if (map) {
      const used = new Set();
      s = rewriteParagraphs(s, map, used);
      const missing = [...map.values()].filter((t) => !used.has(t));
      report.push(`slide${n}: ${used.size}/${map.size} tokens placed${missing.length ? " | NOT FOUND: " + missing.join(", ") : ""}`);
    }
    data = Buffer.from(s, "utf8");
  }
  zout.file(name, data);
}

mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, await zout.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log(report.join("\n"));
console.log("wrote", out);
