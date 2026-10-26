import { describe, expect, it } from "vitest";
import { templateFileId, toAnarockOption } from "../pptxFill/anarockFromCrm";
import type { Template, TemplateRow } from "../types";

const row = (extra: Partial<TemplateRow>): TemplateRow => ({
  id: "1", Name: "Anarock", Slide_Title: "x", Slide_HTML: "", Limit_of_Rows: null, Module_Name: "Accounts", Slide_Order: 1, ...extra,
});

describe("toAnarockOption", () => {
  it("maps Office_Options fields and splits list fields into lines", () => {
    const o = toAnarockOption({
      Name: "315 LARKINS", Option_Type: "Conventional", Micro_Market: "Thane", Quoted_Rent: "325", Lock_in_Period: "2 Years",
      Highlights: "Gym\r\nAuditorium\n\n", Tenant_Profile: "Mr. A\nMs. B", Operational_Hours: "9-5\n9-2",
    });
    expect(o).toMatchObject({ kind: "conventional", name: "315 LARKINS", microMarket: "Thane", rent: "325", lockin: "2 Years" });
    expect(o.highlights).toEqual(["Gym", "Auditorium"]);
    expect(o.tenants).toEqual(["Mr. A", "Ms. B"]);
    expect(o.opHours).toBe("9-5\n9-2");
  });
  it("treats Co-working as its own kind and blanks missing or object values", () => {
    const o = toAnarockOption({ Name: "Galaxy", Option_Type: "Co-working", Operator_Name: "WeWork", Quoted_Price: null, Account: { id: "1", name: "x" } });
    expect(o.kind).toBe("coworking");
    expect(o.operator).toBe("WeWork");
    expect(o.price).toBe("");
  });
});

describe("templateFileId", () => {
  const tpl = (file: unknown): Template => ({ name: "Anarock", slides: [row({ Template_File: file })] });
  it("returns the CRM file id of an uploaded template", () => {
    expect(templateFileId(tpl([{ File_Id__s: "abc123", File_Name__s: "t.pptx" }]))).toBe("abc123");
  });
  it("is empty for an HTML template", () => {
    expect(templateFileId(tpl(null))).toBe("");
    expect(templateFileId(tpl(undefined))).toBe("");
  });
});
