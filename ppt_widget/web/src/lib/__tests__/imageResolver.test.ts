import { describe, expect, it, vi } from "vitest";
import { MAX_IMAGES, fileRef, resolveImages } from "../imageResolver";
import { buildDeck } from "../templateEngine";
import type { GeneratedSlide, Template } from "../types";

const slide = (html: string, title = "S"): GeneratedSlide => ({ title, html, templateRowId: "t1" });
const PNG = "data:image/png;base64,AAAA";

describe("fileRef", () => {
  it("reads the first file id of an Image Upload value", () => {
    expect(fileRef([{ file_Id: "abc123", File_Name: "a.png" }, { file_Id: "zzz" }])).toBe("crmfile:abc123");
  });
  it("reads the real CRM API v8 shape, ignoring the attachment row id", () => {
    expect(fileRef([{ File_Id__s: "h34loc8e193b", File_Name__s: "a.png", id: "7644593000000773034" }])).toBe("crmfile:h34loc8e193b");
  });
  it("does not use the attachment row id when there is no file id", () => {
    expect(fileRef([{ id: "7644593000000773034" }])).toBe("");
  });
  it("returns empty for no file", () => {
    expect(fileRef([])).toBe("");
    expect(fileRef(null)).toBe("");
    expect(fileRef("https://example.com/a.png")).toBe("");
  });
});

describe("engine: image fields (T12)", () => {
  const template: Template = {
    name: "T",
    slides: [{ id: "1", Name: "T", Slide_Title: "Gallery", Slide_HTML: '<img src="#Photo$">', Limit_of_Rows: null, Module_Name: "Property_Options", Slide_Order: 1 }],
  };
  it("emits a crmfile reference for imageupload fields", () => {
    const out = buildDeck({
      template,
      account: { id: "a" },
      related: { Property_Options: [{ id: "p", Photo: [{ file_Id: "f1" }] }] },
      format: { fieldTypes: { Property_Options: { Photo: "imageupload" } } },
    });
    expect(out.slides[0].html).toBe('<img src="crmfile:f1">');
  });
});

describe("resolveImages", () => {
  it("replaces references with data URLs and downloads each file once (FR-26)", async () => {
    const load = vi.fn(async () => PNG);
    const { slides, warnings } = await resolveImages(
      [slide('<img src="crmfile:f1" width="120"><img src="crmfile:f1">'), slide('<img src="crmfile:f2">')],
      load,
    );
    expect(load).toHaveBeenCalledTimes(2);
    expect(slides[0].html).toBe(`<img src="${PNG}" width="120"><img src="${PNG}">`);
    expect(slides[1].html).toBe(`<img src="${PNG}">`);
    expect(warnings).toEqual([]);
  });

  it("removes a picture that fails to load and warns, keeping the rest of the slide (T13, FR-28)", async () => {
    const load = vi.fn(async (id: string) => {
      if (id === "bad") throw new Error("NO_PERMISSION");
      return PNG;
    });
    const { slides, warnings } = await resolveImages([slide('<p>Hi</p><img src="crmfile:bad"><img src="crmfile:ok">', "Gallery")], load);
    expect(slides[0].html).toBe(`<p>Hi</p><img src="${PNG}">`);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ slide: "Gallery" });
    expect(warnings[0].message).toContain("NO_PERMISSION");
  });

  it("removes a picture with an empty field and warns", async () => {
    const { slides, warnings } = await resolveImages([slide('<img src=""><p>x</p>')], vi.fn());
    expect(slides[0].html).toBe("<p>x</p>");
    expect(warnings[0].message).toMatch(/empty/);
  });

  it("leaves ordinary web addresses and data URLs for the exporter", async () => {
    const load = vi.fn();
    const html = `<img src="https://example.com/a.png"><img src="${PNG}">`;
    const { slides, warnings } = await resolveImages([slide(html)], load);
    expect(slides[0].html).toBe(html);
    expect(load).not.toHaveBeenCalled();
    expect(warnings).toEqual([]);
  });

  it("loads at most MAX_IMAGES files and warns for the rest (FR-26)", async () => {
    const load = vi.fn(async () => PNG);
    const html = Array.from({ length: MAX_IMAGES + 2 }, (_, i) => `<img src="crmfile:f${i}">`).join("");
    const { slides, warnings } = await resolveImages([slide(html)], load);
    expect(load).toHaveBeenCalledTimes(MAX_IMAGES);
    expect((slides[0].html.match(/<img/g) ?? []).length).toBe(MAX_IMAGES);
    expect(warnings).toHaveLength(1);
    expect(warnings[0].message).toContain(String(MAX_IMAGES));
  });

  it("never runs more than 4 downloads at once (FR-26)", async () => {
    let running = 0;
    let peak = 0;
    const load = async () => {
      peak = Math.max(peak, ++running);
      await new Promise((r) => setTimeout(r, 5));
      running--;
      return PNG;
    };
    const html = Array.from({ length: 12 }, (_, i) => `<img src="crmfile:f${i}">`).join("");
    await resolveImages([slide(html)], load);
    expect(peak).toBeLessThanOrEqual(4);
  });

  it("strips a stray file reference used outside <img src>", async () => {
    const { slides, warnings } = await resolveImages([slide("<p>crmfile:abc</p>")], vi.fn());
    expect(slides[0].html).toBe("<p></p>");
    expect(warnings).toHaveLength(1);
  });
});
