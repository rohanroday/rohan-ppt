// @vitest-environment jsdom
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { buildPptx } from "../htmlToPptx";
import { resolveImages } from "../imageResolver";
import type { GeneratedSlide } from "../types";

// A 1x1 PNG.
const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

// jsdom never loads images, so `new Image()` would never fire onload and the size lookup would hang.
class FakeImage {
  onload: (() => void) | null = null;
  naturalWidth = 400;
  naturalHeight = 300;
  set src(_: string) {
    setTimeout(() => this.onload?.(), 0);
  }
}

describe("pictures in the exported .pptx (FR-29)", () => {
  it("turns a resolved CRM picture into a native picture object", async () => {
    vi.stubGlobal("Image", FakeImage);
    const slides: GeneratedSlide[] = [
      { title: "Gallery", templateRowId: "t", html: '<h2>Gallery</h2><img src="crmfile:f1" width="120"><p>Caption</p>' },
    ];
    const resolved = await resolveImages(slides, async () => PNG);
    expect(resolved.warnings).toEqual([]);

    const { pptx, warnings } = await buildPptx(resolved.slides, { title: "t" });
    expect(warnings).toEqual([]);
    const zip = await JSZip.loadAsync((await pptx.write({ outputType: "nodebuffer" })) as Buffer);
    const media = Object.keys(zip.files).filter((f) => f.startsWith("ppt/media/") && !f.endsWith("/"));
    expect(media.length).toBeGreaterThan(0);
    const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(xml).toContain("<p:pic>");
  });
});
