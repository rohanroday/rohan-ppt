// Turns the file the Zoho SDK returns into a picture data URL (FR-25, FR-27). Needs the DOM.

const MAX_SIDE = 1280; // FR-27

function sniffImageMime(b: Uint8Array): string | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "image/png";
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "image/jpeg";
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "image/gif";
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45) return "image/webp";
  return null;
}

async function toBlob(file: unknown): Promise<Blob> {
  if (file instanceof Blob) return file;
  if (file instanceof ArrayBuffer) return new Blob([file]);
  if (ArrayBuffer.isView(file)) {
    const copy = new Uint8Array(file.byteLength);
    copy.set(new Uint8Array(file.buffer, file.byteOffset, file.byteLength));
    return new Blob([copy]);
  }
  if (file && typeof file === "object" && typeof (file as Response).blob === "function") return (file as Response).blob();
  throw new Error("CRM returned a file in an unexpected format");
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error ?? new Error("could not read the file"));
    r.readAsDataURL(blob);
  });
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("the file is not a picture the browser can read"));
    img.src = src;
  });
}

/** Re-encodes `dataUrl` when it is over MAX_SIDE or in a format PowerPoint export can't take (WebP). */
async function normalise(dataUrl: string, mime: string): Promise<string> {
  const img = await loadImage(dataUrl);
  const longest = Math.max(img.naturalWidth, img.naturalHeight);
  if (longest <= MAX_SIDE && mime !== "image/webp") return dataUrl;
  const scale = Math.min(1, MAX_SIDE / longest);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
  const g = canvas.getContext("2d");
  if (!g) return dataUrl;
  g.drawImage(img, 0, 0, canvas.width, canvas.height);
  return mime === "image/jpeg" ? canvas.toDataURL("image/jpeg", 0.85) : canvas.toDataURL("image/png");
}

/** `file` is whatever `ZOHO.CRM.API.getFile` returns (Blob, bytes or a Response). */
export async function fileToImageDataUrl(file: unknown): Promise<string> {
  const blob = await toBlob(file);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const mime = sniffImageMime(bytes) ?? (blob.type.startsWith("image/") ? blob.type : null);
  if (!mime) throw new Error("the file is not a picture");
  return normalise(await blobToDataUrl(blob.type === mime ? blob : new Blob([bytes], { type: mime })), mime);
}
