// Builds the Next.js static export and copies it into the ZET `app/` folder.
// ZET serves/packs `app/` and expects the entry page to be `widget.html`.
import { spawnSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readdirSync, renameSync, rmSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const webDir = path.join(root, "web");
const outDir = path.join(webDir, "out");
const appDir = path.join(root, "app");

const build = spawnSync("npm run build", { cwd: webDir, stdio: "inherit", shell: true });
if (build.status !== 0) process.exit(build.status ?? 1);
if (!existsSync(path.join(outDir, "index.html"))) {
  console.error("web/out/index.html not found - is output: 'export' set in next.config.ts?");
  process.exit(1);
}

// Clean app/ but keep ZET's translations folder.
mkdirSync(appDir, { recursive: true });
for (const entry of readdirSync(appDir)) {
  if (entry === "translations") continue;
  rmSync(path.join(appDir, entry), { recursive: true, force: true });
}

cpSync(outDir, appDir, { recursive: true });
renameSync(path.join(appDir, "index.html"), path.join(appDir, "widget.html"));

// ZET only allows [a-zA-Z0-9_-$] in folder names, and the widget is a single screen,
// so drop Next's 404 page and the RSC navigation payloads ('__next.*' folders/files).
for (const entry of ["404", "404.html", "_not-found", "index.txt"]) {
  rmSync(path.join(appDir, entry), { recursive: true, force: true });
}
for (const entry of readdirSync(appDir)) {
  if (entry.startsWith("__next")) rmSync(path.join(appDir, entry), { recursive: true, force: true });
}
console.log("Widget built into app/ (entry: app/widget.html)");
