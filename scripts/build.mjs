import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const builds = path.join(root, "builds");
const extensionFiles = ["accordion.css", "content.js", "options.html", "options.js"];\nconst extensionDirectories = ["icons"];
const requestedTarget = process.argv[2];
const targets = requestedTarget ? [requestedTarget] : ["firefox", "chrome"];

if (targets.some((target) => !["firefox", "chrome"].includes(target))) {
  throw new Error("Usage: node scripts/build.mjs [firefox|chrome]");
}

const baseManifest = JSON.parse(await readFile(path.join(root, "manifest.json"), "utf8"));

for (const target of targets) {
  const output = path.join(builds, target);
  await rm(output, { recursive: true, force: true });
  await mkdir(output, { recursive: true });
  await Promise.all([\n    ...extensionFiles.map((file) => cp(path.join(root, file), path.join(output, file))),\n    ...extensionDirectories.map((directory) => cp(path.join(root, directory), path.join(output, directory), { recursive: true }))\n  ]);

  const manifest = structuredClone(baseManifest);
  if (target === "firefox") {
    Object.assign(manifest, JSON.parse(await readFile(path.join(root, "manifests/firefox.json"), "utf8")));
  }

  await writeFile(path.join(output, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  const archiveName = target === "firefox" ? "InboxRedux-firefox.xpi" : "InboxRedux-chrome.zip";
  const archivePath = path.join(builds, archiveName);
  await rm(archivePath, { force: true });
  await execFileAsync("zip", ["-qr", archivePath, "."], { cwd: output });
  console.log(`Built ${path.relative(root, archivePath)}`);
}
