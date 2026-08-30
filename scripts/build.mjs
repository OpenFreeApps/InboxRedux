import { cp, mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);

async function createArchive(archivePath, sourceDirectory) {
  if (process.platform === "win32") {
    // Windows does not ship the Unix zip command. PowerShell requires a .zip
    // destination, so create one and rename it to .xpi for Firefox afterward.
    const temporaryArchivePath = `${archivePath}.zip`;
    await rm(temporaryArchivePath, { force: true });
    await execFileAsync(
      "powershell.exe",
      [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        [
          "Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem",
          "$archive = [System.IO.Compression.ZipFile]::Open($env:INBOXREDUX_ARCHIVE_PATH, [System.IO.Compression.ZipArchiveMode]::Create)",
          "Get-ChildItem -File -Recurse | ForEach-Object {",
          "  $relativePath = $_.FullName.Substring($PWD.Path.Length + 1).Replace('\\', '/')",
          "  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $relativePath, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null",
          "}",
          "$archive.Dispose()"
        ].join("; ")
      ],
      {
        cwd: sourceDirectory,
        env: { ...process.env, INBOXREDUX_ARCHIVE_PATH: temporaryArchivePath }
      }
    );
    await rename(temporaryArchivePath, archivePath);
    return;
  }

  await execFileAsync("zip", ["-qr", archivePath, "."], { cwd: sourceDirectory });
}
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const builds = path.join(root, "builds");
const extensionFiles = ["accordion.css", "content.js", "options.html", "options.js"];
const extensionDirectories = ["icons"];
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
  await Promise.all([
    ...extensionFiles.map((file) => cp(path.join(root, file), path.join(output, file))),
    ...extensionDirectories.map((directory) => cp(path.join(root, directory), path.join(output, directory), { recursive: true }))
  ]);

  const manifest = structuredClone(baseManifest);
  if (target === "firefox") {
    Object.assign(manifest, JSON.parse(await readFile(path.join(root, "manifests/firefox.json"), "utf8")));
  }

  await writeFile(path.join(output, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  const archiveName = target === "firefox" ? "InboxRedux-firefox.xpi" : "InboxRedux-chrome.zip";
  const archivePath = path.join(builds, archiveName);
  await rm(archivePath, { force: true });
  await createArchive(archivePath, output);
  console.log(`Built ${path.relative(root, archivePath)}`);
}
