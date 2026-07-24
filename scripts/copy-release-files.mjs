import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDirectory, "..");
const distDirectory = path.join(projectRoot, "dist");

fs.copyFileSync(
  path.join(projectRoot, "manifest", "manifest.xml"),
  path.join(distDirectory, "manifest.xml")
);

fs.copyFileSync(
  path.join(projectRoot, "docs", "TEST-CHECKLIST.md"),
  path.join(distDirectory, "TEST-CHECKLIST.md")
);

fs.copyFileSync(
  path.join(projectRoot, "docs", "PROJECT-STATUS.md"),
  path.join(distDirectory, "PROJECT-STATUS.md")
);

process.stdout.write("Copied manifest and release documents to dist.\n");
