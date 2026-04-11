import fs from "node:fs";
import path from "node:path";

const root = process.cwd();

const requiredFiles = [
  ".clasp.json.example",
  "package.json",
  "README.md",
  ".github/workflows/validate.yml",
  "scripts/validate-repo.mjs",
  "src/Code.gs",
  "src/I18nData.gs",
  "src/Index.html",
  "src/appsscript.json"
];

let failed = false;

for (const relativePath of requiredFiles) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) {
    console.error(`Missing required file: ${relativePath}`);
    failed = true;
  }
}

const manifestPath = path.join(root, "src", "appsscript.json");
if (fs.existsSync(manifestPath)) {
  try {
    const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
    if (!manifest.timeZone || !manifest.runtimeVersion) {
      console.error("src/appsscript.json must define timeZone and runtimeVersion.");
      failed = true;
    }
  } catch (error) {
    console.error(`Invalid JSON in src/appsscript.json: ${error.message}`);
    failed = true;
  }
}

const codePath = path.join(root, "src", "Code.gs");
if (fs.existsSync(codePath)) {
  const code = fs.readFileSync(codePath, "utf8");
  if (!/function\s+doGet\s*\(/.test(code)) {
    console.error("src/Code.gs is missing function doGet().");
    failed = true;
  }
}

const htmlPath = path.join(root, "src", "Index.html");
if (fs.existsSync(htmlPath)) {
  const html = fs.readFileSync(htmlPath, "utf8");
  if (!html.includes("const BOOTSTRAP_DATA = <?!= bootstrapData ?>;")) {
    console.error("src/Index.html is missing the Apps Script bootstrap template expression.");
    failed = true;
  }
}

const claspExamplePath = path.join(root, ".clasp.json.example");
if (fs.existsSync(claspExamplePath)) {
  try {
    const claspExample = JSON.parse(fs.readFileSync(claspExamplePath, "utf8"));
    if (claspExample.rootDir !== "src") {
      console.error(".clasp.json.example must point rootDir to src.");
      failed = true;
    }
  } catch (error) {
    console.error(`Invalid JSON in .clasp.json.example: ${error.message}`);
    failed = true;
  }
}

if (failed) {
  process.exit(1);
}

console.log("Repository validation passed.");
