import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";

const pagesUrl = process.env.GITHUB_PAGES_URL;

if (!pagesUrl) {
  throw new Error("GITHUB_PAGES_URL is required, for example https://owner.github.io/repo");
}

const normalizedUrl = pagesUrl.replace(/\/$/, "");
const sourcePath = resolve("public/manifest.xml");
const outputPath = resolve("dist/manifest.xml");
const source = await readFile(sourcePath, "utf8");
const manifest = source.replaceAll("https://localhost:3000", normalizedUrl);

await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, manifest);
