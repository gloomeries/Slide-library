const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(projectRoot, "manifest.xml");
const outputDirectory = path.join(projectRoot, ".tmp");
const outputPath = path.join(outputDirectory, "manifest.xml");
const productionUrl = "https://gloomeries.github.io/Slide-library/";
const developmentUrl = "https://localhost:3000/";

const productionManifest = fs.readFileSync(sourcePath, "utf8");

if (!productionManifest.includes(productionUrl)) {
  throw new Error(`В manifest.xml не найден production-адрес ${productionUrl}`);
}

const developmentManifest = productionManifest.replaceAll(productionUrl, developmentUrl);

fs.mkdirSync(outputDirectory, { recursive: true });
fs.writeFileSync(outputPath, developmentManifest, "utf8");

console.log(`Временный dev-манифест создан: ${path.relative(projectRoot, outputPath)}`);
