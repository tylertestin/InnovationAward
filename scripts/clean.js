const fs = require("fs");
const path = require("path");

const dist = path.join(__dirname, "..", "dist");
if (fs.existsSync(dist)) {
  fs.rmSync(dist, { recursive: true, force: true });
  console.log("Removed dist/");
} else {
  console.log("dist/ does not exist");
}
