import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";

const certificatePath = path.join(
  os.homedir(),
  ".office-addin-dev-certs",
  "ca.crt"
);
const ca = fs.readFileSync(certificatePath);
const targets = [
  "https://localhost:3000/taskpane.html",
  "https://localhost:3000/commands.html",
  "https://localhost:3000/support.html",
  "https://localhost:3000/assets/icon-80.png"
];

function request(target) {
  return new Promise((resolve, reject) => {
    const requestHandle = https.get(target, { ca }, (response) => {
      let bytes = 0;
      response.on("data", (chunk) => {
        bytes += chunk.length;
      });
      response.on("end", () => {
        if (response.statusCode !== 200) {
          reject(new Error(`${target} returned HTTP ${response.statusCode}`));
          return;
        }

        resolve({
          target,
          status: response.statusCode,
          contentType: response.headers["content-type"] ?? "unknown",
          bytes
        });
      });
    });

    requestHandle.on("error", reject);
  });
}

const results = await Promise.all(targets.map(request));

for (const result of results) {
  process.stdout.write(
    `${result.status} ${result.contentType} ${result.bytes}B ${result.target}\n`
  );
}
