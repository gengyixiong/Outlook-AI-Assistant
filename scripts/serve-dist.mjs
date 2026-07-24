import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDirectory, "..");
const distDirectory = path.join(projectRoot, "dist");
const certificateDirectory = path.join(
  os.homedir(),
  ".office-addin-dev-certs"
);

const certificateFiles = {
  ca: path.join(certificateDirectory, "ca.crt"),
  cert: path.join(certificateDirectory, "localhost.crt"),
  key: path.join(certificateDirectory, "localhost.key")
};

for (const [name, certificatePath] of Object.entries(certificateFiles)) {
  if (!fs.existsSync(certificatePath)) {
    throw new Error(
      `Missing ${name} certificate at ${certificatePath}. Run the Office add-in development certificate setup first.`
    );
  }
}

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8"
};

function send(response, statusCode, contentType, body) {
  response.writeHead(statusCode, {
    "Content-Type": contentType,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff"
  });
  response.end(body);
}

const server = https.createServer(
  {
    ca: fs.readFileSync(certificateFiles.ca),
    cert: fs.readFileSync(certificateFiles.cert),
    key: fs.readFileSync(certificateFiles.key)
  },
  (request, response) => {
    try {
      const requestUrl = new URL(request.url ?? "/", "https://localhost");
      const requestedPath =
        requestUrl.pathname === "/"
          ? "taskpane.html"
          : decodeURIComponent(requestUrl.pathname.slice(1));
      const filePath = path.resolve(distDirectory, requestedPath);
      const allowedPrefix = `${distDirectory}${path.sep}`;

      if (filePath !== distDirectory && !filePath.startsWith(allowedPrefix)) {
        send(response, 403, "text/plain; charset=utf-8", "Forbidden");
        return;
      }

      if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
        send(response, 404, "text/plain; charset=utf-8", "Not found");
        return;
      }

      const extension = path.extname(filePath).toLowerCase();
      send(
        response,
        200,
        mimeTypes[extension] ?? "application/octet-stream",
        fs.readFileSync(filePath)
      );
    } catch (error) {
      send(
        response,
        500,
        "text/plain; charset=utf-8",
        error instanceof Error ? error.message : "Internal server error"
      );
    }
  }
);

server.listen(3000, "127.0.0.1", () => {
  process.stdout.write(
    [
      "",
      "Outlook AI Assistant compatibility probe is running.",
      "URL: https://localhost:3000/taskpane.html",
      "Keep this window open while testing Outlook.",
      "Press Ctrl+C to stop.",
      ""
    ].join("\n")
  );
});
