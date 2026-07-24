import { fileURLToPath } from "node:url";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { defineConfig } from "vite";

const projectRoot = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig(({ command }) => {
  const certificateDirectory = path.join(
    os.homedir(),
    ".office-addin-dev-certs"
  );
  const httpsOptions =
    command === "serve"
      ? {
          ca: fs.readFileSync(path.join(certificateDirectory, "ca.crt")),
          cert: fs.readFileSync(
            path.join(certificateDirectory, "localhost.crt")
          ),
          key: fs.readFileSync(path.join(certificateDirectory, "localhost.key"))
        }
      : undefined;

  return {
    root: path.resolve(projectRoot, "src"),
    publicDir: path.resolve(projectRoot, "public"),
    server: {
      host: "127.0.0.1",
      port: 3000,
      strictPort: true,
      https: httpsOptions
    },
    build: {
      outDir: path.resolve(projectRoot, "dist"),
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: path.resolve(projectRoot, "src/taskpane.html"),
          commands: path.resolve(projectRoot, "src/commands.html"),
          support: path.resolve(projectRoot, "src/support.html")
        }
      }
    }
  };
});
