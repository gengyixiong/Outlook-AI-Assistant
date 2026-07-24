import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const outputDirectory = path.resolve(scriptDirectory, "../public/assets");
const sizes = [16, 32, 64, 80, 128];

function crc32(buffer) {
  let crc = 0xffffffff;

  for (const value of buffer) {
    crc ^= value;
    for (let bit = 0; bit < 8; bit += 1) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  }

  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuffer = Buffer.from(type, "ascii");
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])));
  return Buffer.concat([length, typeBuffer, data, checksum]);
}

function distanceToSegment(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const lengthSquared = dx * dx + dy * dy;
  const projection =
    lengthSquared === 0
      ? 0
      : Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / lengthSquared));
  const closestX = x1 + projection * dx;
  const closestY = y1 + projection * dy;
  return Math.hypot(px - closestX, py - closestY);
}

function createIcon(size) {
  const rowLength = size * 4 + 1;
  const raw = Buffer.alloc(rowLength * size);
  const lineWidth = Math.max(1, size * 0.075);

  for (let y = 0; y < size; y += 1) {
    raw[y * rowLength] = 0;

    for (let x = 0; x < size; x += 1) {
      const offset = y * rowLength + 1 + x * 4;
      const nx = x / Math.max(1, size - 1);
      const ny = y / Math.max(1, size - 1);
      const edge = Math.min(x, y, size - 1 - x, size - 1 - y);
      const rounded = edge < size * 0.08 && Math.hypot(
        Math.max(0, size * 0.08 - Math.min(x, size - 1 - x)),
        Math.max(0, size * 0.08 - Math.min(y, size - 1 - y))
      ) > size * 0.08;

      const red = Math.round(10 + 20 * ny);
      const green = Math.round(92 + 45 * nx);
      const blue = Math.round(170 + 38 * (1 - ny));
      const check =
        distanceToSegment(x, y, size * 0.25, size * 0.53, size * 0.43, size * 0.7) < lineWidth ||
        distanceToSegment(x, y, size * 0.43, size * 0.7, size * 0.76, size * 0.31) < lineWidth;

      raw[offset] = check ? 255 : red;
      raw[offset + 1] = check ? 255 : green;
      raw[offset + 2] = check ? 255 : blue;
      raw[offset + 3] = rounded ? 0 : 255;
    }
  }

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header[8] = 8;
  header[9] = 6;
  header[10] = 0;
  header[11] = 0;
  header[12] = 0;

  return Buffer.concat([
    signature,
    chunk("IHDR", header),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

fs.mkdirSync(outputDirectory, { recursive: true });

for (const size of sizes) {
  const outputPath = path.join(outputDirectory, `icon-${size}.png`);
  fs.writeFileSync(outputPath, createIcon(size));
}

process.stdout.write(`Generated ${sizes.length} icons in ${outputDirectory}\n`);
