/**
 * Build script: creates the XPI file using Node.js.
 * Usage: node build-xpi.js
 */

const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const OUTPUT = "exchange-calendar-sync.xpi";
const ROOT = __dirname;

// Files and directories to include
const INCLUDES = [
  "manifest.json",
  "background.html",
  "background.js",
  "ews",
  "sync",
  "api",
  "options",
  "meetingbar",
  "icons",
];

// Collect all files recursively
function collectFiles(basePath, dir) {
  const results = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...collectFiles(basePath, fullPath));
    } else {
      const relPath = path.relative(basePath, fullPath).replace(/\\/g, "/");
      results.push({ fullPath, relPath });
    }
  }
  return results;
}

// Build file list
const files = [];
for (const inc of INCLUDES) {
  const fullPath = path.join(ROOT, inc);
  const stat = fs.statSync(fullPath);
  if (stat.isDirectory()) {
    files.push(...collectFiles(ROOT, fullPath));
  } else {
    files.push({ fullPath, relPath: inc });
  }
}

// --- Minimal ZIP writer (Store method, no compression needed for small files) ---
// ZIP format: https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT

const zlib = require("zlib");

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const localHeaders = [];
const centralHeaders = [];
let offset = 0;

const outputBuf = [];

for (const file of files) {
  const data = fs.readFileSync(file.fullPath);
  const compressed = zlib.deflateRawSync(data);
  const nameBytes = Buffer.from(file.relPath, "utf-8");
  const crc = crc32(data);

  // Local file header
  const local = Buffer.alloc(30 + nameBytes.length);
  local.writeUInt32LE(0x04034b50, 0); // signature
  local.writeUInt16LE(20, 4);         // version needed
  local.writeUInt16LE(0, 6);          // flags
  local.writeUInt16LE(8, 8);          // compression: deflate
  local.writeUInt16LE(0, 10);         // mod time
  local.writeUInt16LE(0, 12);         // mod date
  local.writeUInt32LE(crc, 14);       // crc32
  local.writeUInt32LE(compressed.length, 18); // compressed size
  local.writeUInt32LE(data.length, 22);       // uncompressed size
  local.writeUInt16LE(nameBytes.length, 26);  // filename length
  local.writeUInt16LE(0, 28);                 // extra field length
  nameBytes.copy(local, 30);

  outputBuf.push(local, compressed);

  // Central directory header
  const central = Buffer.alloc(46 + nameBytes.length);
  central.writeUInt32LE(0x02014b50, 0);  // signature
  central.writeUInt16LE(20, 4);          // version made by
  central.writeUInt16LE(20, 6);          // version needed
  central.writeUInt16LE(0, 8);           // flags
  central.writeUInt16LE(8, 10);          // compression: deflate
  central.writeUInt16LE(0, 12);          // mod time
  central.writeUInt16LE(0, 14);          // mod date
  central.writeUInt32LE(crc, 16);        // crc32
  central.writeUInt32LE(compressed.length, 20); // compressed size
  central.writeUInt32LE(data.length, 24);       // uncompressed size
  central.writeUInt16LE(nameBytes.length, 28);  // filename length
  central.writeUInt16LE(0, 30);                 // extra field length
  central.writeUInt16LE(0, 32);                 // file comment length
  central.writeUInt16LE(0, 34);                 // disk number start
  central.writeUInt16LE(0, 36);                 // internal attrs
  central.writeUInt32LE(0, 38);                 // external attrs
  central.writeUInt32LE(offset, 42);             // local header offset
  nameBytes.copy(central, 46);

  centralHeaders.push(central);
  offset += local.length + compressed.length;
}

// Write central directory
const centralStart = offset;
for (const ch of centralHeaders) {
  outputBuf.push(ch);
  offset += ch.length;
}
const centralSize = offset - centralStart;

// End of central directory
const eocd = Buffer.alloc(22);
eocd.writeUInt32LE(0x06054b50, 0);                // signature
eocd.writeUInt16LE(0, 4);                          // disk number
eocd.writeUInt16LE(0, 6);                          // central dir disk
eocd.writeUInt16LE(centralHeaders.length, 8);      // entries on this disk
eocd.writeUInt16LE(centralHeaders.length, 10);     // total entries
eocd.writeUInt32LE(centralSize, 12);               // central dir size
eocd.writeUInt32LE(centralStart, 16);              // central dir offset
eocd.writeUInt16LE(0, 20);                          // comment length
outputBuf.push(eocd);

const result = Buffer.concat(outputBuf);
const outputPath = path.join(ROOT, OUTPUT);
fs.writeFileSync(outputPath, result);

console.log(`Built ${OUTPUT} (${result.length} bytes, ${files.length} files)`);
for (const f of files) {
  console.log(`  ${f.relPath}`);
}
