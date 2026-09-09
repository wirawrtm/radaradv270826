const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

function createZip(outputFile) {
  const zipEntries = [];
  const excludedDirs = ["node_modules", "dist", ".git", "app"];
  const excludedFiles = [outputFile, "radar-advanta.zip", "project.zip"];

  function walk(dir) {
    const files = fs.readdirSync(dir);
    for (const file of files) {
      const fullPath = path.join(dir, file);
      const stat = fs.statSync(fullPath);
      if (stat.isDirectory()) {
        if (!excludedDirs.includes(file)) {
          walk(fullPath);
        }
      } else {
        if (!excludedFiles.includes(file)) {
          const relPath = path.relative(".", fullPath).replace(/\\/g, "/");
          const fileData = fs.readFileSync(fullPath);
          zipEntries.push({ name: relPath, data: fileData, mtime: stat.mtime });
        }
      }
    }
  }

  walk(".");

  // Build ZIP buffer
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  for (const entry of zipEntries) {
    const nameBuffer = Buffer.from(entry.name, "utf8");
    const compressed = zlib.deflateRawSync(entry.data);
    const useCompressed = compressed.length < entry.data.length;
    const compMethod = useCompressed ? 8 : 0;
    const finalData = useCompressed ? compressed : entry.data;

    // CRC32
    const crc = crc32(entry.data);

    // DOS Time & Date
    const d = entry.mtime;
    const dosTime = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xffff;
    const dosDate = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xffff;

    // Local Header (30 bytes + name + data)
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0); // signature
    localHeader.writeUInt16LE(20, 4); // version needed
    localHeader.writeUInt16LE(0x0800, 6); // UTF-8 flag
    localHeader.writeUInt16LE(compMethod, 8); // compression method
    localHeader.writeUInt16LE(dosTime, 10);
    localHeader.writeUInt16LE(dosDate, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(finalData.length, 18); // compressed size
    localHeader.writeUInt32LE(entry.data.length, 22); // uncompressed size
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28); // extra field len

    localHeaders.push(localHeader, nameBuffer, finalData);

    // Central Directory Header (46 bytes + name)
    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0); // signature
    centralHeader.writeUInt16LE(20, 4); // version made by
    centralHeader.writeUInt16LE(20, 6); // version needed
    centralHeader.writeUInt16LE(0x0800, 8); // UTF-8 flag
    centralHeader.writeUInt16LE(compMethod, 10);
    centralHeader.writeUInt16LE(dosTime, 12);
    centralHeader.writeUInt16LE(dosDate, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(finalData.length, 20);
    centralHeader.writeUInt32LE(entry.data.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30); // extra len
    centralHeader.writeUInt16LE(0, 32); // comment len
    centralHeader.writeUInt16LE(0, 34); // disk start
    centralHeader.writeUInt16LE(0, 36); // internal attr
    centralHeader.writeUInt32LE(0x81a40000, 38); // external attr (regular file)
    centralHeader.writeUInt32LE(offset, 42); // relative offset

    centralHeaders.push(centralHeader, nameBuffer);

    offset += localHeader.length + nameBuffer.length + finalData.length;
  }

  const centralDirOffset = offset;
  let centralDirSize = 0;
  for (const b of centralHeaders) {
    centralDirSize += b.length;
  }

  // End of central directory record (22 bytes)
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); // signature
  eocd.writeUInt16LE(0, 4); // disk number
  eocd.writeUInt16LE(0, 6); // start disk
  eocd.writeUInt16LE(zipEntries.length, 8); // entries this disk
  eocd.writeUInt16LE(zipEntries.length, 10); // total entries
  eocd.writeUInt32LE(centralDirSize, 12); // size of central dir
  eocd.writeUInt32LE(centralDirOffset, 16); // offset of central dir
  eocd.writeUInt16LE(0, 20); // comment length

  const allBuffers = [...localHeaders, ...centralHeaders, eocd];
  const fullZip = Buffer.concat(allBuffers);

  fs.mkdirSync(path.dirname(outputFile), { recursive: true });
  fs.writeFileSync(outputFile, fullZip);
  console.log(`ZIP created successfully at ${outputFile} (${(fullZip.length / 1024).toFixed(1)} KB) with ${zipEntries.length} files.`);
}

function crc32(buf) {
  let crc = ~0;
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i]) & 0xff];
  }
  return (~crc) >>> 0;
}

const crcTable = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) {
    c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
  }
  crcTable[i] = c >>> 0;
}

if (require.main === module) {
  const target = process.argv[2] || "public/project.zip";
  createZip(target);
}

module.exports = { createZip };
