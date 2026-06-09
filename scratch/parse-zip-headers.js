const fs = require('fs');
const path = require('path');

function parseZipHeaders(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return;
  }

  const buffer = fs.readFileSync(filePath);
  console.log(`Parsing ZIP headers for: ${filePath} (${buffer.length} bytes)`);

  let offset = 0;
  const entries = [];

  // Scan the buffer for Central Directory Headers: PK\x01\x02 -> 0x02014b50 (little endian)
  while (offset < buffer.length - 46) {
    const sig = buffer.readUInt32LE(offset);
    if (sig === 0x02014b50) {
      const compressionMethod = buffer.readUInt16LE(offset + 10);
      const crc32 = buffer.readUInt32LE(offset + 16);
      const compressedSize = buffer.readUInt32LE(offset + 20);
      const uncompressedSize = buffer.readUInt32LE(offset + 24);
      const fileNameLength = buffer.readUInt16LE(offset + 28);
      const extraFieldLength = buffer.readUInt16LE(offset + 30);
      const fileCommentLength = buffer.readUInt16LE(offset + 32);
      
      const fileName = buffer.toString('utf8', offset + 46, offset + 46 + fileNameLength);

      entries.push({
        name: fileName,
        compressionMethod, // 0 = STORE, 8 = DEFLATE
        crc32: crc32.toString(16).toUpperCase(),
        compressedSize,
        uncompressedSize,
        offset
      });

      offset += 46 + fileNameLength + extraFieldLength + fileCommentLength;
    } else {
      offset++;
    }
  }

  console.log(`Found ${entries.length} central directory entries:`);
  entries.forEach(e => {
    const methodStr = e.compressionMethod === 8 ? 'DEFLATE' : e.compressionMethod === 0 ? 'STORE' : `UNKNOWN(${e.compressionMethod})`;
    console.log(`Entry: ${e.name}`);
    console.log(`  Compression Method: ${methodStr}`);
    console.log(`  Compressed Size:    ${e.compressedSize}`);
    console.log(`  Uncompressed Size:  ${e.uncompressedSize}`);
    console.log(`  CRC32:              ${e.crc32}`);
  });
}

const targetPath = process.argv[2] || path.resolve(__dirname, 'output-debug.pptx');
parseZipHeaders(targetPath);
