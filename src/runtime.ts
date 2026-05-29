function downloadFile(file: ExtractedFile) {
  downloadBytes(file.name, file.mimeType, file.bytes);
}

function downloadBytes(fileName: string, mimeType: string, bytes: Uint8Array) {
  downloadBlob(fileName, new Blob([bytes], { type: mimeType }));
}

function downloadBlob(fileName: string, blob: Blob) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function downloadAll(files: ExtractedFile[]) {
  if (files.length === 0) return;
  const zipBytes = createZip(files);
  const blob = new Blob([zipBytes], { type: "application/zip" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "vba-extractor-output.zip";
  anchor.click();
  URL.revokeObjectURL(url);
}

function resetUi(clearInput = true) {
  currentFiles = [];
  results.innerHTML = "";
  summaryPanel.classList.add("hidden");
  setStatus("Waiting for a document.", "idle");
  if (clearInput) fileInput.value = "";
}

function setStatus(message: string, state: "idle" | "busy" | "ready" | "warn" | "error") {
  statusText.textContent = message;
  statusPanel.dataset.state = state;
}

function findEndOfCentralDirectory(bytes: Uint8Array) {
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset -= 1) {
    if (readU32(bytes, offset) === 0x06054b50) return offset;
  }
  return -1;
}

function isZip(bytes: Uint8Array) {
  return readU32(bytes, 0) === 0x04034b50;
}

function isOle(bytes: Uint8Array) {
  return matches(bytes, 0, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
}

function getSector(bytes: Uint8Array, sectorSize: number, sectorIndex: number) {
  const offset = (sectorIndex + 1) * sectorSize;
  return bytes.subarray(offset, offset + sectorSize);
}

function concatBytes(chunks: Uint8Array[]) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const combined = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    combined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return combined;
}

function readU16(bytes: Uint8Array, offset: number) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes: Uint8Array, offset: number) {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

function readI32(bytes: Uint8Array, offset: number) {
  return readU32(bytes, offset) | 0;
}

function writeU16(bytes: Uint8Array, offset: number, value: number) {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
}

function writeU32(bytes: Uint8Array, offset: number, value: number) {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
  bytes[offset + 2] = (value >>> 16) & 0xff;
  bytes[offset + 3] = (value >>> 24) & 0xff;
}

function decodeText(bytes: Uint8Array) {
  try {
    return new TextDecoder("windows-1252").decode(bytes);
  } catch {
    return new TextDecoder().decode(bytes);
  }
}

function encodeText(text: string) {
  return new TextEncoder().encode(text);
}

function decodeUtf16Le(bytes: Uint8Array) {
  let text = "";
  for (let offset = 0; offset + 1 < bytes.length; offset += 2) {
    const code = bytes[offset] | (bytes[offset + 1] << 8);
    if (code) text += String.fromCharCode(code);
  }
  return text;
}

function matches(bytes: Uint8Array, offset: number, signature: number[]) {
  return signature.every((byte, index) => bytes[offset + index] === byte);
}

function matchesAscii(bytes: Uint8Array, offset: number, value: string) {
  return [...value].every((char, index) => bytes[offset + index] === char.charCodeAt(0));
}

function findBytes(bytes: Uint8Array, pattern: number[], start: number) {
  for (let offset = start; offset <= bytes.length - pattern.length; offset += 1) {
    if (matches(bytes, offset, pattern)) return offset;
  }
  return -1;
}

function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function toHex(value: number, width = 4) {
  return `0x${value.toString(16).toUpperCase().padStart(width, "0")}`;
}

function formatGuid(bytes: Uint8Array) {
  const hex = [...bytes].map((byte) => byte.toString(16).padStart(2, "0"));
  return `{${[
    hex.slice(0, 4).reverse().join(""),
    hex.slice(4, 6).reverse().join(""),
    hex.slice(6, 8).reverse().join(""),
    hex.slice(8, 10).join(""),
    hex.slice(10, 16).join(""),
  ].join("-").toUpperCase()}}`;
}

function safeFileName(value: string) {
  return value.replace(/[^a-z0-9._-]+/gi, "_").replace(/^_+|_+$/g, "").slice(0, 180) || "stream";
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"']/g, (char) => {
    const map: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
    return map[char];
  });
}

function kindRank(kind: ExtractedFile["kind"]) {
  return { frm: 0, vba: 1, media: 2, frx: 3, binary: 4 }[kind];
}

