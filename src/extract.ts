async function extractOffice(bytes: Uint8Array, fileName: string): Promise<ExtractedFile[]> {
  const raw = extractRawVbaOrFormFile(bytes, fileName);
  if (raw) return raw;

  if (isZip(bytes)) {
    const zipEntries = parseZip(bytes);
    const projectBins = zipEntries.filter((entry) => /vbaProject\.bin$/i.test(entry.name));
    const extracted: ExtractedFile[] = [];

    for (const entry of projectBins) {
      const projectBytes = await inflateZipEntry(bytes, entry);
      extracted.push(...extractVbaProject(projectBytes, `${fileName}/${entry.name}`));
    }

    extracted.push(...await extractOfficePackageParts(bytes, zipEntries, fileName));
    return extracted;
  }

  if (isOle(bytes)) {
    return extractVbaProject(bytes, fileName);
  }

  throw new Error("This does not look like a supported Office zip or legacy OLE document.");
}

function extractRawVbaOrFormFile(bytes: Uint8Array, fileName: string): ExtractedFile[] | undefined {
  const extension = fileName.split(".").at(-1)?.toLowerCase();
  if (!extension || !["bas", "cls", "frm", "frx"].includes(extension)) return undefined;

  if (extension === "frx") {
    const baseName = safeFileName(fileName.replace(/\.frx$/i, ""));
    const resource: ExtractedFile = {
      name: fileName,
      kind: "frx",
      bytes,
      analysis: analyzeBinaryStream(bytes, fileName),
      mimeType: "application/octet-stream",
      sourcePath: fileName,
    };
    return [resource, ...extractMediaFromBinary(bytes, baseName).map((file) => ({ ...file, sourcePath: `${fileName} / ${file.sourcePath}` }))];
  }

  const text = decodeText(bytes).replace(/\0+$/g, "");
  return [{
    name: fileName,
    kind: extension === "frm" ? "frm" : "vba",
    bytes,
    text,
    mimeType: "text/plain;charset=windows-1252",
    sourcePath: fileName,
  }];
}

function extractVbaProject(bytes: Uint8Array, sourcePath: string): ExtractedFile[] {
  const cfb = parseCfb(bytes);
  const projectStream = cfb.streams.get("PROJECT") ?? cfb.streams.get("VBA/PROJECT");
  const dirStream = cfb.streams.get("VBA/dir") ?? cfb.streams.get("dir");
  const modules = parseProjectModules(projectStream);
  const dirModules = parseDirModules(cfb.streams.get("VBA/dir") ?? cfb.streams.get("dir"));
  const files: ExtractedFile[] = [];
  const processedSourceStreams = new Set<string>();
  const usedMediaKeys = new Set<string>();

  for (const dirModule of dirModules) {
    const streamEntry = findVbaModuleStream(cfb.streams, dirModule.streamName);
    if (!streamEntry) continue;

    const [path, stream] = streamEntry;
    const projectModule = modules.get(dirModule.name.toLowerCase()) ?? modules.get(dirModule.streamName.toLowerCase());
    const source = extractVbaTextAtOffset(stream, dirModule.textOffset);
    if (!source) continue;

    processedSourceStreams.add(path.toLowerCase());
    files.push(createSourceFile({
      moduleName: projectModule?.name ?? dirModule.name,
      moduleType: projectModule?.type ?? (dirModule.moduleTypeId === 0x0021 ? "module" : "class"),
      source,
      sourcePath: `${sourcePath}/${path} @ ${toHex(dirModule.textOffset)}`,
    }));
  }

  for (const [path, stream] of cfb.streams) {
    const lowerPath = path.toLowerCase();
    if (processedSourceStreams.has(lowerPath)) continue;

    const baseName = path.split("/").at(-1) ?? path;
    const module = modules.get(baseName.toLowerCase());

    if (module && lowerPath.startsWith("vba/")) {
      const source = extractBestVbaText(stream);
      if (!source) continue;

      processedSourceStreams.add(lowerPath);
      files.push(createSourceFile({
        moduleName: module.name,
        moduleType: module.type,
        source,
        sourcePath: `${sourcePath}/${path} (recovered by scan)`,
      }));
    }
  }

  for (const [path, stream] of cfb.streams) {
    const lower = path.toLowerCase();
    if (lower === "project" || lower.startsWith("vba/")) continue;

    const formName = path.split("/")[0] || "form";
    files.push({
      name: `${safeFileName(formName)}-${safeFileName(path.replaceAll("/", "-"))}.frx`,
      kind: "frx",
      bytes: stream,
      analysis: analyzeBinaryStream(stream, path),
      mimeType: "application/octet-stream",
      sourcePath: `${sourcePath}/${path}`,
    });

    const mediaFiles = extractMediaFromBinary(stream, `${safeFileName(formName)}-${safeFileName(path.replaceAll("/", "-"))}`);
    for (const media of mediaFiles) {
      const key = `${path}:${media.name}:${media.bytes.byteLength}`;
      if (usedMediaKeys.has(key)) continue;
      usedMediaKeys.add(key);
      files.push({
        ...media,
        sourcePath: `${sourcePath}/${path}`,
      });
    }
  }

  if (files.length === 0) {
    for (const [path, stream] of cfb.streams) {
      const source = extractBestVbaText(stream);
      if (!source) continue;
      const name = safeFileName(path.split("/").at(-1) ?? "module");
      files.push({
        name: `${name}.bas`,
        kind: "vba",
        bytes: encodeText(source),
        text: source,
        mimeType: "text/plain;charset=windows-1252",
        sourcePath: `${sourcePath}/${path}`,
      });
    }
  }

  const references = parseProjectReferences(projectStream, dirStream);
  if (references.length) {
    files.push({
      name: "vba-project-references.json",
      kind: "binary",
      bytes: encodeText(JSON.stringify(references, null, 2)),
      text: JSON.stringify(references, null, 2),
      mimeType: "application/json",
      sourcePath: `${sourcePath}/[vba project references]`,
    });
  }

  return files.sort((a, b) => kindRank(a.kind) - kindRank(b.kind) || a.name.localeCompare(b.name));
}

async function extractOfficePackageParts(bytes: Uint8Array, zipEntries: ZipEntry[], fileName: string): Promise<ExtractedFile[]> {
  const interesting = zipEntries.filter((entry) =>
    /^word\/media\//i.test(entry.name) ||
    /^ppt\/media\//i.test(entry.name) ||
    /^xl\/media\//i.test(entry.name) ||
    /\/activeX\/.+\.(xml|bin)$/i.test(entry.name) ||
    /\/embeddings\/.+\.(bin|ole|xls|xlsx|doc|docx|ppt|pptx)$/i.test(entry.name),
  );
  const extracted: ExtractedFile[] = [];
  const manifest: Array<{ path: string; kind: string; size: number; note: string }> = [];

  for (const entry of interesting) {
    const entryBytes = await inflateZipEntry(bytes, entry);
    const kind = classifyOfficePart(entry.name);
    manifest.push({ path: entry.name, kind, size: entryBytes.byteLength, note: describeOfficePart(entry.name, entryBytes) });

    extracted.push({
      name: `${safeFileName(entry.name.replaceAll("/", "-"))}`,
      kind: kind === "media" ? "media" : "binary",
      bytes: entryBytes,
      text: kind === "xml" ? decodeText(entryBytes) : undefined,
      analysis: kind === "binary" || kind === "activex" ? analyzeBinaryStream(entryBytes, entry.name) : undefined,
      mimeType: mimeTypeForOfficePart(entry.name),
      sourcePath: `${fileName}/${entry.name}`,
    });

    if (kind === "binary" || kind === "activex") {
      const nestedMedia = extractMediaFromBinary(entryBytes, safeFileName(entry.name.replaceAll("/", "-")));
      for (const media of nestedMedia) {
        extracted.push({
          ...media,
          sourcePath: `${fileName}/${entry.name} / ${media.sourcePath}`,
        });
      }
    }
  }

  if (manifest.length) {
    extracted.push({
      name: "office-package-manifest.json",
      kind: "binary",
      bytes: encodeText(JSON.stringify({ source: fileName, parts: manifest }, null, 2)),
      text: JSON.stringify({ source: fileName, parts: manifest }, null, 2),
      mimeType: "application/json",
      sourcePath: `${fileName}/[office package manifest]`,
    });
  }

  return extracted;
}

function classifyOfficePart(path: string) {
  if (/\/media\//i.test(path)) return "media";
  if (/\.xml$/i.test(path)) return "xml";
  if (/\/activeX\//i.test(path)) return "activex";
  return "binary";
}

function describeOfficePart(path: string, bytes: Uint8Array) {
  if (/\/activeX\//i.test(path)) return "ActiveX control package part";
  if (/\/media\//i.test(path)) return "Office document media asset";
  if (isOle(bytes)) return "Embedded OLE compound file";
  return "Office embedded package part";
}

function mimeTypeForOfficePart(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith(".xml")) return "application/xml";
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".bmp")) return "image/bmp";
  if (lower.endsWith(".wmf")) return "image/wmf";
  if (lower.endsWith(".emf")) return "image/emf";
  return "application/octet-stream";
}

function parseProjectModules(projectStream?: Uint8Array) {
  const modules = new Map<string, ProjectModule>();
  if (!projectStream) return modules;

  const text = decodeText(projectStream);
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    const [key, value] = line.split("=", 2);
    if (!key || !value) continue;

    if (key === "Module") modules.set(value.toLowerCase(), { name: value, type: "module" });
    if (key === "Class") modules.set(value.toLowerCase(), { name: value, type: "class" });
    if (key === "BaseClass") modules.set(value.toLowerCase(), { name: value, type: "form" });
    if (key === "Document") {
      const documentName = value.split("/")[0];
      modules.set(documentName.toLowerCase(), { name: documentName, type: "document" });
    }
  }

  return modules;
}

function parseProjectReferences(projectStream?: Uint8Array, dirStream?: Uint8Array): ProjectReferenceModel[] {
  const references = [...parseProjectStreamReferences(projectStream), ...parseDirProjectReferences(dirStream)];
  const seen = new Set<string>();
  return references.filter((reference) => {
    const key = `${reference.source}:${reference.kind}:${reference.name}:${reference.libId}:${reference.raw}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function parseProjectStreamReferences(projectStream?: Uint8Array): ProjectReferenceModel[] {
  if (!projectStream) return [];
  const references: ProjectReferenceModel[] = [];
  const text = decodeText(projectStream);

  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    const [key, value = ""] = line.split("=", 2);
    if (!["Reference", "Object", "Package"].includes(key)) continue;

    const guid = normalizeGuidString(value);
    const libParts = value.split("#").map((part) => part.trim()).filter(Boolean);
    const libId = value.match(/\*?\\G\{[^}]+\}#[^#]+#[^#]+#?[^#]*/i)?.[0] ?? guid;
    references.push({
      source: "PROJECT",
      kind: key === "Reference" ? "registered" : key.toLowerCase() as ProjectReferenceModel["kind"],
      name: inferReferenceName(value),
      libId,
      guid,
      version: libParts.find((part) => /^\d+\.\d+$/.test(part)),
      path: libParts.find((part) => /[\\/]|\.dll|\.ocx|\.tlb|\.olb/i.test(part)),
      raw: line,
    });
  }

  return references;
}

function parseDirProjectReferences(dirStream?: Uint8Array): ProjectReferenceModel[] {
  if (!dirStream) return [];
  let dir: Uint8Array;
  try {
    dir = decompressVba(dirStream);
  } catch {
    return [];
  }

  const modulesStart = findProjectModulesRecord(dir);
  const referenceBytes = modulesStart > 0 ? dir.subarray(0, modulesStart) : dir;
  const references: ProjectReferenceModel[] = [];
  let cursor = 0;
  let pendingName: string | undefined;
  let pendingOriginalLibId: string | undefined;

  while (cursor + 6 <= referenceBytes.length) {
    if (readU16(referenceBytes, cursor) === 0x0016) {
      const parsedName = readReferenceNameRecord(referenceBytes, cursor);
      if (parsedName) {
        pendingName = parsedName.name;
        cursor = parsedName.next;
        continue;
      }
    }

    let record;
    try {
      record = readRecord(referenceBytes, cursor);
    } catch {
      cursor += 1;
      continue;
    }

    if (record.id === 0x0016) {
      pendingName = parseReferenceNameRecord(record.payload);
    } else if (record.id === 0x0033) {
      pendingOriginalLibId = parseReferenceLibIdRecord(record.payload, "direct");
    } else if ([0x002f, 0x0030, 0x000d, 0x000e].includes(record.id)) {
      const parsed = parseDirReferenceRecord(record.id, record.payload);
      const raw = parsed.libId ?? pendingOriginalLibId;
      if (raw || pendingName) {
        const guid = normalizeGuidString(raw ?? "");
        const lib = parseLibId(raw);
        references.push({
          source: "dir",
          kind: classifyDirReferenceRecord(record.id),
          name: pendingName,
          libId: raw,
          guid,
          version: parsed.version ?? lib.version,
          path: parsed.path ?? lib.path,
          raw: `record ${toHex(record.id, 4)} ${raw ?? ""}`.trim(),
        });
        pendingName = undefined;
        pendingOriginalLibId = undefined;
      }
    } else if (pendingOriginalLibId && record.id !== 0x0016) {
      const raw = pendingOriginalLibId;
      const guid = normalizeGuidString(raw);
      const lib = parseLibId(raw);
      references.push({
        source: "dir",
        kind: "control",
        name: pendingName,
        libId: raw,
        guid,
        version: lib.version,
        path: lib.path,
        raw: `record 0x0033 ${raw}`,
      });
      pendingName = undefined;
      pendingOriginalLibId = undefined;
    }
    cursor = record.next;
  }

  return references;
}

function parseReferenceNameRecord(payload: Uint8Array) {
  if (payload.length < 6) return cleanReferenceText(decodeText(payload));
  const size = readU32(payload, 0);
  if (size > 0 && size <= payload.length - 4) {
    const unicodeSizeOffset = 4 + size + 2;
    if (unicodeSizeOffset + 4 <= payload.length) {
      const unicodeSize = readU32(payload, unicodeSizeOffset);
      const unicodeStart = unicodeSizeOffset + 4;
      if (unicodeSize > 0 && unicodeStart + unicodeSize <= payload.length) {
        const unicode = decodeUtf16Le(payload.subarray(unicodeStart, unicodeStart + unicodeSize)).trim();
        if (unicode) return unicode;
      }
    }
    return cleanReferenceText(decodeText(payload.subarray(4, 4 + size)));
  }
  return cleanReferenceText(decodeText(payload));
}

function readReferenceNameRecord(bytes: Uint8Array, cursor: number) {
  if (cursor + 6 > bytes.length || readU16(bytes, cursor) !== 0x0016) return undefined;
  const sizeOfName = readU32(bytes, cursor + 2);
  const nameStart = cursor + 6;
  const reservedOffset = nameStart + sizeOfName;
  if (sizeOfName > bytes.length - nameStart || reservedOffset + 6 > bytes.length) return undefined;
  if (readU16(bytes, reservedOffset) !== 0x003e) return undefined;
  const sizeOfNameUnicode = readU32(bytes, reservedOffset + 2);
  const unicodeStart = reservedOffset + 6;
  const next = unicodeStart + sizeOfNameUnicode;
  if (sizeOfNameUnicode > bytes.length - unicodeStart) return undefined;
  const unicode = decodeUtf16Le(bytes.subarray(unicodeStart, next)).trim();
  const mbcs = cleanReferenceText(decodeText(bytes.subarray(nameStart, reservedOffset)));
  return {
    name: unicode || mbcs,
    next,
  };
}

function parseDirReferenceRecord(id: number, payload: Uint8Array) {
  if (id === 0x000d) {
    const libId = parseReferenceLibIdRecord(payload, "length-prefixed");
    return { libId, ...parseLibId(libId) };
  }

  if (id === 0x000e) {
    const absolute = parseReferenceLibIdRecord(payload, "length-prefixed");
    let relative: string | undefined;
    if (absolute) {
      const relativeOffset = 4 + byteLength(absolute);
      if (relativeOffset + 4 <= payload.length) relative = parseReferenceLibIdRecord(payload.subarray(relativeOffset), "length-prefixed");
    }
    const libId = relative || absolute;
    return { libId, ...parseLibId(libId) };
  }

  const libId = parseReferenceLibIdRecord(payload, "length-prefixed") ?? parseReferenceLibIdRecord(payload, "direct");
  return { libId, ...parseLibId(libId) };
}

function parseReferenceLibIdRecord(payload: Uint8Array, mode: "direct" | "length-prefixed") {
  let raw = "";
  if (mode === "length-prefixed" && payload.length >= 4) {
    const size = readU32(payload, 0);
    if (size > 0 && size <= payload.length - 4) raw = decodeText(payload.subarray(4, 4 + size));
  }
  if (!raw && mode === "direct") raw = decodeText(payload);
  const cleaned = normalizeReferenceLibId(cleanReferenceText(raw));
  return cleaned || undefined;
}

function classifyDirReferenceRecord(id: number): ProjectReferenceModel["kind"] {
  if (id === 0x002f || id === 0x0030) return "control";
  if (id === 0x000d) return "registered";
  if (id === 0x000e) return "project";
  return "unknown";
}

function inferReferenceName(value: string) {
  const parts = value.split("#").map((part) => part.trim()).filter(Boolean);
  const candidate = parts.find((part) => /^[A-Za-z_][A-Za-z0-9_. -]*$/.test(part) && !/^\d+\.\d+$/.test(part));
  return candidate;
}

function normalizeReferenceLibId(value: string) {
  const marker = value.search(/\*?\\G\{/i);
  const normalized = marker >= 0 ? value.slice(marker) : value;
  return normalized.replace(/[^\x20-\x7e]+$/g, "").trim();
}

function parseLibId(value: string | undefined) {
  if (!value) return {};
  const parts = value.split("#");
  return {
    version: parts.find((part) => /^\d+\.\d+$/.test(part)),
    path: parts.find((part) => /[A-Z]:\\|\/|\.dll|\.ocx|\.tlb|\.olb/i.test(part)),
  };
}

function cleanReferenceText(value: string) {
  return value.replace(/\0/g, "").replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]+/g, "").trim();
}

function byteLength(value: string) {
  return encodeText(value).byteLength;
}

function parseDirModules(dirStream?: Uint8Array): DirModule[] {
  if (!dirStream) return [];

  let dir: Uint8Array;
  try {
    dir = decompressVba(dirStream);
  } catch {
    return [];
  }

  const start = findProjectModulesRecord(dir);
  if (start < 0) return [];

  try {
    let cursor = start;
    const projectModules = readRecord(dir, cursor);
    const count = readU16(projectModules.payload, 0);
    cursor = projectModules.next;

    const projectCookie = readRecord(dir, cursor);
    if (projectCookie.id !== 0x0013) return [];
    cursor = projectCookie.next;

    const modules: DirModule[] = [];
    for (let index = 0; index < count && cursor < dir.length; index += 1) {
      const parsed = readDirModule(dir, cursor);
      if (!parsed) break;
      modules.push(parsed.module);
      cursor = parsed.next;
    }
    return modules;
  } catch {
    return [];
  }
}

function findProjectModulesRecord(dir: Uint8Array) {
  for (let offset = 0; offset + 16 <= dir.length; offset += 1) {
    if (readU16(dir, offset) !== 0x000f || readU32(dir, offset + 2) !== 0x00000002) continue;
    const count = readU16(dir, offset + 6);
    const nextRecord = offset + 8;
    if (count > 0 && count < 1000 && readU16(dir, nextRecord) === 0x0013 && readU32(dir, nextRecord + 2) === 0x00000002) {
      return offset;
    }
  }
  return -1;
}

function readDirModule(dir: Uint8Array, start: number) {
  let cursor = start;

  const nameRecord = readRecord(dir, cursor);
  if (nameRecord.id !== 0x0019) return undefined;
  const name = decodeText(nameRecord.payload);
  cursor = nameRecord.next;

  const maybeNameUnicode = readRecord(dir, cursor);
  if (maybeNameUnicode.id === 0x0047) cursor = maybeNameUnicode.next;

  const streamNameRecord = readRecord(dir, cursor);
  if (streamNameRecord.id !== 0x001a) return undefined;
  const streamName = decodeText(streamNameRecord.payload);
  cursor = streamNameRecord.next;

  for (const expectedId of [0x0032, 0x001c, 0x0048]) {
    const record = readRecord(dir, cursor);
    if (record.id !== expectedId) return undefined;
    cursor = record.next;
  }

  const offsetRecord = readRecord(dir, cursor);
  if (offsetRecord.id !== 0x0031 || offsetRecord.payload.byteLength < 4) return undefined;
  const textOffset = readU32(offsetRecord.payload, 0);
  cursor = offsetRecord.next;

  for (const expectedId of [0x001e, 0x002c]) {
    const record = readRecord(dir, cursor);
    if (record.id !== expectedId) return undefined;
    cursor = record.next;
  }

  let moduleTypeId: number | undefined;
  while (cursor + 2 <= dir.length) {
    const id = readU16(dir, cursor);
    cursor += 2;

    if (id === 0x0021 || id === 0x0022) {
      moduleTypeId = id;
      cursor += 4;
      continue;
    }

    if (id === 0x0025 || id === 0x0028) {
      cursor += 4;
      continue;
    }

    if (id === 0x002b) {
      cursor += 4;
      break;
    }

    return undefined;
  }

  return {
    module: { name, streamName, textOffset, moduleTypeId },
    next: cursor,
  };
}

function readRecord(bytes: Uint8Array, offset: number) {
  if (offset + 6 > bytes.length) throw new Error("Record is out of bounds.");
  const id = readU16(bytes, offset);
  const size = readU32(bytes, offset + 2);
  const payloadStart = offset + 6;
  const next = payloadStart + size;
  if (next > bytes.length) throw new Error("Record payload is out of bounds.");
  return { id, size, payload: bytes.subarray(payloadStart, next), next };
}

function findVbaModuleStream(streams: Map<string, Uint8Array>, streamName: string): [string, Uint8Array] | undefined {
  const wanted = `vba/${streamName}`.toLowerCase();
  for (const entry of streams) {
    if (entry[0].toLowerCase() === wanted) return entry;
  }
  return undefined;
}

function createSourceFile(input: { moduleName: string; moduleType: ProjectModule["type"]; source: string; sourcePath: string }): ExtractedFile {
  const extension = input.moduleType === "form" ? "frm" : input.moduleType === "class" ? "cls" : "bas";
  return {
    name: `${input.moduleName}.${extension}`,
    kind: extension === "frm" ? "frm" : "vba",
    bytes: encodeText(input.source),
    text: input.source,
    mimeType: "text/plain;charset=windows-1252",
    sourcePath: input.sourcePath,
  };
}

function extractVbaTextAtOffset(stream: Uint8Array, offset: number) {
  if (offset < 0 || offset >= stream.length) return undefined;
  try {
    return decodeText(decompressVba(stream.subarray(offset))).replace(/\0+$/g, "");
  } catch {
    return undefined;
  }
}

function extractBestVbaText(stream: Uint8Array): string | undefined {
  let best = "";
  let bestScore = 0;
  for (let offset = 0; offset < stream.byteLength; offset += 1) {
    if (stream[offset] !== 0x01) continue;
    try {
      const text = decodeText(decompressVba(stream.subarray(offset)));
      const score = scoreVbaText(text);
      if (score > bestScore) {
        best = text;
        bestScore = score;
      }
    } catch {
      // Many byte offsets can look like a VBA container signature. Ignore misses.
    }
  }

  return bestScore > 8 ? best.replace(/\0+$/g, "") : undefined;
}

function decompressVba(container: Uint8Array): Uint8Array {
  if (container[0] !== 0x01) throw new Error("Invalid compressed VBA container signature.");

  const output: number[] = [];
  let cursor = 1;

  while (cursor + 2 <= container.length) {
    const header = readU16(container, cursor);
    cursor += 2;

    const chunkSize = (header & 0x0fff) + 3;
    const chunkEnd = Math.min(container.length, cursor + chunkSize - 2);
    const signature = (header >> 12) & 0x07;
    const compressed = (header & 0x8000) !== 0;

    if (signature !== 0x03) break;

    const chunkStart = output.length;
    if (!compressed) {
      while (cursor < chunkEnd) output.push(container[cursor++]);
      continue;
    }

    while (cursor < chunkEnd) {
      const flags = container[cursor++];
      for (let bit = 0; bit < 8 && cursor < chunkEnd; bit += 1) {
        if ((flags & (1 << bit)) === 0) {
          output.push(container[cursor++]);
          continue;
        }

        if (cursor + 1 >= chunkEnd) break;
        const token = readU16(container, cursor);
        cursor += 2;

        const chunkPosition = output.length - chunkStart;
        const bitCount = Math.max(4, Math.ceil(Math.log2(Math.max(chunkPosition, 1))));
        const lengthMask = 0xffff >> bitCount;
        const length = (token & lengthMask) + 3;
        const offset = (token >> (16 - bitCount)) + 1;
        const copySource = output.length - offset;

        for (let i = 0; i < length; i += 1) {
          output.push(output[copySource + i]);
        }
      }
    }

    cursor = chunkEnd;
  }

  return new Uint8Array(output);
}

function scoreVbaText(text: string) {
  let score = 0;
  if (/Attribute\s+VB_Name/i.test(text)) score += 50;
  if (/\b(Sub|Function|Property|Option|Dim|Private|Public|End)\b/i.test(text)) score += 20;
  if (/Begin\s+VB\./i.test(text)) score += 20;

  const printable = [...text.slice(0, 2000)].filter((char) => {
    const code = char.charCodeAt(0);
    return code === 9 || code === 10 || code === 13 || (code >= 32 && code < 127);
  }).length;

  return score + printable / Math.max(text.length, 1);
}

function extractMediaFromBinary(bytes: Uint8Array, baseName: string): ExtractedFile[] {
  const found: ExtractedFile[] = [];
  const seen = new Set<string>();

  for (let offset = 0; offset < bytes.length; offset += 1) {
    const match = detectMedia(bytes, offset);
    if (!match) continue;

    const mediaBytes = match.bytes ?? bytes.slice(offset, offset + match.length);
    const key = `${offset}:${mediaBytes.byteLength}:${match.extension}`;
    if (seen.has(key)) continue;
    seen.add(key);
    found.push({
      name: `${baseName}-${String(found.length + 1).padStart(2, "0")}.${match.extension}`,
      kind: "media",
      bytes: mediaBytes,
      mimeType: match.mimeType,
      sourcePath: match.label ? `${match.label} at ${toHex(offset)}` : "",
    });
    offset += Math.max(match.length - 1, 0);
  }

  for (const nested of extractNestedOleMedia(bytes, baseName, found.length)) {
    const key = `${nested.sourcePath}:${nested.bytes.byteLength}:${nested.name}`;
    if (seen.has(key)) continue;
    seen.add(key);
    found.push(nested);
  }

  return found;
}

function extractNestedOleMedia(bytes: Uint8Array, baseName: string, startIndex: number) {
  const found: ExtractedFile[] = [];

  for (let offset = 0; offset + 8 <= bytes.length; offset += 1) {
    if (!isOle(bytes.subarray(offset, offset + 8))) continue;

    try {
      const nested = parseCfb(bytes.subarray(offset));
      for (const [streamPath, streamBytes] of nested.streams) {
        const nestedMedia = extractMediaFromBinary(streamBytes, `${baseName}-${safeFileName(streamPath.replaceAll("/", "-"))}`);
        for (const media of nestedMedia) {
          found.push({
            ...media,
            name: `${baseName}-${String(startIndex + found.length + 1).padStart(2, "0")}-${safeFileName(streamPath.replaceAll("/", "-"))}.${media.name.split(".").at(-1) ?? "bin"}`,
            sourcePath: `Nested OLE at ${toHex(offset)} / ${printableStreamName(streamPath)} / ${media.sourcePath}`,
          });
        }
      }
      offset += 511;
    } catch {
      // OLE signatures can appear inside arbitrary data; ignore malformed candidates.
    }
  }

  return found;
}

function detectMedia(bytes: Uint8Array, offset: number): MediaMatch | undefined {
  const guidAndPicture = detectGuidAndPicture(bytes, offset);
  if (guidAndPicture) return guidAndPicture;

  const stdPicture = detectStdPicture(bytes, offset);
  if (stdPicture) return stdPicture;

  if (matches(bytes, offset, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    const end = findBytes(bytes, [0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82], offset + 8);
    if (end > -1) return { extension: "png", mimeType: "image/png", length: end + 8 - offset, label: "PNG image" };
  }

  if (bytes[offset] === 0xff && bytes[offset + 1] === 0xd8) {
    const end = findBytes(bytes, [0xff, 0xd9], offset + 2);
    if (end > -1) return { extension: "jpg", mimeType: "image/jpeg", length: end + 2 - offset, label: "JPEG image" };
  }

  if (matchesAscii(bytes, offset, "GIF87a") || matchesAscii(bytes, offset, "GIF89a")) {
    const end = bytes.indexOf(0x3b, offset + 13);
    if (end > -1) return { extension: "gif", mimeType: "image/gif", length: end + 1 - offset, label: "GIF image" };
  }

  if (matchesAscii(bytes, offset, "BM") && offset + 14 < bytes.length) {
    const length = readU32(bytes, offset + 2);
    if (length > 14 && offset + length <= bytes.length) return { extension: "bmp", mimeType: "image/bmp", length, label: "BMP image" };
  }

  if (offset + 6 < bytes.length && readU16(bytes, offset) === 0 && [1, 2].includes(readU16(bytes, offset + 2))) {
    const count = readU16(bytes, offset + 4);
    if (count > 0 && count < 64 && offset + 6 + count * 16 <= bytes.length) {
      let length = 6 + count * 16;
      for (let i = 0; i < count; i += 1) {
        const entry = offset + 6 + i * 16;
        const size = readU32(bytes, entry + 8);
        const imageOffset = readU32(bytes, entry + 12);
        length = Math.max(length, imageOffset + size);
      }
      if (length > 0 && offset + length <= bytes.length) {
        const isIcon = readU16(bytes, offset + 2) === 1;
        return { extension: isIcon ? "ico" : "cur", mimeType: "image/x-icon", length, label: isIcon ? "ICO icon" : "CUR cursor" };
      }
    }
  }

  if (matchesAscii(bytes, offset, "II*\0") || matchesAscii(bytes, offset, "MM\0*")) {
    const next = findNextLikelyMedia(bytes, offset + 8);
    const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
    if (length > 32) return { extension: "tif", mimeType: "image/tiff", length, label: "TIFF image" };
  }

  if (matches(bytes, offset, [0xd7, 0xcd, 0xc6, 0x9a]) && offset + 22 < bytes.length) {
    const next = findNextLikelyMedia(bytes, offset + 22);
    const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
    return { extension: "wmf", mimeType: "image/wmf", length, label: "Placeable WMF image" };
  }

  if (readU32(bytes, offset) === 1 && offset + 88 < bytes.length && matchesAscii(bytes, offset + 40, " EMF")) {
    const length = readU32(bytes, offset + 48);
    if (length > 88 && offset + length <= bytes.length) {
      return { extension: "emf", mimeType: "image/emf", length, label: "EMF image" };
    }
  }

  const dib = buildBmpFromDib(bytes, offset);
  if (dib) {
    return { extension: "bmp", mimeType: "image/bmp", length: dib.sourceLength, bytes: dib.bmp, label: "DIB image rebuilt as BMP" };
  }

  return undefined;
}

function detectGuidAndPicture(bytes: Uint8Array, offset: number): MediaMatch | undefined {
  if (!matches(bytes, offset, CLSID_STD_PICTURE)) return undefined;
  const picture = detectStdPicture(bytes, offset + 16);
  if (!picture) return undefined;
  return {
    ...picture,
    length: 16 + picture.length,
    label: `MS-OFORMS GuidAndPicture ${picture.label ?? "picture"}`,
  };
}

function detectStdPicture(bytes: Uint8Array, offset: number): MediaMatch | undefined {
  if (offset + 8 > bytes.length || readU32(bytes, offset) !== 0x0000746c) return undefined;

  const size = readU32(bytes, offset + 4);
  const payloadStart = offset + 8;
  const payloadEnd = payloadStart + size;
  if (size === 0 || payloadEnd > bytes.length) return undefined;

  const payload = bytes.subarray(payloadStart, payloadEnd);
  const decoded = decodePicturePayload(payload);
  if (!decoded) {
    return {
      extension: "picture.bin",
      mimeType: "application/octet-stream",
      length: 8 + size,
      bytes: payload.slice(),
      label: "MS-OFORMS StdPicture payload",
    };
  }

  return {
    ...decoded,
    length: 8 + size,
    label: `MS-OFORMS StdPicture ${decoded.label ?? decoded.extension.toUpperCase()}`,
  };
}

function decodePicturePayload(payload: Uint8Array): Omit<MediaMatch, "length"> | undefined {
  const direct = detectDirectImagePayload(payload);
  if (direct) return direct;

  const dib = buildBmpFromDib(payload, 0);
  if (dib) {
    return {
      extension: "bmp",
      mimeType: "image/bmp",
      bytes: dib.bmp,
      label: "DIB rebuilt as BMP",
    };
  }

  if (isStandardWmf(payload, 0)) {
    return {
      extension: "wmf",
      mimeType: "image/wmf",
      bytes: payload.slice(),
      label: "WMF image",
    };
  }

  if (isLikelyEmf(payload, 0)) {
    return {
      extension: "emf",
      mimeType: "image/emf",
      bytes: payload.slice(),
      label: "EMF image",
    };
  }

  return undefined;
}

function detectDirectImagePayload(payload: Uint8Array): Omit<MediaMatch, "length"> | undefined {
  const media = detectMediaWithoutOFormsWrappers(payload, 0);
  if (!media) return undefined;
  return {
    extension: media.extension,
    mimeType: media.mimeType,
    bytes: media.bytes ?? payload.slice(0, media.length),
    label: media.label,
  };
}

function detectMediaWithoutOFormsWrappers(bytes: Uint8Array, offset: number): MediaMatch | undefined {
  if (matches(bytes, offset, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    const end = findBytes(bytes, [0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82], offset + 8);
    if (end > -1) return { extension: "png", mimeType: "image/png", length: end + 8 - offset, label: "PNG image" };
  }

  if (bytes[offset] === 0xff && bytes[offset + 1] === 0xd8) {
    const end = findBytes(bytes, [0xff, 0xd9], offset + 2);
    if (end > -1) return { extension: "jpg", mimeType: "image/jpeg", length: end + 2 - offset, label: "JPEG image" };
  }

  if (matchesAscii(bytes, offset, "GIF87a") || matchesAscii(bytes, offset, "GIF89a")) {
    const end = bytes.indexOf(0x3b, offset + 13);
    if (end > -1) return { extension: "gif", mimeType: "image/gif", length: end + 1 - offset, label: "GIF image" };
  }

  if (matchesAscii(bytes, offset, "BM") && offset + 14 < bytes.length) {
    const length = readU32(bytes, offset + 2);
    if (length > 14 && offset + length <= bytes.length) return { extension: "bmp", mimeType: "image/bmp", length, label: "BMP image" };
  }

  const icon = detectIconOrCursor(bytes, offset);
  if (icon) return icon;

  if (matchesAscii(bytes, offset, "II*\0") || matchesAscii(bytes, offset, "MM\0*")) {
    const next = findNextLikelyMedia(bytes, offset + 8);
    const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
    if (length > 32) return { extension: "tif", mimeType: "image/tiff", length, label: "TIFF image" };
  }

  if (matches(bytes, offset, [0xd7, 0xcd, 0xc6, 0x9a])) {
    const next = findNextLikelyMedia(bytes, offset + 22);
    const length = Math.min((next > -1 ? next : bytes.length) - offset, bytes.length - offset);
    return { extension: "wmf", mimeType: "image/wmf", length, label: "Placeable WMF image" };
  }

  if (isLikelyEmf(bytes, offset)) {
    const length = readU32(bytes, offset + 48);
    if (length > 88 && offset + length <= bytes.length) {
      return { extension: "emf", mimeType: "image/emf", length, label: "EMF image" };
    }
  }

  return undefined;
}

function findNextLikelyMedia(bytes: Uint8Array, start: number) {
  for (let offset = start; offset < bytes.length; offset += 1) {
    if (
      matches(bytes, offset, [0x89, 0x50, 0x4e, 0x47]) ||
      (bytes[offset] === 0xff && bytes[offset + 1] === 0xd8) ||
      matchesAscii(bytes, offset, "GIF8") ||
      matchesAscii(bytes, offset, "BM")
    ) {
      return offset;
    }
  }
  return -1;
}

function detectIconOrCursor(bytes: Uint8Array, offset: number): MediaMatch | undefined {
  if (offset + 6 >= bytes.length || readU16(bytes, offset) !== 0 || ![1, 2].includes(readU16(bytes, offset + 2))) {
    return undefined;
  }

  const count = readU16(bytes, offset + 4);
  if (count <= 0 || count >= 64 || offset + 6 + count * 16 > bytes.length) return undefined;

  let length = 6 + count * 16;
  for (let i = 0; i < count; i += 1) {
    const entry = offset + 6 + i * 16;
    const size = readU32(bytes, entry + 8);
    const imageOffset = readU32(bytes, entry + 12);
    if (size === 0 || imageOffset < 6 + count * 16) return undefined;
    length = Math.max(length, imageOffset + size);
  }

  if (offset + length > bytes.length) return undefined;
  const isIcon = readU16(bytes, offset + 2) === 1;
  return { extension: isIcon ? "ico" : "cur", mimeType: "image/x-icon", length, label: isIcon ? "ICO icon" : "CUR cursor" };
}

function isStandardWmf(bytes: Uint8Array, offset: number) {
  if (offset + 18 > bytes.length) return false;
  const fileType = readU16(bytes, offset);
  const headerSizeWords = readU16(bytes, offset + 2);
  const windowsVersion = readU16(bytes, offset + 4);
  const fileSizeWords = readU32(bytes, offset + 6);
  return [1, 2].includes(fileType) && headerSizeWords === 9 && windowsVersion >= 0x0100 && fileSizeWords > 9 && offset + fileSizeWords * 2 <= bytes.length;
}

function isLikelyEmf(bytes: Uint8Array, offset: number) {
  return readU32(bytes, offset) === 1 && offset + 88 < bytes.length && matchesAscii(bytes, offset + 40, " EMF");
}

function analyzeBinaryStream(bytes: Uint8Array, path: string): BinaryAnalysis {
  const signatures = collectSignatures(bytes);
  const strings = collectStrings(bytes);
  const guids = collectGuids(bytes);
  const oforms = parseOFormsStream(bytes, path, strings);
  const entropy = calculateEntropy(bytes);
  const zeroBytes = bytes.reduce((count, byte) => count + (byte === 0 ? 1 : 0), 0);
  const pathParts = path.split("/");
  const streamName = pathParts.at(-1) ?? path;
  const storagePath = pathParts.slice(0, -1).join("/") || "(root)";

  return {
    title: describeFormStream(path),
    summary: [
      { label: "Storage", value: storagePath },
      { label: "Stream", value: printableStreamName(streamName) },
      { label: "Size", value: formatBytes(bytes.byteLength) },
      { label: "Entropy", value: `${entropy.toFixed(2)} bits/byte` },
      { label: "Zero bytes", value: `${zeroBytes} (${bytes.length ? ((zeroBytes / bytes.length) * 100).toFixed(1) : "0.0"}%)` },
      { label: "Media candidates", value: String(signatures.filter((item) => /image|bitmap|icon|cursor/i.test(item.label)).length) },
      { label: "Text strings", value: String(strings.length) },
      { label: "GUID candidates", value: String(guids.length) },
      { label: "MS-OFORMS", value: oforms ? `${oforms.kind} (${oforms.confidence})` : "No structured match" },
    ],
    stringCount: strings.length,
    guidCount: guids.length,
    signatures: signatures.slice(0, 24),
    oforms,
  };
}

function describeFormStream(path: string) {
  const name = path.split("/").at(-1) ?? path;
  if (name === "f") return "Form/control properties stream";
  if (name === "o") return "Object/property data stream";
  if (name === "x") return "Extended form data stream";
  if (name === "\x01CompObj") return "COM compound object metadata";
  if (name === "\x03VBFrame") return "VB frame metadata";
  if (/^i\d+$/i.test(name)) return "Embedded control storage";
  if (/^PROJECT/i.test(name)) return "VBA project metadata stream";
  return "Binary form resource stream";
}

function printableStreamName(value: string) {
  return value.replace(/\x01/g, "0x01 ").replace(/\x03/g, "0x03 ");
}

function collectSignatures(bytes: Uint8Array) {
  const signatures: BinaryAnalysis["signatures"] = [];
  for (let offset = 0; offset < bytes.length; offset += 1) {
    if (readU32(bytes, offset) === 0x0000746c && offset + 8 <= bytes.length) {
      const pictureSize = readU32(bytes, offset + 4);
      if (pictureSize > 0 && offset + 8 + pictureSize <= bytes.length) {
        signatures.push({ offset, label: "MS-OFORMS StdPicture", detail: `${formatBytes(pictureSize)} picture payload` });
        offset += 7;
        continue;
      }
    }

    const media = detectMedia(bytes, offset);
    if (media) {
      signatures.push({
        offset,
        label: media.label ?? `${media.extension.toUpperCase()} media`,
        detail: `${formatBytes(media.bytes?.byteLength ?? media.length)}${media.bytes ? " extracted/rebuilt" : ""}`,
      });
      offset += Math.max(media.length - 1, 0);
      continue;
    }

    if (isOle(bytes.subarray(offset, offset + 8))) {
      signatures.push({ offset, label: "Nested OLE compound file", detail: "D0 CF 11 E0 A1 B1 1A E1" });
    }
  }
  return signatures;
}

function collectStrings(bytes: Uint8Array): ExtractedString[] {
  const strings: ExtractedString[] = [];
  const ascii = /[\x20-\x7e]{4,}/g;
  const asciiText = Array.from(bytes, (byte) => (byte >= 0x20 && byte <= 0x7e ? String.fromCharCode(byte) : "\n")).join("");
  let match: RegExpExecArray | null;
  while ((match = ascii.exec(asciiText))) {
    strings.push({ encoding: "ASCII", offset: match.index, value: match[0] });
  }

  let start = -1;
  let chars = "";
  for (let offset = 0; offset + 1 < bytes.length; offset += 2) {
    const code = readU16(bytes, offset);
    if (code >= 0x20 && code <= 0x7e) {
      if (start < 0) start = offset;
      chars += String.fromCharCode(code);
      continue;
    }
    if (chars.length >= 4) strings.push({ encoding: "UTF-16LE", offset: start, value: chars });
    start = -1;
    chars = "";
  }
  if (chars.length >= 4) strings.push({ encoding: "UTF-16LE", offset: start, value: chars });

  return strings
    .filter((item, index, all) => all.findIndex((candidate) => candidate.encoding === item.encoding && candidate.offset === item.offset && candidate.value === item.value) === index)
    .sort((a, b) => a.offset - b.offset);
}

function collectGuids(bytes: Uint8Array) {
  const guids: BinaryAnalysis["guids"] = [];
  for (let offset = 0; offset + 16 <= bytes.length; offset += 1) {
    const chunk = bytes.subarray(offset, offset + 16);
    const nonZero = chunk.some((byte) => byte !== 0);
    const nonFF = chunk.some((byte) => byte !== 0xff);
    if (!nonZero || !nonFF) continue;

    const variant = chunk[8] & 0xc0;
    const version = (chunk[7] >> 4) & 0x0f;
    if (variant !== 0x80 || version > 5) continue;

    guids.push({ offset, value: formatGuid(chunk) });
    offset += 15;
  }
  return guids;
}

function labelGuid(guid: string) {
  return KNOWN_GUID_LABELS[guid.toUpperCase()];
}

function labelProgId(value: string) {
  const normalized = value.trim();
  return KNOWN_PROGID_LABELS[normalized.toLowerCase()];
}

function labelPossibleIdentifier(value: string) {
  const progLabel = labelProgId(value);
  if (progLabel) return `${value} (${progLabel})`;

  const guidMatch = value.match(/\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}/i);
  if (guidMatch) {
    const guid = guidMatch[0].toUpperCase();
    const guidLabel = labelGuid(guid);
    if (guidLabel) return value.replace(guidMatch[0], `${guid} (${guidLabel})`);
  }

  return value;
}

function parseOFormsStream(bytes: Uint8Array, path: string, strings: ExtractedString[]): OFormsAnalysis | undefined {
  const streamName = path.split("/").at(-1) ?? path;

  if (streamName === "f") {
    const record = parseOFormsControlRecord(bytes, 0, "FormControl", "form", strings);
    if (!record) return undefined;
    return {
      kind: "Form stream",
      confidence: record.size <= bytes.length ? "high" : "medium",
      records: [record],
      notes: [
        "Parsed as an MS-OFORMS FormControl header. Property values with variable sizing are decoded only when they can be identified confidently.",
        "Embedded site/control records can also be represented in this stream after the top-level FormControl data.",
      ],
    };
  }

  if (streamName === "o") {
    const records = scanOFormsObjectRecords(bytes, strings);
    if (records.length === 0) return undefined;
    return {
      kind: "Object stream",
      confidence: records.length > 1 || records[0].offset === 0 ? "medium" : "low",
      records,
      notes: ["Scanned for MS-OFORMS control records inside the object stream. Offsets are best-effort when the stream contains mixed payloads."],
    };
  }

  if (streamName === "x") {
    return parseOFormsExtendedStream(bytes);
  }

  if (streamName === "\x01CompObj") {
    return parseCompObjStream(bytes);
  }

  if (streamName === "\x03VBFrame") {
    return {
      kind: "VBFrame metadata",
      confidence: "medium",
      records: [{
        offset: 0,
        type: "VBFrame stream",
        size: bytes.byteLength,
        properties: [
          { name: "Size", value: formatBytes(bytes.byteLength) },
          ...strings.slice(0, 8).map((item) => ({ name: `${item.encoding} string @ ${toHex(item.offset)}`, value: item.value })),
        ],
      }],
      notes: ["VBFrame streams are Office Forms metadata used by frame-like containers."],
    };
  }

  return undefined;
}

function parseOFormsControlRecord(bytes: Uint8Array, offset: number, type: string, maskKind: "form" | "site" | "generic", strings: ExtractedString[]) {
  if (offset + 8 > bytes.length) return undefined;
  const minor = bytes[offset];
  const major = bytes[offset + 1];
  const cb = readU16(bytes, offset + 2);
  const propMask = readU32(bytes, offset + 4);
  const end = offset + 4 + cb;

  if (minor !== 0 || ![0x02, 0x04].includes(major) || cb < 4 || cb > bytes.length - offset) return undefined;

  const properties: Array<{ name: string; value: string }> = [
    { name: "MinorVersion", value: toHex(minor, 2) },
    { name: "MajorVersion", value: toHex(major, 2) },
    { name: "cb", value: `${cb} bytes` },
    { name: "PropMask", value: `${toHex(propMask, 8)} (${decodeMaskFlags(propMask, maskKind).join(", ") || "no known flags"})` },
  ];

  properties.push(...decodeOFormsLayoutProperties(bytes, offset, end, propMask, maskKind));
  properties.push(...decodeCommonOFormsValues(bytes, offset, end, propMask, maskKind, strings));

  return {
    offset,
    type,
    size: Math.min(end - offset, bytes.length - offset),
    properties,
  };
}

function scanOFormsObjectRecords(bytes: Uint8Array, strings: ExtractedString[]) {
  const records: OFormsAnalysis["records"] = [];
  for (let offset = 0; offset + 8 <= bytes.length; offset += 1) {
    const record = parseOFormsControlRecord(bytes, offset, offset === 0 ? "OleSiteConcrete control" : "Nested OleSite/control record", "site", strings);
    if (!record) continue;
    records.push(record);
    offset += Math.max(record.size - 1, 0);
    if (records.length >= 32) break;
  }
  return records;
}

function parseOFormsExtendedStream(bytes: Uint8Array): OFormsAnalysis | undefined {
  if (bytes.byteLength < 8) return undefined;
  const records: OFormsAnalysis["records"] = [];
  for (let offset = 0; offset + 8 <= bytes.length; offset += 8) {
    records.push({
      offset,
      type: "PageProperties / extended data candidate",
      size: Math.min(8, bytes.length - offset),
      properties: [
        { name: "Value1", value: toHex(readU32(bytes, offset), 8) },
        { name: "Value2", value: toHex(readU32(bytes, offset + 4), 8) },
      ],
    });
  }
  return {
    kind: "Extended stream",
    confidence: "low",
    records,
    notes: ["The x stream is used by some container controls such as MultiPage. Values are shown as page/extended-data candidates."],
  };
}

function parseCompObjStream(bytes: Uint8Array): OFormsAnalysis {
  const strings = collectStrings(bytes);
  return {
    kind: "COM compound object",
    confidence: strings.some((item) => /Microsoft Forms/i.test(item.value)) ? "high" : "medium",
    records: [{
      offset: 0,
      type: "CompObj",
      size: bytes.byteLength,
      properties: strings.map((item) => ({ name: `${item.encoding} string @ ${toHex(item.offset)}`, value: labelPossibleIdentifier(item.value) })),
    }],
    notes: ["CompObj identifies embedded COM/OLE object metadata for the form or control."],
  };
}

function decodeMaskFlags(mask: number, kind: "form" | "site" | "generic") {
  const flags = kind === "form" ? FORM_PROP_FLAGS : kind === "site" ? SITE_PROP_FLAGS : GENERIC_CONTROL_FLAGS;
  return flags.filter((flag) => (mask & (1 << flag.bit)) !== 0).map((flag) => flag.name);
}

function decodeCommonOFormsValues(
  bytes: Uint8Array,
  offset: number,
  end: number,
  propMask: number,
  maskKind: "form" | "site" | "generic",
  strings: ExtractedString[],
) {
  const properties: Array<{ name: string; value: string }> = [];
  const recordStrings = strings.filter((item) => item.offset >= offset + 8 && item.offset < end);
  const streamDataStrings = strings.filter((item) => item.offset >= end && item.offset < Math.min(bytes.length, end + 128));

  properties.push(...decodeOrderedOFormsStringProperties(propMask, maskKind, recordStrings, streamDataStrings));

  for (const item of recordStrings.slice(0, 12)) {
    properties.push({ name: `Record string @ ${toHex(item.offset)}`, value: labelPossibleIdentifier(item.value) });
  }

  for (const item of streamDataStrings.slice(0, 8)) {
    properties.push({ name: `StreamData string @ ${toHex(item.offset)}`, value: labelPossibleIdentifier(item.value) });
  }

  const fontGuidOffset = findBytes(bytes, [0x03, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11, 0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51], end);
  if (fontGuidOffset > -1) {
    properties.push({ name: "Font", value: `StdFont GUID at ${toHex(fontGuidOffset)} (MSForms StdFont)` });
  }

  const pictureOffset = findStdPicture(bytes, end);
  if (pictureOffset > -1) {
    properties.push({ name: "Picture", value: `StdPicture at ${toHex(pictureOffset)} (${formatBytes(readU32(bytes, pictureOffset + 4))})` });
  }

  if (maskKind === "form") {
    const extraStart = guessExtraDataStart(bytes, offset, end, propMask, FORM_PROP_FLAGS);
    if (extraStart > -1) {
      properties.push({ name: "ExtraDataBlock candidate", value: `${toHex(extraStart)}-${toHex(end)}` });
    }
  }

  return properties;
}

function decodeOFormsLayoutProperties(bytes: Uint8Array, offset: number, end: number, propMask: number, maskKind: "form" | "site" | "generic") {
  const properties: Array<{ name: string; value: string }> = [];
  if (maskKind === "form") {
    const sizes = findFmSizeCandidates(bytes, offset + 8, end);
    const selected = selectFormSizeCandidate(sizes);
    if ((propMask & (1 << 10)) !== 0 && selected) {
      properties.push({ name: "Decoded DisplayedSize", value: formatSizeHint(selected) });
    }
    if ((propMask & (1 << 11)) !== 0 && sizes.length > 1) {
      const logical = sizes.find((candidate) => candidate !== selected && candidate.width <= 40_000 && candidate.height <= 40_000);
      if (logical) properties.push({ name: "Decoded LogicalSize", value: formatSizeHint(logical) });
    }
  }

  if (maskKind === "site" || maskKind === "generic") {
    const positions = findFmPositionCandidates(bytes, offset + 8, Math.min(bytes.length, end + 96));
    const selected = selectSitePositionCandidate(positions);
    if ((propMask & (1 << 2)) !== 0 && selected) {
      properties.push({ name: "Decoded SitePosition", value: formatPositionHint(selected) });
    }
  }

  return properties;
}

function decodeOrderedOFormsStringProperties(
  propMask: number,
  maskKind: "form" | "site" | "generic",
  recordStrings: ExtractedString[],
  streamDataStrings: ExtractedString[],
) {
  const properties: Array<{ name: string; value: string }> = [];
  const candidates = [...recordStrings, ...streamDataStrings]
    .map((item) => ({ ...item, clean: item.value.replace(/\s+\(.+\)$/, "") }))
    .filter((item) => isUsefulDecodedPropertyString(item.clean));

  if (maskKind === "form") {
    if ((propMask & (1 << 19)) !== 0) {
      const caption = candidates.find((item) => isHumanLabel(item.clean));
      if (caption) properties.push({ name: "Decoded Caption", value: `${labelPossibleIdentifier(caption.value)} (medium)` });
    }

    const extraLabels = candidates
      .filter((item) => isHumanLabel(item.clean))
      .slice(1, 8)
      .map((item, index) => ({ name: `Decoded Child/Page Label ${index + 1}`, value: `${labelPossibleIdentifier(item.value)} (low)` }));
    properties.push(...extraLabels);
  }

  if (maskKind === "site" || maskKind === "generic") {
    const ordered: Array<{ bit: number; name: string; isString: boolean }> = [
      { bit: 0, name: "Decoded Name", isString: true },
      { bit: 1, name: "Decoded Tag", isString: true },
      { bit: 2, name: "Decoded SitePosition", isString: false },
      { bit: 3, name: "Decoded ControlTipText", isString: true },
      { bit: 4, name: "Decoded RuntimeLicKey", isString: true },
      { bit: 5, name: "Decoded ControlSource", isString: true },
      { bit: 6, name: "Decoded RowSource", isString: true },
    ];
    let stringIndex = 0;
    for (const property of ordered) {
      if ((propMask & (1 << property.bit)) === 0 || !property.isString) continue;
      const candidate = candidates[stringIndex];
      stringIndex += 1;
      if (!candidate) continue;
      properties.push({ name: property.name, value: `${labelPossibleIdentifier(candidate.value)} (low)` });
    }
  }

  return properties;
}

function isUsefulDecodedPropertyString(value: string) {
  if (!value || value.length > 160) return false;
  if (/^Tahoma$/i.test(value)) return false;
  if (/^[-=]{5,}$/.test(value)) return false;
  if (/^'/.test(value) || /https?:\/\//i.test(value)) return false;
  if (/copyright|permission is hereby|software/i.test(value)) return false;
  return true;
}

function findFmSizeCandidates(bytes: Uint8Array, start: number, end: number) {
  const candidates: Array<{ offset: number; width: number; height: number; confidence: "medium" | "low" }> = [];
  for (let offset = start; offset + 8 <= end; offset += 4) {
    const width = readU32(bytes, offset);
    const height = readU32(bytes, offset + 4);
    if (width < 100 || height < 100 || width > 32_000 || height > 32_000) continue;
    if (width === 32_000 || height === 32_000) continue;
    if (width === 0xffff || height === 0xffff) continue;
    const confidence = width < 25_000 && height < 25_000 ? "medium" : "low";
    candidates.push({ offset, width, height, confidence });
  }
  return candidates;
}

function selectFormSizeCandidate(candidates: Array<{ offset: number; width: number; height: number; confidence: "medium" | "low" }>) {
  return candidates
    .filter((candidate) => candidate.confidence !== "low")
    .sort((a, b) => b.width * b.height - a.width * a.height)[0] ?? candidates.sort((a, b) => b.width * b.height - a.width * a.height)[0];
}

function findFmPositionCandidates(bytes: Uint8Array, start: number, end: number) {
  const candidates: Array<{ offset: number; left: number; top: number; confidence: "medium" | "low" }> = [];
  for (let offset = start; offset + 8 <= end; offset += 4) {
    const left = readU32(bytes, offset);
    const top = readU32(bytes, offset + 4);
    if (left > 25_000 || top > 25_000) continue;
    if (left === 0xffff || top === 0xffff) continue;
    const confidence = left < 15_000 && top < 15_000 ? "medium" : "low";
    candidates.push({ offset, left, top, confidence });
  }
  return candidates;
}

function selectSitePositionCandidate(candidates: Array<{ offset: number; left: number; top: number; confidence: "medium" | "low" }>) {
  return candidates
    .filter((candidate) => candidate.confidence !== "low")
    .sort((a, b) => scorePositionCandidate(b) - scorePositionCandidate(a))[0] ?? candidates[0];
}

function scorePositionCandidate(candidate: { left: number; top: number; confidence: "medium" | "low" }) {
  let score = candidate.confidence === "medium" ? 3 : 0;
  if (candidate.left > 0 || candidate.top > 0) score += 1;
  if (candidate.left < 10_000 && candidate.top < 10_000) score += 1;
  return score;
}

function formatSizeHint(size: { offset: number; width: number; height: number; confidence: string }) {
  return `${size.width} x ${size.height} twips @ ${toHex(size.offset)} (${size.confidence})`;
}

function formatPositionHint(position: { offset: number; left: number; top: number; confidence: string }) {
  return `${position.left}, ${position.top} twips @ ${toHex(position.offset)} (${position.confidence})`;
}

function guessExtraDataStart(bytes: Uint8Array, offset: number, end: number, propMask: number, flags: Array<{ bit: number; name: string }>) {
  const firstRecordString = collectStrings(bytes.subarray(offset, end))
    .map((item) => item.offset + offset)
    .filter((stringOffset) => stringOffset >= offset + 8)
    .sort((a, b) => a - b)[0];

  if (firstRecordString !== undefined) return firstRecordString;

  const hasExtra = flags.some((flag) => ["DisplayedSize", "LogicalSize", "ScrollPosition", "Caption"].includes(flag.name) && (propMask & (1 << flag.bit)) !== 0);
  return hasExtra ? Math.max(offset + 8, end - 24) : -1;
}

function findStdPicture(bytes: Uint8Array, start: number) {
  for (let offset = start; offset + 8 <= bytes.length; offset += 1) {
    if (readU32(bytes, offset) !== 0x0000746c) continue;
    const size = readU32(bytes, offset + 4);
    if (size > 0 && offset + 8 + size <= bytes.length) return offset;
  }
  return -1;
}

const FORM_PROP_FLAGS = [
  { bit: 1, name: "BackColor" },
  { bit: 2, name: "ForeColor" },
  { bit: 3, name: "NextAvailableID" },
  { bit: 6, name: "BooleanProperties" },
  { bit: 7, name: "BorderStyle" },
  { bit: 8, name: "MousePointer" },
  { bit: 9, name: "ScrollBars" },
  { bit: 10, name: "DisplayedSize" },
  { bit: 11, name: "LogicalSize" },
  { bit: 12, name: "ScrollPosition" },
  { bit: 13, name: "GroupCount" },
  { bit: 15, name: "MouseIcon" },
  { bit: 16, name: "Cycle" },
  { bit: 17, name: "SpecialEffect" },
  { bit: 18, name: "BorderColor" },
  { bit: 19, name: "Caption" },
  { bit: 20, name: "Font" },
  { bit: 21, name: "Picture" },
  { bit: 22, name: "Zoom" },
  { bit: 23, name: "PictureAlignment" },
  { bit: 24, name: "PictureTiling" },
  { bit: 25, name: "PictureSizeMode" },
  { bit: 26, name: "ShapeCookie" },
  { bit: 27, name: "DrawBuffer" },
];

const SITE_PROP_FLAGS = [
  { bit: 0, name: "Name" },
  { bit: 1, name: "Tag" },
  { bit: 2, name: "Position" },
  { bit: 3, name: "ControlTipText" },
  { bit: 4, name: "RuntimeLicKey" },
  { bit: 5, name: "ControlSource" },
  { bit: 6, name: "RowSource" },
  { bit: 7, name: "Enabled/visibility flags" },
];

const GENERIC_CONTROL_FLAGS = [
  { bit: 0, name: "Property0" },
  { bit: 1, name: "Property1" },
  { bit: 2, name: "Property2" },
  { bit: 3, name: "Property3" },
  { bit: 4, name: "Property4" },
  { bit: 5, name: "Property5" },
  { bit: 6, name: "Property6" },
  { bit: 7, name: "Property7" },
  { bit: 8, name: "Property8" },
  { bit: 9, name: "Property9" },
  { bit: 10, name: "Size/position" },
  { bit: 11, name: "Caption/name string" },
  { bit: 12, name: "Font" },
  { bit: 13, name: "Picture/icon" },
  { bit: 14, name: "Mouse/icon data" },
  { bit: 15, name: "Extra data" },
  { bit: 16, name: "Property16" },
  { bit: 17, name: "Property17" },
  { bit: 18, name: "Property18" },
  { bit: 19, name: "Property19" },
  { bit: 20, name: "Property20" },
  { bit: 21, name: "Property21" },
  { bit: 22, name: "Property22" },
  { bit: 23, name: "Property23" },
  { bit: 24, name: "Property24" },
  { bit: 25, name: "Property25" },
  { bit: 26, name: "Property26" },
  { bit: 27, name: "Property27" },
  { bit: 28, name: "Property28" },
  { bit: 29, name: "Property29" },
  { bit: 30, name: "Property30" },
  { bit: 31, name: "Property31" },
];

const KNOWN_GUID_LABELS: Record<string, string> = {
  "{00020430-0000-0000-C000-000000000046}": "OLE Automation / stdole",
  "{00020813-0000-0000-C000-000000000046}": "Microsoft Excel Object Library",
  "{00020905-0000-0000-C000-000000000046}": "Microsoft Word Object Library",
  "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}": "Microsoft Office Object Library",
  "{0002E157-0000-0000-C000-000000000046}": "Microsoft Visual Basic for Applications Extensibility",
  "{00062FFF-0000-0000-C000-000000000046}": "Microsoft Outlook Object Library",
  "{91493440-5A91-11CF-8700-00AA0060263B}": "Microsoft PowerPoint Object Library",
  "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}": "Microsoft Access Object Library",
  "{0D452EE1-E08F-101A-852E-02608C4D0BB4}": "Microsoft Forms 2.0 Object Library",
  "{0BE35203-8F91-11CE-9DE3-00AA004BB851}": "MSForms StdFont",
  "{0BE35204-8F91-11CE-9DE3-00AA004BB851}": "MSForms StdPicture",
  "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}": "Microsoft Forms 2.0 Form",
  "{46E31370-3F7A-11CE-BED6-00AA00611080}": "Microsoft Forms 2.0 MultiPage",
  "{978C9E23-D4B0-11CE-BF2D-00AA003F40D0}": "Microsoft Forms 2.0 Label",
  "{EAE50EB0-4A62-11CE-BED6-00AA00611080}": "Microsoft Forms 2.0 TextBox",
  "{A8BD21D30-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 ComboBox",
  "{B7D5D0F0-9D69-11CE-A98A-00AA00602F79}": "Microsoft Forms 2.0 ListBox",
  "{F8D77580-0F09-11D0-AA61-3C284E000000}": "Microsoft Forms 2.0 Image",
  "{79176FB0-B7F2-11CE-97EF-00AA006D2776}": "Microsoft Forms 2.0 ScrollBar",
  "{79176FB1-B7F2-11CE-97EF-00AA006D2776}": "Microsoft Forms 2.0 SpinButton",
  "{A8A6A960-7B3B-11D0-BB40-00A0C90F2744}": "Microsoft Forms 2.0 TabStrip",
  "{47B0DFC7-B7A3-11CE-A9F4-00AA006D2776}": "Microsoft Forms 2.0 MultiPage",
  "{6E182020-F460-11CE-9BCD-00AA00608E01}": "Microsoft Forms 2.0 Frame",
  "{D7053240-CE69-11CD-A777-00DD01143C57}": "Microsoft Forms 2.0 CommandButton",
  "{D7053241-CE69-11CD-A777-00DD01143C57}": "Microsoft Forms 2.0 OptionButton",
  "{D7053242-CE69-11CD-A777-00DD01143C57}": "Microsoft Forms 2.0 ToggleButton",
  "{D7053243-CE69-11CD-A777-00DD01143C57}": "Microsoft Forms 2.0 CommandButton",
  "{8BD21D10-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 TextBox",
  "{8BD21D20-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 ListBox",
  "{8BD21D30-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 ComboBox",
  "{8BD21D40-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 CheckBox",
  "{8BD21D50-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 OptionButton",
  "{8BD21D60-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 ToggleButton",
  "{8BD21D70-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 Frame",
  "{8BD21D80-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 CommandButton",
  "{8BD21D90-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 TabStrip",
  "{8BD21DA0-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 Image",
  "{8BD21DB0-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 ScrollBar",
  "{8BD21DC0-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 SpinButton",
  "{8BD21DD0-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 Label",
};

const COMMON_FORMS_CONTROLS = [
  ["Forms.Form.1", "Microsoft Forms 2.0 UserForm"],
  ["Forms.Frame.1", "Microsoft Forms 2.0 Frame control"],
  ["Forms.MultiPage.1", "Microsoft Forms 2.0 MultiPage control"],
  ["Forms.Page.1", "Microsoft Forms 2.0 Page control"],
  ["Forms.TabStrip.1", "Microsoft Forms 2.0 TabStrip control"],
  ["Forms.Label.1", "Microsoft Forms 2.0 Label control"],
  ["Forms.TextBox.1", "Microsoft Forms 2.0 TextBox control"],
  ["Forms.ComboBox.1", "Microsoft Forms 2.0 ComboBox control"],
  ["Forms.ListBox.1", "Microsoft Forms 2.0 ListBox control"],
  ["Forms.CheckBox.1", "Microsoft Forms 2.0 CheckBox control"],
  ["Forms.OptionButton.1", "Microsoft Forms 2.0 OptionButton control"],
  ["Forms.ToggleButton.1", "Microsoft Forms 2.0 ToggleButton control"],
  ["Forms.CommandButton.1", "Microsoft Forms 2.0 CommandButton control"],
  ["Forms.Image.1", "Microsoft Forms 2.0 Image control"],
  ["Forms.ScrollBar.1", "Microsoft Forms 2.0 ScrollBar control"],
  ["Forms.SpinButton.1", "Microsoft Forms 2.0 SpinButton control"],
  ["Forms.HTMLCheckBox.1", "Microsoft Forms 2.0 HTML CheckBox control"],
  ["Forms.HTMLHidden.1", "Microsoft Forms 2.0 HTML Hidden control"],
  ["Forms.HTMLImage.1", "Microsoft Forms 2.0 HTML Image control"],
  ["Forms.HTMLOption.1", "Microsoft Forms 2.0 HTML Option control"],
  ["Forms.HTMLPassword.1", "Microsoft Forms 2.0 HTML Password control"],
  ["Forms.HTMLReset.1", "Microsoft Forms 2.0 HTML Reset control"],
  ["Forms.HTMLSelect.1", "Microsoft Forms 2.0 HTML Select control"],
  ["Forms.HTMLSubmit.1", "Microsoft Forms 2.0 HTML Submit control"],
  ["Forms.HTMLText.1", "Microsoft Forms 2.0 HTML Text control"],
  ["Forms.HTMLTextArea.1", "Microsoft Forms 2.0 HTML TextArea control"],
] as const;

const KNOWN_PROGID_LABELS: Record<string, string> = Object.fromEntries(
  COMMON_FORMS_CONTROLS.map(([progId, label]) => [progId.toLowerCase(), label]),
);

function buildBmpFromDib(bytes: Uint8Array, offset: number) {
  if (offset + 40 > bytes.length) return undefined;
  const headerSize = readU32(bytes, offset);
  if (![40, 52, 56, 108, 124].includes(headerSize)) return undefined;

  const width = readI32(bytes, offset + 4);
  const height = readI32(bytes, offset + 8);
  const planes = readU16(bytes, offset + 12);
  const bitCount = readU16(bytes, offset + 14);
  const compression = readU32(bytes, offset + 16);
  const declaredImageSize = readU32(bytes, offset + 20);
  const colorsUsed = readU32(bytes, offset + 32);

  if (width <= 0 || Math.abs(height) <= 0 || width > 20000 || Math.abs(height) > 20000) return undefined;
  if (planes !== 1 || ![1, 4, 8, 16, 24, 32].includes(bitCount) || compression > 6) return undefined;

  const paletteEntries = bitCount <= 8 ? colorsUsed || 1 << bitCount : 0;
  const pixelOffset = headerSize + paletteEntries * 4;
  const rowSize = Math.floor((bitCount * width + 31) / 32) * 4;
  const calculatedImageSize = rowSize * Math.abs(height);
  const imageSize = declaredImageSize || calculatedImageSize;
  const sourceLength = pixelOffset + imageSize;
  if (sourceLength <= headerSize || offset + sourceLength > bytes.length) return undefined;

  const bmp = new Uint8Array(14 + sourceLength);
  bmp[0] = 0x42;
  bmp[1] = 0x4d;
  writeU32(bmp, 2, bmp.length);
  writeU32(bmp, 10, 14 + pixelOffset);
  bmp.set(bytes.subarray(offset, offset + sourceLength), 14);
  return { bmp, sourceLength };
}

function hexDump(bytes: Uint8Array, start: number, length: number) {
  const lines: string[] = [];
  for (let offset = start; offset < start + length; offset += 16) {
    const chunk = bytes.subarray(offset, Math.min(offset + 16, start + length));
    const hex = [...chunk].map((byte) => byte.toString(16).padStart(2, "0")).join(" ").padEnd(47, " ");
    const ascii = [...chunk].map((byte) => (byte >= 0x20 && byte <= 0x7e ? String.fromCharCode(byte) : ".")).join("");
    lines.push(`${toHex(offset, 8)}  ${hex}  ${ascii}`);
  }
  return lines.join("\n");
}

function calculateEntropy(bytes: Uint8Array) {
  if (bytes.length === 0) return 0;
  const counts = new Array<number>(256).fill(0);
  for (const byte of bytes) counts[byte] += 1;
  return counts.reduce((entropy, count) => {
    if (count === 0) return entropy;
    const probability = count / bytes.length;
    return entropy - probability * Math.log2(probability);
  }, 0);
}

function parseCfb(bytes: Uint8Array) {
  if (!isOle(bytes)) throw new Error("Invalid OLE compound document.");

  const sectorSize = 1 << readU16(bytes, 30);
  const miniSectorSize = 1 << readU16(bytes, 32);
  const firstDirectorySector = readI32(bytes, 48);
  const miniStreamCutoff = readU32(bytes, 56);
  const firstMiniFatSector = readI32(bytes, 60);
  const miniFatSectorCount = readU32(bytes, 64);
  const firstDifatSector = readI32(bytes, 68);
  const difatSectorCount = readU32(bytes, 72);

  const difat: number[] = [];
  for (let offset = 76; offset < 512; offset += 4) {
    const sector = readI32(bytes, offset);
    if (sector >= 0) difat.push(sector);
  }

  let difatSector = firstDifatSector;
  for (let i = 0; i < difatSectorCount && difatSector >= 0; i += 1) {
    const sector = getSector(bytes, sectorSize, difatSector);
    for (let offset = 0; offset < sectorSize - 4; offset += 4) {
      const fatSector = readI32(sector, offset);
      if (fatSector >= 0) difat.push(fatSector);
    }
    difatSector = readI32(sector, sectorSize - 4);
  }

  const fat: number[] = [];
  for (const sectorIndex of difat) {
    const sector = getSector(bytes, sectorSize, sectorIndex);
    for (let offset = 0; offset < sectorSize; offset += 4) {
      fat.push(readI32(sector, offset));
    }
  }

  const readChain = (startSector: number) => {
    const chunks: Uint8Array[] = [];
    const visited = new Set<number>();
    let sector = startSector;

    while (sector >= 0 && sector !== END_OF_CHAIN && !visited.has(sector)) {
      visited.add(sector);
      chunks.push(getSector(bytes, sectorSize, sector));
      const next = fat[sector];
      if (next === undefined || next === FREE_SECTOR || next === FAT_SECTOR || next === DIFAT_SECTOR) break;
      sector = next;
    }

    return concatBytes(chunks);
  };

  const directoryBytes = readChain(firstDirectorySector);
  const entries: DirectoryEntry[] = [];

  for (let offset = 0, index = 0; offset + 128 <= directoryBytes.length; offset += 128, index += 1) {
    const entryBytes = directoryBytes.subarray(offset, offset + 128);
    const nameLength = readU16(entryBytes, 64);
    const rawName = entryBytes.subarray(0, Math.max(0, nameLength - 2));
    const name = decodeUtf16Le(rawName);
    const type = entryBytes[66];

    if (!name || type === 0) continue;

    entries[index] = {
      index,
      name,
      type,
      left: readI32(entryBytes, 68),
      right: readI32(entryBytes, 72),
      child: readI32(entryBytes, 76),
      startSector: readI32(entryBytes, 116),
      size: readU32(entryBytes, 120),
      path: name,
    };
  }

  const root = entries.find((entry) => entry.type === 5);
  const miniStream = root ? readChain(root.startSector).subarray(0, root.size) : new Uint8Array();
  const miniFatBytes = firstMiniFatSector >= 0 && miniFatSectorCount ? readChain(firstMiniFatSector) : new Uint8Array();
  const miniFat: number[] = [];
  for (let offset = 0; offset + 4 <= miniFatBytes.length; offset += 4) miniFat.push(readI32(miniFatBytes, offset));

  const readMiniChain = (startSector: number, size: number) => {
    const chunks: Uint8Array[] = [];
    const visited = new Set<number>();
    let sector = startSector;

    while (sector >= 0 && sector !== END_OF_CHAIN && !visited.has(sector)) {
      visited.add(sector);
      const offset = sector * miniSectorSize;
      chunks.push(miniStream.subarray(offset, offset + miniSectorSize));
      const next = miniFat[sector];
      if (next === undefined || next === FREE_SECTOR) break;
      sector = next;
    }

    return concatBytes(chunks).subarray(0, size);
  };

  const readStream = (entry: DirectoryEntry) => {
    if (entry.size === 0 || entry.startSector < 0) return new Uint8Array();
    if (entry.size < miniStreamCutoff && root && entry.index !== root.index) {
      return readMiniChain(entry.startSector, entry.size);
    }
    return readChain(entry.startSector).subarray(0, entry.size);
  };

  const streams = new Map<string, Uint8Array>();
  const walk = (index: number, parent: string) => {
    const entry = entries[index];
    if (!entry) return;
    if (entry.left >= 0) walk(entry.left, parent);
    const path = parent ? `${parent}/${entry.name}` : entry.name;
    entry.path = path;
    if (entry.type === 1 && entry.child >= 0) walk(entry.child, path);
    if (entry.type === 2) streams.set(path.replace(/^Root Entry\//, ""), readStream(entry));
    if (entry.right >= 0) walk(entry.right, parent);
  };

  if (root?.child !== undefined && root.child >= 0) walk(root.child, "");

  return { entries, streams };
}

function parseZip(bytes: Uint8Array): ZipEntry[] {
  const eocdOffset = findEndOfCentralDirectory(bytes);
  if (eocdOffset < 0) throw new Error("Could not find the Office zip directory.");

  const entriesCount = readU16(bytes, eocdOffset + 10);
  const centralDirectoryOffset = readU32(bytes, eocdOffset + 16);
  const entries: ZipEntry[] = [];
  let cursor = centralDirectoryOffset;

  for (let i = 0; i < entriesCount; i += 1) {
    if (readU32(bytes, cursor) !== 0x02014b50) throw new Error("Invalid zip central directory.");

    const method = readU16(bytes, cursor + 10);
    const compressedSize = readU32(bytes, cursor + 20);
    const uncompressedSize = readU32(bytes, cursor + 24);
    const nameLength = readU16(bytes, cursor + 28);
    const extraLength = readU16(bytes, cursor + 30);
    const commentLength = readU16(bytes, cursor + 32);
    const localHeaderOffset = readU32(bytes, cursor + 42);
    const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));

    entries.push({ name, method, compressedSize, uncompressedSize, localHeaderOffset });
    cursor += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

async function inflateZipEntry(bytes: Uint8Array, entry: ZipEntry) {
  const cursor = entry.localHeaderOffset;
  if (readU32(bytes, cursor) !== 0x04034b50) throw new Error(`Invalid local zip header for ${entry.name}.`);

  const nameLength = readU16(bytes, cursor + 26);
  const extraLength = readU16(bytes, cursor + 28);
  const dataStart = cursor + 30 + nameLength + extraLength;
  const compressed = bytes.subarray(dataStart, dataStart + entry.compressedSize);

  if (entry.method === 0) return compressed;
  if (entry.method !== 8) throw new Error(`${entry.name} uses unsupported zip compression method ${entry.method}.`);
  if (!("DecompressionStream" in window)) throw new Error("This browser does not support in-browser zip decompression.");

  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const inflated = new Uint8Array(await new Response(stream).arrayBuffer());
  if (entry.uncompressedSize && inflated.byteLength !== entry.uncompressedSize) {
    console.warn(`Unexpected uncompressed size for ${entry.name}.`);
  }
  return inflated;
}
