// Generated from split TypeScript sources by scripts/build-main.js.
// Edit the .ts source files, then run: node scripts/build-main.js
// ---- types.ts ----
const CLSID_STD_PICTURE = [
    0x04,
    0x52,
    0xe3,
    0x0b,
    0x91,
    0x8f,
    0xce,
    0x11,
    0x9d,
    0xe3,
    0x00,
    0xaa,
    0x00,
    0x4b,
    0xb8,
    0x51
];
const END_OF_CHAIN = -2;
const FREE_SECTOR = -1;
const FAT_SECTOR = -3;
const DIFAT_SECTOR = -4;


// ---- app.ts ----
const fileInput = document.querySelector("#fileInput");
const dropZone = document.querySelector("#dropZone");
const clearButton = document.querySelector("#clearButton");
const downloadAllButton = document.querySelector("#downloadAllButton");
const statusText = document.querySelector("#statusText");
const statusPanel = document.querySelector("#statusPanel");
const summaryPanel = document.querySelector("#summaryPanel");
const moduleCount = document.querySelector("#moduleCount");
const formCount = document.querySelector("#formCount");
const frxCount = document.querySelector("#frxCount");
const results = document.querySelector("#results");
let currentFiles = [];
fileInput.addEventListener("change", ()=>{
    const file = fileInput.files?.[0];
    if (file) void handleFile(file);
});
clearButton.addEventListener("click", resetUi);
downloadAllButton.addEventListener("click", ()=>downloadAll(currentFiles));
dropZone.addEventListener("dragover", (event)=>{
    event.preventDefault();
    dropZone.classList.add("dragging");
});
dropZone.addEventListener("dragleave", ()=>dropZone.classList.remove("dragging"));
dropZone.addEventListener("drop", (event)=>{
    event.preventDefault();
    dropZone.classList.remove("dragging");
    const file = event.dataTransfer?.files[0];
    if (file) void handleFile(file);
});
async function handleFile(file) {
    resetUi(false);
    setStatus(`Reading ${file.name}...`, "busy");
    try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const extracted = await extractOffice(bytes, file.name);
        currentFiles = extracted;
        renderResults(extracted);
        if (extracted.length === 0) {
            setStatus("No VBA project, form streams, or embedded FRX media were found.", "warn");
            return;
        }
        const media = extracted.filter((item)=>item.kind === "media").length;
        setStatus(`Extracted ${extracted.length} item${extracted.length === 1 ? "" : "s"}${media ? `, including ${media} embedded media file${media === 1 ? "" : "s"}` : ""}.`, "ready");
    } catch (error) {
        console.error(error);
        setStatus(error instanceof Error ? error.message : "Could not parse this file.", "error");
    }
}


// ---- extract.ts ----
async function extractOffice(bytes, fileName) {
    const raw = extractRawVbaOrFormFile(bytes, fileName);
    if (raw) return raw;
    if (isZip(bytes)) {
        const zipEntries = parseZip(bytes);
        const projectBins = zipEntries.filter((entry)=>/vbaProject\.bin$/i.test(entry.name));
        const extracted = [];
        for (const entry of projectBins){
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
function extractRawVbaOrFormFile(bytes, fileName) {
    const extension = fileName.split(".").at(-1)?.toLowerCase();
    if (!extension || ![
        "bas",
        "cls",
        "frm",
        "frx"
    ].includes(extension)) return undefined;
    if (extension === "frx") {
        const baseName = safeFileName(fileName.replace(/\.frx$/i, ""));
        const resource = {
            name: fileName,
            kind: "frx",
            bytes,
            analysis: analyzeBinaryStream(bytes, fileName),
            mimeType: "application/octet-stream",
            sourcePath: fileName
        };
        return [
            resource,
            ...extractMediaFromBinary(bytes, baseName).map((file)=>({
                    ...file,
                    sourcePath: `${fileName} / ${file.sourcePath}`
                }))
        ];
    }
    const text = decodeText(bytes).replace(/\0+$/g, "");
    return [
        {
            name: fileName,
            kind: extension === "frm" ? "frm" : "vba",
            bytes,
            text,
            mimeType: "text/plain;charset=windows-1252",
            sourcePath: fileName
        }
    ];
}
function extractVbaProject(bytes, sourcePath) {
    const cfb = parseCfb(bytes);
    const projectStream = cfb.streams.get("PROJECT") ?? cfb.streams.get("VBA/PROJECT");
    const dirStream = cfb.streams.get("VBA/dir") ?? cfb.streams.get("dir");
    const modules = parseProjectModules(projectStream);
    const dirModules = parseDirModules(cfb.streams.get("VBA/dir") ?? cfb.streams.get("dir"));
    const files = [];
    const processedSourceStreams = new Set();
    const usedMediaKeys = new Set();
    for (const dirModule of dirModules){
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
            sourcePath: `${sourcePath}/${path} @ ${toHex(dirModule.textOffset)}`
        }));
    }
    for (const [path, stream] of cfb.streams){
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
                sourcePath: `${sourcePath}/${path} (recovered by scan)`
            }));
        }
    }
    for (const [path, stream] of cfb.streams){
        const lower = path.toLowerCase();
        if (lower === "project" || lower.startsWith("vba/")) continue;
        const formName = path.split("/")[0] || "form";
        files.push({
            name: `${safeFileName(formName)}-${safeFileName(path.replaceAll("/", "-"))}.frx`,
            kind: "frx",
            bytes: stream,
            analysis: analyzeBinaryStream(stream, path),
            mimeType: "application/octet-stream",
            sourcePath: `${sourcePath}/${path}`
        });
        const mediaFiles = extractMediaFromBinary(stream, `${safeFileName(formName)}-${safeFileName(path.replaceAll("/", "-"))}`);
        for (const media of mediaFiles){
            const key = `${path}:${media.name}:${media.bytes.byteLength}`;
            if (usedMediaKeys.has(key)) continue;
            usedMediaKeys.add(key);
            files.push({
                ...media,
                sourcePath: `${sourcePath}/${path}`
            });
        }
    }
    if (files.length === 0) {
        for (const [path, stream] of cfb.streams){
            const source = extractBestVbaText(stream);
            if (!source) continue;
            const name = safeFileName(path.split("/").at(-1) ?? "module");
            files.push({
                name: `${name}.bas`,
                kind: "vba",
                bytes: encodeText(source),
                text: source,
                mimeType: "text/plain;charset=windows-1252",
                sourcePath: `${sourcePath}/${path}`
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
            sourcePath: `${sourcePath}/[vba project references]`
        });
    }
    return files.sort((a, b)=>kindRank(a.kind) - kindRank(b.kind) || a.name.localeCompare(b.name));
}
async function extractOfficePackageParts(bytes, zipEntries, fileName) {
    const interesting = zipEntries.filter((entry)=>/^word\/media\//i.test(entry.name) || /^ppt\/media\//i.test(entry.name) || /^xl\/media\//i.test(entry.name) || /\/activeX\/.+\.(xml|bin)$/i.test(entry.name) || /\/embeddings\/.+\.(bin|ole|xls|xlsx|doc|docx|ppt|pptx)$/i.test(entry.name));
    const extracted = [];
    const manifest = [];
    for (const entry of interesting){
        const entryBytes = await inflateZipEntry(bytes, entry);
        const kind = classifyOfficePart(entry.name);
        manifest.push({
            path: entry.name,
            kind,
            size: entryBytes.byteLength,
            note: describeOfficePart(entry.name, entryBytes)
        });
        extracted.push({
            name: `${safeFileName(entry.name.replaceAll("/", "-"))}`,
            kind: kind === "media" ? "media" : "binary",
            bytes: entryBytes,
            text: kind === "xml" ? decodeText(entryBytes) : undefined,
            analysis: kind === "binary" || kind === "activex" ? analyzeBinaryStream(entryBytes, entry.name) : undefined,
            mimeType: mimeTypeForOfficePart(entry.name),
            sourcePath: `${fileName}/${entry.name}`
        });
        if (kind === "binary" || kind === "activex") {
            const nestedMedia = extractMediaFromBinary(entryBytes, safeFileName(entry.name.replaceAll("/", "-")));
            for (const media of nestedMedia){
                extracted.push({
                    ...media,
                    sourcePath: `${fileName}/${entry.name} / ${media.sourcePath}`
                });
            }
        }
    }
    if (manifest.length) {
        extracted.push({
            name: "office-package-manifest.json",
            kind: "binary",
            bytes: encodeText(JSON.stringify({
                source: fileName,
                parts: manifest
            }, null, 2)),
            text: JSON.stringify({
                source: fileName,
                parts: manifest
            }, null, 2),
            mimeType: "application/json",
            sourcePath: `${fileName}/[office package manifest]`
        });
    }
    return extracted;
}
function classifyOfficePart(path) {
    if (/\/media\//i.test(path)) return "media";
    if (/\.xml$/i.test(path)) return "xml";
    if (/\/activeX\//i.test(path)) return "activex";
    return "binary";
}
function describeOfficePart(path, bytes) {
    if (/\/activeX\//i.test(path)) return "ActiveX control package part";
    if (/\/media\//i.test(path)) return "Office document media asset";
    if (isOle(bytes)) return "Embedded OLE compound file";
    return "Office embedded package part";
}
function mimeTypeForOfficePart(path) {
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
function parseProjectModules(projectStream) {
    const modules = new Map();
    if (!projectStream) return modules;
    const text = decodeText(projectStream);
    for (const rawLine of text.split(/\r?\n/)){
        const line = rawLine.trim();
        const [key, value] = line.split("=", 2);
        if (!key || !value) continue;
        if (key === "Module") modules.set(value.toLowerCase(), {
            name: value,
            type: "module"
        });
        if (key === "Class") modules.set(value.toLowerCase(), {
            name: value,
            type: "class"
        });
        if (key === "BaseClass") modules.set(value.toLowerCase(), {
            name: value,
            type: "form"
        });
        if (key === "Document") {
            const documentName = value.split("/")[0];
            modules.set(documentName.toLowerCase(), {
                name: documentName,
                type: "document"
            });
        }
    }
    return modules;
}
function parseProjectReferences(projectStream, dirStream) {
    const references = [
        ...parseProjectStreamReferences(projectStream),
        ...parseDirProjectReferences(dirStream)
    ];
    const seen = new Set();
    return references.filter((reference)=>{
        const key = `${reference.source}:${reference.kind}:${reference.name}:${reference.libId}:${reference.raw}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}
function parseProjectStreamReferences(projectStream) {
    if (!projectStream) return [];
    const references = [];
    const text = decodeText(projectStream);
    for (const rawLine of text.split(/\r?\n/)){
        const line = rawLine.trim();
        if (!line) continue;
        const [key, value = ""] = line.split("=", 2);
        if (![
            "Reference",
            "Object",
            "Package"
        ].includes(key)) continue;
        const guid = normalizeGuidString(value);
        const libParts = value.split("#").map((part)=>part.trim()).filter(Boolean);
        const libId = value.match(/\*?\\G\{[^}]+\}#[^#]+#[^#]+#?[^#]*/i)?.[0] ?? guid;
        references.push({
            source: "PROJECT",
            kind: key === "Reference" ? "registered" : key.toLowerCase(),
            name: inferReferenceName(value),
            libId,
            guid,
            version: libParts.find((part)=>/^\d+\.\d+$/.test(part)),
            path: libParts.find((part)=>/[\\/]|\.dll|\.ocx|\.tlb|\.olb/i.test(part)),
            raw: line
        });
    }
    return references;
}
function parseDirProjectReferences(dirStream) {
    if (!dirStream) return [];
    let dir;
    try {
        dir = decompressVba(dirStream);
    } catch  {
        return [];
    }
    const modulesStart = findProjectModulesRecord(dir);
    const referenceBytes = modulesStart > 0 ? dir.subarray(0, modulesStart) : dir;
    const references = [];
    let cursor = 0;
    let pendingName;
    let pendingOriginalLibId;
    while(cursor + 6 <= referenceBytes.length){
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
        } catch  {
            cursor += 1;
            continue;
        }
        if (record.id === 0x0016) {
            pendingName = parseReferenceNameRecord(record.payload);
        } else if (record.id === 0x0033) {
            pendingOriginalLibId = parseReferenceLibIdRecord(record.payload, "direct");
        } else if ([
            0x002f,
            0x0030,
            0x000d,
            0x000e
        ].includes(record.id)) {
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
                    raw: `record ${toHex(record.id, 4)} ${raw ?? ""}`.trim()
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
                raw: `record 0x0033 ${raw}`
            });
            pendingName = undefined;
            pendingOriginalLibId = undefined;
        }
        cursor = record.next;
    }
    return references;
}
function parseReferenceNameRecord(payload) {
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
function readReferenceNameRecord(bytes, cursor) {
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
        next
    };
}
function parseDirReferenceRecord(id, payload) {
    if (id === 0x000d) {
        const libId = parseReferenceLibIdRecord(payload, "length-prefixed");
        return {
            libId,
            ...parseLibId(libId)
        };
    }
    if (id === 0x000e) {
        const absolute = parseReferenceLibIdRecord(payload, "length-prefixed");
        let relative;
        if (absolute) {
            const relativeOffset = 4 + byteLength(absolute);
            if (relativeOffset + 4 <= payload.length) relative = parseReferenceLibIdRecord(payload.subarray(relativeOffset), "length-prefixed");
        }
        const libId = relative || absolute;
        return {
            libId,
            ...parseLibId(libId)
        };
    }
    const libId = parseReferenceLibIdRecord(payload, "length-prefixed") ?? parseReferenceLibIdRecord(payload, "direct");
    return {
        libId,
        ...parseLibId(libId)
    };
}
function parseReferenceLibIdRecord(payload, mode) {
    let raw = "";
    if (mode === "length-prefixed" && payload.length >= 4) {
        const size = readU32(payload, 0);
        if (size > 0 && size <= payload.length - 4) raw = decodeText(payload.subarray(4, 4 + size));
    }
    if (!raw && mode === "direct") raw = decodeText(payload);
    const cleaned = normalizeReferenceLibId(cleanReferenceText(raw));
    return cleaned || undefined;
}
function classifyDirReferenceRecord(id) {
    if (id === 0x002f || id === 0x0030) return "control";
    if (id === 0x000d) return "registered";
    if (id === 0x000e) return "project";
    return "unknown";
}
function inferReferenceName(value) {
    const parts = value.split("#").map((part)=>part.trim()).filter(Boolean);
    const candidate = parts.find((part)=>/^[A-Za-z_][A-Za-z0-9_. -]*$/.test(part) && !/^\d+\.\d+$/.test(part));
    return candidate;
}
function normalizeReferenceLibId(value) {
    const marker = value.search(/\*?\\G\{/i);
    const normalized = marker >= 0 ? value.slice(marker) : value;
    return normalized.replace(/[^\x20-\x7e]+$/g, "").trim();
}
function parseLibId(value) {
    if (!value) return {};
    const parts = value.split("#");
    return {
        version: parts.find((part)=>/^\d+\.\d+$/.test(part)),
        path: parts.find((part)=>/[A-Z]:\\|\/|\.dll|\.ocx|\.tlb|\.olb/i.test(part))
    };
}
function cleanReferenceText(value) {
    return value.replace(/\0/g, "").replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]+/g, "").trim();
}
function byteLength(value) {
    return encodeText(value).byteLength;
}
function parseDirModules(dirStream) {
    if (!dirStream) return [];
    let dir;
    try {
        dir = decompressVba(dirStream);
    } catch  {
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
        const modules = [];
        for(let index = 0; index < count && cursor < dir.length; index += 1){
            const parsed = readDirModule(dir, cursor);
            if (!parsed) break;
            modules.push(parsed.module);
            cursor = parsed.next;
        }
        return modules;
    } catch  {
        return [];
    }
}
function findProjectModulesRecord(dir) {
    for(let offset = 0; offset + 16 <= dir.length; offset += 1){
        if (readU16(dir, offset) !== 0x000f || readU32(dir, offset + 2) !== 0x00000002) continue;
        const count = readU16(dir, offset + 6);
        const nextRecord = offset + 8;
        if (count > 0 && count < 1000 && readU16(dir, nextRecord) === 0x0013 && readU32(dir, nextRecord + 2) === 0x00000002) {
            return offset;
        }
    }
    return -1;
}
function readDirModule(dir, start) {
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
    for (const expectedId of [
        0x0032,
        0x001c,
        0x0048
    ]){
        const record = readRecord(dir, cursor);
        if (record.id !== expectedId) return undefined;
        cursor = record.next;
    }
    const offsetRecord = readRecord(dir, cursor);
    if (offsetRecord.id !== 0x0031 || offsetRecord.payload.byteLength < 4) return undefined;
    const textOffset = readU32(offsetRecord.payload, 0);
    cursor = offsetRecord.next;
    for (const expectedId of [
        0x001e,
        0x002c
    ]){
        const record = readRecord(dir, cursor);
        if (record.id !== expectedId) return undefined;
        cursor = record.next;
    }
    let moduleTypeId;
    while(cursor + 2 <= dir.length){
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
        module: {
            name,
            streamName,
            textOffset,
            moduleTypeId
        },
        next: cursor
    };
}
function readRecord(bytes, offset) {
    if (offset + 6 > bytes.length) throw new Error("Record is out of bounds.");
    const id = readU16(bytes, offset);
    const size = readU32(bytes, offset + 2);
    const payloadStart = offset + 6;
    const next = payloadStart + size;
    if (next > bytes.length) throw new Error("Record payload is out of bounds.");
    return {
        id,
        size,
        payload: bytes.subarray(payloadStart, next),
        next
    };
}
function findVbaModuleStream(streams, streamName) {
    const wanted = `vba/${streamName}`.toLowerCase();
    for (const entry of streams){
        if (entry[0].toLowerCase() === wanted) return entry;
    }
    return undefined;
}
function createSourceFile(input) {
    const extension = input.moduleType === "form" ? "frm" : input.moduleType === "class" ? "cls" : "bas";
    return {
        name: `${input.moduleName}.${extension}`,
        kind: extension === "frm" ? "frm" : "vba",
        bytes: encodeText(input.source),
        text: input.source,
        mimeType: "text/plain;charset=windows-1252",
        sourcePath: input.sourcePath
    };
}
function extractVbaTextAtOffset(stream, offset) {
    if (offset < 0 || offset >= stream.length) return undefined;
    try {
        return decodeText(decompressVba(stream.subarray(offset))).replace(/\0+$/g, "");
    } catch  {
        return undefined;
    }
}
function extractBestVbaText(stream) {
    let best = "";
    let bestScore = 0;
    for(let offset = 0; offset < stream.byteLength; offset += 1){
        if (stream[offset] !== 0x01) continue;
        try {
            const text = decodeText(decompressVba(stream.subarray(offset)));
            const score = scoreVbaText(text);
            if (score > bestScore) {
                best = text;
                bestScore = score;
            }
        } catch  {}
    }
    return bestScore > 8 ? best.replace(/\0+$/g, "") : undefined;
}
function decompressVba(container) {
    if (container[0] !== 0x01) throw new Error("Invalid compressed VBA container signature.");
    const output = [];
    let cursor = 1;
    while(cursor + 2 <= container.length){
        const header = readU16(container, cursor);
        cursor += 2;
        const chunkSize = (header & 0x0fff) + 3;
        const chunkEnd = Math.min(container.length, cursor + chunkSize - 2);
        const signature = header >> 12 & 0x07;
        const compressed = (header & 0x8000) !== 0;
        if (signature !== 0x03) break;
        const chunkStart = output.length;
        if (!compressed) {
            while(cursor < chunkEnd)output.push(container[cursor++]);
            continue;
        }
        while(cursor < chunkEnd){
            const flags = container[cursor++];
            for(let bit = 0; bit < 8 && cursor < chunkEnd; bit += 1){
                if ((flags & 1 << bit) === 0) {
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
                const offset = (token >> 16 - bitCount) + 1;
                const copySource = output.length - offset;
                for(let i = 0; i < length; i += 1){
                    output.push(output[copySource + i]);
                }
            }
        }
        cursor = chunkEnd;
    }
    return new Uint8Array(output);
}
function scoreVbaText(text) {
    let score = 0;
    if (/Attribute\s+VB_Name/i.test(text)) score += 50;
    if (/\b(Sub|Function|Property|Option|Dim|Private|Public|End)\b/i.test(text)) score += 20;
    if (/Begin\s+VB\./i.test(text)) score += 20;
    const printable = [
        ...text.slice(0, 2000)
    ].filter((char)=>{
        const code = char.charCodeAt(0);
        return code === 9 || code === 10 || code === 13 || code >= 32 && code < 127;
    }).length;
    return score + printable / Math.max(text.length, 1);
}
function extractMediaFromBinary(bytes, baseName) {
    const found = [];
    const seen = new Set();
    for(let offset = 0; offset < bytes.length; offset += 1){
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
            sourcePath: match.label ? `${match.label} at ${toHex(offset)}` : ""
        });
        offset += Math.max(match.length - 1, 0);
    }
    for (const nested of extractNestedOleMedia(bytes, baseName, found.length)){
        const key = `${nested.sourcePath}:${nested.bytes.byteLength}:${nested.name}`;
        if (seen.has(key)) continue;
        seen.add(key);
        found.push(nested);
    }
    return found;
}
function extractNestedOleMedia(bytes, baseName, startIndex) {
    const found = [];
    for(let offset = 0; offset + 8 <= bytes.length; offset += 1){
        if (!isOle(bytes.subarray(offset, offset + 8))) continue;
        try {
            const nested = parseCfb(bytes.subarray(offset));
            for (const [streamPath, streamBytes] of nested.streams){
                const nestedMedia = extractMediaFromBinary(streamBytes, `${baseName}-${safeFileName(streamPath.replaceAll("/", "-"))}`);
                for (const media of nestedMedia){
                    found.push({
                        ...media,
                        name: `${baseName}-${String(startIndex + found.length + 1).padStart(2, "0")}-${safeFileName(streamPath.replaceAll("/", "-"))}.${media.name.split(".").at(-1) ?? "bin"}`,
                        sourcePath: `Nested OLE at ${toHex(offset)} / ${printableStreamName(streamPath)} / ${media.sourcePath}`
                    });
                }
            }
            offset += 511;
        } catch  {}
    }
    return found;
}
function detectMedia(bytes, offset) {
    const guidAndPicture = detectGuidAndPicture(bytes, offset);
    if (guidAndPicture) return guidAndPicture;
    const stdPicture = detectStdPicture(bytes, offset);
    if (stdPicture) return stdPicture;
    if (matches(bytes, offset, [
        0x89,
        0x50,
        0x4e,
        0x47,
        0x0d,
        0x0a,
        0x1a,
        0x0a
    ])) {
        const end = findBytes(bytes, [
            0x49,
            0x45,
            0x4e,
            0x44,
            0xae,
            0x42,
            0x60,
            0x82
        ], offset + 8);
        if (end > -1) return {
            extension: "png",
            mimeType: "image/png",
            length: end + 8 - offset,
            label: "PNG image"
        };
    }
    if (bytes[offset] === 0xff && bytes[offset + 1] === 0xd8) {
        const end = findBytes(bytes, [
            0xff,
            0xd9
        ], offset + 2);
        if (end > -1) return {
            extension: "jpg",
            mimeType: "image/jpeg",
            length: end + 2 - offset,
            label: "JPEG image"
        };
    }
    if (matchesAscii(bytes, offset, "GIF87a") || matchesAscii(bytes, offset, "GIF89a")) {
        const end = bytes.indexOf(0x3b, offset + 13);
        if (end > -1) return {
            extension: "gif",
            mimeType: "image/gif",
            length: end + 1 - offset,
            label: "GIF image"
        };
    }
    if (matchesAscii(bytes, offset, "BM") && offset + 14 < bytes.length) {
        const length = readU32(bytes, offset + 2);
        if (length > 14 && offset + length <= bytes.length) return {
            extension: "bmp",
            mimeType: "image/bmp",
            length,
            label: "BMP image"
        };
    }
    if (offset + 6 < bytes.length && readU16(bytes, offset) === 0 && [
        1,
        2
    ].includes(readU16(bytes, offset + 2))) {
        const count = readU16(bytes, offset + 4);
        if (count > 0 && count < 64 && offset + 6 + count * 16 <= bytes.length) {
            let length = 6 + count * 16;
            for(let i = 0; i < count; i += 1){
                const entry = offset + 6 + i * 16;
                const size = readU32(bytes, entry + 8);
                const imageOffset = readU32(bytes, entry + 12);
                length = Math.max(length, imageOffset + size);
            }
            if (length > 0 && offset + length <= bytes.length) {
                const isIcon = readU16(bytes, offset + 2) === 1;
                return {
                    extension: isIcon ? "ico" : "cur",
                    mimeType: "image/x-icon",
                    length,
                    label: isIcon ? "ICO icon" : "CUR cursor"
                };
            }
        }
    }
    if (matchesAscii(bytes, offset, "II*\0") || matchesAscii(bytes, offset, "MM\0*")) {
        const next = findNextLikelyMedia(bytes, offset + 8);
        const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
        if (length > 32) return {
            extension: "tif",
            mimeType: "image/tiff",
            length,
            label: "TIFF image"
        };
    }
    if (matches(bytes, offset, [
        0xd7,
        0xcd,
        0xc6,
        0x9a
    ]) && offset + 22 < bytes.length) {
        const next = findNextLikelyMedia(bytes, offset + 22);
        const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
        return {
            extension: "wmf",
            mimeType: "image/wmf",
            length,
            label: "Placeable WMF image"
        };
    }
    if (readU32(bytes, offset) === 1 && offset + 88 < bytes.length && matchesAscii(bytes, offset + 40, " EMF")) {
        const length = readU32(bytes, offset + 48);
        if (length > 88 && offset + length <= bytes.length) {
            return {
                extension: "emf",
                mimeType: "image/emf",
                length,
                label: "EMF image"
            };
        }
    }
    const dib = buildBmpFromDib(bytes, offset);
    if (dib) {
        return {
            extension: "bmp",
            mimeType: "image/bmp",
            length: dib.sourceLength,
            bytes: dib.bmp,
            label: "DIB image rebuilt as BMP"
        };
    }
    return undefined;
}
function detectGuidAndPicture(bytes, offset) {
    if (!matches(bytes, offset, CLSID_STD_PICTURE)) return undefined;
    const picture = detectStdPicture(bytes, offset + 16);
    if (!picture) return undefined;
    return {
        ...picture,
        length: 16 + picture.length,
        label: `MS-OFORMS GuidAndPicture ${picture.label ?? "picture"}`
    };
}
function detectStdPicture(bytes, offset) {
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
            label: "MS-OFORMS StdPicture payload"
        };
    }
    return {
        ...decoded,
        length: 8 + size,
        label: `MS-OFORMS StdPicture ${decoded.label ?? decoded.extension.toUpperCase()}`
    };
}
function decodePicturePayload(payload) {
    const direct = detectDirectImagePayload(payload);
    if (direct) return direct;
    const dib = buildBmpFromDib(payload, 0);
    if (dib) {
        return {
            extension: "bmp",
            mimeType: "image/bmp",
            bytes: dib.bmp,
            label: "DIB rebuilt as BMP"
        };
    }
    if (isStandardWmf(payload, 0)) {
        return {
            extension: "wmf",
            mimeType: "image/wmf",
            bytes: payload.slice(),
            label: "WMF image"
        };
    }
    if (isLikelyEmf(payload, 0)) {
        return {
            extension: "emf",
            mimeType: "image/emf",
            bytes: payload.slice(),
            label: "EMF image"
        };
    }
    return undefined;
}
function detectDirectImagePayload(payload) {
    const media = detectMediaWithoutOFormsWrappers(payload, 0);
    if (!media) return undefined;
    return {
        extension: media.extension,
        mimeType: media.mimeType,
        bytes: media.bytes ?? payload.slice(0, media.length),
        label: media.label
    };
}
function detectMediaWithoutOFormsWrappers(bytes, offset) {
    if (matches(bytes, offset, [
        0x89,
        0x50,
        0x4e,
        0x47,
        0x0d,
        0x0a,
        0x1a,
        0x0a
    ])) {
        const end = findBytes(bytes, [
            0x49,
            0x45,
            0x4e,
            0x44,
            0xae,
            0x42,
            0x60,
            0x82
        ], offset + 8);
        if (end > -1) return {
            extension: "png",
            mimeType: "image/png",
            length: end + 8 - offset,
            label: "PNG image"
        };
    }
    if (bytes[offset] === 0xff && bytes[offset + 1] === 0xd8) {
        const end = findBytes(bytes, [
            0xff,
            0xd9
        ], offset + 2);
        if (end > -1) return {
            extension: "jpg",
            mimeType: "image/jpeg",
            length: end + 2 - offset,
            label: "JPEG image"
        };
    }
    if (matchesAscii(bytes, offset, "GIF87a") || matchesAscii(bytes, offset, "GIF89a")) {
        const end = bytes.indexOf(0x3b, offset + 13);
        if (end > -1) return {
            extension: "gif",
            mimeType: "image/gif",
            length: end + 1 - offset,
            label: "GIF image"
        };
    }
    if (matchesAscii(bytes, offset, "BM") && offset + 14 < bytes.length) {
        const length = readU32(bytes, offset + 2);
        if (length > 14 && offset + length <= bytes.length) return {
            extension: "bmp",
            mimeType: "image/bmp",
            length,
            label: "BMP image"
        };
    }
    const icon = detectIconOrCursor(bytes, offset);
    if (icon) return icon;
    if (matchesAscii(bytes, offset, "II*\0") || matchesAscii(bytes, offset, "MM\0*")) {
        const next = findNextLikelyMedia(bytes, offset + 8);
        const length = Math.min((next > -1 ? next : offset + 2_000_000) - offset, bytes.length - offset);
        if (length > 32) return {
            extension: "tif",
            mimeType: "image/tiff",
            length,
            label: "TIFF image"
        };
    }
    if (matches(bytes, offset, [
        0xd7,
        0xcd,
        0xc6,
        0x9a
    ])) {
        const next = findNextLikelyMedia(bytes, offset + 22);
        const length = Math.min((next > -1 ? next : bytes.length) - offset, bytes.length - offset);
        return {
            extension: "wmf",
            mimeType: "image/wmf",
            length,
            label: "Placeable WMF image"
        };
    }
    if (isLikelyEmf(bytes, offset)) {
        const length = readU32(bytes, offset + 48);
        if (length > 88 && offset + length <= bytes.length) {
            return {
                extension: "emf",
                mimeType: "image/emf",
                length,
                label: "EMF image"
            };
        }
    }
    return undefined;
}
function findNextLikelyMedia(bytes, start) {
    for(let offset = start; offset < bytes.length; offset += 1){
        if (matches(bytes, offset, [
            0x89,
            0x50,
            0x4e,
            0x47
        ]) || bytes[offset] === 0xff && bytes[offset + 1] === 0xd8 || matchesAscii(bytes, offset, "GIF8") || matchesAscii(bytes, offset, "BM")) {
            return offset;
        }
    }
    return -1;
}
function detectIconOrCursor(bytes, offset) {
    if (offset + 6 >= bytes.length || readU16(bytes, offset) !== 0 || ![
        1,
        2
    ].includes(readU16(bytes, offset + 2))) {
        return undefined;
    }
    const count = readU16(bytes, offset + 4);
    if (count <= 0 || count >= 64 || offset + 6 + count * 16 > bytes.length) return undefined;
    let length = 6 + count * 16;
    for(let i = 0; i < count; i += 1){
        const entry = offset + 6 + i * 16;
        const size = readU32(bytes, entry + 8);
        const imageOffset = readU32(bytes, entry + 12);
        if (size === 0 || imageOffset < 6 + count * 16) return undefined;
        length = Math.max(length, imageOffset + size);
    }
    if (offset + length > bytes.length) return undefined;
    const isIcon = readU16(bytes, offset + 2) === 1;
    return {
        extension: isIcon ? "ico" : "cur",
        mimeType: "image/x-icon",
        length,
        label: isIcon ? "ICO icon" : "CUR cursor"
    };
}
function isStandardWmf(bytes, offset) {
    if (offset + 18 > bytes.length) return false;
    const fileType = readU16(bytes, offset);
    const headerSizeWords = readU16(bytes, offset + 2);
    const windowsVersion = readU16(bytes, offset + 4);
    const fileSizeWords = readU32(bytes, offset + 6);
    return [
        1,
        2
    ].includes(fileType) && headerSizeWords === 9 && windowsVersion >= 0x0100 && fileSizeWords > 9 && offset + fileSizeWords * 2 <= bytes.length;
}
function isLikelyEmf(bytes, offset) {
    return readU32(bytes, offset) === 1 && offset + 88 < bytes.length && matchesAscii(bytes, offset + 40, " EMF");
}
function analyzeBinaryStream(bytes, path) {
    const signatures = collectSignatures(bytes);
    const strings = collectStrings(bytes);
    const guids = collectGuids(bytes);
    const oforms = parseOFormsStream(bytes, path, strings);
    const entropy = calculateEntropy(bytes);
    const zeroBytes = bytes.reduce((count, byte)=>count + (byte === 0 ? 1 : 0), 0);
    const pathParts = path.split("/");
    const streamName = pathParts.at(-1) ?? path;
    const storagePath = pathParts.slice(0, -1).join("/") || "(root)";
    return {
        title: describeFormStream(path),
        summary: [
            {
                label: "Storage",
                value: storagePath
            },
            {
                label: "Stream",
                value: printableStreamName(streamName)
            },
            {
                label: "Size",
                value: formatBytes(bytes.byteLength)
            },
            {
                label: "Entropy",
                value: `${entropy.toFixed(2)} bits/byte`
            },
            {
                label: "Zero bytes",
                value: `${zeroBytes} (${bytes.length ? (zeroBytes / bytes.length * 100).toFixed(1) : "0.0"}%)`
            },
            {
                label: "Media candidates",
                value: String(signatures.filter((item)=>/image|bitmap|icon|cursor/i.test(item.label)).length)
            },
            {
                label: "Text strings",
                value: String(strings.length)
            },
            {
                label: "GUID candidates",
                value: String(guids.length)
            },
            {
                label: "MS-OFORMS",
                value: oforms ? `${oforms.kind} (${oforms.confidence})` : "No structured match"
            }
        ],
        stringCount: strings.length,
        guidCount: guids.length,
        signatures: signatures.slice(0, 24),
        oforms
    };
}
function describeFormStream(path) {
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
function printableStreamName(value) {
    return value.replace(/\x01/g, "0x01 ").replace(/\x03/g, "0x03 ");
}
function collectSignatures(bytes) {
    const signatures = [];
    for(let offset = 0; offset < bytes.length; offset += 1){
        if (readU32(bytes, offset) === 0x0000746c && offset + 8 <= bytes.length) {
            const pictureSize = readU32(bytes, offset + 4);
            if (pictureSize > 0 && offset + 8 + pictureSize <= bytes.length) {
                signatures.push({
                    offset,
                    label: "MS-OFORMS StdPicture",
                    detail: `${formatBytes(pictureSize)} picture payload`
                });
                offset += 7;
                continue;
            }
        }
        const media = detectMedia(bytes, offset);
        if (media) {
            signatures.push({
                offset,
                label: media.label ?? `${media.extension.toUpperCase()} media`,
                detail: `${formatBytes(media.bytes?.byteLength ?? media.length)}${media.bytes ? " extracted/rebuilt" : ""}`
            });
            offset += Math.max(media.length - 1, 0);
            continue;
        }
        if (isOle(bytes.subarray(offset, offset + 8))) {
            signatures.push({
                offset,
                label: "Nested OLE compound file",
                detail: "D0 CF 11 E0 A1 B1 1A E1"
            });
        }
    }
    return signatures;
}
function collectStrings(bytes) {
    const strings = [];
    const ascii = /[\x20-\x7e]{4,}/g;
    const asciiText = Array.from(bytes, (byte)=>byte >= 0x20 && byte <= 0x7e ? String.fromCharCode(byte) : "\n").join("");
    let match;
    while(match = ascii.exec(asciiText)){
        strings.push({
            encoding: "ASCII",
            offset: match.index,
            value: match[0]
        });
    }
    let start = -1;
    let chars = "";
    for(let offset = 0; offset + 1 < bytes.length; offset += 2){
        const code = readU16(bytes, offset);
        if (code >= 0x20 && code <= 0x7e) {
            if (start < 0) start = offset;
            chars += String.fromCharCode(code);
            continue;
        }
        if (chars.length >= 4) strings.push({
            encoding: "UTF-16LE",
            offset: start,
            value: chars
        });
        start = -1;
        chars = "";
    }
    if (chars.length >= 4) strings.push({
        encoding: "UTF-16LE",
        offset: start,
        value: chars
    });
    return strings.filter((item, index, all)=>all.findIndex((candidate)=>candidate.encoding === item.encoding && candidate.offset === item.offset && candidate.value === item.value) === index).sort((a, b)=>a.offset - b.offset);
}
function collectGuids(bytes) {
    const guids = [];
    for(let offset = 0; offset + 16 <= bytes.length; offset += 1){
        const chunk = bytes.subarray(offset, offset + 16);
        const nonZero = chunk.some((byte)=>byte !== 0);
        const nonFF = chunk.some((byte)=>byte !== 0xff);
        if (!nonZero || !nonFF) continue;
        const variant = chunk[8] & 0xc0;
        const version = chunk[7] >> 4 & 0x0f;
        if (variant !== 0x80 || version > 5) continue;
        guids.push({
            offset,
            value: formatGuid(chunk)
        });
        offset += 15;
    }
    return guids;
}
function labelGuid(guid) {
    return KNOWN_GUID_LABELS[guid.toUpperCase()];
}
function labelProgId(value) {
    const normalized = value.trim();
    return KNOWN_PROGID_LABELS[normalized.toLowerCase()];
}
function labelPossibleIdentifier(value) {
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
function parseOFormsStream(bytes, path, strings) {
    const streamName = path.split("/").at(-1) ?? path;
    if (streamName === "f") {
        const record = parseOFormsControlRecord(bytes, 0, "FormControl", "form", strings);
        if (!record) return undefined;
        return {
            kind: "Form stream",
            confidence: record.size <= bytes.length ? "high" : "medium",
            records: [
                record
            ],
            notes: [
                "Parsed as an MS-OFORMS FormControl header. Property values with variable sizing are decoded only when they can be identified confidently.",
                "Embedded site/control records can also be represented in this stream after the top-level FormControl data."
            ]
        };
    }
    if (streamName === "o") {
        const records = scanOFormsObjectRecords(bytes, strings);
        if (records.length === 0) return undefined;
        return {
            kind: "Object stream",
            confidence: records.length > 1 || records[0].offset === 0 ? "medium" : "low",
            records,
            notes: [
                "Scanned for MS-OFORMS control records inside the object stream. Offsets are best-effort when the stream contains mixed payloads."
            ]
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
            records: [
                {
                    offset: 0,
                    type: "VBFrame stream",
                    size: bytes.byteLength,
                    properties: [
                        {
                            name: "Size",
                            value: formatBytes(bytes.byteLength)
                        },
                        ...strings.slice(0, 8).map((item)=>({
                                name: `${item.encoding} string @ ${toHex(item.offset)}`,
                                value: item.value
                            }))
                    ]
                }
            ],
            notes: [
                "VBFrame streams are Office Forms metadata used by frame-like containers."
            ]
        };
    }
    return undefined;
}
function parseOFormsControlRecord(bytes, offset, type, maskKind, strings) {
    if (offset + 8 > bytes.length) return undefined;
    const minor = bytes[offset];
    const major = bytes[offset + 1];
    const cb = readU16(bytes, offset + 2);
    const propMask = readU32(bytes, offset + 4);
    const end = offset + 4 + cb;
    if (minor !== 0 || ![
        0x02,
        0x04
    ].includes(major) || cb < 4 || cb > bytes.length - offset) return undefined;
    const properties = [
        {
            name: "MinorVersion",
            value: toHex(minor, 2)
        },
        {
            name: "MajorVersion",
            value: toHex(major, 2)
        },
        {
            name: "cb",
            value: `${cb} bytes`
        },
        {
            name: "PropMask",
            value: `${toHex(propMask, 8)} (${decodeMaskFlags(propMask, maskKind).join(", ") || "no known flags"})`
        }
    ];
    properties.push(...decodeOFormsLayoutProperties(bytes, offset, end, propMask, maskKind));
    properties.push(...decodeCommonOFormsValues(bytes, offset, end, propMask, maskKind, strings));
    return {
        offset,
        type,
        size: Math.min(end - offset, bytes.length - offset),
        properties
    };
}
function scanOFormsObjectRecords(bytes, strings) {
    const records = [];
    for(let offset = 0; offset + 8 <= bytes.length; offset += 1){
        const record = parseOFormsControlRecord(bytes, offset, offset === 0 ? "OleSiteConcrete control" : "Nested OleSite/control record", "site", strings);
        if (!record) continue;
        records.push(record);
        offset += Math.max(record.size - 1, 0);
        if (records.length >= 32) break;
    }
    return records;
}
function parseOFormsExtendedStream(bytes) {
    if (bytes.byteLength < 8) return undefined;
    const records = [];
    for(let offset = 0; offset + 8 <= bytes.length; offset += 8){
        records.push({
            offset,
            type: "PageProperties / extended data candidate",
            size: Math.min(8, bytes.length - offset),
            properties: [
                {
                    name: "Value1",
                    value: toHex(readU32(bytes, offset), 8)
                },
                {
                    name: "Value2",
                    value: toHex(readU32(bytes, offset + 4), 8)
                }
            ]
        });
    }
    return {
        kind: "Extended stream",
        confidence: "low",
        records,
        notes: [
            "The x stream is used by some container controls such as MultiPage. Values are shown as page/extended-data candidates."
        ]
    };
}
function parseCompObjStream(bytes) {
    const strings = collectStrings(bytes);
    return {
        kind: "COM compound object",
        confidence: strings.some((item)=>/Microsoft Forms/i.test(item.value)) ? "high" : "medium",
        records: [
            {
                offset: 0,
                type: "CompObj",
                size: bytes.byteLength,
                properties: strings.map((item)=>({
                        name: `${item.encoding} string @ ${toHex(item.offset)}`,
                        value: labelPossibleIdentifier(item.value)
                    }))
            }
        ],
        notes: [
            "CompObj identifies embedded COM/OLE object metadata for the form or control."
        ]
    };
}
function decodeMaskFlags(mask, kind) {
    const flags = kind === "form" ? FORM_PROP_FLAGS : kind === "site" ? SITE_PROP_FLAGS : GENERIC_CONTROL_FLAGS;
    return flags.filter((flag)=>(mask & 1 << flag.bit) !== 0).map((flag)=>flag.name);
}
function decodeCommonOFormsValues(bytes, offset, end, propMask, maskKind, strings) {
    const properties = [];
    const recordStrings = strings.filter((item)=>item.offset >= offset + 8 && item.offset < end);
    const streamDataStrings = strings.filter((item)=>item.offset >= end && item.offset < Math.min(bytes.length, end + 128));
    properties.push(...decodeOrderedOFormsStringProperties(propMask, maskKind, recordStrings, streamDataStrings));
    for (const item of recordStrings.slice(0, 12)){
        properties.push({
            name: `Record string @ ${toHex(item.offset)}`,
            value: labelPossibleIdentifier(item.value)
        });
    }
    for (const item of streamDataStrings.slice(0, 8)){
        properties.push({
            name: `StreamData string @ ${toHex(item.offset)}`,
            value: labelPossibleIdentifier(item.value)
        });
    }
    const fontGuidOffset = findBytes(bytes, [
        0x03,
        0x52,
        0xe3,
        0x0b,
        0x91,
        0x8f,
        0xce,
        0x11,
        0x9d,
        0xe3,
        0x00,
        0xaa,
        0x00,
        0x4b,
        0xb8,
        0x51
    ], end);
    if (fontGuidOffset > -1) {
        properties.push({
            name: "Font",
            value: `StdFont GUID at ${toHex(fontGuidOffset)} (MSForms StdFont)`
        });
    }
    const pictureOffset = findStdPicture(bytes, end);
    if (pictureOffset > -1) {
        properties.push({
            name: "Picture",
            value: `StdPicture at ${toHex(pictureOffset)} (${formatBytes(readU32(bytes, pictureOffset + 4))})`
        });
    }
    if (maskKind === "form") {
        const extraStart = guessExtraDataStart(bytes, offset, end, propMask, FORM_PROP_FLAGS);
        if (extraStart > -1) {
            properties.push({
                name: "ExtraDataBlock candidate",
                value: `${toHex(extraStart)}-${toHex(end)}`
            });
        }
    }
    return properties;
}
function decodeOFormsLayoutProperties(bytes, offset, end, propMask, maskKind) {
    const properties = [];
    if (maskKind === "form") {
        const sizes = findFmSizeCandidates(bytes, offset + 8, end);
        const selected = selectFormSizeCandidate(sizes);
        if ((propMask & 1 << 10) !== 0 && selected) {
            properties.push({
                name: "Decoded DisplayedSize",
                value: formatSizeHint(selected)
            });
        }
        if ((propMask & 1 << 11) !== 0 && sizes.length > 1) {
            const logical = sizes.find((candidate)=>candidate !== selected && candidate.width <= 40_000 && candidate.height <= 40_000);
            if (logical) properties.push({
                name: "Decoded LogicalSize",
                value: formatSizeHint(logical)
            });
        }
    }
    if (maskKind === "site" || maskKind === "generic") {
        const positions = findFmPositionCandidates(bytes, offset + 8, Math.min(bytes.length, end + 96));
        const selected = selectSitePositionCandidate(positions);
        if ((propMask & 1 << 2) !== 0 && selected) {
            properties.push({
                name: "Decoded SitePosition",
                value: formatPositionHint(selected)
            });
        }
    }
    return properties;
}
function decodeOrderedOFormsStringProperties(propMask, maskKind, recordStrings, streamDataStrings) {
    const properties = [];
    const candidates = [
        ...recordStrings,
        ...streamDataStrings
    ].map((item)=>({
            ...item,
            clean: item.value.replace(/\s+\(.+\)$/, "")
        })).filter((item)=>isUsefulDecodedPropertyString(item.clean));
    if (maskKind === "form") {
        if ((propMask & 1 << 19) !== 0) {
            const caption = candidates.find((item)=>isHumanLabel(item.clean));
            if (caption) properties.push({
                name: "Decoded Caption",
                value: `${labelPossibleIdentifier(caption.value)} (medium)`
            });
        }
        const extraLabels = candidates.filter((item)=>isHumanLabel(item.clean)).slice(1, 8).map((item, index)=>({
                name: `Decoded Child/Page Label ${index + 1}`,
                value: `${labelPossibleIdentifier(item.value)} (low)`
            }));
        properties.push(...extraLabels);
    }
    if (maskKind === "site" || maskKind === "generic") {
        const ordered = [
            {
                bit: 0,
                name: "Decoded Name",
                isString: true
            },
            {
                bit: 1,
                name: "Decoded Tag",
                isString: true
            },
            {
                bit: 2,
                name: "Decoded SitePosition",
                isString: false
            },
            {
                bit: 3,
                name: "Decoded ControlTipText",
                isString: true
            },
            {
                bit: 4,
                name: "Decoded RuntimeLicKey",
                isString: true
            },
            {
                bit: 5,
                name: "Decoded ControlSource",
                isString: true
            },
            {
                bit: 6,
                name: "Decoded RowSource",
                isString: true
            }
        ];
        let stringIndex = 0;
        for (const property of ordered){
            if ((propMask & 1 << property.bit) === 0 || !property.isString) continue;
            const candidate = candidates[stringIndex];
            stringIndex += 1;
            if (!candidate) continue;
            properties.push({
                name: property.name,
                value: `${labelPossibleIdentifier(candidate.value)} (low)`
            });
        }
    }
    return properties;
}
function isUsefulDecodedPropertyString(value) {
    if (!value || value.length > 160) return false;
    if (/^Tahoma$/i.test(value)) return false;
    if (/^[-=]{5,}$/.test(value)) return false;
    if (/^'/.test(value) || /https?:\/\//i.test(value)) return false;
    if (/copyright|permission is hereby|software/i.test(value)) return false;
    return true;
}
function findFmSizeCandidates(bytes, start, end) {
    const candidates = [];
    for(let offset = start; offset + 8 <= end; offset += 4){
        const width = readU32(bytes, offset);
        const height = readU32(bytes, offset + 4);
        if (width < 100 || height < 100 || width > 32_000 || height > 32_000) continue;
        if (width === 32_000 || height === 32_000) continue;
        if (width === 0xffff || height === 0xffff) continue;
        const confidence = width < 25_000 && height < 25_000 ? "medium" : "low";
        candidates.push({
            offset,
            width,
            height,
            confidence
        });
    }
    return candidates;
}
function selectFormSizeCandidate(candidates) {
    return candidates.filter((candidate)=>candidate.confidence !== "low").sort((a, b)=>b.width * b.height - a.width * a.height)[0] ?? candidates.sort((a, b)=>b.width * b.height - a.width * a.height)[0];
}
function findFmPositionCandidates(bytes, start, end) {
    const candidates = [];
    for(let offset = start; offset + 8 <= end; offset += 4){
        const left = readU32(bytes, offset);
        const top = readU32(bytes, offset + 4);
        if (left > 25_000 || top > 25_000) continue;
        if (left === 0xffff || top === 0xffff) continue;
        const confidence = left < 15_000 && top < 15_000 ? "medium" : "low";
        candidates.push({
            offset,
            left,
            top,
            confidence
        });
    }
    return candidates;
}
function selectSitePositionCandidate(candidates) {
    return candidates.filter((candidate)=>candidate.confidence !== "low").sort((a, b)=>scorePositionCandidate(b) - scorePositionCandidate(a))[0] ?? candidates[0];
}
function scorePositionCandidate(candidate) {
    let score = candidate.confidence === "medium" ? 3 : 0;
    if (candidate.left > 0 || candidate.top > 0) score += 1;
    if (candidate.left < 10_000 && candidate.top < 10_000) score += 1;
    return score;
}
function formatSizeHint(size) {
    return `${size.width} x ${size.height} twips @ ${toHex(size.offset)} (${size.confidence})`;
}
function formatPositionHint(position) {
    return `${position.left}, ${position.top} twips @ ${toHex(position.offset)} (${position.confidence})`;
}
function guessExtraDataStart(bytes, offset, end, propMask, flags) {
    const firstRecordString = collectStrings(bytes.subarray(offset, end)).map((item)=>item.offset + offset).filter((stringOffset)=>stringOffset >= offset + 8).sort((a, b)=>a - b)[0];
    if (firstRecordString !== undefined) return firstRecordString;
    const hasExtra = flags.some((flag)=>[
            "DisplayedSize",
            "LogicalSize",
            "ScrollPosition",
            "Caption"
        ].includes(flag.name) && (propMask & 1 << flag.bit) !== 0);
    return hasExtra ? Math.max(offset + 8, end - 24) : -1;
}
function findStdPicture(bytes, start) {
    for(let offset = start; offset + 8 <= bytes.length; offset += 1){
        if (readU32(bytes, offset) !== 0x0000746c) continue;
        const size = readU32(bytes, offset + 4);
        if (size > 0 && offset + 8 + size <= bytes.length) return offset;
    }
    return -1;
}
const FORM_PROP_FLAGS = [
    {
        bit: 1,
        name: "BackColor"
    },
    {
        bit: 2,
        name: "ForeColor"
    },
    {
        bit: 3,
        name: "NextAvailableID"
    },
    {
        bit: 6,
        name: "BooleanProperties"
    },
    {
        bit: 7,
        name: "BorderStyle"
    },
    {
        bit: 8,
        name: "MousePointer"
    },
    {
        bit: 9,
        name: "ScrollBars"
    },
    {
        bit: 10,
        name: "DisplayedSize"
    },
    {
        bit: 11,
        name: "LogicalSize"
    },
    {
        bit: 12,
        name: "ScrollPosition"
    },
    {
        bit: 13,
        name: "GroupCount"
    },
    {
        bit: 15,
        name: "MouseIcon"
    },
    {
        bit: 16,
        name: "Cycle"
    },
    {
        bit: 17,
        name: "SpecialEffect"
    },
    {
        bit: 18,
        name: "BorderColor"
    },
    {
        bit: 19,
        name: "Caption"
    },
    {
        bit: 20,
        name: "Font"
    },
    {
        bit: 21,
        name: "Picture"
    },
    {
        bit: 22,
        name: "Zoom"
    },
    {
        bit: 23,
        name: "PictureAlignment"
    },
    {
        bit: 24,
        name: "PictureTiling"
    },
    {
        bit: 25,
        name: "PictureSizeMode"
    },
    {
        bit: 26,
        name: "ShapeCookie"
    },
    {
        bit: 27,
        name: "DrawBuffer"
    }
];
const SITE_PROP_FLAGS = [
    {
        bit: 0,
        name: "Name"
    },
    {
        bit: 1,
        name: "Tag"
    },
    {
        bit: 2,
        name: "Position"
    },
    {
        bit: 3,
        name: "ControlTipText"
    },
    {
        bit: 4,
        name: "RuntimeLicKey"
    },
    {
        bit: 5,
        name: "ControlSource"
    },
    {
        bit: 6,
        name: "RowSource"
    },
    {
        bit: 7,
        name: "Enabled/visibility flags"
    }
];
const GENERIC_CONTROL_FLAGS = [
    {
        bit: 0,
        name: "Property0"
    },
    {
        bit: 1,
        name: "Property1"
    },
    {
        bit: 2,
        name: "Property2"
    },
    {
        bit: 3,
        name: "Property3"
    },
    {
        bit: 4,
        name: "Property4"
    },
    {
        bit: 5,
        name: "Property5"
    },
    {
        bit: 6,
        name: "Property6"
    },
    {
        bit: 7,
        name: "Property7"
    },
    {
        bit: 8,
        name: "Property8"
    },
    {
        bit: 9,
        name: "Property9"
    },
    {
        bit: 10,
        name: "Size/position"
    },
    {
        bit: 11,
        name: "Caption/name string"
    },
    {
        bit: 12,
        name: "Font"
    },
    {
        bit: 13,
        name: "Picture/icon"
    },
    {
        bit: 14,
        name: "Mouse/icon data"
    },
    {
        bit: 15,
        name: "Extra data"
    },
    {
        bit: 16,
        name: "Property16"
    },
    {
        bit: 17,
        name: "Property17"
    },
    {
        bit: 18,
        name: "Property18"
    },
    {
        bit: 19,
        name: "Property19"
    },
    {
        bit: 20,
        name: "Property20"
    },
    {
        bit: 21,
        name: "Property21"
    },
    {
        bit: 22,
        name: "Property22"
    },
    {
        bit: 23,
        name: "Property23"
    },
    {
        bit: 24,
        name: "Property24"
    },
    {
        bit: 25,
        name: "Property25"
    },
    {
        bit: 26,
        name: "Property26"
    },
    {
        bit: 27,
        name: "Property27"
    },
    {
        bit: 28,
        name: "Property28"
    },
    {
        bit: 29,
        name: "Property29"
    },
    {
        bit: 30,
        name: "Property30"
    },
    {
        bit: 31,
        name: "Property31"
    }
];
const KNOWN_GUID_LABELS = {
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
    "{8BD21DD0-EC42-11CE-9E0D-00AA006002F3}": "Microsoft Forms 2.0 Label"
};
const COMMON_FORMS_CONTROLS = [
    [
        "Forms.Form.1",
        "Microsoft Forms 2.0 UserForm"
    ],
    [
        "Forms.Frame.1",
        "Microsoft Forms 2.0 Frame control"
    ],
    [
        "Forms.MultiPage.1",
        "Microsoft Forms 2.0 MultiPage control"
    ],
    [
        "Forms.Page.1",
        "Microsoft Forms 2.0 Page control"
    ],
    [
        "Forms.TabStrip.1",
        "Microsoft Forms 2.0 TabStrip control"
    ],
    [
        "Forms.Label.1",
        "Microsoft Forms 2.0 Label control"
    ],
    [
        "Forms.TextBox.1",
        "Microsoft Forms 2.0 TextBox control"
    ],
    [
        "Forms.ComboBox.1",
        "Microsoft Forms 2.0 ComboBox control"
    ],
    [
        "Forms.ListBox.1",
        "Microsoft Forms 2.0 ListBox control"
    ],
    [
        "Forms.CheckBox.1",
        "Microsoft Forms 2.0 CheckBox control"
    ],
    [
        "Forms.OptionButton.1",
        "Microsoft Forms 2.0 OptionButton control"
    ],
    [
        "Forms.ToggleButton.1",
        "Microsoft Forms 2.0 ToggleButton control"
    ],
    [
        "Forms.CommandButton.1",
        "Microsoft Forms 2.0 CommandButton control"
    ],
    [
        "Forms.Image.1",
        "Microsoft Forms 2.0 Image control"
    ],
    [
        "Forms.ScrollBar.1",
        "Microsoft Forms 2.0 ScrollBar control"
    ],
    [
        "Forms.SpinButton.1",
        "Microsoft Forms 2.0 SpinButton control"
    ],
    [
        "Forms.HTMLCheckBox.1",
        "Microsoft Forms 2.0 HTML CheckBox control"
    ],
    [
        "Forms.HTMLHidden.1",
        "Microsoft Forms 2.0 HTML Hidden control"
    ],
    [
        "Forms.HTMLImage.1",
        "Microsoft Forms 2.0 HTML Image control"
    ],
    [
        "Forms.HTMLOption.1",
        "Microsoft Forms 2.0 HTML Option control"
    ],
    [
        "Forms.HTMLPassword.1",
        "Microsoft Forms 2.0 HTML Password control"
    ],
    [
        "Forms.HTMLReset.1",
        "Microsoft Forms 2.0 HTML Reset control"
    ],
    [
        "Forms.HTMLSelect.1",
        "Microsoft Forms 2.0 HTML Select control"
    ],
    [
        "Forms.HTMLSubmit.1",
        "Microsoft Forms 2.0 HTML Submit control"
    ],
    [
        "Forms.HTMLText.1",
        "Microsoft Forms 2.0 HTML Text control"
    ],
    [
        "Forms.HTMLTextArea.1",
        "Microsoft Forms 2.0 HTML TextArea control"
    ]
];
const KNOWN_PROGID_LABELS = Object.fromEntries(COMMON_FORMS_CONTROLS.map(([progId, label])=>[
        progId.toLowerCase(),
        label
    ]));
function buildBmpFromDib(bytes, offset) {
    if (offset + 40 > bytes.length) return undefined;
    const headerSize = readU32(bytes, offset);
    if (![
        40,
        52,
        56,
        108,
        124
    ].includes(headerSize)) return undefined;
    const width = readI32(bytes, offset + 4);
    const height = readI32(bytes, offset + 8);
    const planes = readU16(bytes, offset + 12);
    const bitCount = readU16(bytes, offset + 14);
    const compression = readU32(bytes, offset + 16);
    const declaredImageSize = readU32(bytes, offset + 20);
    const colorsUsed = readU32(bytes, offset + 32);
    if (width <= 0 || Math.abs(height) <= 0 || width > 20000 || Math.abs(height) > 20000) return undefined;
    if (planes !== 1 || ![
        1,
        4,
        8,
        16,
        24,
        32
    ].includes(bitCount) || compression > 6) return undefined;
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
    return {
        bmp,
        sourceLength
    };
}
function hexDump(bytes, start, length) {
    const lines = [];
    for(let offset = start; offset < start + length; offset += 16){
        const chunk = bytes.subarray(offset, Math.min(offset + 16, start + length));
        const hex = [
            ...chunk
        ].map((byte)=>byte.toString(16).padStart(2, "0")).join(" ").padEnd(47, " ");
        const ascii = [
            ...chunk
        ].map((byte)=>byte >= 0x20 && byte <= 0x7e ? String.fromCharCode(byte) : ".").join("");
        lines.push(`${toHex(offset, 8)}  ${hex}  ${ascii}`);
    }
    return lines.join("\n");
}
function calculateEntropy(bytes) {
    if (bytes.length === 0) return 0;
    const counts = new Array(256).fill(0);
    for (const byte of bytes)counts[byte] += 1;
    return counts.reduce((entropy, count)=>{
        if (count === 0) return entropy;
        const probability = count / bytes.length;
        return entropy - probability * Math.log2(probability);
    }, 0);
}
function parseCfb(bytes) {
    if (!isOle(bytes)) throw new Error("Invalid OLE compound document.");
    const sectorSize = 1 << readU16(bytes, 30);
    const miniSectorSize = 1 << readU16(bytes, 32);
    const firstDirectorySector = readI32(bytes, 48);
    const miniStreamCutoff = readU32(bytes, 56);
    const firstMiniFatSector = readI32(bytes, 60);
    const miniFatSectorCount = readU32(bytes, 64);
    const firstDifatSector = readI32(bytes, 68);
    const difatSectorCount = readU32(bytes, 72);
    const difat = [];
    for(let offset = 76; offset < 512; offset += 4){
        const sector = readI32(bytes, offset);
        if (sector >= 0) difat.push(sector);
    }
    let difatSector = firstDifatSector;
    for(let i = 0; i < difatSectorCount && difatSector >= 0; i += 1){
        const sector = getSector(bytes, sectorSize, difatSector);
        for(let offset = 0; offset < sectorSize - 4; offset += 4){
            const fatSector = readI32(sector, offset);
            if (fatSector >= 0) difat.push(fatSector);
        }
        difatSector = readI32(sector, sectorSize - 4);
    }
    const fat = [];
    for (const sectorIndex of difat){
        const sector = getSector(bytes, sectorSize, sectorIndex);
        for(let offset = 0; offset < sectorSize; offset += 4){
            fat.push(readI32(sector, offset));
        }
    }
    const readChain = (startSector)=>{
        const chunks = [];
        const visited = new Set();
        let sector = startSector;
        while(sector >= 0 && sector !== END_OF_CHAIN && !visited.has(sector)){
            visited.add(sector);
            chunks.push(getSector(bytes, sectorSize, sector));
            const next = fat[sector];
            if (next === undefined || next === FREE_SECTOR || next === FAT_SECTOR || next === DIFAT_SECTOR) break;
            sector = next;
        }
        return concatBytes(chunks);
    };
    const directoryBytes = readChain(firstDirectorySector);
    const entries = [];
    for(let offset = 0, index = 0; offset + 128 <= directoryBytes.length; offset += 128, index += 1){
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
            path: name
        };
    }
    const root = entries.find((entry)=>entry.type === 5);
    const miniStream = root ? readChain(root.startSector).subarray(0, root.size) : new Uint8Array();
    const miniFatBytes = firstMiniFatSector >= 0 && miniFatSectorCount ? readChain(firstMiniFatSector) : new Uint8Array();
    const miniFat = [];
    for(let offset = 0; offset + 4 <= miniFatBytes.length; offset += 4)miniFat.push(readI32(miniFatBytes, offset));
    const readMiniChain = (startSector, size)=>{
        const chunks = [];
        const visited = new Set();
        let sector = startSector;
        while(sector >= 0 && sector !== END_OF_CHAIN && !visited.has(sector)){
            visited.add(sector);
            const offset = sector * miniSectorSize;
            chunks.push(miniStream.subarray(offset, offset + miniSectorSize));
            const next = miniFat[sector];
            if (next === undefined || next === FREE_SECTOR) break;
            sector = next;
        }
        return concatBytes(chunks).subarray(0, size);
    };
    const readStream = (entry)=>{
        if (entry.size === 0 || entry.startSector < 0) return new Uint8Array();
        if (entry.size < miniStreamCutoff && root && entry.index !== root.index) {
            return readMiniChain(entry.startSector, entry.size);
        }
        return readChain(entry.startSector).subarray(0, entry.size);
    };
    const streams = new Map();
    const walk = (index, parent)=>{
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
    return {
        entries,
        streams
    };
}
function parseZip(bytes) {
    const eocdOffset = findEndOfCentralDirectory(bytes);
    if (eocdOffset < 0) throw new Error("Could not find the Office zip directory.");
    const entriesCount = readU16(bytes, eocdOffset + 10);
    const centralDirectoryOffset = readU32(bytes, eocdOffset + 16);
    const entries = [];
    let cursor = centralDirectoryOffset;
    for(let i = 0; i < entriesCount; i += 1){
        if (readU32(bytes, cursor) !== 0x02014b50) throw new Error("Invalid zip central directory.");
        const method = readU16(bytes, cursor + 10);
        const compressedSize = readU32(bytes, cursor + 20);
        const uncompressedSize = readU32(bytes, cursor + 24);
        const nameLength = readU16(bytes, cursor + 28);
        const extraLength = readU16(bytes, cursor + 30);
        const commentLength = readU16(bytes, cursor + 32);
        const localHeaderOffset = readU32(bytes, cursor + 42);
        const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
        entries.push({
            name,
            method,
            compressedSize,
            uncompressedSize,
            localHeaderOffset
        });
        cursor += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
}
async function inflateZipEntry(bytes, entry) {
    const cursor = entry.localHeaderOffset;
    if (readU32(bytes, cursor) !== 0x04034b50) throw new Error(`Invalid local zip header for ${entry.name}.`);
    const nameLength = readU16(bytes, cursor + 26);
    const extraLength = readU16(bytes, cursor + 28);
    const dataStart = cursor + 30 + nameLength + extraLength;
    const compressed = bytes.subarray(dataStart, dataStart + entry.compressedSize);
    if (entry.method === 0) return compressed;
    if (entry.method !== 8) throw new Error(`${entry.name} uses unsupported zip compression method ${entry.method}.`);
    if (!("DecompressionStream" in window)) throw new Error("This browser does not support in-browser zip decompression.");
    const stream = new Blob([
        compressed
    ]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    const inflated = new Uint8Array(await new Response(stream).arrayBuffer());
    if (entry.uncompressedSize && inflated.byteLength !== entry.uncompressedSize) {
        console.warn(`Unexpected uncompressed size for ${entry.name}.`);
    }
    return inflated;
}


// ---- results-ui.ts ----
function renderResults(files) {
    results.innerHTML = "";
    summaryPanel.classList.toggle("hidden", files.length === 0);
    moduleCount.textContent = String(files.filter((file)=>file.kind === "vba").length);
    formCount.textContent = String(files.filter((file)=>file.kind === "frm").length);
    frxCount.textContent = String(files.filter((file)=>file.kind === "frx").length);
    for (const group of buildResultGroups(files)){
        results.append(renderResultGroup(group));
    }
}
function buildResultGroups(files) {
    const groups = new Map();
    for (const file of files){
        const owner = getResultOwner(file);
        let group = groups.get(owner);
        if (!group) {
            group = {
                name: owner,
                code: [],
                resources: [],
                media: [],
                other: []
            };
            groups.set(owner, group);
        }
        if (file.kind === "vba" || file.kind === "frm") group.code.push(file);
        else if (file.kind === "frx") group.resources.push(file);
        else if (file.kind === "media") group.media.push(file);
        else group.other.push(file);
    }
    for (const group of groups.values()){
        group.designer = buildDesignerSummary(group);
    }
    return [
        ...groups.values()
    ].sort((a, b)=>{
        const aHasForm = a.code.some((file)=>file.kind === "frm") || a.resources.length > 0;
        const bHasForm = b.code.some((file)=>file.kind === "frm") || b.resources.length > 0;
        if (aHasForm !== bHasForm) return aHasForm ? -1 : 1;
        return a.name.localeCompare(b.name);
    });
}
function getResultOwner(file) {
    const streamPath = extractInternalStreamPath(file.sourcePath);
    if (streamPath && !streamPath.toLowerCase().startsWith("vba/")) {
        const owner = streamPath.split("/")[0];
        return /^project/i.test(owner) ? "Project metadata" : printableStreamName(owner);
    }
    if (/\/(word|xl|ppt)\/(media|activeX|embeddings)\//i.test(file.sourcePath)) {
        const match = file.sourcePath.match(/\/(word|xl|ppt)\/(media|activeX|embeddings)\//i);
        return `Office ${match?.[2] ?? "package"} parts`;
    }
    if (file.name === "office-package-manifest.json") return "Office package parts";
    return file.name.replace(/\.(bas|cls|frm|frx|png|jpg|jpeg|gif|bmp|tif|ico|cur|wmf|emf)$/i, "");
}
function extractInternalStreamPath(sourcePath) {
    const marker = "vbaProject.bin/";
    const markerIndex = sourcePath.indexOf(marker);
    if (markerIndex < 0) return undefined;
    return sourcePath.slice(markerIndex + marker.length).replace(/\s+@\s+0x[0-9a-f]+.*$/i, "").replace(/\s+\(recovered by scan\)$/i, "");
}
function renderResultGroup(group) {
    const section = document.createElement("section");
    section.className = "result-group";
    const header = document.createElement("header");
    header.className = "group-header";
    const titleWrap = document.createElement("div");
    const title = document.createElement("h2");
    title.textContent = group.name;
    const subtitle = document.createElement("p");
    subtitle.textContent = describeGroup(group);
    titleWrap.append(title, subtitle);
    header.append(titleWrap);
    section.append(header);
    if (group.designer) {
        section.append(renderDesignerSummary(group.designer));
    }
    appendGroupSection(section, "Code", group.code);
    appendGroupSection(section, "Form resources", group.resources);
    appendGroupSection(section, "Recovered media", group.media);
    appendGroupSection(section, "Other", group.other);
    return section;
}
function describeGroup(group) {
    const parts = [
        countLabel(group.code.length, "code file"),
        countLabel(group.resources.length, "resource stream"),
        countLabel(group.media.length, "media file"),
        countLabel(group.other.length, "other file")
    ].filter(Boolean);
    return parts.join(" · ") || "No extracted files";
}
function countLabel(count, label) {
    if (!count) return "";
    return `${count} ${label}${count === 1 ? "" : "s"}`;
}
function appendGroupSection(parent, title, files) {
    if (files.length === 0) return;
    const section = document.createElement("section");
    section.className = "group-subsection";
    const heading = document.createElement("h3");
    heading.textContent = title;
    section.append(heading);
    const list = document.createElement("div");
    list.className = "group-items";
    for (const file of files.sort((a, b)=>kindRank(a.kind) - kindRank(b.kind) || getDisplayPath(a).localeCompare(getDisplayPath(b)))){
        list.append(renderFileCard(file));
    }
    section.append(list);
    parent.append(section);
}
function renderFileCard(file) {
    const card = document.createElement("article");
    card.className = "result-card";
    const header = document.createElement("header");
    header.className = "result-header";
    header.innerHTML = `
    <div>
      <h4>${escapeHtml(file.name)}</h4>
      <p>${escapeHtml(getDisplayPath(file))} · ${formatBytes(file.bytes.byteLength)}</p>
    </div>
  `;
    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.textContent = "Download";
    downloadButton.addEventListener("click", ()=>downloadFile(file));
    header.append(downloadButton);
    card.append(header);
    if (file.text) {
        const pre = document.createElement("pre");
        pre.textContent = file.text;
        card.append(pre);
    } else if (file.analysis) {
        card.append(renderAnalysis(file.analysis, file.bytes));
    } else if (canPreviewImage(file.mimeType)) {
        const url = URL.createObjectURL(new Blob([
            file.bytes
        ], {
            type: file.mimeType
        }));
        const image = document.createElement("img");
        image.src = url;
        image.alt = file.name;
        image.className = "preview-image";
        image.addEventListener("load", ()=>URL.revokeObjectURL(url), {
            once: true
        });
        card.append(image);
    } else {
        const binaryNote = document.createElement("p");
        binaryNote.className = "binary-note";
        binaryNote.textContent = "Binary stream extracted. Download it to inspect or recover additional form resources.";
        card.append(binaryNote);
    }
    return card;
}
function canPreviewImage(mimeType) {
    return [
        "image/png",
        "image/jpeg",
        "image/gif",
        "image/bmp",
        "image/x-icon"
    ].includes(mimeType);
}


// ---- zip-export.ts ----
function createZip(files) {
    const groups = buildResultGroups(files);
    const appModel = buildApplicationModel(files, groups);
    const appSummary = renderApplicationSummary(appModel);
    const rebuildBrief = renderLlmRebuildBrief(appModel);
    const migrationChecklist = renderMigrationChecklist(appModel);
    const callGraph = renderCallGraph(appModel);
    const dependencyReport = renderDependencyReport(appModel);
    const projectReferencesReport = renderProjectReferencesReport(appModel);
    const frontendPlan = renderFrontendImplementationPlan(appModel);
    const testPlan = renderMigrationTestPlan(appModel);
    const layoutModel = buildLayoutModel(appModel);
    const traceabilityMap = buildTraceabilityMap(files, appModel);
    const activeXControls = appModel.documentControls;
    const activeXPersistence = appModel.documentControls.filter((control)=>control.persistenceAnalysis).map((control)=>({
            id: control.id,
            xmlPath: control.xmlPath,
            binPath: control.binPath,
            analysis: control.persistenceAnalysis
        }));
    const visualPreviewArtifacts = createVisualPreviewArtifacts(groups);
    const validationReport = renderValidationReport(files, appModel);
    const appArtifacts = [
        {
            path: "application-model.json",
            bytes: encodeText(JSON.stringify(appModel, null, 2))
        },
        {
            path: "application-summary.md",
            bytes: encodeText(appSummary)
        },
        {
            path: "llm-rebuild-brief.md",
            bytes: encodeText(rebuildBrief)
        },
        {
            path: "migration-checklist.md",
            bytes: encodeText(migrationChecklist)
        },
        {
            path: "call-graph.md",
            bytes: encodeText(callGraph)
        },
        {
            path: "dependency-report.md",
            bytes: encodeText(dependencyReport)
        },
        {
            path: "vba-project-references.md",
            bytes: encodeText(projectReferencesReport)
        },
        {
            path: "frontend-implementation-plan.md",
            bytes: encodeText(frontendPlan)
        },
        {
            path: "migration-test-plan.md",
            bytes: encodeText(testPlan)
        },
        {
            path: "layout-model.json",
            bytes: encodeText(JSON.stringify(layoutModel, null, 2))
        },
        {
            path: "traceability-map.json",
            bytes: encodeText(JSON.stringify(traceabilityMap, null, 2))
        },
        {
            path: "activex-controls.json",
            bytes: encodeText(JSON.stringify(activeXControls, null, 2))
        },
        {
            path: "activex-persistence.json",
            bytes: encodeText(JSON.stringify(activeXPersistence, null, 2))
        },
        {
            path: "validation-report.md",
            bytes: encodeText(validationReport)
        }
    ];
    const summaryFiles = groups.filter((group)=>group.designer).map((group)=>({
            owner: group.name,
            bytes: encodeText(JSON.stringify(group.designer, null, 2))
        }));
    const extractedEntries = files.map((file, index)=>({
            path: getZipPath(file, index),
            bytes: file.bytes,
            crc: crc32(file.bytes)
        }));
    const summaryEntries = summaryFiles.map((summary)=>{
        const path = `${safeFileName(summary.owner)}/designer-summary/${safeFileName(summary.owner)}.designer.json`;
        return {
            path,
            bytes: summary.bytes,
            crc: crc32(summary.bytes)
        };
    });
    const appEntries = appArtifacts.map((artifact)=>({
            ...artifact,
            crc: crc32(artifact.bytes)
        }));
    const visualPreviewEntries = visualPreviewArtifacts.map((artifact)=>({
            ...artifact,
            crc: crc32(artifact.bytes)
        }));
    const procedureEntries = createProcedureZipEntries(files, appModel);
    const entries = [
        ...appEntries,
        ...visualPreviewEntries,
        ...procedureEntries,
        ...extractedEntries,
        ...summaryEntries
    ];
    const localParts = [];
    const centralParts = [];
    let offset = 0;
    for (const entry of entries){
        const name = encodeText(entry.path);
        const timeDate = getDosDateTime(new Date());
        const local = new Uint8Array(30 + name.length);
        writeU32(local, 0, 0x04034b50);
        writeU16(local, 4, 20);
        writeU16(local, 6, 0x0800);
        writeU16(local, 8, 0);
        writeU16(local, 10, timeDate.time);
        writeU16(local, 12, timeDate.date);
        writeU32(local, 14, entry.crc);
        writeU32(local, 18, entry.bytes.length);
        writeU32(local, 22, entry.bytes.length);
        writeU16(local, 26, name.length);
        local.set(name, 30);
        localParts.push(local, entry.bytes);
        const central = new Uint8Array(46 + name.length);
        writeU32(central, 0, 0x02014b50);
        writeU16(central, 4, 0x0314);
        writeU16(central, 6, 20);
        writeU16(central, 8, 0x0800);
        writeU16(central, 10, 0);
        writeU16(central, 12, timeDate.time);
        writeU16(central, 14, timeDate.date);
        writeU32(central, 16, entry.crc);
        writeU32(central, 20, entry.bytes.length);
        writeU32(central, 24, entry.bytes.length);
        writeU16(central, 28, name.length);
        writeU32(central, 42, offset);
        central.set(name, 46);
        centralParts.push(central);
        offset += local.length + entry.bytes.length;
    }
    const centralDirectoryOffset = offset;
    const centralDirectorySize = centralParts.reduce((sum, part)=>sum + part.length, 0);
    const end = new Uint8Array(22);
    writeU32(end, 0, 0x06054b50);
    writeU16(end, 8, entries.length);
    writeU16(end, 10, entries.length);
    writeU32(end, 12, centralDirectorySize);
    writeU32(end, 16, centralDirectoryOffset);
    return concatBytes([
        ...localParts,
        ...centralParts,
        end
    ]);
}
function getZipPath(file, index) {
    const owner = safeFileName(getResultOwner(file));
    const category = file.kind === "media" ? "media" : file.kind === "frx" ? "resources" : file.kind === "frm" || file.kind === "vba" ? "code" : /office-package-manifest|\/(word|xl|ppt)\//i.test(file.sourcePath) ? "office-package" : "other";
    return `${owner}/${category}/${String(index + 1).padStart(3, "0")}-${safeFileName(file.name)}`;
}
function createProcedureZipEntries(files, model) {
    const entries = [];
    for (const module of model.modules){
        const source = files.find((file)=>file.name === module.fileName)?.text;
        if (!source) continue;
        const lines = source.split(/\r?\n/);
        module.procedures.forEach((procedure, index)=>{
            const end = procedure.lineEnd ?? procedure.lineStart;
            const body = lines.slice(procedure.lineStart - 1, end).join("\n");
            const metadata = [
                `' Module: ${module.name}`,
                `' Procedure: ${procedure.name}`,
                `' Kind: ${procedure.kind}`,
                `' Scope: ${procedure.scope}`,
                `' Lines: ${procedure.lineStart}-${end}`,
                `' Calls: ${procedure.calls.join(", ") || "none"}`,
                `' Uses: ${procedure.uses.join(", ") || "none"}`,
                ""
            ].join("\n");
            const bytes = encodeText(`${metadata}${body}\n`);
            entries.push({
                path: `procedure-chunks/${safeFileName(module.name)}/${String(index + 1).padStart(3, "0")}-${safeFileName(procedure.name)}.vba`,
                bytes,
                crc: crc32(bytes)
            });
        });
    }
    return entries;
}
function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes)crc = crc >>> 8 ^ CRC32_TABLE[(crc ^ byte) & 0xff];
    return (crc ^ 0xffffffff) >>> 0;
}
function makeCrc32Table() {
    const table = [];
    for(let i = 0; i < 256; i += 1){
        let value = i;
        for(let bit = 0; bit < 8; bit += 1)value = value & 1 ? 0xedb88320 ^ value >>> 1 : value >>> 1;
        table.push(value >>> 0);
    }
    return table;
}
const CRC32_TABLE = makeCrc32Table();
function getDosDateTime(date) {
    const year = Math.max(1980, date.getFullYear());
    return {
        time: date.getHours() << 11 | date.getMinutes() << 5 | Math.floor(date.getSeconds() / 2),
        date: year - 1980 << 9 | date.getMonth() + 1 << 5 | date.getDate()
    };
}


// ---- model.ts ----
function buildApplicationModel(files, groups = buildResultGroups(files)) {
    const modules = files.filter((file)=>file.text && (file.kind === "vba" || file.kind === "frm")).map((file)=>buildModuleModel(file));
    const projectReferences = parseExtractedProjectReferences(files);
    const forms = groups.filter((group)=>group.designer).map((group)=>buildFormModel(group.designer));
    const documentControls = buildActiveXControls(files);
    enrichProcedureCalls(modules, files);
    linkEventsToControls(modules, forms);
    addInferredEventControls(modules, forms);
    linkEventsToControls(modules, forms);
    enrichModuleReferences(modules, files);
    return {
        generatedAt: new Date().toISOString(),
        sourceFiles: files.map((file)=>({
                name: file.name,
                kind: file.kind,
                path: getZipPath(file, files.indexOf(file)),
                size: file.bytes.byteLength
            })),
        modules,
        projectReferences,
        forms,
        documentControls,
        dependencies: uniqueDependencies([
            ...detectDependencies(modules, files),
            ...projectReferences.map((reference)=>({
                    category: "VBA reference",
                    value: reference.name ?? reference.libId ?? reference.raw,
                    source: reference.source,
                    reason: `${reference.kind} project reference`
                }))
        ]),
        assets: files.filter((file)=>file.kind === "media").map((file)=>({
                name: file.name,
                mimeType: file.mimeType,
                size: file.bytes.byteLength,
                sourcePath: file.sourcePath
            })),
        migrationNotes: [
            "VBA code is extracted verbatim. Business logic should be reviewed before translation because Office object model calls often imply workbook state.",
            "Form/control layout is reconstructed from MS-OFORMS and VBFrame data where available; raw FRX streams remain in the archive for audit.",
            "Event links are inferred from VBA naming conventions such as ControlName_Click.",
            "Dependencies are heuristic detections intended to help an LLM identify integrations, not a security verdict."
        ]
    };
}
function parseExtractedProjectReferences(files) {
    const referenceFile = files.find((file)=>file.name === "vba-project-references.json" && file.text);
    if (!referenceFile?.text) return [];
    try {
        return JSON.parse(referenceFile.text);
    } catch  {
        return [];
    }
}
function buildModuleModel(file) {
    const source = file.text ?? "";
    const procedures = parseProcedures(source);
    const declarations = source.split(/\r?\n/).filter((line)=>/^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\b/i.test(line)).map((line)=>line.trim());
    const variables = parseVariableDeclarations(source);
    const constants = parseConstantDeclarations(source);
    const kind = file.name.endsWith(".frm") ? "form" : file.name.endsWith(".cls") ? "class" : /ThisWorkbook|Sheet\d+/i.test(file.name) ? "document" : "standard";
    return {
        name: file.name.replace(/\.(bas|cls|frm)$/i, ""),
        fileName: file.name,
        kind,
        sourcePath: file.sourcePath,
        procedures,
        declarations,
        variables,
        constants,
        events: procedures.map(procedureToEvent).filter((event)=>Boolean(event)),
        references: [],
        riskMarkers: detectRiskMarkers(source)
    };
}
function parseProcedures(source) {
    const lines = source.split(/\r?\n/);
    const procedures = [];
    const startRegex = /^\s*(?:(Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\([^)]*\))?\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
    const endRegex = /^\s*End\s+(Sub|Function|Property)\b/i;
    let current;
    for(let index = 0; index < lines.length; index += 1){
        const line = lines[index];
        const start = line.match(startRegex);
        if (start) {
            current = {
                name: start[4],
                kind: start[3].startsWith("Property") ? "Property" : start[3],
                scope: start[1] ?? "Implicit",
                signature: line.trim(),
                parameters: parseParameters(start[5] ?? ""),
                returnType: start[6],
                lineStart: index + 1,
                calls: [],
                uses: []
            };
            procedures.push(current);
            continue;
        }
        if (current && endRegex.test(line)) {
            current.lineEnd = index + 1;
            current = undefined;
        }
    }
    return procedures;
}
function parseParameters(parameterList) {
    return parameterList.replace(/^\(|\)$/g, "").split(",").map((parameter)=>parameter.trim()).filter(Boolean);
}
function parseVariableDeclarations(source) {
    const declarations = [];
    const lines = source.split(/\r?\n/);
    const regex = /^\s*(Public|Private|Friend|Dim|Static|Global)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
    for(let index = 0; index < lines.length; index += 1){
        const line = lines[index];
        if (/\b(Sub|Function|Property)\b/i.test(line)) continue;
        const match = line.match(regex);
        if (!match || /^Declare$/i.test(match[2])) continue;
        declarations.push({
            name: match[2],
            scope: match[1],
            type: match[3],
            line: index + 1,
            statement: line.trim()
        });
    }
    return declarations;
}
function parseConstantDeclarations(source) {
    const declarations = [];
    const lines = source.split(/\r?\n/);
    const regex = /^\s*(Public|Private|Friend)?\s*Const\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
    for(let index = 0; index < lines.length; index += 1){
        const match = lines[index].match(regex);
        if (!match) continue;
        declarations.push({
            name: match[2],
            scope: match[1] ?? "Implicit",
            type: match[3],
            line: index + 1,
            statement: lines[index].trim()
        });
    }
    return declarations;
}
function procedureToEvent(procedure) {
    const match = procedure.name.match(/^(.+)_([A-Za-z][A-Za-z0-9]*)$/);
    if (!match) return undefined;
    return {
        procedure: procedure.name,
        controlName: match[1],
        eventName: match[2]
    };
}
function enrichProcedureCalls(modules, files) {
    const procedureNames = new Set(modules.flatMap((module)=>module.procedures.map((procedure)=>procedure.name)));
    for (const module of modules){
        const source = files.find((file)=>file.name === module.fileName)?.text ?? "";
        const lines = source.split(/\r?\n/);
        for (const procedure of module.procedures){
            const body = lines.slice(procedure.lineStart - 1, procedure.lineEnd ?? procedure.lineStart + 80).join("\n");
            procedure.calls = [
                ...procedureNames
            ].filter((name)=>name !== procedure.name && new RegExp(`\\b${escapeRegExp(name)}\\b`, "i").test(body)).sort((a, b)=>a.localeCompare(b));
            procedure.uses = detectProcedureUses(body);
        }
    }
}
function enrichModuleReferences(modules, files) {
    for (const module of modules){
        const source = files.find((file)=>file.name === module.fileName)?.text ?? "";
        module.references = modules.filter((candidate)=>candidate.name !== module.name).filter((candidate)=>new RegExp(`\\b${escapeRegExp(candidate.name)}\\b`, "i").test(source)).map((candidate)=>candidate.name).sort((a, b)=>a.localeCompare(b));
    }
}
function detectRiskMarkers(source) {
    const markers = [];
    const checks = [
        {
            category: "TODO",
            reason: "Developer note may indicate incomplete or special-case logic.",
            regex: /\b(TODO|FIXME|HACK|XXX)\b/i
        },
        {
            category: "Error handling",
            reason: "Broad error handling can hide behavior that needs explicit frontend states.",
            regex: /\bOn\s+Error\s+Resume\s+Next\b/i
        },
        {
            category: "Global state",
            reason: "Global/public mutable state should become explicit app state.",
            regex: /^\s*(Public|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s+/i
        },
        {
            category: "Dynamic execution",
            reason: "Dynamic execution or indirection may require manual migration review.",
            regex: /\b(Application\.Run|CallByName|Evaluate\s*\()/i
        },
        {
            category: "External process",
            reason: "External process execution cannot be directly migrated to browser frontend code.",
            regex: /\b(Shell|WScript\.Shell|CreateObject\s*\(\s*"WScript\.Shell")\b/i
        },
        {
            category: "File system",
            reason: "Local file access needs replacement with browser file APIs or backend services.",
            regex: /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Kill\s+|MkDir\s+|RmDir\s+)\b/i
        },
        {
            category: "WinAPI",
            reason: "WinAPI calls do not run in browser frontend code.",
            regex: /\bDeclare\s+(PtrSafe\s+)?(Sub|Function)\b/i
        }
    ];
    source.split(/\r?\n/).forEach((line, index)=>{
        for (const check of checks){
            if (check.regex.test(line)) {
                markers.push({
                    category: check.category,
                    line: index + 1,
                    text: line.trim(),
                    reason: check.reason
                });
            }
        }
    });
    return markers;
}
function detectProcedureUses(body) {
    const uses = new Set();
    const checks = [
        [
            "Excel object model",
            /\b(Application|Workbook|Worksheet|Range|Cells|Rows|Columns|Sheets|ActiveCell|Selection)\b/i
        ],
        [
            "UserForm/UI state",
            /\b(Me\.|Controls\(|\.Caption|\.Value|\.Visible|\.Enabled|\.ListIndex|\.AddItem|\.Clear)\b/i
        ],
        [
            "File system",
            /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Dir\s*\(|Kill\s+|MkDir\s+|RmDir\s+)\b/i
        ],
        [
            "HTTP/network",
            /\b(XMLHTTP|WinHttp|ServerXMLHTTP|WebRequest|URLDownloadToFile)\b/i
        ],
        [
            "Database/ADO",
            /\b(ADODB|DAO\.|Recordset|Connection|OpenRecordset)\b/i
        ],
        [
            "Shell/process",
            /\b(Shell|WScript\.Shell|Run\s*\()\b/i
        ],
        [
            "WinAPI",
            /\b(Declare\s+(PtrSafe\s+)?(Sub|Function)|AddressOf|LongPtr|LongLong)\b/i
        ],
        [
            "Error handling",
            /\b(On\s+Error|Err\.)\b/i
        ]
    ];
    for (const [label, regex] of checks){
        if (regex.test(body)) uses.add(label);
    }
    return [
        ...uses
    ].sort((a, b)=>a.localeCompare(b));
}
function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function buildFormModel(summary) {
    return {
        name: summary.formName,
        properties: summary.frame,
        controls: flattenControls(summary.controls)
    };
}
function buildActiveXControls(files) {
    const activeXXmlFiles = files.filter((file)=>/\/activeX\/activeX\d+\.xml$/i.test(file.sourcePath) && file.text);
    const activeXBinFiles = files.filter((file)=>/\/activeX\/activeX\d+\.bin$/i.test(file.sourcePath));
    return activeXXmlFiles.map((xmlFile)=>{
        const base = xmlFile.sourcePath.match(/activeX(\d+)\.xml$/i)?.[1] ?? String(activeXXmlFiles.indexOf(xmlFile) + 1);
        const binFile = activeXBinFiles.find((file)=>new RegExp(`activeX${base}\\.bin$`, "i").test(file.sourcePath));
        const properties = parseActiveXXmlProperties(xmlFile.text ?? "");
        const classId = properties.classid ?? properties.classId ?? properties.clsid;
        const normalizedClassId = normalizeGuidString(classId);
        const persistenceAnalysis = parseActiveXBinPersistence(binFile);
        const label = persistenceAnalysis?.compObj?.userType ?? (normalizedClassId ? labelGuid(normalizedClassId) : undefined);
        const persistenceProperties = Object.fromEntries((persistenceAnalysis?.properties ?? []).map((property)=>[
                `bin.${property.name}`,
                property.value
            ]));
        return {
            id: `activeX${base}`,
            xmlPath: xmlFile.sourcePath,
            binPath: binFile?.sourcePath,
            classId: normalizedClassId ?? classId,
            label,
            persistence: properties.persistence ?? properties.persistStorage,
            properties: {
                ...properties,
                ...persistenceProperties
            },
            persistenceAnalysis,
            sourceFiles: [
                xmlFile.sourcePath,
                binFile?.sourcePath
            ].filter((value)=>Boolean(value))
        };
    });
}
function parseActiveXBinPersistence(binFile) {
    if (!binFile) return undefined;
    const bytes = binFile.bytes;
    const media = extractMediaFromBinary(bytes, safeFileName(binFile.name)).map((file)=>({
            name: file.name,
            mimeType: file.mimeType,
            size: file.bytes.byteLength,
            sourcePath: file.sourcePath
        }));
    if (bytes.byteLength === 0) {
        return {
            format: "empty",
            confidence: "high",
            size: 0,
            streams: [],
            properties: [],
            media,
            warnings: [
                "ActiveX binary part is empty."
            ]
        };
    }
    if (!isOle(bytes)) {
        return {
            format: "raw-stream",
            confidence: "medium",
            size: bytes.byteLength,
            streams: [
                summarizeActiveXPersistenceStream("(raw)", bytes)
            ],
            properties: inferActiveXContentsProperties(bytes, "(raw)", undefined),
            media,
            warnings: [
                "Binary persistence part is not an OLE/CFB storage; parsed as a raw persistStream payload."
            ]
        };
    }
    try {
        const cfb = parseCfb(bytes);
        const streams = [
            ...cfb.streams
        ].map(([path, streamBytes])=>summarizeActiveXPersistenceStream(path, streamBytes));
        const compObjEntry = [
            ...cfb.streams
        ].find(([path])=>/compobj/i.test(path));
        const contentsEntry = [
            ...cfb.streams
        ].find(([path])=>/(^|\/)contents$/i.test(printableStreamName(path)));
        const compObj = compObjEntry ? parseActiveXCompObjStream(compObjEntry[1]) : undefined;
        const properties = [
            ...compObj ? activeXCompObjProperties(compObj, compObjEntry?.[0] ?? "CompObj") : [],
            ...contentsEntry ? inferActiveXContentsProperties(contentsEntry[1], contentsEntry[0], compObj?.userType) : []
        ];
        return {
            format: "cfb-storage",
            confidence: compObj || contentsEntry ? "high" : "medium",
            size: bytes.byteLength,
            streams,
            compObj,
            properties,
            media,
            warnings: contentsEntry ? [] : [
                "No contents stream was found in the ActiveX storage."
            ]
        };
    } catch (error) {
        return {
            format: "unknown",
            confidence: "low",
            size: bytes.byteLength,
            streams: [
                summarizeActiveXPersistenceStream("(unparsed)", bytes)
            ],
            properties: [],
            media,
            warnings: [
                `Could not parse ActiveX OLE storage: ${error instanceof Error ? error.message : String(error)}`
            ]
        };
    }
}
function summarizeActiveXPersistenceStream(path, bytes) {
    const strings = collectStrings(bytes).slice(0, 16).map((item)=>labelPossibleIdentifier(item.value));
    const guids = collectGuids(bytes).slice(0, 16).map((item)=>{
        const label = labelGuid(item.value);
        return label ? `${item.value} (${label})` : item.value;
    });
    const signatures = collectSignatures(bytes).slice(0, 12).map((signature)=>`${signature.label} at ${toHex(signature.offset)}${signature.detail ? ` (${signature.detail})` : ""}`);
    const oforms = parseOFormsStream(bytes, path, collectStrings(bytes));
    return {
        path: printableStreamName(path),
        size: bytes.byteLength,
        strings,
        guids,
        signatures,
        oforms: oforms ? `${oforms.kind} (${oforms.confidence})` : undefined
    };
}
function parseActiveXCompObjStream(bytes) {
    const strings = collectStrings(bytes).map((item)=>item.value).filter(Boolean);
    const clsid = bytes.byteLength >= 28 ? formatGuid(bytes.subarray(12, 28)) : undefined;
    return {
        clsid,
        userType: strings.find((value)=>/Microsoft Forms|Control|Button|Box|Label|Image|Object/i.test(value)),
        clipboardFormat: strings.find((value)=>/Embedded Object|Object|Control/i.test(value) && !/Microsoft Forms/i.test(value)),
        progId: strings.find((value)=>/^[A-Za-z][A-Za-z0-9_.]+(\.\d+)?$/.test(value))
    };
}
function activeXCompObjProperties(compObj, source) {
    const properties = [];
    if (compObj.clsid) {
        properties.push({
            name: "CompObj CLSID",
            value: labelPossibleIdentifier(compObj.clsid),
            source,
            confidence: "high"
        });
    }
    if (compObj.userType) properties.push({
        name: "UserType",
        value: compObj.userType,
        source,
        confidence: "high"
    });
    if (compObj.clipboardFormat) properties.push({
        name: "ClipboardFormat",
        value: compObj.clipboardFormat,
        source,
        confidence: "medium"
    });
    if (compObj.progId) properties.push({
        name: "ProgID",
        value: labelPossibleIdentifier(compObj.progId),
        source,
        confidence: "high"
    });
    return properties;
}
function inferActiveXContentsProperties(bytes, source, userType) {
    const properties = [];
    const strings = collectStrings(bytes).filter((item)=>item.value.length > 1);
    if (bytes.byteLength >= 8) {
        const minor = bytes[0];
        const major = bytes[1];
        const recordSize = readU16(bytes, 2);
        const propMask = readU32(bytes, 4);
        if (minor === 0 && [
            2,
            4
        ].includes(major) && recordSize >= 4 && recordSize <= bytes.byteLength) {
            properties.push({
                name: "PersistedRecordVersion",
                value: `${major}.${minor}`,
                source,
                confidence: "medium"
            });
            properties.push({
                name: "PersistedRecordSize",
                value: `${recordSize} bytes`,
                source,
                confidence: "medium"
            });
            properties.push({
                name: "PersistedPropMask",
                value: `${toHex(propMask, 8)} (${describeActiveXPropertyMask(propMask, userType).join(", ") || "unknown flags"})`,
                source,
                confidence: "medium"
            });
        }
    }
    const usefulStrings = strings.filter((item)=>isUsefulDecodedPropertyString(item.value));
    const caption = usefulStrings.find((item)=>isHumanLabel(item.value));
    if (caption) properties.push({
        name: "Caption",
        value: caption.value,
        source: `${source}@${toHex(caption.offset)}`,
        confidence: "high"
    });
    const font = usefulStrings.find((item)=>/^(Aptos|Arial|Calibri|Cambria|Courier New|MS Sans Serif|Segoe UI|Tahoma|Times New Roman|Verdana)$/i.test(item.value));
    if (font) properties.push({
        name: "FontName",
        value: font.value,
        source: `${source}@${toHex(font.offset)}`,
        confidence: "high"
    });
    for (const color of findOleColorCandidates(bytes).slice(0, 8)){
        properties.push({
            name: color.name,
            value: color.value,
            source: `${source}@${toHex(color.offset)}`,
            confidence: color.confidence
        });
    }
    const extraStrings = usefulStrings.filter((item)=>item !== caption && item !== font).slice(0, 10);
    for (const item of extraStrings){
        properties.push({
            name: "StringProperty",
            value: labelPossibleIdentifier(item.value),
            source: `${source}@${toHex(item.offset)}`,
            confidence: "low"
        });
    }
    return properties;
}
function describeActiveXPropertyMask(propMask, userType) {
    const type = userType?.toLowerCase() ?? "";
    const generic = [
        {
            bit: 2,
            name: "Color/visual property"
        },
        {
            bit: 3,
            name: "Caption/string data"
        },
        {
            bit: 5,
            name: "Font/property data"
        },
        {
            bit: 7,
            name: "Enabled/visibility state"
        },
        {
            bit: 10,
            name: "Picture/icon data"
        },
        {
            bit: 11,
            name: "Mouse/icon data"
        }
    ];
    if (type.includes("commandbutton")) {
        generic[1] = {
            bit: 3,
            name: "Caption"
        };
        generic[2] = {
            bit: 5,
            name: "Font"
        };
    }
    return generic.filter((flag)=>(propMask & 1 << flag.bit) !== 0).map((flag)=>flag.name);
}
function findOleColorCandidates(bytes) {
    const colors = [];
    for(let offset = 0; offset + 4 <= bytes.length; offset += 4){
        const value = readU32(bytes, offset);
        const high = value >>> 24;
        if (high === 0x80) {
            colors.push({
                offset,
                name: "SystemColor",
                value: `${toHex(value, 8)} (${describeSystemOleColor(value & 0xff)})`,
                confidence: "medium"
            });
        } else if (high === 0 && value !== 0 && value <= 0x00ffffff) {
            colors.push({
                offset,
                name: "RgbColor",
                value: `${toHex(value, 8)} (#${(value & 0xffffff).toString(16).padStart(6, "0")})`,
                confidence: "low"
            });
        }
    }
    return colors;
}
function describeSystemOleColor(index) {
    const labels = {
        0x00: "scrollbar",
        0x05: "window background",
        0x08: "window text",
        0x0f: "button face",
        0x12: "button text",
        0x14: "highlight text",
        0x15: "button shadow"
    };
    return labels[index] ?? `system color ${index}`;
}
function parseActiveXXmlProperties(xml) {
    const properties = {};
    const attrRegex = /(?:^|\s)([A-Za-z_:][A-Za-z0-9_.:-]*)\s*=\s*"([^"]*)"/g;
    let match;
    while(match = attrRegex.exec(xml)){
        const key = match[1].replace(/^.*:/, "");
        properties[key] = labelPossibleIdentifier(match[2]);
    }
    const tagRegex = /<([A-Za-z_:][A-Za-z0-9_.:-]*)[^>]*>([^<]+)<\/\1>/g;
    while(match = tagRegex.exec(xml)){
        const key = match[1].replace(/^.*:/, "");
        properties[key] = labelPossibleIdentifier(match[2].trim());
    }
    const relationshipId = xml.match(/r:id="([^"]+)"/i)?.[1];
    if (relationshipId) properties.relationshipId = relationshipId;
    return properties;
}
function normalizeGuidString(value) {
    const match = value?.match(/\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?/i);
    if (!match) return undefined;
    const raw = match[0].replace(/[{}]/g, "").toUpperCase();
    return `{${raw}}`;
}
function flattenControls(controls, parentPath) {
    return controls.flatMap((control)=>[
            {
                id: control.id,
                path: control.path,
                name: control.name,
                caption: control.caption,
                type: control.type,
                progId: control.progId,
                bounds: control.bounds,
                sourceStreams: control.sourceStreams,
                properties: control.properties,
                parentPath
            },
            ...flattenControls(control.children, control.path)
        ]);
}
function linkEventsToControls(modules, forms) {
    const controls = forms.flatMap((form)=>form.controls.map((control)=>({
                form,
                control
            })));
    for (const module of modules){
        for (const event of module.events){
            const linked = controls.find(({ form, control })=>form.name === module.name && [
                    control.name,
                    control.caption,
                    control.id,
                    ...Object.values(control.properties)
                ].filter(Boolean).some((value)=>normalizeIdentifier(String(value)) === normalizeIdentifier(event.controlName)));
            if (linked) {
                event.linkedControlPath = linked.control.path;
                event.linkedControlType = linked.control.type;
            }
        }
    }
}
function addInferredEventControls(modules, forms) {
    for (const form of forms){
        const module = modules.find((candidate)=>candidate.name === form.name);
        if (!module) continue;
        for (const event of module.events){
            const exists = form.controls.some((control)=>[
                    control.name,
                    control.caption,
                    control.id,
                    ...Object.values(control.properties)
                ].filter(Boolean).some((value)=>normalizeIdentifier(String(value)) === normalizeIdentifier(event.controlName)));
            if (exists) continue;
            form.controls.push({
                id: event.controlName,
                path: `${form.name}/inferred/${event.controlName}`,
                name: event.controlName,
                type: inferControlTypeFromName(event.controlName),
                properties: {
                    InferredFromEvent: event.procedure,
                    Confidence: "event-name heuristic"
                },
                sourceStreams: []
            });
        }
    }
}
function inferControlTypeFromName(name) {
    const lower = name.toLowerCase();
    if (lower.startsWith("cmd") || lower.startsWith("btn") || lower.includes("button")) return "Inferred command button";
    if (lower.startsWith("chk") || lower.includes("checkbox")) return "Inferred checkbox";
    if (lower.startsWith("opt") || lower.includes("option")) return "Inferred option button";
    if (lower.startsWith("txt") || lower.includes("textbox")) return "Inferred text box";
    if (lower.startsWith("cmb") || lower.includes("combo")) return "Inferred combo box";
    if (lower.startsWith("lst") || lower.includes("list")) return "Inferred list box";
    if (lower.startsWith("lbl") || lower.includes("label")) return "Inferred label";
    return "Inferred MSForms control";
}
function normalizeIdentifier(value) {
    return value.replace(/\s+\(.+\)$/, "").replace(/[^a-z0-9_]/gi, "").toLowerCase();
}
function detectDependencies(modules, files) {
    const dependencies = [];
    const patterns = [
        {
            category: "WinAPI",
            reason: "Declare statement",
            regex: /\bDeclare\s+(PtrSafe\s+)?(Sub|Function)\s+([A-Za-z0-9_]+)/ig
        },
        {
            category: "COM automation",
            reason: "CreateObject call",
            regex: /\bCreateObject\s*\(\s*"([^"]+)"/ig
        },
        {
            category: "COM automation",
            reason: "GetObject call",
            regex: /\bGetObject\s*\([^)]*"([^"]+)"/ig
        },
        {
            category: "File system",
            reason: "FileSystemObject or file IO usage",
            regex: /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Kill\s+|MkDir\s+|RmDir\s+|Dir\s*\()/ig
        },
        {
            category: "Shell/process",
            reason: "Shell execution",
            regex: /\b(Shell|WScript\.Shell|Run\s*\()/ig
        },
        {
            category: "HTTP/network",
            reason: "HTTP client usage",
            regex: /\b(XMLHTTP|WinHttp|ServerXMLHTTP|WebRequest|InternetOpen|URLDownloadToFile)\b/ig
        },
        {
            category: "Database",
            reason: "Database/ADO usage",
            regex: /\b(ADODB|DAO\.|Recordset|ConnectionString|OpenRecordset)\b/ig
        },
        {
            category: "Office automation",
            reason: "Office object model usage",
            regex: /\b(Excel\.|Word\.|Outlook\.|PowerPoint\.|Application\.|Workbook|Worksheet|Range\(|Cells\()\b/ig
        },
        {
            category: "Registry",
            reason: "Registry access",
            regex: /\b(RegRead|RegWrite|RegDelete|GetSetting|SaveSetting|DeleteSetting)\b/ig
        }
    ];
    for (const module of modules){
        const source = module.procedures.map((procedure)=>procedure.signature).join("\n") + "\n" + module.declarations.join("\n");
        const fullSource = files.find((file)=>file.name === module.fileName)?.text ?? source;
        for (const pattern of patterns){
            for (const match of fullSource.matchAll(pattern.regex)){
                const value = match[1] || match[3] || match[0];
                dependencies.push({
                    category: pattern.category,
                    value: value.trim(),
                    source: module.fileName,
                    reason: pattern.reason
                });
            }
        }
    }
    return uniqueDependencies(dependencies);
}
function uniqueDependencies(dependencies) {
    const seen = new Set();
    return dependencies.filter((dependency)=>{
        const key = `${dependency.category}:${dependency.value}:${dependency.source}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}
function renderApplicationSummary(model) {
    const lines = [
        "# Application Migration Summary",
        "",
        `Generated: ${model.generatedAt}`,
        "",
        "## Inventory",
        "",
        `- Source files: ${model.sourceFiles.length}`,
        `- Modules: ${model.modules.length}`,
        `- Project references: ${model.projectReferences.length}`,
        `- Forms: ${model.forms.length}`,
        `- Controls: ${model.forms.reduce((sum, form)=>sum + form.controls.length, 0)}`,
        `- Document ActiveX controls: ${model.documentControls.length}`,
        `- Procedures: ${model.modules.reduce((sum, module)=>sum + module.procedures.length, 0)}`,
        `- Dependencies: ${model.dependencies.length}`,
        `- Assets: ${model.assets.length}`,
        "",
        "## Forms",
        ""
    ];
    for (const form of model.forms){
        lines.push(`### ${form.name}`, "");
        for (const [key, value] of Object.entries(form.properties).slice(0, 12))lines.push(`- ${key}: ${value}`);
        if (form.controls.length) {
            lines.push("- Controls:");
            for (const control of form.controls){
                lines.push(`  - ${control.name || control.caption || control.id}: ${control.type ?? control.progId ?? "unknown control"} (${control.path})`);
            }
        }
        lines.push("");
    }
    if (model.documentControls.length) {
        lines.push("## Document ActiveX Controls", "");
        for (const control of model.documentControls){
            lines.push(`- ${control.id}: ${control.label ?? control.classId ?? "Unknown ActiveX control"}`);
            lines.push(`  - XML: ${control.xmlPath}`);
            if (control.binPath) lines.push(`  - Binary persistence: ${control.binPath}`);
            if (control.persistenceAnalysis?.compObj?.progId) lines.push(`  - ProgID: ${control.persistenceAnalysis.compObj.progId}`);
            const caption = control.persistenceAnalysis?.properties.find((property)=>property.name === "Caption");
            if (caption) lines.push(`  - Caption: ${caption.value}`);
        }
        lines.push("");
    }
    if (model.projectReferences.length) {
        lines.push("## VBA Project References", "");
        for (const reference of model.projectReferences){
            lines.push(`- ${reference.name ?? reference.kind}: ${reference.libId ?? reference.raw}`);
            if (reference.guid) lines.push(`  - GUID: ${reference.guid}${labelGuid(reference.guid) ? ` (${labelGuid(reference.guid)})` : ""}`);
            if (reference.path) lines.push(`  - Path: ${reference.path}`);
        }
        lines.push("");
    }
    lines.push("## Modules", "");
    for (const module of model.modules){
        lines.push(`### ${module.name}`, "");
        lines.push(`- Kind: ${module.kind}`);
        lines.push(`- Procedures: ${module.procedures.length}`);
        lines.push(`- Variables: ${module.variables.length}`);
        lines.push(`- Constants: ${module.constants.length}`);
        if (module.references.length) lines.push(`- References modules: ${module.references.join(", ")}`);
        if (module.riskMarkers.length) lines.push(`- Risk markers: ${module.riskMarkers.length}`);
        const uses = [
            ...new Set(module.procedures.flatMap((procedure)=>procedure.uses))
        ];
        if (uses.length) lines.push(`- Uses: ${uses.join(", ")}`);
        if (module.events.length) {
            lines.push("- Events:");
            for (const event of module.events){
                lines.push(`  - ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : ""}`);
            }
        }
        lines.push("");
    }
    lines.push("## Dependencies", "");
    if (model.dependencies.length === 0) lines.push("- No high-confidence dependencies detected.");
    for (const dependency of model.dependencies){
        lines.push(`- ${dependency.category}: ${dependency.value} (${dependency.source}, ${dependency.reason})`);
    }
    lines.push("", "## Migration Notes", "");
    for (const note of model.migrationNotes)lines.push(`- ${note}`);
    return `${lines.join("\n")}\n`;
}
function renderMigrationChecklist(model) {
    const riskMarkers = model.modules.flatMap((module)=>module.riskMarkers.map((risk)=>({
                module: module.name,
                ...risk
            })));
    const unlinkedEvents = model.modules.flatMap((module)=>module.events.filter((event)=>!event.linkedControlPath).map((event)=>({
                module: module.name,
                ...event
            })));
    const globalState = model.modules.flatMap((module)=>module.variables.filter((variable)=>/^(Public|Global)$/i.test(variable.scope)).map((variable)=>({
                module: module.name,
                ...variable
            })));
    const lines = [
        "# Migration Checklist",
        "",
        "## UI Reconstruction",
        "",
        `- [ ] Recreate ${model.forms.length} form(s) as frontend views/components.`,
        `- [ ] Recreate ${model.forms.reduce((sum, form)=>sum + form.controls.length, 0)} control(s), including inferred event-only controls.`,
        `- [ ] Recreate or intentionally discard ${model.documentControls.length} document-level ActiveX control(s).`,
        `- [ ] Verify layout properties from designer summaries against screenshots or Office if available.`,
        `- [ ] Wire ${model.modules.reduce((sum, module)=>sum + module.events.length, 0)} event handler(s) to frontend callbacks.`,
        "",
        "## Business Logic",
        "",
        `- [ ] Port ${model.modules.reduce((sum, module)=>sum + module.procedures.length, 0)} VBA procedure(s).`,
        `- [ ] Replace ${globalState.length} public/global variable(s) with explicit app state.`,
        "- [ ] Preserve procedure names in comments or metadata for traceability.",
        "- [ ] Add tests around translated calculations, validation, and state transitions.",
        "",
        "## Dependencies",
        ""
    ];
    if (model.dependencies.length === 0) {
        lines.push("- [ ] No high-confidence dependencies detected; still review source code manually.");
    } else {
        for (const dependency of model.dependencies){
            lines.push(`- [ ] Replace or implement ${dependency.category}: ${dependency.value} (${dependency.source}).`);
        }
    }
    lines.push("", "## Review Items", "");
    if (riskMarkers.length === 0) {
        lines.push("- [ ] No risk markers detected by heuristic scan.");
    } else {
        for (const risk of riskMarkers.slice(0, 80)){
            lines.push(`- [ ] ${risk.module}:${risk.line} ${risk.category} - ${risk.reason}`);
        }
    }
    lines.push("", "## Unlinked Events", "");
    if (unlinkedEvents.length === 0) {
        lines.push("- [x] All detected event handlers are linked or inferred.");
    } else {
        for (const event of unlinkedEvents){
            lines.push(`- [ ] ${event.module}.${event.procedure} (${event.controlName}.${event.eventName}) needs manual control mapping.`);
        }
    }
    return `${lines.join("\n")}\n`;
}
function renderCallGraph(model) {
    const lines = [
        "# VBA Call Graph",
        "",
        "This graph is inferred from procedure-name references in procedure bodies. Treat it as a navigation aid, not a complete compiler graph.",
        "",
        "## Mermaid",
        "",
        "```mermaid",
        "graph TD"
    ];
    let edgeCount = 0;
    for (const module of model.modules){
        for (const procedure of module.procedures){
            const from = graphNodeId(`${module.name}.${procedure.name}`);
            if (procedure.calls.length === 0) {
                lines.push(`  ${from}["${escapeMermaidLabel(`${module.name}.${procedure.name}`)}"]`);
            }
            for (const call of procedure.calls){
                const targetModule = model.modules.find((candidate)=>candidate.procedures.some((candidateProcedure)=>candidateProcedure.name === call));
                const toLabel = `${targetModule?.name ?? "unknown"}.${call}`;
                const to = graphNodeId(toLabel);
                lines.push(`  ${from}["${escapeMermaidLabel(`${module.name}.${procedure.name}`)}"] --> ${to}["${escapeMermaidLabel(toLabel)}"]`);
                edgeCount += 1;
            }
        }
    }
    lines.push("```", "", `Edges: ${edgeCount}`, "", "## Event Entry Points", "");
    for (const module of model.modules.filter((candidate)=>candidate.events.length)){
        lines.push(`### ${module.name}`, "");
        for (const event of module.events){
            lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : ""}`);
        }
        lines.push("");
    }
    return `${lines.join("\n")}\n`;
}
function renderDependencyReport(model) {
    const lines = [
        "# Dependency Report",
        "",
        "This report highlights integrations and platform-specific calls that usually need redesign when moving VBA to a browser frontend.",
        ""
    ];
    const byCategory = groupBy(model.dependencies, (dependency)=>dependency.category);
    if (byCategory.size === 0) {
        lines.push("No high-confidence dependencies detected.", "");
    }
    for (const [category, dependencies] of byCategory){
        lines.push(`## ${category}`, "");
        for (const dependency of dependencies){
            lines.push(`- ${dependency.value} in ${dependency.source}: ${dependency.reason}`);
        }
        lines.push("");
    }
    lines.push("## Risk Markers", "");
    for (const module of model.modules){
        if (module.riskMarkers.length === 0) continue;
        lines.push(`### ${module.name}`, "");
        for (const risk of module.riskMarkers.slice(0, 60)){
            lines.push(`- Line ${risk.line}: ${risk.category} - ${risk.reason}`);
            lines.push(`  - ${risk.text}`);
        }
        lines.push("");
    }
    return `${lines.join("\n")}\n`;
}
function renderProjectReferencesReport(model) {
    const lines = [
        "# VBA Project References",
        "",
        "This report is generated from the textual PROJECT stream and the decompressed MS-OVBA dir stream reference records.",
        "",
        `Total references: ${model.projectReferences.length}`,
        ""
    ];
    if (model.projectReferences.length === 0) {
        lines.push("No VBA project references were decoded.", "");
        return `${lines.join("\n")}\n`;
    }
    const byKind = groupBy(model.projectReferences, (reference)=>reference.kind);
    for (const [kind, references] of byKind){
        lines.push(`## ${titleCase(kind)} References`, "");
        for (const reference of references){
            const title = reference.name ?? reference.guid ?? reference.libId ?? reference.raw;
            lines.push(`- ${title}`);
            lines.push(`  - Source: ${reference.source}`);
            if (reference.guid) lines.push(`  - GUID: ${reference.guid}${labelGuid(reference.guid) ? ` (${labelGuid(reference.guid)})` : ""}`);
            if (reference.version) lines.push(`  - Version: ${reference.version}`);
            if (reference.path) lines.push(`  - Path: ${reference.path}`);
            if (reference.libId) lines.push(`  - LibId: ${reference.libId}`);
            lines.push(`  - Raw: ${reference.raw}`);
        }
        lines.push("");
    }
    lines.push("## Migration Notes", "");
    lines.push("- Registered references usually map to Office/VBA object-model dependencies that need browser replacements or backend adapters.");
    lines.push("- Control references usually indicate MSForms or third-party ActiveX controls; rebuild these as native HTML controls where possible.");
    lines.push("- Project references indicate another VBA project must be extracted too for a complete migration.");
    return `${lines.join("\n")}\n`;
}
function titleCase(value) {
    return value.replace(/\b[a-z]/g, (letter)=>letter.toUpperCase());
}
function renderFrontendImplementationPlan(model) {
    const lines = [
        "# Frontend Implementation Plan",
        "",
        "This is a generated starting plan for rebuilding the extracted VBA application as a frontend website.",
        "",
        "## Suggested Architecture",
        "",
        "- Use a single-page frontend unless the workbook contains clearly separate business workflows.",
        "- Model each VBA UserForm as a page, dialog, or component.",
        "- Model each standard module as a service/helper module.",
        "- Model class modules as domain classes, state machines, or service objects.",
        "- Replace implicit workbook globals with explicit state objects.",
        "- Route all file/network/database/Office integrations through adapter interfaces.",
        "",
        "## State Model Candidates",
        ""
    ];
    const globalVariables = model.modules.flatMap((module)=>module.variables.filter((variable)=>/^(Public|Global)$/i.test(variable.scope)).map((variable)=>({
                module,
                variable
            })));
    if (globalVariables.length === 0) {
        lines.push("- No public/global VBA variables detected. Derive state from forms, controls, and procedure data flow.");
    } else {
        for (const { module, variable } of globalVariables.slice(0, 80)){
            lines.push(`- ${module.name}.${variable.name}${variable.type ? `: ${variable.type}` : ""} -> app state candidate`);
        }
    }
    lines.push("", "## UI Components", "");
    for (const form of model.forms){
        lines.push(`### ${form.name}`, "");
        lines.push(`- Component: \`${form.name}View\``);
        if (form.properties.Caption) lines.push(`- Title: ${form.properties.Caption}`);
        if (form.controls.length) {
            lines.push("- Child components/controls:");
            for (const control of form.controls){
                lines.push(`  - ${control.name || control.caption || control.id}: ${mapControlToFrontend(control)}`);
            }
        }
        lines.push("");
    }
    lines.push("## Services", "");
    for (const module of model.modules.filter((candidate)=>candidate.kind === "standard" || candidate.kind === "class")){
        lines.push(`- \`${module.name}Service\`: port ${module.procedures.length} procedure(s) from ${module.fileName}.`);
    }
    lines.push("", "## Event Wiring", "");
    for (const module of model.modules.filter((candidate)=>candidate.events.length)){
        lines.push(`### ${module.name}`, "");
        for (const event of module.events){
            lines.push(`- Wire ${event.linkedControlPath ?? event.controlName} ${event.eventName} -> \`${event.procedure}\`.`);
        }
        lines.push("");
    }
    lines.push("## Integration Adapters", "");
    const categories = [
        ...new Set(model.dependencies.map((dependency)=>dependency.category))
    ];
    if (categories.length === 0) lines.push("- No adapters detected by heuristic scan.");
    for (const category of categories){
        lines.push(`- ${category}: create an adapter boundary and decide browser-only vs backend implementation.`);
    }
    return `${lines.join("\n")}\n`;
}
function buildLayoutModel(model) {
    return {
        generatedAt: model.generatedAt,
        units: "twips",
        note: "Bounds are best-effort reconstructions from MS-OFORMS streams. Verify against original Office forms before pixel-perfect rebuilds.",
        documentControls: model.documentControls.map((control)=>({
                id: control.id,
                label: control.label,
                classId: control.classId,
                xmlPath: control.xmlPath,
                binPath: control.binPath,
                properties: control.properties
            })),
        views: model.forms.map((form)=>({
                name: form.name,
                caption: form.properties.Caption,
                clientWidth: parseNumber(form.properties.ClientWidth),
                clientHeight: parseNumber(form.properties.ClientHeight),
                controls: form.controls.map((control)=>({
                        path: control.path,
                        name: control.name,
                        caption: control.caption,
                        type: control.type,
                        progId: control.progId,
                        bounds: control.bounds,
                        parentPath: control.parentPath
                    }))
            }))
    };
}
function buildTraceabilityMap(files, model) {
    const fileEntries = files.map((file, index)=>({
            extractedName: file.name,
            kind: file.kind,
            zipPath: getZipPath(file, index),
            sourcePath: file.sourcePath,
            size: file.bytes.byteLength
        }));
    return {
        generatedAt: model.generatedAt,
        files: fileEntries,
        modules: model.modules.map((module)=>({
                name: module.name,
                fileName: module.fileName,
                sourcePath: module.sourcePath,
                zipPath: fileEntries.find((file)=>file.extractedName === module.fileName)?.zipPath,
                procedures: module.procedures.map((procedure, index)=>({
                        name: procedure.name,
                        lines: {
                            start: procedure.lineStart,
                            end: procedure.lineEnd
                        },
                        chunkPath: `procedure-chunks/${safeFileName(module.name)}/${String(index + 1).padStart(3, "0")}-${safeFileName(procedure.name)}.vba`
                    }))
            })),
        forms: model.forms.map((form)=>({
                name: form.name,
                designerSummaryPath: `${safeFileName(form.name)}/designer-summary/${safeFileName(form.name)}.designer.json`,
                controls: form.controls.map((control)=>({
                        path: control.path,
                        name: control.name,
                        caption: control.caption,
                        type: control.type,
                        sourceStreams: control.sourceStreams,
                        parentPath: control.parentPath
                    }))
            })),
        documentControls: model.documentControls.map((control)=>({
                id: control.id,
                label: control.label,
                classId: control.classId,
                xmlPath: control.xmlPath,
                binPath: control.binPath,
                sourceFiles: control.sourceFiles
            })),
        generatedArtifacts: [
            "application-model.json",
            "application-summary.md",
            "llm-rebuild-brief.md",
            "migration-checklist.md",
            "call-graph.md",
            "dependency-report.md",
            "vba-project-references.md",
            "frontend-implementation-plan.md",
            "migration-test-plan.md",
            "layout-model.json",
            "traceability-map.json",
            "activex-controls.json",
            "activex-persistence.json",
            "visual-previews/*.svg",
            "visual-previews/*.html",
            "vba-project-references.json"
        ]
    };
}
function renderValidationReport(files, model) {
    const sourceFiles = files.filter((file)=>file.kind === "vba" || file.kind === "frm");
    const fallbackFiles = sourceFiles.filter((file)=>file.sourcePath.includes("recovered by scan"));
    const controls = model.forms.flatMap((form)=>form.controls);
    const inferredControls = controls.filter((control)=>control.path.includes("/inferred/"));
    const typedControls = controls.filter((control)=>control.type || control.progId);
    const namedControls = controls.filter((control)=>control.name || control.caption);
    const boundedControls = controls.filter((control)=>control.bounds);
    const events = model.modules.flatMap((module)=>module.events);
    const linkedEvents = events.filter((event)=>event.linkedControlPath);
    const lowConfidenceProperties = controls.flatMap((control)=>Object.entries(control.properties).filter(([, value])=>/\(low\)$/i.test(value)).map(([key, value])=>({
                control: control.path,
                key,
                value
            })));
    const media = files.filter((file)=>file.kind === "media");
    const activeXParts = files.filter((file)=>/\/activeX\//i.test(file.sourcePath));
    const manifest = files.find((file)=>file.name === "office-package-manifest.json");
    const riskMarkers = model.modules.flatMap((module)=>module.riskMarkers.map((risk)=>({
                module: module.name,
                ...risk
            })));
    const hasVbaProjectOutput = sourceFiles.length > 0;
    const scoreItems = [
        sourceFiles.length > 0 || media.length > 0 || activeXParts.length > 0,
        fallbackFiles.length === 0,
        events.length === 0 || linkedEvents.length / events.length >= 0.75,
        controls.length === 0 || typedControls.length / controls.length >= 0.75,
        controls.length === 0 || inferredControls.length / controls.length <= 0.5
    ];
    const score = Math.round(scoreItems.filter(Boolean).length / scoreItems.length * 100);
    const lines = [
        "# Validation Report",
        "",
        `Overall extraction confidence: ${score}%`,
        `VBA migration package: ${hasVbaProjectOutput ? "yes" : "no VBA source extracted"}`,
        "",
        "## Extraction Coverage",
        "",
        `- Extracted files: ${files.length}`,
        `- VBA/FRM source files: ${sourceFiles.length}`,
        `- Source files recovered by scan fallback: ${fallbackFiles.length}`,
        `- Forms with designer summaries: ${model.forms.length}`,
        `- Controls: ${controls.length}`,
        `- Document ActiveX controls: ${model.documentControls.length}`,
        `- VBA project references: ${model.projectReferences.length}`,
        `- Media assets: ${media.length}`,
        `- Office package manifest: ${manifest ? "yes" : "no"}`,
        "",
        "## Model Quality",
        "",
        `- Typed controls: ${typedControls.length}/${controls.length}`,
        `- Named/captioned controls: ${namedControls.length}/${controls.length}`,
        `- Controls with bounds: ${boundedControls.length}/${controls.length}`,
        `- Inferred controls: ${inferredControls.length}/${controls.length}`,
        `- Linked events: ${linkedEvents.length}/${events.length}`,
        `- Low-confidence decoded properties: ${lowConfidenceProperties.length}`,
        `- Risk markers: ${riskMarkers.length}`,
        "",
        "## Human Review Required",
        ""
    ];
    const reviewItems = [];
    if (!hasVbaProjectOutput) reviewItems.push("No VBA/FRM source was extracted. This package is useful for Office media/ActiveX migration context, but not a complete VBA application rebuild.");
    if (fallbackFiles.length) reviewItems.push("Some source modules used byte-scan fallback instead of exact dir-stream offsets.");
    if (inferredControls.length) reviewItems.push("Some controls were inferred from event handler names rather than decoded from form streams.");
    if (events.length !== linkedEvents.length) reviewItems.push("Some event handlers could not be linked to decoded/inferred controls.");
    if (controls.length !== boundedControls.length) reviewItems.push("Some controls have no trustworthy layout bounds.");
    if (lowConfidenceProperties.length) reviewItems.push("Some decoded control properties are marked low-confidence.");
    if (activeXParts.length && model.documentControls.length === 0) reviewItems.push("ActiveX package parts were found but no documentControls were decoded.");
    if (riskMarkers.length) reviewItems.push("Risk markers indicate platform-specific or migration-sensitive VBA code.");
    if (reviewItems.length === 0) reviewItems.push("No major validation warnings detected by heuristics.");
    for (const item of reviewItems)lines.push(`- ${item}`);
    lines.push("", "## Unlinked Events", "");
    const unlinkedEvents = events.filter((event)=>!event.linkedControlPath);
    if (unlinkedEvents.length === 0) lines.push("- None.");
    for (const event of unlinkedEvents.slice(0, 80))lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}`);
    lines.push("", "## Inferred Controls", "");
    if (inferredControls.length === 0) lines.push("- None.");
    for (const control of inferredControls.slice(0, 80))lines.push(`- ${control.path}: ${control.type ?? "unknown"} from ${control.properties.InferredFromEvent ?? "event heuristic"}`);
    lines.push("", "## Low-Confidence Properties", "");
    if (lowConfidenceProperties.length === 0) lines.push("- None.");
    for (const property of lowConfidenceProperties.slice(0, 80))lines.push(`- ${property.control}: ${property.key} = ${property.value}`);
    lines.push("", "## Fallback Source Files", "");
    if (fallbackFiles.length === 0) lines.push("- None.");
    for (const file of fallbackFiles)lines.push(`- ${file.name}: ${file.sourcePath}`);
    lines.push("", "## Media And ActiveX", "");
    lines.push(`- Media assets recovered: ${media.length}`);
    lines.push(`- ActiveX package parts extracted: ${activeXParts.length}`);
    lines.push(`- ActiveX controls modeled: ${model.documentControls.length}`);
    lines.push(`- ActiveX binary persistence parsed: ${model.documentControls.filter((control)=>control.persistenceAnalysis).length}`);
    return `${lines.join("\n")}\n`;
}
function parseNumber(value) {
    if (!value) return undefined;
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : undefined;
}
function mapControlToFrontend(control) {
    const type = `${control.type ?? ""} ${control.progId ?? ""}`.toLowerCase();
    if (type.includes("commandbutton")) return "button";
    if (type.includes("checkbox")) return "checkbox";
    if (type.includes("optionbutton")) return "radio button";
    if (type.includes("textbox")) return "text input";
    if (type.includes("combobox")) return "select/combobox";
    if (type.includes("listbox")) return "list/select";
    if (type.includes("multipage") || type.includes("tabstrip")) return "tabs";
    if (type.includes("frame")) return "fieldset/panel";
    if (type.includes("image")) return "image";
    if (type.includes("label")) return "text label";
    if (type.includes("scrollbar")) return "scrollbar or range control";
    if (type.includes("spinbutton")) return "number stepper";
    return "custom component";
}
function renderMigrationTestPlan(model) {
    const lines = [
        "# Migration Test Plan",
        "",
        "Use this generated plan to validate that the frontend rebuild preserves VBA behavior.",
        "",
        "## Smoke Tests",
        "",
        "- [ ] Application loads without console errors.",
        "- [ ] Every reconstructed form/view can open.",
        "- [ ] Every visible control can render with a recognizable label/type.",
        "- [ ] Recovered media assets load or have intentional replacements.",
        "",
        "## Event Tests",
        ""
    ];
    for (const module of model.modules.filter((candidate)=>candidate.events.length)){
        lines.push(`### ${module.name}`, "");
        for (const event of module.events){
            lines.push(`- [ ] Trigger ${event.controlName}.${event.eventName} and verify behavior from \`${event.procedure}\`.`);
        }
        lines.push("");
    }
    lines.push("## Procedure Tests", "");
    for (const module of model.modules){
        const candidates = module.procedures.filter((procedure)=>procedure.kind === "Function" || procedure.uses.length || procedure.calls.length);
        if (candidates.length === 0) continue;
        lines.push(`### ${module.name}`, "");
        for (const procedure of candidates.slice(0, 80)){
            const notes = [
                procedure.uses.join(", "),
                procedure.calls.length ? `calls ${procedure.calls.join(", ")}` : ""
            ].filter(Boolean).join("; ");
            lines.push(`- [ ] ${procedure.name}${notes ? ` (${notes})` : ""}`);
        }
        lines.push("");
    }
    lines.push("## Dependency Tests", "");
    if (model.dependencies.length === 0) {
        lines.push("- [ ] No detected dependency tests. Review source manually for hidden integrations.");
    } else {
        for (const dependency of model.dependencies){
            lines.push(`- [ ] Validate replacement for ${dependency.category}: ${dependency.value} (${dependency.source}).`);
        }
    }
    lines.push("", "## Regression Data", "");
    lines.push("- [ ] Collect representative workbook/document inputs from the business owner.");
    lines.push("- [ ] Capture expected outputs from the original VBA app before retiring it.");
    lines.push("- [ ] Add automated tests for calculations and transformations before UI polish.");
    return `${lines.join("\n")}\n`;
}
function graphNodeId(value) {
    return `n_${value.replace(/[^a-z0-9_]/gi, "_")}`;
}
function escapeMermaidLabel(value) {
    return value.replace(/"/g, "'");
}
function groupBy(items, keyFn) {
    const grouped = new Map();
    for (const item of items){
        const key = keyFn(item);
        grouped.set(key, [
            ...grouped.get(key) ?? [],
            item
        ]);
    }
    return grouped;
}


// ---- designer-ui.ts ----
function renderLlmRebuildBrief(model) {
    const eventCount = model.modules.reduce((sum, module)=>sum + module.events.length, 0);
    const linkedEventCount = model.modules.reduce((sum, module)=>sum + module.events.filter((event)=>event.linkedControlPath).length, 0);
    const lines = [
        "# LLM Rebuild Brief",
        "",
        "You are rebuilding a legacy Microsoft Office VBA application as a frontend website.",
        "Use this archive as source material. Preserve business behavior first; improve UX only after behavior is understood.",
        "",
        "## Read Order",
        "",
        "1. Read `application-summary.md` for the overview.",
        "2. Read `application-model.json` for structured modules, forms, controls, procedures, events, dependencies, and assets.",
        "3. Read files under each `code/` folder for the original VBA source.",
        "4. Read each `designer-summary/*.designer.json` for reconstructed form/control structure.",
        "5. Inspect `resources/` and `media/` only when UI details or embedded assets are unclear.",
        "",
        "## Rebuild Goals",
        "",
        "- Convert UserForms into web components or pages.",
        "- Convert MSForms controls into native HTML controls.",
        "- Convert event handlers into frontend event callbacks.",
        "- Convert workbook/document object model calls into explicit application state, API calls, file imports, or generated data tables.",
        "- Keep external dependencies visible and ask for replacements when a dependency cannot run in the browser.",
        "",
        "## Inventory",
        "",
        `- Forms: ${model.forms.length}`,
        `- Controls: ${model.forms.reduce((sum, form)=>sum + form.controls.length, 0)}`,
        `- Modules: ${model.modules.length}`,
        `- Procedures: ${model.modules.reduce((sum, module)=>sum + module.procedures.length, 0)}`,
        `- Event handlers: ${eventCount}`,
        `- Linked event handlers: ${linkedEventCount}`,
        `- Detected dependencies: ${model.dependencies.length}`,
        `- VBA project references: ${model.projectReferences.length}`,
        `- Recovered assets: ${model.assets.length}`,
        "",
        "## Suggested Frontend Mapping",
        "",
        "- `Forms.TextBox.1` -> `<input>` or `<textarea>` depending on multiline behavior.",
        "- `Forms.CommandButton.1` -> `<button>`.",
        "- `Forms.CheckBox.1` -> `<input type=\"checkbox\">`.",
        "- `Forms.OptionButton.1` -> radio group.",
        "- `Forms.ComboBox.1` -> `<select>` or combobox.",
        "- `Forms.ListBox.1` -> list/select component.",
        "- `Forms.Frame.1` -> fieldset/panel.",
        "- `Forms.MultiPage.1` and `Forms.TabStrip.1` -> tabs.",
        "- `Forms.Image.1` -> image component using recovered `media/` assets when available.",
        "- Document ActiveX controls -> inspect `activex-controls.json`; rebuild as native frontend controls when class labels are recognized.",
        "",
        "## Event Handler Map",
        ""
    ];
    for (const module of model.modules.filter((candidate)=>candidate.events.length)){
        lines.push(`### ${module.name}`, "");
        for (const event of module.events){
            lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : " -> unlinked"}`);
        }
        lines.push("");
    }
    lines.push("## Dependency Warnings", "");
    if (model.dependencies.length === 0) {
        lines.push("- No high-confidence external dependencies were detected.");
    } else {
        for (const dependency of model.dependencies){
            lines.push(`- ${dependency.category}: ${dependency.value} from ${dependency.source}. Migration concern: ${dependency.reason}.`);
        }
    }
    lines.push("", "## Implementation Advice", "", "- Start by recreating the UI shape from `forms` and `controls` in `application-model.json`.", "- Then wire linked event handlers using the original VBA code as behavior reference.", "- Replace global workbook state with explicit frontend state stores.", "- Treat dependency detections as integration requirements. Do not silently drop them.", "- Preserve original names in comments or metadata so reviewers can trace VBA behavior to the new code.", "");
    return lines.join("\n");
}
function getDisplayPath(file) {
    const streamPath = extractInternalStreamPath(file.sourcePath);
    return streamPath ? printableStreamName(streamPath) : file.sourcePath;
}
function buildDesignerSummary(group) {
    const hasFormCode = group.code.some((file)=>file.kind === "frm");
    if (!hasFormCode && group.resources.length === 0) return undefined;
    if (group.name === "Project metadata") return undefined;
    const frame = parseVbFrameProperties(group.resources);
    const controlsByPath = new Map();
    const rootControls = [];
    for (const resource of group.resources){
        const internalPath = extractInternalStreamPath(resource.sourcePath);
        if (!internalPath) continue;
        const parts = internalPath.split("/");
        if (parts.length < 2 || parts[0] !== group.name) continue;
        for(let depth = 2; depth <= parts.length - 1; depth += 1){
            const controlPath = parts.slice(0, depth).join("/");
            const controlId = parts[depth - 1];
            if (!/^i\d+$/i.test(controlId)) continue;
            let control = controlsByPath.get(controlPath);
            if (!control) {
                control = {
                    id: controlId,
                    path: controlPath,
                    properties: {},
                    sourceStreams: [],
                    children: []
                };
                controlsByPath.set(controlPath, control);
                const parentPath = parts.slice(0, depth - 1).join("/");
                const parent = controlsByPath.get(parentPath);
                if (parent) parent.children.push(control);
                else rootControls.push(control);
            }
        }
        const ownerPath = parts.slice(0, -1).join("/");
        const control = controlsByPath.get(ownerPath);
        if (control) mergeControlResource(control, resource);
    }
    for (const control of controlsByPath.values()){
        finalizeControlBounds(control);
    }
    const directFormProps = extractResourceProperties(group.resources.filter((resource)=>{
        const internalPath = extractInternalStreamPath(resource.sourcePath);
        return internalPath === `${group.name}/f` || internalPath === `${group.name}/\x01CompObj`;
    }));
    for (const [key, value] of Object.entries(directFormProps)){
        if (!frame[key]) frame[key] = value;
    }
    return {
        formName: group.name,
        frame,
        controls: rootControls
    };
}
function parseVbFrameProperties(resources) {
    const frame = {};
    const vbFrame = resources.find((resource)=>extractInternalStreamPath(resource.sourcePath)?.endsWith("/\x03VBFrame"));
    if (!vbFrame) return frame;
    const text = decodeText(vbFrame.bytes);
    const beginMatch = text.match(/Begin\s+\{[^}]+\}\s+([^\s\r\n]+)/i);
    if (beginMatch) frame.Name = beginMatch[1];
    for (const line of text.split(/\r?\n/)){
        const match = line.match(/^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
        if (!match) continue;
        frame[match[1]] = match[2].replace(/^"|"$/g, "");
    }
    return frame;
}
function mergeControlResource(control, resource) {
    const internalPath = extractInternalStreamPath(resource.sourcePath);
    const streamName = internalPath?.split("/").at(-1);
    if (!streamName) return;
    control.sourceStreams.push(internalPath ?? resource.sourcePath);
    const properties = extractResourceProperties([
        resource
    ]);
    const bounds = inferBoundsFromResource(resource);
    if (bounds && !control.bounds) control.bounds = bounds;
    for (const [key, value] of Object.entries(properties)){
        control.properties[key] = value;
    }
    if (streamName === "\x01CompObj") {
        const progId = Object.values(properties).find((value)=>/^Forms\./i.test(value.replace(/\s+\(.+\)$/, "")));
        if (progId) {
            control.progId = progId.replace(/\s+\(.+\)$/, "");
            control.type = labelProgId(control.progId) ?? progId;
        }
        const microsoftFormsName = Object.values(properties).find((value)=>/Microsoft Forms 2\.0/i.test(value));
        if (!control.type && microsoftFormsName) control.type = microsoftFormsName;
    }
    const decodedName = stripConfidence(properties["Decoded Name"]);
    const decodedCaption = stripConfidence(properties["Decoded Caption"]);
    if (decodedName && isHumanLabel(decodedName) && propertyConfidence(properties["Decoded Name"]) !== "low") control.name = decodedName;
    if (decodedCaption && isHumanLabel(decodedCaption)) control.caption = decodedCaption;
    const captionCandidate = Object.entries(properties).find(([key, value])=>key.startsWith("Label ") && isHumanLabel(value));
    if (!control.caption && captionCandidate) control.caption = captionCandidate[1].replace(/\s+\(.+\)$/, "");
    if (!control.name && control.caption) control.name = control.caption;
}
function stripConfidence(value) {
    return value?.replace(/\s+\((?:high|medium|low)\)$/i, "");
}
function propertyConfidence(value) {
    return value?.match(/\((high|medium|low)\)$/i)?.[1].toLowerCase();
}
function finalizeControlBounds(control) {
    const size = parseSizeHint(control.properties["Decoded DisplayedSize"]) ?? parseSizeHint(control.properties["Decoded LogicalSize"]);
    const position = parsePositionHint(control.properties["Decoded SitePosition"]);
    if (size && position) {
        control.bounds = {
            left: position.left,
            top: position.top,
            width: size.width,
            height: size.height,
            unit: "twips",
            confidence: size.confidence === "medium" && position.confidence === "medium" ? "medium" : "low"
        };
    }
}
function parseSizeHint(value) {
    const match = value?.match(/(\d+)\s+x\s+(\d+)\s+twips.*\((medium|low)\)/i);
    if (!match) return undefined;
    return {
        width: Number(match[1]),
        height: Number(match[2]),
        confidence: match[3]
    };
}
function parsePositionHint(value) {
    const match = value?.match(/(\d+),\s*(\d+)\s+twips.*\((medium|low)\)/i);
    if (!match) return undefined;
    return {
        left: Number(match[1]),
        top: Number(match[2]),
        confidence: match[3]
    };
}
function inferBoundsFromResource(resource) {
    const internalPath = extractInternalStreamPath(resource.sourcePath);
    if (!internalPath?.endsWith("/f")) return undefined;
    const bytes = resource.bytes;
    if (bytes.length < 32) return undefined;
    const candidates = [];
    for(let offset = 8; offset + 16 <= Math.min(bytes.length, 160); offset += 4){
        const a = readU32(bytes, offset);
        const b = readU32(bytes, offset + 4);
        const c = readU32(bytes, offset + 8);
        const d = readU32(bytes, offset + 12);
        if ([
            a,
            b,
            c,
            d
        ].every((value)=>value >= 0 && value < 60_000) && c > 0 && d > 0) {
            const looksLikeSize = c >= 30 && d >= 30 && c < 32_000 && d < 32_000;
            const looksLikePosition = a < 25_000 && b < 25_000;
            if (looksLikeSize) {
                candidates.push({
                    left: a,
                    top: b,
                    width: c,
                    height: d,
                    unit: "twips",
                    confidence: looksLikePosition && offset <= 64 ? "medium" : "low"
                });
            }
        }
    }
    const best = candidates.sort((a, b)=>scoreBounds(b) - scoreBounds(a))[0];
    return best?.confidence === "low" && (best.left > 25_000 || best.top > 25_000) ? undefined : best;
}
function scoreBounds(bounds) {
    let score = 0;
    if (bounds.width > bounds.height) score += 1;
    if (bounds.width > 500 && bounds.height > 300) score += 2;
    if (bounds.left < 20_000 && bounds.top < 20_000) score += 1;
    if (bounds.confidence === "medium") score += 2;
    return score;
}
function extractResourceProperties(resources) {
    const properties = {};
    for (const resource of resources){
        const records = resource.analysis?.oforms?.records ?? [];
        for (const record of records){
            for (const property of record.properties){
                const value = labelPossibleIdentifier(property.value);
                if (property.name === "PropMask") {
                    properties[`${record.type} PropMask`] = value;
                    continue;
                }
                if (/^(Font|Picture|MouseIcon)$/i.test(property.name)) {
                    properties[property.name] = value;
                    continue;
                }
                if (/^Decoded /i.test(property.name)) {
                    properties[property.name] = value;
                    continue;
                }
                if (/string @/i.test(property.name) && (/Microsoft Forms/i.test(value) || /^Forms\./i.test(value))) {
                    properties[property.name] = value;
                    continue;
                }
                if (/string @/i.test(property.name) && isHumanLabel(value)) {
                    properties[`Label ${Object.keys(properties).length}`] = value;
                }
            }
        }
    }
    return properties;
}
function isHumanLabel(value) {
    const plain = value.replace(/\s+\(.+\)$/, "");
    if (/^Forms\./i.test(plain) || /Embedded Object/i.test(plain) || /Microsoft Forms/i.test(plain)) return false;
    if (/^Tahoma$/i.test(plain) || /^0x[0-9a-f]+$/i.test(plain) || /^\{[0-9A-F-]+\}$/i.test(plain)) return false;
    if (/^'/.test(plain) || /^[-=]{5,}$/.test(plain) || /https?:\/\//i.test(plain) || /copyright|permission is hereby/i.test(plain)) return false;
    return /^[\w .'-]{2,80}$/i.test(plain);
}
function renderDesignerSummary(summary) {
    const section = document.createElement("section");
    section.className = "designer-summary";
    const header = document.createElement("div");
    header.className = "designer-summary-header";
    const title = document.createElement("h3");
    title.textContent = "Form designer summary";
    const meta = document.createElement("p");
    meta.textContent = `${Object.keys(summary.frame).length} form properties · ${countControls(summary.controls)} controls`;
    header.append(title, meta);
    section.append(header);
    if (Object.keys(summary.frame).length) {
        section.append(renderKeyValueGrid(summary.frame));
    }
    if (summary.controls.length) {
        section.append(renderVisualLayoutPreview(summary));
        const tree = document.createElement("div");
        tree.className = "control-tree";
        for (const control of summary.controls)tree.append(renderControlNode(control));
        section.append(tree);
    }
    return section;
}
function renderKeyValueGrid(values) {
    const dl = document.createElement("dl");
    dl.className = "designer-grid";
    for (const [key, value] of Object.entries(values).slice(0, 24)){
        const dt = document.createElement("dt");
        dt.textContent = key;
        const dd = document.createElement("dd");
        dd.textContent = labelPossibleIdentifier(value);
        dl.append(dt, dd);
    }
    return dl;
}
function renderLayoutPreview(summary) {
    const wrapper = document.createElement("div");
    wrapper.className = "layout-preview-wrap";
    const title = document.createElement("h4");
    title.textContent = "Best-effort layout preview";
    wrapper.append(title);
    const controls = flattenDesignerControls(summary.controls);
    const bounded = controls.filter((control)=>control.bounds);
    if (bounded.length === 0) {
        const empty = document.createElement("p");
        empty.className = "analysis-empty";
        empty.textContent = "No usable bounds detected yet.";
        wrapper.append(empty);
        return wrapper;
    }
    const maxRight = Math.max(...bounded.map((control)=>control.bounds.left + control.bounds.width));
    const maxBottom = Math.max(...bounded.map((control)=>control.bounds.top + control.bounds.height));
    const scale = Math.min(1, 760 / Math.max(maxRight, 1), 420 / Math.max(maxBottom, 1));
    const canvas = document.createElement("div");
    canvas.className = "layout-preview";
    canvas.style.width = `${Math.max(320, maxRight * scale + 24)}px`;
    canvas.style.height = `${Math.max(180, maxBottom * scale + 24)}px`;
    for (const control of bounded){
        const bounds = control.bounds;
        const item = document.createElement("div");
        item.className = "layout-control";
        item.style.left = `${bounds.left * scale + 8}px`;
        item.style.top = `${bounds.top * scale + 8}px`;
        item.style.width = `${Math.max(36, bounds.width * scale)}px`;
        item.style.height = `${Math.max(24, bounds.height * scale)}px`;
        item.title = `${control.path} · ${bounds.confidence}`;
        item.textContent = control.caption || control.name || control.id;
        canvas.append(item);
    }
    wrapper.append(canvas);
    return wrapper;
}
function flattenDesignerControls(controls) {
    return controls.flatMap((control)=>[
            control,
            ...flattenDesignerControls(control.children)
        ]);
}
function renderVisualLayoutPreview(summary) {
    const wrapper = document.createElement("div");
    wrapper.className = "layout-preview-wrap";
    const header = document.createElement("div");
    header.className = "layout-preview-header";
    const title = document.createElement("h4");
    title.textContent = "Visual form preview";
    const actions = document.createElement("div");
    actions.className = "layout-preview-actions";
    header.append(title, actions);
    wrapper.append(header);
    const preview = buildVisualPreviewModel(summary);
    if (!preview) {
        const empty = document.createElement("p");
        empty.className = "analysis-empty";
        empty.textContent = "No usable bounds detected yet.";
        wrapper.append(empty);
        return wrapper;
    }
    const svgButton = document.createElement("button");
    svgButton.type = "button";
    svgButton.textContent = "SVG";
    svgButton.addEventListener("click", ()=>downloadBytes(`${safeFileName(summary.formName)}-preview.svg`, "image/svg+xml", encodeText(renderVisualPreviewSvg(preview))));
    actions.append(svgButton);
    const pngButton = document.createElement("button");
    pngButton.type = "button";
    pngButton.textContent = "PNG";
    pngButton.addEventListener("click", async ()=>{
        const blob = await renderVisualPreviewPng(preview);
        downloadBlob(`${safeFileName(summary.formName)}-preview.png`, blob);
    });
    actions.append(pngButton);
    const canvas = document.createElement("div");
    canvas.className = "layout-preview";
    canvas.style.width = `${preview.width}px`;
    canvas.style.height = `${preview.height}px`;
    for (const control of preview.controls){
        const item = document.createElement("div");
        item.className = `layout-control layout-control-${control.kind}`;
        item.style.left = `${control.left}px`;
        item.style.top = `${control.top}px`;
        item.style.width = `${control.width}px`;
        item.style.height = `${control.height}px`;
        item.title = `${control.path} - ${control.confidence}`;
        item.append(renderVisualControlContents(control));
        canvas.append(item);
    }
    wrapper.append(canvas);
    return wrapper;
}
function buildVisualPreviewModel(summary) {
    const controls = flattenDesignerControls(summary.controls).filter((control)=>control.bounds).map((control)=>({
            control,
            bounds: control.bounds
        }));
    if (controls.length === 0) return undefined;
    const maxRight = Math.max(...controls.map(({ bounds })=>bounds.left + bounds.width));
    const maxBottom = Math.max(...controls.map(({ bounds })=>bounds.top + bounds.height));
    const scale = Math.min(1, 900 / Math.max(maxRight, 1), 560 / Math.max(maxBottom, 1));
    return {
        name: summary.formName,
        width: Math.ceil(Math.max(360, maxRight * scale + 24)),
        height: Math.ceil(Math.max(220, maxBottom * scale + 24)),
        scale,
        controls: controls.map(({ control, bounds })=>({
                id: control.id,
                path: control.path,
                kind: getVisualControlKind(control),
                label: control.caption || control.name || control.id,
                left: Math.round(bounds.left * scale + 8),
                top: Math.round(bounds.top * scale + 8),
                width: Math.round(Math.max(36, bounds.width * scale)),
                height: Math.round(Math.max(24, bounds.height * scale)),
                confidence: bounds.confidence,
                type: control.type,
                progId: control.progId
            }))
    };
}
function getVisualControlKind(control) {
    const value = `${control.type ?? ""} ${control.progId ?? ""} ${control.label ?? ""} ${control.caption ?? ""} ${control.name ?? ""}`.toLowerCase();
    if (value.includes("commandbutton")) return "button";
    if (value.includes("checkbox")) return "checkbox";
    if (value.includes("optionbutton")) return "radio";
    if (value.includes("togglebutton")) return "toggle";
    if (value.includes("textbox")) return "textbox";
    if (value.includes("combobox")) return "combo";
    if (value.includes("listbox")) return "listbox";
    if (value.includes("multipage") || value.includes("tabstrip")) return "tabs";
    if (value.includes("frame")) return "frame";
    if (value.includes("image")) return "image";
    if (value.includes("label")) return "label";
    if (value.includes("scrollbar")) return "scrollbar";
    if (value.includes("spinbutton")) return "spin";
    return "unknown";
}
function renderVisualControlContents(control) {
    const fragment = document.createDocumentFragment();
    const label = document.createElement("span");
    label.className = "layout-control-text";
    label.textContent = control.label;
    if (control.kind === "checkbox" || control.kind === "radio") {
        const mark = document.createElement("span");
        mark.className = `layout-control-mark ${control.kind === "radio" ? "layout-control-mark-round" : ""}`;
        fragment.append(mark, label);
        return fragment;
    }
    if (control.kind === "combo" || control.kind === "spin") {
        const affordance = document.createElement("span");
        affordance.className = "layout-control-affordance";
        affordance.textContent = control.kind === "spin" ? "up/dn" : "v";
        fragment.append(label, affordance);
        return fragment;
    }
    if (control.kind === "image") {
        const imageIcon = document.createElement("span");
        imageIcon.className = "layout-control-image-icon";
        imageIcon.textContent = "img";
        fragment.append(imageIcon, label);
        return fragment;
    }
    fragment.append(label);
    return fragment;
}
function createVisualPreviewArtifacts(groups) {
    const artifacts = [];
    for (const group of groups){
        if (!group.designer) continue;
        const preview = buildVisualPreviewModel(group.designer);
        if (!preview) continue;
        const base = `visual-previews/${safeFileName(group.designer.formName)}`;
        artifacts.push({
            path: `${base}.svg`,
            bytes: encodeText(renderVisualPreviewSvg(preview))
        });
        artifacts.push({
            path: `${base}.html`,
            bytes: encodeText(renderVisualPreviewHtml(preview))
        });
    }
    return artifacts;
}
function renderVisualPreviewSvg(preview) {
    const controls = preview.controls.map((control)=>renderVisualControlSvg(control)).join("\n");
    return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${preview.width}" height="${preview.height}" viewBox="0 0 ${preview.width} ${preview.height}" role="img" aria-label="${escapeXml(preview.name)} visual preview">
  <defs>
    <pattern id="grid" width="20" height="20" patternUnits="userSpaceOnUse">
      <path d="M 20 0 L 0 0 0 20" fill="none" stroke="#edf4f3" stroke-width="1"/>
    </pattern>
  </defs>
  <rect width="100%" height="100%" fill="#ffffff"/>
  <rect width="100%" height="100%" fill="url(#grid)"/>
  <text x="10" y="18" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="#607178">${escapeXml(preview.name)}</text>
${controls}
</svg>
`;
}
function renderVisualControlSvg(control) {
    const label = truncateLabel(control.label, Math.max(6, Math.floor(control.width / 7)));
    const style = visualControlStyle(control.kind);
    const textY = control.top + Math.max(15, Math.min(control.height - 6, control.height / 2 + 4));
    const textX = control.left + ([
        "checkbox",
        "radio"
    ].includes(control.kind) ? 24 : 8);
    const parts = [
        `  <g data-path="${escapeXml(control.path)}" data-kind="${escapeXml(control.kind)}">`,
        `    <rect x="${control.left}" y="${control.top}" width="${control.width}" height="${control.height}" rx="${style.radius}" fill="${style.fill}" stroke="${style.stroke}" stroke-width="1"/>`
    ];
    if (control.kind === "checkbox" || control.kind === "radio") {
        const markX = control.left + 7;
        const markY = control.top + Math.max(6, control.height / 2 - 6);
        if (control.kind === "radio") parts.push(`    <circle cx="${markX + 6}" cy="${markY + 6}" r="6" fill="#ffffff" stroke="#526168" stroke-width="1"/>`);
        else parts.push(`    <rect x="${markX}" y="${markY}" width="12" height="12" rx="2" fill="#ffffff" stroke="#526168" stroke-width="1"/>`);
    }
    if (control.kind === "combo" || control.kind === "spin") {
        parts.push(`    <rect x="${control.left + control.width - 22}" y="${control.top + 1}" width="21" height="${Math.max(0, control.height - 2)}" fill="#e5ecee" stroke="#c9d5d8" stroke-width="1"/>`);
        parts.push(`    <text x="${control.left + control.width - 15}" y="${textY}" font-family="Segoe UI, Arial, sans-serif" font-size="9" fill="#526168">${control.kind === "spin" ? "+-" : "v"}</text>`);
    }
    if (control.kind === "image") {
        parts.push(`    <path d="M ${control.left + 8} ${control.top + control.height - 8} L ${control.left + control.width / 2} ${control.top + control.height / 2} L ${control.left + control.width - 8} ${control.top + control.height - 8}" fill="none" stroke="#607178" stroke-width="1.5"/>`);
    }
    parts.push(`    <text x="${textX}" y="${textY}" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="${style.text}">${escapeXml(label)}</text>`);
    parts.push("  </g>");
    return parts.join("\n");
}
function renderVisualPreviewHtml(preview) {
    return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(preview.name)} visual preview</title>
  <style>
    body { margin: 24px; background: #eef2f3; color: #172026; font-family: "Segoe UI", Arial, sans-serif; }
    .preview { position: relative; width: ${preview.width}px; height: ${preview.height}px; overflow: hidden; border: 1px solid #c9d5d8; border-radius: 6px; background: linear-gradient(#edf4f3 1px, transparent 1px), linear-gradient(90deg, #edf4f3 1px, transparent 1px), #fff; background-size: 20px 20px; }
    .control { position: absolute; display: flex; align-items: center; gap: 6px; overflow: hidden; padding: 2px 6px; border: 1px solid #1e6b66; border-radius: 4px; background: rgba(217, 235, 232, .9); color: #173532; font-size: 11px; white-space: nowrap; }
    .button { justify-content: center; background: linear-gradient(#f9fbfb, #dce7e9); border-color: #8a9ca1; color: #172026; }
    .label { border-color: transparent; background: transparent; justify-content: flex-start; }
    .textbox, .combo, .listbox { background: #fff; border-color: #8a9ca1; justify-content: flex-start; }
    .frame, .tabs { background: rgba(255,255,255,.38); border-color: #8fb9b4; align-items: flex-start; justify-content: flex-start; }
    .image { background: #f4f7f8; border-style: dashed; color: #607178; }
  </style>
</head>
<body>
  <h1>${escapeHtml(preview.name)}</h1>
  <div class="preview">
    ${preview.controls.map((control)=>`<div class="control ${escapeHtml(control.kind)}" title="${escapeHtml(control.path)}" style="left:${control.left}px;top:${control.top}px;width:${control.width}px;height:${control.height}px">${escapeHtml(control.label)}</div>`).join("\n    ")}
  </div>
</body>
</html>
`;
}
async function renderVisualPreviewPng(preview) {
    const canvas = document.createElement("canvas");
    canvas.width = preview.width * 2;
    canvas.height = preview.height * 2;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas is not available.");
    ctx.scale(2, 2);
    drawVisualPreviewCanvas(ctx, preview);
    return new Promise((resolve, reject)=>{
        canvas.toBlob((blob)=>blob ? resolve(blob) : reject(new Error("Could not render PNG.")), "image/png");
    });
}
function drawVisualPreviewCanvas(ctx, preview) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, preview.width, preview.height);
    ctx.strokeStyle = "#edf4f3";
    ctx.lineWidth = 1;
    for(let x = 0; x <= preview.width; x += 20){
        ctx.beginPath();
        ctx.moveTo(x, 0);
        ctx.lineTo(x, preview.height);
        ctx.stroke();
    }
    for(let y = 0; y <= preview.height; y += 20){
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(preview.width, y);
        ctx.stroke();
    }
    ctx.font = "11px Segoe UI, Arial, sans-serif";
    ctx.fillStyle = "#607178";
    ctx.fillText(preview.name, 10, 18);
    for (const control of preview.controls)drawVisualControlCanvas(ctx, control);
}
function drawVisualControlCanvas(ctx, control) {
    const style = visualControlStyle(control.kind);
    ctx.fillStyle = style.fill;
    ctx.strokeStyle = style.stroke;
    roundedRect(ctx, control.left, control.top, control.width, control.height, style.radius);
    ctx.fill();
    ctx.stroke();
    let textX = control.left + 7;
    if (control.kind === "checkbox" || control.kind === "radio") {
        const markX = control.left + 7;
        const markY = control.top + Math.max(6, control.height / 2 - 6);
        ctx.fillStyle = "#ffffff";
        ctx.strokeStyle = "#526168";
        if (control.kind === "radio") {
            ctx.beginPath();
            ctx.arc(markX + 6, markY + 6, 6, 0, Math.PI * 2);
            ctx.fill();
            ctx.stroke();
        } else {
            roundedRect(ctx, markX, markY, 12, 12, 2);
            ctx.fill();
            ctx.stroke();
        }
        textX += 18;
    }
    if (control.kind === "combo" || control.kind === "spin") {
        ctx.fillStyle = "#e5ecee";
        ctx.fillRect(control.left + control.width - 22, control.top + 1, 21, Math.max(0, control.height - 2));
        ctx.strokeStyle = "#c9d5d8";
        ctx.strokeRect(control.left + control.width - 22, control.top + 1, 21, Math.max(0, control.height - 2));
    }
    ctx.fillStyle = style.text;
    ctx.font = "11px Segoe UI, Arial, sans-serif";
    const label = truncateLabel(control.label, Math.max(4, Math.floor((control.width - (textX - control.left) - 4) / 7)));
    ctx.fillText(label, textX, control.top + Math.max(15, Math.min(control.height - 6, control.height / 2 + 4)));
}
function visualControlStyle(kind) {
    const styles = {
        button: {
            fill: "#eef3f4",
            stroke: "#8a9ca1",
            text: "#172026",
            radius: 4
        },
        label: {
            fill: "rgba(255,255,255,0)",
            stroke: "rgba(255,255,255,0)",
            text: "#172026",
            radius: 0
        },
        textbox: {
            fill: "#ffffff",
            stroke: "#8a9ca1",
            text: "#172026",
            radius: 2
        },
        combo: {
            fill: "#ffffff",
            stroke: "#8a9ca1",
            text: "#172026",
            radius: 2
        },
        listbox: {
            fill: "#ffffff",
            stroke: "#8a9ca1",
            text: "#172026",
            radius: 2
        },
        frame: {
            fill: "rgba(255,255,255,0.52)",
            stroke: "#8fb9b4",
            text: "#526168",
            radius: 3
        },
        tabs: {
            fill: "rgba(255,255,255,0.7)",
            stroke: "#8fb9b4",
            text: "#526168",
            radius: 3
        },
        image: {
            fill: "#f4f7f8",
            stroke: "#8a9ca1",
            text: "#607178",
            radius: 3
        },
        checkbox: {
            fill: "rgba(255,255,255,0)",
            stroke: "rgba(255,255,255,0)",
            text: "#172026",
            radius: 0
        },
        radio: {
            fill: "rgba(255,255,255,0)",
            stroke: "rgba(255,255,255,0)",
            text: "#172026",
            radius: 0
        }
    };
    return styles[kind] ?? {
        fill: "rgba(217,235,232,0.88)",
        stroke: "#1e6b66",
        text: "#173532",
        radius: 4
    };
}
function roundedRect(ctx, x, y, width, height, radius) {
    const r = Math.min(radius, width / 2, height / 2);
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + width - r, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + r);
    ctx.lineTo(x + width, y + height - r);
    ctx.quadraticCurveTo(x + width, y + height, x + width - r, y + height);
    ctx.lineTo(x + r, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
}
function truncateLabel(value, maxLength) {
    if (value.length <= maxLength) return value;
    return `${value.slice(0, Math.max(1, maxLength - 3))}...`;
}
function escapeXml(value) {
    return value.replace(/[<>&"']/g, (char)=>({
            "<": "&lt;",
            ">": "&gt;",
            "&": "&amp;",
            '"': "&quot;",
            "'": "&apos;"
        })[char] ?? char);
}
function renderControlNode(control) {
    const node = document.createElement("details");
    node.className = "control-node";
    node.open = true;
    const summary = document.createElement("summary");
    const label = control.name || control.caption || control.id;
    summary.textContent = `${label}${control.type ? ` · ${control.type}` : ""}`;
    node.append(summary);
    const props = {
        ID: control.id,
        Path: printableStreamName(control.path)
    };
    if (control.progId) props.ProgID = control.progId;
    if (control.caption) props.Caption = control.caption;
    if (control.bounds) props.Bounds = `${control.bounds.left}, ${control.bounds.top}, ${control.bounds.width}, ${control.bounds.height} ${control.bounds.unit} (${control.bounds.confidence})`;
    Object.assign(props, control.properties);
    node.append(renderKeyValueGrid(props));
    for (const child of control.children){
        node.append(renderControlNode(child));
    }
    return node;
}
function countControls(controls) {
    return controls.reduce((sum, control)=>sum + 1 + countControls(control.children), 0);
}


// ---- analysis-ui.ts ----
function renderAnalysis(analysis, bytes) {
    const section = document.createElement("section");
    section.className = "analysis";
    const title = document.createElement("h3");
    title.textContent = analysis.title;
    section.append(title);
    const summary = document.createElement("dl");
    summary.className = "analysis-grid";
    for (const item of analysis.summary){
        const label = document.createElement("dt");
        label.textContent = item.label;
        const value = document.createElement("dd");
        value.textContent = item.value;
        summary.append(label, value);
    }
    section.append(summary);
    section.append(renderAnalysisTable("Detected structures", [
        "Offset",
        "Type",
        "Details"
    ], analysis.signatures.map((item)=>[
            toHex(item.offset),
            item.label,
            item.detail
        ])));
    if (analysis.oforms) {
        section.append(renderOFormsAnalysis(analysis.oforms));
    }
    section.append(renderLazyStrings(analysis, bytes));
    section.append(renderLazyGuids(analysis, bytes));
    const details = document.createElement("details");
    details.className = "hex-details";
    const summaryText = document.createElement("summary");
    summaryText.textContent = "Full hex dump";
    const pre = document.createElement("pre");
    pre.dataset.loaded = "false";
    pre.textContent = "Open to load hex dump.";
    details.addEventListener("toggle", ()=>{
        if (!details.open || pre.dataset.loaded === "true") return;
        pre.textContent = bytes.length ? hexDump(bytes, 0, bytes.length) : "(empty stream)";
        pre.dataset.loaded = "true";
    }, {
        once: false
    });
    details.append(summaryText, pre);
    section.append(details);
    return section;
}
function renderOFormsAnalysis(oforms) {
    const wrapper = document.createElement("section");
    wrapper.className = "oforms-analysis";
    const heading = document.createElement("h4");
    heading.textContent = `MS-OFORMS parse: ${oforms.kind}`;
    wrapper.append(heading);
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    for (const header of [
        "Offset",
        "Record",
        "Size",
        "Properties"
    ]){
        const cell = document.createElement("th");
        cell.textContent = header;
        headerRow.append(cell);
    }
    thead.append(headerRow);
    table.append(thead);
    const tbody = document.createElement("tbody");
    for (const record of oforms.records){
        const row = document.createElement("tr");
        for (const value of [
            toHex(record.offset),
            record.type,
            formatBytes(record.size)
        ]){
            const cell = document.createElement("td");
            cell.textContent = value;
            row.append(cell);
        }
        const propertiesCell = document.createElement("td");
        const list = document.createElement("dl");
        list.className = "property-list";
        for (const property of record.properties){
            const label = document.createElement("dt");
            label.textContent = property.name;
            const value = document.createElement("dd");
            value.textContent = labelPossibleIdentifier(property.value);
            list.append(label, value);
        }
        propertiesCell.append(list);
        row.append(propertiesCell);
        tbody.append(row);
    }
    table.append(tbody);
    wrapper.append(table);
    if (oforms.notes.length) {
        const notes = document.createElement("ul");
        notes.className = "analysis-notes";
        for (const note of oforms.notes){
            const item = document.createElement("li");
            item.textContent = note;
            notes.append(item);
        }
        wrapper.append(notes);
    }
    return wrapper;
}
function renderLazyGuids(analysis, bytes) {
    const wrapper = document.createElement("div");
    wrapper.className = "analysis-table-wrap";
    const heading = document.createElement("h4");
    heading.textContent = `GUID / CLSID candidates (${analysis.guidCount})`;
    wrapper.append(heading);
    if (analysis.guidCount === 0) {
        const empty = document.createElement("p");
        empty.className = "analysis-empty";
        empty.textContent = "No high-confidence values found.";
        wrapper.append(empty);
        return wrapper;
    }
    const details = document.createElement("details");
    details.className = "lazy-details";
    const summary = document.createElement("summary");
    summary.textContent = `Show all ${analysis.guidCount} GUID / CLSID candidates`;
    const container = document.createElement("div");
    container.dataset.loaded = "false";
    container.className = "lazy-table-container";
    container.textContent = "Open to load all GUID / CLSID candidates.";
    details.addEventListener("toggle", ()=>{
        if (!details.open || container.dataset.loaded === "true") return;
        container.textContent = "";
        container.append(renderAnalysisTable("", [
            "Offset",
            "Value",
            "Label"
        ], collectGuids(bytes).map((item)=>[
                toHex(item.offset),
                item.value,
                labelGuid(item.value) ?? ""
            ])));
        container.dataset.loaded = "true";
    });
    details.append(summary, container);
    wrapper.append(details);
    return wrapper;
}
function renderLazyStrings(analysis, bytes) {
    const wrapper = document.createElement("div");
    wrapper.className = "analysis-table-wrap";
    const heading = document.createElement("h4");
    heading.textContent = `Strings (${analysis.stringCount})`;
    wrapper.append(heading);
    if (analysis.stringCount === 0) {
        const empty = document.createElement("p");
        empty.className = "analysis-empty";
        empty.textContent = "No high-confidence values found.";
        wrapper.append(empty);
        return wrapper;
    }
    const details = document.createElement("details");
    details.className = "lazy-details";
    const summary = document.createElement("summary");
    summary.textContent = `Show all ${analysis.stringCount} strings`;
    const container = document.createElement("div");
    container.dataset.loaded = "false";
    container.className = "lazy-table-container";
    container.textContent = "Open to load all strings.";
    details.addEventListener("toggle", ()=>{
        if (!details.open || container.dataset.loaded === "true") return;
        container.textContent = "";
        container.append(renderStringTable(collectStrings(bytes)));
        container.dataset.loaded = "true";
    });
    details.append(summary, container);
    wrapper.append(details);
    return wrapper;
}
function renderStringTable(strings) {
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    for (const header of [
        "Offset",
        "Encoding",
        "Value"
    ]){
        const cell = document.createElement("th");
        cell.textContent = header;
        headerRow.append(cell);
    }
    thead.append(headerRow);
    table.append(thead);
    const tbody = document.createElement("tbody");
    for (const item of strings){
        const row = document.createElement("tr");
        for (const value of [
            toHex(item.offset),
            item.encoding,
            item.value
        ]){
            const cell = document.createElement("td");
            cell.textContent = value;
            row.append(cell);
        }
        tbody.append(row);
    }
    table.append(tbody);
    return table;
}
function renderAnalysisTable(title, headers, rows) {
    const wrapper = document.createElement("div");
    wrapper.className = "analysis-table-wrap";
    const heading = document.createElement("h4");
    heading.textContent = title;
    wrapper.append(heading);
    if (rows.length === 0) {
        const empty = document.createElement("p");
        empty.className = "analysis-empty";
        empty.textContent = "No high-confidence values found.";
        wrapper.append(empty);
        return wrapper;
    }
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    for (const header of headers){
        const cell = document.createElement("th");
        cell.textContent = header;
        headerRow.append(cell);
    }
    thead.append(headerRow);
    table.append(thead);
    const tbody = document.createElement("tbody");
    for (const row of rows){
        const tableRow = document.createElement("tr");
        for (const value of row){
            const cell = document.createElement("td");
            cell.textContent = value;
            tableRow.append(cell);
        }
        tbody.append(tableRow);
    }
    table.append(tbody);
    wrapper.append(table);
    return wrapper;
}


// ---- runtime.ts ----
function downloadFile(file) {
    downloadBytes(file.name, file.mimeType, file.bytes);
}
function downloadBytes(fileName, mimeType, bytes) {
    downloadBlob(fileName, new Blob([
        bytes
    ], {
        type: mimeType
    }));
}
function downloadBlob(fileName, blob) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
}
function downloadAll(files) {
    if (files.length === 0) return;
    const zipBytes = createZip(files);
    const blob = new Blob([
        zipBytes
    ], {
        type: "application/zip"
    });
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
function setStatus(message, state) {
    statusText.textContent = message;
    statusPanel.dataset.state = state;
}
function findEndOfCentralDirectory(bytes) {
    for(let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset -= 1){
        if (readU32(bytes, offset) === 0x06054b50) return offset;
    }
    return -1;
}
function isZip(bytes) {
    return readU32(bytes, 0) === 0x04034b50;
}
function isOle(bytes) {
    return matches(bytes, 0, [
        0xd0,
        0xcf,
        0x11,
        0xe0,
        0xa1,
        0xb1,
        0x1a,
        0xe1
    ]);
}
function getSector(bytes, sectorSize, sectorIndex) {
    const offset = (sectorIndex + 1) * sectorSize;
    return bytes.subarray(offset, offset + sectorSize);
}
function concatBytes(chunks) {
    const total = chunks.reduce((sum, chunk)=>sum + chunk.byteLength, 0);
    const combined = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks){
        combined.set(chunk, offset);
        offset += chunk.byteLength;
    }
    return combined;
}
function readU16(bytes, offset) {
    return bytes[offset] | bytes[offset + 1] << 8;
}
function readU32(bytes, offset) {
    return (bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24) >>> 0;
}
function readI32(bytes, offset) {
    return readU32(bytes, offset) | 0;
}
function writeU16(bytes, offset, value) {
    bytes[offset] = value & 0xff;
    bytes[offset + 1] = value >>> 8 & 0xff;
}
function writeU32(bytes, offset, value) {
    bytes[offset] = value & 0xff;
    bytes[offset + 1] = value >>> 8 & 0xff;
    bytes[offset + 2] = value >>> 16 & 0xff;
    bytes[offset + 3] = value >>> 24 & 0xff;
}
function decodeText(bytes) {
    try {
        return new TextDecoder("windows-1252").decode(bytes);
    } catch  {
        return new TextDecoder().decode(bytes);
    }
}
function encodeText(text) {
    return new TextEncoder().encode(text);
}
function decodeUtf16Le(bytes) {
    let text = "";
    for(let offset = 0; offset + 1 < bytes.length; offset += 2){
        const code = bytes[offset] | bytes[offset + 1] << 8;
        if (code) text += String.fromCharCode(code);
    }
    return text;
}
function matches(bytes, offset, signature) {
    return signature.every((byte, index)=>bytes[offset + index] === byte);
}
function matchesAscii(bytes, offset, value) {
    return [
        ...value
    ].every((char, index)=>bytes[offset + index] === char.charCodeAt(0));
}
function findBytes(bytes, pattern, start) {
    for(let offset = start; offset <= bytes.length - pattern.length; offset += 1){
        if (matches(bytes, offset, pattern)) return offset;
    }
    return -1;
}
function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
function toHex(value, width = 4) {
    return `0x${value.toString(16).toUpperCase().padStart(width, "0")}`;
}
function formatGuid(bytes) {
    const hex = [
        ...bytes
    ].map((byte)=>byte.toString(16).padStart(2, "0"));
    return `{${[
        hex.slice(0, 4).reverse().join(""),
        hex.slice(4, 6).reverse().join(""),
        hex.slice(6, 8).reverse().join(""),
        hex.slice(8, 10).join(""),
        hex.slice(10, 16).join("")
    ].join("-").toUpperCase()}}`;
}
function safeFileName(value) {
    return value.replace(/[^a-z0-9._-]+/gi, "_").replace(/^_+|_+$/g, "").slice(0, 180) || "stream";
}
function escapeHtml(value) {
    return value.replace(/[&<>"']/g, (char)=>{
        const map = {
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#39;"
        };
        return map[char];
    });
}
function kindRank(kind) {
    return ({
        frm: 0,
        vba: 1,
        media: 2,
        frx: 3,
        binary: 4
    })[kind];
}

