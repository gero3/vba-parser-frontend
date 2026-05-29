function buildApplicationModel(files: ExtractedFile[], groups = buildResultGroups(files)): ApplicationModel {
  const modules = files
    .filter((file) => file.text && (file.kind === "vba" || file.kind === "frm"))
    .map((file) => buildModuleModel(file));
  const projectReferences = parseExtractedProjectReferences(files);
  const forms = groups
    .filter((group) => group.designer)
    .map((group) => buildFormModel(group.designer!));
  const documentControls = buildActiveXControls(files);

  enrichProcedureCalls(modules, files);
  linkEventsToControls(modules, forms);
  addInferredEventControls(modules, forms);
  linkEventsToControls(modules, forms);
  enrichModuleReferences(modules, files);

  return {
    generatedAt: new Date().toISOString(),
    sourceFiles: files.map((file) => ({
      name: file.name,
      kind: file.kind,
      path: getZipPath(file, files.indexOf(file)),
      size: file.bytes.byteLength,
    })),
    modules,
    projectReferences,
    forms,
    documentControls,
    dependencies: uniqueDependencies([
      ...detectDependencies(modules, files),
      ...projectReferences.map((reference) => ({
        category: "VBA reference",
        value: reference.name ?? reference.libId ?? reference.raw,
        source: reference.source,
        reason: `${reference.kind} project reference`,
      })),
    ]),
    assets: files
      .filter((file) => file.kind === "media")
      .map((file) => ({ name: file.name, mimeType: file.mimeType, size: file.bytes.byteLength, sourcePath: file.sourcePath })),
    migrationNotes: [
      "VBA code is extracted verbatim. Business logic should be reviewed before translation because Office object model calls often imply workbook state.",
      "Form/control layout is reconstructed from MS-OFORMS and VBFrame data where available; raw FRX streams remain in the archive for audit.",
      "Event links are inferred from VBA naming conventions such as ControlName_Click.",
      "Dependencies are heuristic detections intended to help an LLM identify integrations, not a security verdict.",
    ],
  };
}

function parseExtractedProjectReferences(files: ExtractedFile[]): ProjectReferenceModel[] {
  const referenceFile = files.find((file) => file.name === "vba-project-references.json" && file.text);
  if (!referenceFile?.text) return [];
  try {
    return JSON.parse(referenceFile.text) as ProjectReferenceModel[];
  } catch {
    return [];
  }
}

function buildModuleModel(file: ExtractedFile): ModuleModel {
  const source = file.text ?? "";
  const procedures = parseProcedures(source);
  const declarations = source
    .split(/\r?\n/)
    .filter((line) => /^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\b/i.test(line))
    .map((line) => line.trim());
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
    events: procedures.map(procedureToEvent).filter((event): event is EventModel => Boolean(event)),
    references: [],
    riskMarkers: detectRiskMarkers(source),
  };
}

function parseProcedures(source: string): ProcedureModel[] {
  const lines = source.split(/\r?\n/);
  const procedures: ProcedureModel[] = [];
  const startRegex = /^\s*(?:(Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\([^)]*\))?\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
  const endRegex = /^\s*End\s+(Sub|Function|Property)\b/i;
  let current: ProcedureModel | undefined;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const start = line.match(startRegex);
    if (start) {
      current = {
        name: start[4],
        kind: start[3].startsWith("Property") ? "Property" : start[3] as "Sub" | "Function",
        scope: (start[1] as ProcedureModel["scope"]) ?? "Implicit",
        signature: line.trim(),
        parameters: parseParameters(start[5] ?? ""),
        returnType: start[6],
        lineStart: index + 1,
        calls: [],
        uses: [],
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

function parseParameters(parameterList: string) {
  return parameterList
    .replace(/^\(|\)$/g, "")
    .split(",")
    .map((parameter) => parameter.trim())
    .filter(Boolean);
}

function parseVariableDeclarations(source: string): DeclarationModel[] {
  const declarations: DeclarationModel[] = [];
  const lines = source.split(/\r?\n/);
  const regex = /^\s*(Public|Private|Friend|Dim|Static|Global)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (/\b(Sub|Function|Property)\b/i.test(line)) continue;
    const match = line.match(regex);
    if (!match || /^Declare$/i.test(match[2])) continue;
    declarations.push({
      name: match[2],
      scope: match[1],
      type: match[3],
      line: index + 1,
      statement: line.trim(),
    });
  }
  return declarations;
}

function parseConstantDeclarations(source: string): DeclarationModel[] {
  const declarations: DeclarationModel[] = [];
  const lines = source.split(/\r?\n/);
  const regex = /^\s*(Public|Private|Friend)?\s*Const\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?/i;
  for (let index = 0; index < lines.length; index += 1) {
    const match = lines[index].match(regex);
    if (!match) continue;
    declarations.push({
      name: match[2],
      scope: match[1] ?? "Implicit",
      type: match[3],
      line: index + 1,
      statement: lines[index].trim(),
    });
  }
  return declarations;
}

function procedureToEvent(procedure: ProcedureModel): EventModel | undefined {
  const match = procedure.name.match(/^(.+)_([A-Za-z][A-Za-z0-9]*)$/);
  if (!match) return undefined;
  return {
    procedure: procedure.name,
    controlName: match[1],
    eventName: match[2],
  };
}

function enrichProcedureCalls(modules: ModuleModel[], files: ExtractedFile[]) {
  const procedureNames = new Set(modules.flatMap((module) => module.procedures.map((procedure) => procedure.name)));

  for (const module of modules) {
    const source = files.find((file) => file.name === module.fileName)?.text ?? "";
    const lines = source.split(/\r?\n/);
    for (const procedure of module.procedures) {
      const body = lines.slice(procedure.lineStart - 1, procedure.lineEnd ?? procedure.lineStart + 80).join("\n");
      procedure.calls = [...procedureNames]
        .filter((name) => name !== procedure.name && new RegExp(`\\b${escapeRegExp(name)}\\b`, "i").test(body))
        .sort((a, b) => a.localeCompare(b));
      procedure.uses = detectProcedureUses(body);
    }
  }
}

function enrichModuleReferences(modules: ModuleModel[], files: ExtractedFile[]) {
  for (const module of modules) {
    const source = files.find((file) => file.name === module.fileName)?.text ?? "";
    module.references = modules
      .filter((candidate) => candidate.name !== module.name)
      .filter((candidate) => new RegExp(`\\b${escapeRegExp(candidate.name)}\\b`, "i").test(source))
      .map((candidate) => candidate.name)
      .sort((a, b) => a.localeCompare(b));
  }
}

function detectRiskMarkers(source: string): RiskMarker[] {
  const markers: RiskMarker[] = [];
  const checks: Array<{ category: string; reason: string; regex: RegExp }> = [
    { category: "TODO", reason: "Developer note may indicate incomplete or special-case logic.", regex: /\b(TODO|FIXME|HACK|XXX)\b/i },
    { category: "Error handling", reason: "Broad error handling can hide behavior that needs explicit frontend states.", regex: /\bOn\s+Error\s+Resume\s+Next\b/i },
    { category: "Global state", reason: "Global/public mutable state should become explicit app state.", regex: /^\s*(Public|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s+/i },
    { category: "Dynamic execution", reason: "Dynamic execution or indirection may require manual migration review.", regex: /\b(Application\.Run|CallByName|Evaluate\s*\()/i },
    { category: "External process", reason: "External process execution cannot be directly migrated to browser frontend code.", regex: /\b(Shell|WScript\.Shell|CreateObject\s*\(\s*"WScript\.Shell")\b/i },
    { category: "File system", reason: "Local file access needs replacement with browser file APIs or backend services.", regex: /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Kill\s+|MkDir\s+|RmDir\s+)\b/i },
    { category: "WinAPI", reason: "WinAPI calls do not run in browser frontend code.", regex: /\bDeclare\s+(PtrSafe\s+)?(Sub|Function)\b/i },
  ];

  source.split(/\r?\n/).forEach((line, index) => {
    for (const check of checks) {
      if (check.regex.test(line)) {
        markers.push({ category: check.category, line: index + 1, text: line.trim(), reason: check.reason });
      }
    }
  });

  return markers;
}

function detectProcedureUses(body: string) {
  const uses = new Set<string>();
  const checks: Array<[string, RegExp]> = [
    ["Excel object model", /\b(Application|Workbook|Worksheet|Range|Cells|Rows|Columns|Sheets|ActiveCell|Selection)\b/i],
    ["UserForm/UI state", /\b(Me\.|Controls\(|\.Caption|\.Value|\.Visible|\.Enabled|\.ListIndex|\.AddItem|\.Clear)\b/i],
    ["File system", /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Dir\s*\(|Kill\s+|MkDir\s+|RmDir\s+)\b/i],
    ["HTTP/network", /\b(XMLHTTP|WinHttp|ServerXMLHTTP|WebRequest|URLDownloadToFile)\b/i],
    ["Database/ADO", /\b(ADODB|DAO\.|Recordset|Connection|OpenRecordset)\b/i],
    ["Shell/process", /\b(Shell|WScript\.Shell|Run\s*\()\b/i],
    ["WinAPI", /\b(Declare\s+(PtrSafe\s+)?(Sub|Function)|AddressOf|LongPtr|LongLong)\b/i],
    ["Error handling", /\b(On\s+Error|Err\.)\b/i],
  ];

  for (const [label, regex] of checks) {
    if (regex.test(body)) uses.add(label);
  }
  return [...uses].sort((a, b) => a.localeCompare(b));
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function buildFormModel(summary: DesignerSummary): FormModel {
  return {
    name: summary.formName,
    properties: summary.frame,
    controls: flattenControls(summary.controls),
  };
}

function buildActiveXControls(files: ExtractedFile[]): ActiveXControlModel[] {
  const activeXXmlFiles = files.filter((file) => /\/activeX\/activeX\d+\.xml$/i.test(file.sourcePath) && file.text);
  const activeXBinFiles = files.filter((file) => /\/activeX\/activeX\d+\.bin$/i.test(file.sourcePath));

  return activeXXmlFiles.map((xmlFile) => {
    const base = xmlFile.sourcePath.match(/activeX(\d+)\.xml$/i)?.[1] ?? String(activeXXmlFiles.indexOf(xmlFile) + 1);
    const binFile = activeXBinFiles.find((file) => new RegExp(`activeX${base}\\.bin$`, "i").test(file.sourcePath));
    const properties = parseActiveXXmlProperties(xmlFile.text ?? "");
    const classId = properties.classid ?? properties.classId ?? properties.clsid;
    const normalizedClassId = normalizeGuidString(classId);
    const persistenceAnalysis = parseActiveXBinPersistence(binFile);
    const label = persistenceAnalysis?.compObj?.userType ?? (normalizedClassId ? labelGuid(normalizedClassId) : undefined);
    const persistenceProperties = Object.fromEntries(
      (persistenceAnalysis?.properties ?? []).map((property) => [`bin.${property.name}`, property.value]),
    );

    return {
      id: `activeX${base}`,
      xmlPath: xmlFile.sourcePath,
      binPath: binFile?.sourcePath,
      classId: normalizedClassId ?? classId,
      label,
      persistence: properties.persistence ?? properties.persistStorage,
      properties: { ...properties, ...persistenceProperties },
      persistenceAnalysis,
      sourceFiles: [xmlFile.sourcePath, binFile?.sourcePath].filter((value): value is string => Boolean(value)),
    };
  });
}

function parseActiveXBinPersistence(binFile: ExtractedFile | undefined): ActiveXPersistenceAnalysis | undefined {
  if (!binFile) return undefined;
  const bytes = binFile.bytes;
  const media = extractMediaFromBinary(bytes, safeFileName(binFile.name)).map((file) => ({
    name: file.name,
    mimeType: file.mimeType,
    size: file.bytes.byteLength,
    sourcePath: file.sourcePath,
  }));

  if (bytes.byteLength === 0) {
    return { format: "empty", confidence: "high", size: 0, streams: [], properties: [], media, warnings: ["ActiveX binary part is empty."] };
  }

  if (!isOle(bytes)) {
    return {
      format: "raw-stream",
      confidence: "medium",
      size: bytes.byteLength,
      streams: [summarizeActiveXPersistenceStream("(raw)", bytes)],
      properties: inferActiveXContentsProperties(bytes, "(raw)", undefined),
      media,
      warnings: ["Binary persistence part is not an OLE/CFB storage; parsed as a raw persistStream payload."],
    };
  }

  try {
    const cfb = parseCfb(bytes);
    const streams = [...cfb.streams].map(([path, streamBytes]) => summarizeActiveXPersistenceStream(path, streamBytes));
    const compObjEntry = [...cfb.streams].find(([path]) => /compobj/i.test(path));
    const contentsEntry = [...cfb.streams].find(([path]) => /(^|\/)contents$/i.test(printableStreamName(path)));
    const compObj = compObjEntry ? parseActiveXCompObjStream(compObjEntry[1]) : undefined;
    const properties = [
      ...(compObj ? activeXCompObjProperties(compObj, compObjEntry?.[0] ?? "CompObj") : []),
      ...(contentsEntry ? inferActiveXContentsProperties(contentsEntry[1], contentsEntry[0], compObj?.userType) : []),
    ];

    return {
      format: "cfb-storage",
      confidence: compObj || contentsEntry ? "high" : "medium",
      size: bytes.byteLength,
      streams,
      compObj,
      properties,
      media,
      warnings: contentsEntry ? [] : ["No contents stream was found in the ActiveX storage."],
    };
  } catch (error) {
    return {
      format: "unknown",
      confidence: "low",
      size: bytes.byteLength,
      streams: [summarizeActiveXPersistenceStream("(unparsed)", bytes)],
      properties: [],
      media,
      warnings: [`Could not parse ActiveX OLE storage: ${error instanceof Error ? error.message : String(error)}`],
    };
  }
}

function summarizeActiveXPersistenceStream(path: string, bytes: Uint8Array): ActiveXPersistenceStream {
  const strings = collectStrings(bytes).slice(0, 16).map((item) => labelPossibleIdentifier(item.value));
  const guids = collectGuids(bytes).slice(0, 16).map((item) => {
    const label = labelGuid(item.value);
    return label ? `${item.value} (${label})` : item.value;
  });
  const signatures = collectSignatures(bytes).slice(0, 12).map((signature) => `${signature.label} at ${toHex(signature.offset)}${signature.detail ? ` (${signature.detail})` : ""}`);
  const oforms = parseOFormsStream(bytes, path, collectStrings(bytes));

  return {
    path: printableStreamName(path),
    size: bytes.byteLength,
    strings,
    guids,
    signatures,
    oforms: oforms ? `${oforms.kind} (${oforms.confidence})` : undefined,
  };
}

function parseActiveXCompObjStream(bytes: Uint8Array): ActiveXPersistenceAnalysis["compObj"] {
  const strings = collectStrings(bytes).map((item) => item.value).filter(Boolean);
  const clsid = bytes.byteLength >= 28 ? formatGuid(bytes.subarray(12, 28)) : undefined;
  return {
    clsid,
    userType: strings.find((value) => /Microsoft Forms|Control|Button|Box|Label|Image|Object/i.test(value)),
    clipboardFormat: strings.find((value) => /Embedded Object|Object|Control/i.test(value) && !/Microsoft Forms/i.test(value)),
    progId: strings.find((value) => /^[A-Za-z][A-Za-z0-9_.]+(\.\d+)?$/.test(value)),
  };
}

function activeXCompObjProperties(compObj: NonNullable<ActiveXPersistenceAnalysis["compObj"]>, source: string): ActiveXPersistenceAnalysis["properties"] {
  const properties: ActiveXPersistenceAnalysis["properties"] = [];
  if (compObj.clsid) {
    properties.push({
      name: "CompObj CLSID",
      value: labelPossibleIdentifier(compObj.clsid),
      source,
      confidence: "high",
    });
  }
  if (compObj.userType) properties.push({ name: "UserType", value: compObj.userType, source, confidence: "high" });
  if (compObj.clipboardFormat) properties.push({ name: "ClipboardFormat", value: compObj.clipboardFormat, source, confidence: "medium" });
  if (compObj.progId) properties.push({ name: "ProgID", value: labelPossibleIdentifier(compObj.progId), source, confidence: "high" });
  return properties;
}

function inferActiveXContentsProperties(bytes: Uint8Array, source: string, userType: string | undefined): ActiveXPersistenceAnalysis["properties"] {
  const properties: ActiveXPersistenceAnalysis["properties"] = [];
  const strings = collectStrings(bytes).filter((item) => item.value.length > 1);
  if (bytes.byteLength >= 8) {
    const minor = bytes[0];
    const major = bytes[1];
    const recordSize = readU16(bytes, 2);
    const propMask = readU32(bytes, 4);
    if (minor === 0 && [2, 4].includes(major) && recordSize >= 4 && recordSize <= bytes.byteLength) {
      properties.push({ name: "PersistedRecordVersion", value: `${major}.${minor}`, source, confidence: "medium" });
      properties.push({ name: "PersistedRecordSize", value: `${recordSize} bytes`, source, confidence: "medium" });
      properties.push({ name: "PersistedPropMask", value: `${toHex(propMask, 8)} (${describeActiveXPropertyMask(propMask, userType).join(", ") || "unknown flags"})`, source, confidence: "medium" });
    }
  }

  const usefulStrings = strings.filter((item) => isUsefulDecodedPropertyString(item.value));
  const caption = usefulStrings.find((item) => isHumanLabel(item.value));
  if (caption) properties.push({ name: "Caption", value: caption.value, source: `${source}@${toHex(caption.offset)}`, confidence: "high" });

  const font = usefulStrings.find((item) => /^(Aptos|Arial|Calibri|Cambria|Courier New|MS Sans Serif|Segoe UI|Tahoma|Times New Roman|Verdana)$/i.test(item.value));
  if (font) properties.push({ name: "FontName", value: font.value, source: `${source}@${toHex(font.offset)}`, confidence: "high" });

  for (const color of findOleColorCandidates(bytes).slice(0, 8)) {
    properties.push({ name: color.name, value: color.value, source: `${source}@${toHex(color.offset)}`, confidence: color.confidence });
  }

  const extraStrings = usefulStrings
    .filter((item) => item !== caption && item !== font)
    .slice(0, 10);
  for (const item of extraStrings) {
    properties.push({ name: "StringProperty", value: labelPossibleIdentifier(item.value), source: `${source}@${toHex(item.offset)}`, confidence: "low" });
  }

  return properties;
}

function describeActiveXPropertyMask(propMask: number, userType: string | undefined) {
  const type = userType?.toLowerCase() ?? "";
  const generic = [
    { bit: 2, name: "Color/visual property" },
    { bit: 3, name: "Caption/string data" },
    { bit: 5, name: "Font/property data" },
    { bit: 7, name: "Enabled/visibility state" },
    { bit: 10, name: "Picture/icon data" },
    { bit: 11, name: "Mouse/icon data" },
  ];
  if (type.includes("commandbutton")) {
    generic[1] = { bit: 3, name: "Caption" };
    generic[2] = { bit: 5, name: "Font" };
  }
  return generic.filter((flag) => (propMask & (1 << flag.bit)) !== 0).map((flag) => flag.name);
}

function findOleColorCandidates(bytes: Uint8Array) {
  const colors: Array<{ offset: number; name: string; value: string; confidence: "medium" | "low" }> = [];
  for (let offset = 0; offset + 4 <= bytes.length; offset += 4) {
    const value = readU32(bytes, offset);
    const high = value >>> 24;
    if (high === 0x80) {
      colors.push({ offset, name: "SystemColor", value: `${toHex(value, 8)} (${describeSystemOleColor(value & 0xff)})`, confidence: "medium" });
    } else if (high === 0 && value !== 0 && value <= 0x00ffffff) {
      colors.push({ offset, name: "RgbColor", value: `${toHex(value, 8)} (#${(value & 0xffffff).toString(16).padStart(6, "0")})`, confidence: "low" });
    }
  }
  return colors;
}

function describeSystemOleColor(index: number) {
  const labels: Record<number, string> = {
    0x00: "scrollbar",
    0x05: "window background",
    0x08: "window text",
    0x0f: "button face",
    0x12: "button text",
    0x14: "highlight text",
    0x15: "button shadow",
  };
  return labels[index] ?? `system color ${index}`;
}

function parseActiveXXmlProperties(xml: string) {
  const properties: Record<string, string> = {};
  const attrRegex = /(?:^|\s)([A-Za-z_:][A-Za-z0-9_.:-]*)\s*=\s*"([^"]*)"/g;
  let match: RegExpExecArray | null;
  while ((match = attrRegex.exec(xml))) {
    const key = match[1].replace(/^.*:/, "");
    properties[key] = labelPossibleIdentifier(match[2]);
  }

  const tagRegex = /<([A-Za-z_:][A-Za-z0-9_.:-]*)[^>]*>([^<]+)<\/\1>/g;
  while ((match = tagRegex.exec(xml))) {
    const key = match[1].replace(/^.*:/, "");
    properties[key] = labelPossibleIdentifier(match[2].trim());
  }

  const relationshipId = xml.match(/r:id="([^"]+)"/i)?.[1];
  if (relationshipId) properties.relationshipId = relationshipId;

  return properties;
}

function normalizeGuidString(value: string | undefined) {
  const match = value?.match(/\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?/i);
  if (!match) return undefined;
  const raw = match[0].replace(/[{}]/g, "").toUpperCase();
  return `{${raw}}`;
}

function flattenControls(controls: DesignerControl[], parentPath?: string): FlatControlModel[] {
  return controls.flatMap((control) => [
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
      parentPath,
    },
    ...flattenControls(control.children, control.path),
  ]);
}

function linkEventsToControls(modules: ModuleModel[], forms: FormModel[]) {
  const controls = forms.flatMap((form) => form.controls.map((control) => ({ form, control })));
  for (const module of modules) {
    for (const event of module.events) {
      const linked = controls.find(({ form, control }) =>
        form.name === module.name &&
        [control.name, control.caption, control.id, ...Object.values(control.properties)].filter(Boolean).some((value) => normalizeIdentifier(String(value)) === normalizeIdentifier(event.controlName)),
      );
      if (linked) {
        event.linkedControlPath = linked.control.path;
        event.linkedControlType = linked.control.type;
      }
    }
  }
}

function addInferredEventControls(modules: ModuleModel[], forms: FormModel[]) {
  for (const form of forms) {
    const module = modules.find((candidate) => candidate.name === form.name);
    if (!module) continue;

    for (const event of module.events) {
      const exists = form.controls.some((control) =>
        [control.name, control.caption, control.id, ...Object.values(control.properties)].filter(Boolean).some((value) => normalizeIdentifier(String(value)) === normalizeIdentifier(event.controlName)),
      );
      if (exists) continue;

      form.controls.push({
        id: event.controlName,
        path: `${form.name}/inferred/${event.controlName}`,
        name: event.controlName,
        type: inferControlTypeFromName(event.controlName),
        properties: {
          InferredFromEvent: event.procedure,
          Confidence: "event-name heuristic",
        },
        sourceStreams: [],
      });
    }
  }
}

function inferControlTypeFromName(name: string) {
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

function normalizeIdentifier(value: string) {
  return value.replace(/\s+\(.+\)$/, "").replace(/[^a-z0-9_]/gi, "").toLowerCase();
}

function detectDependencies(modules: ModuleModel[], files: ExtractedFile[]): DependencyModel[] {
  const dependencies: DependencyModel[] = [];
  const patterns: Array<{ category: string; reason: string; regex: RegExp }> = [
    { category: "WinAPI", reason: "Declare statement", regex: /\bDeclare\s+(PtrSafe\s+)?(Sub|Function)\s+([A-Za-z0-9_]+)/ig },
    { category: "COM automation", reason: "CreateObject call", regex: /\bCreateObject\s*\(\s*"([^"]+)"/ig },
    { category: "COM automation", reason: "GetObject call", regex: /\bGetObject\s*\([^)]*"([^"]+)"/ig },
    { category: "File system", reason: "FileSystemObject or file IO usage", regex: /\b(FileSystemObject|Open\s+[^\r\n]+For\s+|Kill\s+|MkDir\s+|RmDir\s+|Dir\s*\()/ig },
    { category: "Shell/process", reason: "Shell execution", regex: /\b(Shell|WScript\.Shell|Run\s*\()/ig },
    { category: "HTTP/network", reason: "HTTP client usage", regex: /\b(XMLHTTP|WinHttp|ServerXMLHTTP|WebRequest|InternetOpen|URLDownloadToFile)\b/ig },
    { category: "Database", reason: "Database/ADO usage", regex: /\b(ADODB|DAO\.|Recordset|ConnectionString|OpenRecordset)\b/ig },
    { category: "Office automation", reason: "Office object model usage", regex: /\b(Excel\.|Word\.|Outlook\.|PowerPoint\.|Application\.|Workbook|Worksheet|Range\(|Cells\()\b/ig },
    { category: "Registry", reason: "Registry access", regex: /\b(RegRead|RegWrite|RegDelete|GetSetting|SaveSetting|DeleteSetting)\b/ig },
  ];

  for (const module of modules) {
    const source = module.procedures.map((procedure) => procedure.signature).join("\n") + "\n" + module.declarations.join("\n");
    const fullSource = files.find((file) => file.name === module.fileName)?.text ?? source;
    for (const pattern of patterns) {
      for (const match of fullSource.matchAll(pattern.regex)) {
        const value = match[1] || match[3] || match[0];
        dependencies.push({ category: pattern.category, value: value.trim(), source: module.fileName, reason: pattern.reason });
      }
    }
  }

  return uniqueDependencies(dependencies);
}

function uniqueDependencies(dependencies: DependencyModel[]) {
  const seen = new Set<string>();
  return dependencies.filter((dependency) => {
    const key = `${dependency.category}:${dependency.value}:${dependency.source}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function renderApplicationSummary(model: ApplicationModel) {
  const lines: string[] = [
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
    `- Controls: ${model.forms.reduce((sum, form) => sum + form.controls.length, 0)}`,
    `- Document ActiveX controls: ${model.documentControls.length}`,
    `- Procedures: ${model.modules.reduce((sum, module) => sum + module.procedures.length, 0)}`,
    `- Dependencies: ${model.dependencies.length}`,
    `- Assets: ${model.assets.length}`,
    "",
    "## Forms",
    "",
  ];

  for (const form of model.forms) {
    lines.push(`### ${form.name}`, "");
    for (const [key, value] of Object.entries(form.properties).slice(0, 12)) lines.push(`- ${key}: ${value}`);
    if (form.controls.length) {
      lines.push("- Controls:");
      for (const control of form.controls) {
        lines.push(`  - ${control.name || control.caption || control.id}: ${control.type ?? control.progId ?? "unknown control"} (${control.path})`);
      }
    }
    lines.push("");
  }

  if (model.documentControls.length) {
    lines.push("## Document ActiveX Controls", "");
    for (const control of model.documentControls) {
      lines.push(`- ${control.id}: ${control.label ?? control.classId ?? "Unknown ActiveX control"}`);
      lines.push(`  - XML: ${control.xmlPath}`);
      if (control.binPath) lines.push(`  - Binary persistence: ${control.binPath}`);
      if (control.persistenceAnalysis?.compObj?.progId) lines.push(`  - ProgID: ${control.persistenceAnalysis.compObj.progId}`);
      const caption = control.persistenceAnalysis?.properties.find((property) => property.name === "Caption");
      if (caption) lines.push(`  - Caption: ${caption.value}`);
    }
    lines.push("");
  }

  if (model.projectReferences.length) {
    lines.push("## VBA Project References", "");
    for (const reference of model.projectReferences) {
      lines.push(`- ${reference.name ?? reference.kind}: ${reference.libId ?? reference.raw}`);
      if (reference.guid) lines.push(`  - GUID: ${reference.guid}${labelGuid(reference.guid) ? ` (${labelGuid(reference.guid)})` : ""}`);
      if (reference.path) lines.push(`  - Path: ${reference.path}`);
    }
    lines.push("");
  }

  lines.push("## Modules", "");
  for (const module of model.modules) {
    lines.push(`### ${module.name}`, "");
    lines.push(`- Kind: ${module.kind}`);
    lines.push(`- Procedures: ${module.procedures.length}`);
    lines.push(`- Variables: ${module.variables.length}`);
    lines.push(`- Constants: ${module.constants.length}`);
    if (module.references.length) lines.push(`- References modules: ${module.references.join(", ")}`);
    if (module.riskMarkers.length) lines.push(`- Risk markers: ${module.riskMarkers.length}`);
    const uses = [...new Set(module.procedures.flatMap((procedure) => procedure.uses))];
    if (uses.length) lines.push(`- Uses: ${uses.join(", ")}`);
    if (module.events.length) {
      lines.push("- Events:");
      for (const event of module.events) {
        lines.push(`  - ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : ""}`);
      }
    }
    lines.push("");
  }

  lines.push("## Dependencies", "");
  if (model.dependencies.length === 0) lines.push("- No high-confidence dependencies detected.");
  for (const dependency of model.dependencies) {
    lines.push(`- ${dependency.category}: ${dependency.value} (${dependency.source}, ${dependency.reason})`);
  }

  lines.push("", "## Migration Notes", "");
  for (const note of model.migrationNotes) lines.push(`- ${note}`);

  return `${lines.join("\n")}\n`;
}

function renderMigrationChecklist(model: ApplicationModel) {
  const riskMarkers = model.modules.flatMap((module) => module.riskMarkers.map((risk) => ({ module: module.name, ...risk })));
  const unlinkedEvents = model.modules.flatMap((module) => module.events.filter((event) => !event.linkedControlPath).map((event) => ({ module: module.name, ...event })));
  const globalState = model.modules.flatMap((module) =>
    module.variables
      .filter((variable) => /^(Public|Global)$/i.test(variable.scope))
      .map((variable) => ({ module: module.name, ...variable })),
  );

  const lines = [
    "# Migration Checklist",
    "",
    "## UI Reconstruction",
    "",
    `- [ ] Recreate ${model.forms.length} form(s) as frontend views/components.`,
    `- [ ] Recreate ${model.forms.reduce((sum, form) => sum + form.controls.length, 0)} control(s), including inferred event-only controls.`,
    `- [ ] Recreate or intentionally discard ${model.documentControls.length} document-level ActiveX control(s).`,
    `- [ ] Verify layout properties from designer summaries against screenshots or Office if available.`,
    `- [ ] Wire ${model.modules.reduce((sum, module) => sum + module.events.length, 0)} event handler(s) to frontend callbacks.`,
    "",
    "## Business Logic",
    "",
    `- [ ] Port ${model.modules.reduce((sum, module) => sum + module.procedures.length, 0)} VBA procedure(s).`,
    `- [ ] Replace ${globalState.length} public/global variable(s) with explicit app state.`,
    "- [ ] Preserve procedure names in comments or metadata for traceability.",
    "- [ ] Add tests around translated calculations, validation, and state transitions.",
    "",
    "## Dependencies",
    "",
  ];

  if (model.dependencies.length === 0) {
    lines.push("- [ ] No high-confidence dependencies detected; still review source code manually.");
  } else {
    for (const dependency of model.dependencies) {
      lines.push(`- [ ] Replace or implement ${dependency.category}: ${dependency.value} (${dependency.source}).`);
    }
  }

  lines.push("", "## Review Items", "");
  if (riskMarkers.length === 0) {
    lines.push("- [ ] No risk markers detected by heuristic scan.");
  } else {
    for (const risk of riskMarkers.slice(0, 80)) {
      lines.push(`- [ ] ${risk.module}:${risk.line} ${risk.category} - ${risk.reason}`);
    }
  }

  lines.push("", "## Unlinked Events", "");
  if (unlinkedEvents.length === 0) {
    lines.push("- [x] All detected event handlers are linked or inferred.");
  } else {
    for (const event of unlinkedEvents) {
      lines.push(`- [ ] ${event.module}.${event.procedure} (${event.controlName}.${event.eventName}) needs manual control mapping.`);
    }
  }

  return `${lines.join("\n")}\n`;
}

function renderCallGraph(model: ApplicationModel) {
  const lines = [
    "# VBA Call Graph",
    "",
    "This graph is inferred from procedure-name references in procedure bodies. Treat it as a navigation aid, not a complete compiler graph.",
    "",
    "## Mermaid",
    "",
    "```mermaid",
    "graph TD",
  ];

  let edgeCount = 0;
  for (const module of model.modules) {
    for (const procedure of module.procedures) {
      const from = graphNodeId(`${module.name}.${procedure.name}`);
      if (procedure.calls.length === 0) {
        lines.push(`  ${from}["${escapeMermaidLabel(`${module.name}.${procedure.name}`)}"]`);
      }
      for (const call of procedure.calls) {
        const targetModule = model.modules.find((candidate) => candidate.procedures.some((candidateProcedure) => candidateProcedure.name === call));
        const toLabel = `${targetModule?.name ?? "unknown"}.${call}`;
        const to = graphNodeId(toLabel);
        lines.push(`  ${from}["${escapeMermaidLabel(`${module.name}.${procedure.name}`)}"] --> ${to}["${escapeMermaidLabel(toLabel)}"]`);
        edgeCount += 1;
      }
    }
  }

  lines.push("```", "", `Edges: ${edgeCount}`, "", "## Event Entry Points", "");
  for (const module of model.modules.filter((candidate) => candidate.events.length)) {
    lines.push(`### ${module.name}`, "");
    for (const event of module.events) {
      lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : ""}`);
    }
    lines.push("");
  }

  return `${lines.join("\n")}\n`;
}

function renderDependencyReport(model: ApplicationModel) {
  const lines = [
    "# Dependency Report",
    "",
    "This report highlights integrations and platform-specific calls that usually need redesign when moving VBA to a browser frontend.",
    "",
  ];

  const byCategory = groupBy(model.dependencies, (dependency) => dependency.category);
  if (byCategory.size === 0) {
    lines.push("No high-confidence dependencies detected.", "");
  }

  for (const [category, dependencies] of byCategory) {
    lines.push(`## ${category}`, "");
    for (const dependency of dependencies) {
      lines.push(`- ${dependency.value} in ${dependency.source}: ${dependency.reason}`);
    }
    lines.push("");
  }

  lines.push("## Risk Markers", "");
  for (const module of model.modules) {
    if (module.riskMarkers.length === 0) continue;
    lines.push(`### ${module.name}`, "");
    for (const risk of module.riskMarkers.slice(0, 60)) {
      lines.push(`- Line ${risk.line}: ${risk.category} - ${risk.reason}`);
      lines.push(`  - ${risk.text}`);
    }
    lines.push("");
  }

  return `${lines.join("\n")}\n`;
}

function renderProjectReferencesReport(model: ApplicationModel) {
  const lines = [
    "# VBA Project References",
    "",
    "This report is generated from the textual PROJECT stream and the decompressed MS-OVBA dir stream reference records.",
    "",
    `Total references: ${model.projectReferences.length}`,
    "",
  ];

  if (model.projectReferences.length === 0) {
    lines.push("No VBA project references were decoded.", "");
    return `${lines.join("\n")}\n`;
  }

  const byKind = groupBy(model.projectReferences, (reference) => reference.kind);
  for (const [kind, references] of byKind) {
    lines.push(`## ${titleCase(kind)} References`, "");
    for (const reference of references) {
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

function titleCase(value: string) {
  return value.replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function renderFrontendImplementationPlan(model: ApplicationModel) {
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
    "",
  ];

  const globalVariables = model.modules.flatMap((module) => module.variables.filter((variable) => /^(Public|Global)$/i.test(variable.scope)).map((variable) => ({ module, variable })));
  if (globalVariables.length === 0) {
    lines.push("- No public/global VBA variables detected. Derive state from forms, controls, and procedure data flow.");
  } else {
    for (const { module, variable } of globalVariables.slice(0, 80)) {
      lines.push(`- ${module.name}.${variable.name}${variable.type ? `: ${variable.type}` : ""} -> app state candidate`);
    }
  }

  lines.push("", "## UI Components", "");
  for (const form of model.forms) {
    lines.push(`### ${form.name}`, "");
    lines.push(`- Component: \`${form.name}View\``);
    if (form.properties.Caption) lines.push(`- Title: ${form.properties.Caption}`);
    if (form.controls.length) {
      lines.push("- Child components/controls:");
      for (const control of form.controls) {
        lines.push(`  - ${control.name || control.caption || control.id}: ${mapControlToFrontend(control)}`);
      }
    }
    lines.push("");
  }

  lines.push("## Services", "");
  for (const module of model.modules.filter((candidate) => candidate.kind === "standard" || candidate.kind === "class")) {
    lines.push(`- \`${module.name}Service\`: port ${module.procedures.length} procedure(s) from ${module.fileName}.`);
  }

  lines.push("", "## Event Wiring", "");
  for (const module of model.modules.filter((candidate) => candidate.events.length)) {
    lines.push(`### ${module.name}`, "");
    for (const event of module.events) {
      lines.push(`- Wire ${event.linkedControlPath ?? event.controlName} ${event.eventName} -> \`${event.procedure}\`.`);
    }
    lines.push("");
  }

  lines.push("## Integration Adapters", "");
  const categories = [...new Set(model.dependencies.map((dependency) => dependency.category))];
  if (categories.length === 0) lines.push("- No adapters detected by heuristic scan.");
  for (const category of categories) {
    lines.push(`- ${category}: create an adapter boundary and decide browser-only vs backend implementation.`);
  }

  return `${lines.join("\n")}\n`;
}

function buildLayoutModel(model: ApplicationModel) {
  return {
    generatedAt: model.generatedAt,
    units: "twips",
    note: "Bounds are best-effort reconstructions from MS-OFORMS streams. Verify against original Office forms before pixel-perfect rebuilds.",
    documentControls: model.documentControls.map((control) => ({
      id: control.id,
      label: control.label,
      classId: control.classId,
      xmlPath: control.xmlPath,
      binPath: control.binPath,
      properties: control.properties,
    })),
    views: model.forms.map((form) => ({
      name: form.name,
      caption: form.properties.Caption,
      clientWidth: parseNumber(form.properties.ClientWidth),
      clientHeight: parseNumber(form.properties.ClientHeight),
      controls: form.controls.map((control) => ({
        path: control.path,
        name: control.name,
        caption: control.caption,
        type: control.type,
        progId: control.progId,
        bounds: control.bounds,
        parentPath: control.parentPath,
      })),
    })),
  };
}

function buildTraceabilityMap(files: ExtractedFile[], model: ApplicationModel) {
  const fileEntries = files.map((file, index) => ({
    extractedName: file.name,
    kind: file.kind,
    zipPath: getZipPath(file, index),
    sourcePath: file.sourcePath,
    size: file.bytes.byteLength,
  }));

  return {
    generatedAt: model.generatedAt,
    files: fileEntries,
    modules: model.modules.map((module) => ({
      name: module.name,
      fileName: module.fileName,
      sourcePath: module.sourcePath,
      zipPath: fileEntries.find((file) => file.extractedName === module.fileName)?.zipPath,
      procedures: module.procedures.map((procedure, index) => ({
        name: procedure.name,
        lines: { start: procedure.lineStart, end: procedure.lineEnd },
        chunkPath: `procedure-chunks/${safeFileName(module.name)}/${String(index + 1).padStart(3, "0")}-${safeFileName(procedure.name)}.vba`,
      })),
    })),
    forms: model.forms.map((form) => ({
      name: form.name,
      designerSummaryPath: `${safeFileName(form.name)}/designer-summary/${safeFileName(form.name)}.designer.json`,
      controls: form.controls.map((control) => ({
        path: control.path,
        name: control.name,
        caption: control.caption,
        type: control.type,
        sourceStreams: control.sourceStreams,
        parentPath: control.parentPath,
      })),
    })),
    documentControls: model.documentControls.map((control) => ({
      id: control.id,
      label: control.label,
      classId: control.classId,
      xmlPath: control.xmlPath,
      binPath: control.binPath,
      sourceFiles: control.sourceFiles,
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
      "vba-project-references.json",
    ],
  };
}

function renderValidationReport(files: ExtractedFile[], model: ApplicationModel) {
  const sourceFiles = files.filter((file) => file.kind === "vba" || file.kind === "frm");
  const fallbackFiles = sourceFiles.filter((file) => file.sourcePath.includes("recovered by scan"));
  const controls = model.forms.flatMap((form) => form.controls);
  const inferredControls = controls.filter((control) => control.path.includes("/inferred/"));
  const typedControls = controls.filter((control) => control.type || control.progId);
  const namedControls = controls.filter((control) => control.name || control.caption);
  const boundedControls = controls.filter((control) => control.bounds);
  const events = model.modules.flatMap((module) => module.events);
  const linkedEvents = events.filter((event) => event.linkedControlPath);
  const lowConfidenceProperties = controls.flatMap((control) =>
    Object.entries(control.properties)
      .filter(([, value]) => /\(low\)$/i.test(value))
      .map(([key, value]) => ({ control: control.path, key, value })),
  );
  const media = files.filter((file) => file.kind === "media");
  const activeXParts = files.filter((file) => /\/activeX\//i.test(file.sourcePath));
  const manifest = files.find((file) => file.name === "office-package-manifest.json");
  const riskMarkers = model.modules.flatMap((module) => module.riskMarkers.map((risk) => ({ module: module.name, ...risk })));
  const hasVbaProjectOutput = sourceFiles.length > 0;
  const scoreItems = [
    sourceFiles.length > 0 || media.length > 0 || activeXParts.length > 0,
    fallbackFiles.length === 0,
    events.length === 0 || linkedEvents.length / events.length >= 0.75,
    controls.length === 0 || typedControls.length / controls.length >= 0.75,
    controls.length === 0 || inferredControls.length / controls.length <= 0.5,
  ];
  const score = Math.round((scoreItems.filter(Boolean).length / scoreItems.length) * 100);

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
    "",
  ];

  const reviewItems: string[] = [];
  if (!hasVbaProjectOutput) reviewItems.push("No VBA/FRM source was extracted. This package is useful for Office media/ActiveX migration context, but not a complete VBA application rebuild.");
  if (fallbackFiles.length) reviewItems.push("Some source modules used byte-scan fallback instead of exact dir-stream offsets.");
  if (inferredControls.length) reviewItems.push("Some controls were inferred from event handler names rather than decoded from form streams.");
  if (events.length !== linkedEvents.length) reviewItems.push("Some event handlers could not be linked to decoded/inferred controls.");
  if (controls.length !== boundedControls.length) reviewItems.push("Some controls have no trustworthy layout bounds.");
  if (lowConfidenceProperties.length) reviewItems.push("Some decoded control properties are marked low-confidence.");
  if (activeXParts.length && model.documentControls.length === 0) reviewItems.push("ActiveX package parts were found but no documentControls were decoded.");
  if (riskMarkers.length) reviewItems.push("Risk markers indicate platform-specific or migration-sensitive VBA code.");
  if (reviewItems.length === 0) reviewItems.push("No major validation warnings detected by heuristics.");
  for (const item of reviewItems) lines.push(`- ${item}`);

  lines.push("", "## Unlinked Events", "");
  const unlinkedEvents = events.filter((event) => !event.linkedControlPath);
  if (unlinkedEvents.length === 0) lines.push("- None.");
  for (const event of unlinkedEvents.slice(0, 80)) lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}`);

  lines.push("", "## Inferred Controls", "");
  if (inferredControls.length === 0) lines.push("- None.");
  for (const control of inferredControls.slice(0, 80)) lines.push(`- ${control.path}: ${control.type ?? "unknown"} from ${control.properties.InferredFromEvent ?? "event heuristic"}`);

  lines.push("", "## Low-Confidence Properties", "");
  if (lowConfidenceProperties.length === 0) lines.push("- None.");
  for (const property of lowConfidenceProperties.slice(0, 80)) lines.push(`- ${property.control}: ${property.key} = ${property.value}`);

  lines.push("", "## Fallback Source Files", "");
  if (fallbackFiles.length === 0) lines.push("- None.");
  for (const file of fallbackFiles) lines.push(`- ${file.name}: ${file.sourcePath}`);

  lines.push("", "## Media And ActiveX", "");
  lines.push(`- Media assets recovered: ${media.length}`);
  lines.push(`- ActiveX package parts extracted: ${activeXParts.length}`);
  lines.push(`- ActiveX controls modeled: ${model.documentControls.length}`);
  lines.push(`- ActiveX binary persistence parsed: ${model.documentControls.filter((control) => control.persistenceAnalysis).length}`);

  return `${lines.join("\n")}\n`;
}

function parseNumber(value: string | undefined) {
  if (!value) return undefined;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function mapControlToFrontend(control: FlatControlModel) {
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

function renderMigrationTestPlan(model: ApplicationModel) {
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
    "",
  ];

  for (const module of model.modules.filter((candidate) => candidate.events.length)) {
    lines.push(`### ${module.name}`, "");
    for (const event of module.events) {
      lines.push(`- [ ] Trigger ${event.controlName}.${event.eventName} and verify behavior from \`${event.procedure}\`.`);
    }
    lines.push("");
  }

  lines.push("## Procedure Tests", "");
  for (const module of model.modules) {
    const candidates = module.procedures.filter((procedure) => procedure.kind === "Function" || procedure.uses.length || procedure.calls.length);
    if (candidates.length === 0) continue;
    lines.push(`### ${module.name}`, "");
    for (const procedure of candidates.slice(0, 80)) {
      const notes = [procedure.uses.join(", "), procedure.calls.length ? `calls ${procedure.calls.join(", ")}` : ""].filter(Boolean).join("; ");
      lines.push(`- [ ] ${procedure.name}${notes ? ` (${notes})` : ""}`);
    }
    lines.push("");
  }

  lines.push("## Dependency Tests", "");
  if (model.dependencies.length === 0) {
    lines.push("- [ ] No detected dependency tests. Review source manually for hidden integrations.");
  } else {
    for (const dependency of model.dependencies) {
      lines.push(`- [ ] Validate replacement for ${dependency.category}: ${dependency.value} (${dependency.source}).`);
    }
  }

  lines.push("", "## Regression Data", "");
  lines.push("- [ ] Collect representative workbook/document inputs from the business owner.");
  lines.push("- [ ] Capture expected outputs from the original VBA app before retiring it.");
  lines.push("- [ ] Add automated tests for calculations and transformations before UI polish.");

  return `${lines.join("\n")}\n`;
}

function graphNodeId(value: string) {
  return `n_${value.replace(/[^a-z0-9_]/gi, "_")}`;
}

function escapeMermaidLabel(value: string) {
  return value.replace(/"/g, "'");
}

function groupBy<T>(items: T[], keyFn: (item: T) => string) {
  const grouped = new Map<string, T[]>();
  for (const item of items) {
    const key = keyFn(item);
    grouped.set(key, [...(grouped.get(key) ?? []), item]);
  }
  return grouped;
}
