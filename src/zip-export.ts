import type { ApplicationModel, ExtractedFile } from "./types";
import { concatBytes, encodeText, safeFileName, writeU16, writeU32 } from "./runtime";
import { buildResultGroups, getResultOwner } from "./results-ui";
import { buildApplicationModel, buildLayoutModel, buildTraceabilityMap, renderApplicationSummary, renderCallGraph, renderDependencyReport, renderFrontendImplementationPlan, renderMigrationChecklist, renderMigrationTestPlan, renderProjectReferencesReport, renderValidationReport } from "./model";
import { createVisualPreviewArtifacts, renderLlmRebuildBrief } from "./designer-ui";

export function createZip(files: ExtractedFile[]) {
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
  const activeXPersistence = appModel.documentControls
    .filter((control) => control.persistenceAnalysis)
    .map((control) => ({ id: control.id, xmlPath: control.xmlPath, binPath: control.binPath, analysis: control.persistenceAnalysis }));
  const visualPreviewArtifacts = createVisualPreviewArtifacts(groups);
  const validationReport = renderValidationReport(files, appModel);
  const appArtifacts = [
    { path: "application-model.json", bytes: encodeText(JSON.stringify(appModel, null, 2)) },
    { path: "application-summary.md", bytes: encodeText(appSummary) },
    { path: "llm-rebuild-brief.md", bytes: encodeText(rebuildBrief) },
    { path: "migration-checklist.md", bytes: encodeText(migrationChecklist) },
    { path: "call-graph.md", bytes: encodeText(callGraph) },
    { path: "dependency-report.md", bytes: encodeText(dependencyReport) },
    { path: "vba-project-references.md", bytes: encodeText(projectReferencesReport) },
    { path: "frontend-implementation-plan.md", bytes: encodeText(frontendPlan) },
    { path: "migration-test-plan.md", bytes: encodeText(testPlan) },
    { path: "layout-model.json", bytes: encodeText(JSON.stringify(layoutModel, null, 2)) },
    { path: "traceability-map.json", bytes: encodeText(JSON.stringify(traceabilityMap, null, 2)) },
    { path: "activex-controls.json", bytes: encodeText(JSON.stringify(activeXControls, null, 2)) },
    { path: "activex-persistence.json", bytes: encodeText(JSON.stringify(activeXPersistence, null, 2)) },
    { path: "validation-report.md", bytes: encodeText(validationReport) },
  ];
  const summaryFiles = groups
    .filter((group) => group.designer)
    .map((group) => ({
      owner: group.name,
      bytes: encodeText(JSON.stringify(group.designer, null, 2)),
    }));
  const extractedEntries = files.map((file, index) => ({
    path: getZipPath(file, index),
    bytes: file.bytes,
    crc: crc32(file.bytes),
  }));
  const summaryEntries = summaryFiles.map((summary) => {
    const path = `${safeFileName(summary.owner)}/designer-summary/${safeFileName(summary.owner)}.designer.json`;
    return { path, bytes: summary.bytes, crc: crc32(summary.bytes) };
  });
  const appEntries = appArtifacts.map((artifact) => ({ ...artifact, crc: crc32(artifact.bytes) }));
  const visualPreviewEntries = visualPreviewArtifacts.map((artifact) => ({ ...artifact, crc: crc32(artifact.bytes) }));
  const procedureEntries = createProcedureZipEntries(files, appModel);
  const entries = [...appEntries, ...visualPreviewEntries, ...procedureEntries, ...extractedEntries, ...summaryEntries];
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
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
  const centralDirectorySize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, entries.length);
  writeU16(end, 10, entries.length);
  writeU32(end, 12, centralDirectorySize);
  writeU32(end, 16, centralDirectoryOffset);

  return concatBytes([...localParts, ...centralParts, end]);
}

export function getZipPath(file: ExtractedFile, index: number) {
  const owner = safeFileName(getResultOwner(file));
  const category = file.kind === "media" ? "media" : file.kind === "frx" ? "resources" : file.kind === "frm" || file.kind === "vba" ? "code" : /office-package-manifest|\/(word|xl|ppt)\//i.test(file.sourcePath) ? "office-package" : "other";
  return `${owner}/${category}/${String(index + 1).padStart(3, "0")}-${safeFileName(file.name)}`;
}

export function createProcedureZipEntries(files: ExtractedFile[], model: ApplicationModel) {
  const entries: Array<{ path: string; bytes: Uint8Array; crc: number }> = [];

  for (const module of model.modules) {
    const source = files.find((file) => file.name === module.fileName)?.text;
    if (!source) continue;
    const lines = source.split(/\r?\n/);

    module.procedures.forEach((procedure, index) => {
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
        "",
      ].join("\n");
      const bytes = encodeText(`${metadata}${body}\n`);
      entries.push({
        path: `procedure-chunks/${safeFileName(module.name)}/${String(index + 1).padStart(3, "0")}-${safeFileName(procedure.name)}.vba`,
        bytes,
        crc: crc32(bytes),
      });
    });
  }

  return entries;
}

export function crc32(bytes: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ byte) & 0xff];
  return (crc ^ 0xffffffff) >>> 0;
}

export function makeCrc32Table() {
  const table: number[] = [];
  for (let i = 0; i < 256; i += 1) {
    let value = i;
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table.push(value >>> 0);
  }
  return table;
}

export const CRC32_TABLE = makeCrc32Table();

export function getDosDateTime(date: Date) {
  const year = Math.max(1980, date.getFullYear());
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}
