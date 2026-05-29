export type ExtractedFile = {
  name: string;
  kind: "vba" | "frm" | "frx" | "media" | "binary";
  bytes: Uint8Array;
  text?: string;
  analysis?: BinaryAnalysis;
  mimeType: string;
  sourcePath: string;
};

export type OfficeZipEntryData = {
  entry: ZipEntry;
  bytes: Uint8Array;
};

export type ProjectReferenceModel = {
  source: "PROJECT" | "dir";
  kind: "registered" | "control" | "project" | "package" | "object" | "unknown";
  name?: string;
  libId?: string;
  guid?: string;
  version?: string;
  path?: string;
  raw: string;
};

export type BinaryAnalysis = {
  title: string;
  summary: Array<{ label: string; value: string }>;
  stringCount: number;
  guidCount: number;
  signatures: Array<{ offset: number; label: string; detail: string }>;
  oforms?: OFormsAnalysis;
};

export type OFormsAnalysis = {
  kind: string;
  confidence: "high" | "medium" | "low";
  records: Array<{
    offset: number;
    type: string;
    size: number;
    properties: Array<{ name: string; value: string }>;
  }>;
  notes: string[];
};

export type ExtractedString = {
  encoding: string;
  offset: number;
  value: string;
};

export type ProjectModule = {
  name: string;
  type: "module" | "class" | "form" | "document";
};

export type DirModule = {
  name: string;
  streamName: string;
  textOffset: number;
  moduleTypeId?: number;
};

export type ResultGroup = {
  name: string;
  code: ExtractedFile[];
  resources: ExtractedFile[];
  media: ExtractedFile[];
  other: ExtractedFile[];
  designer?: DesignerSummary;
};

export type DesignerSummary = {
  formName: string;
  frame: Record<string, string>;
  controls: DesignerControl[];
};

export type DesignerControl = {
  id: string;
  path: string;
  type?: string;
  progId?: string;
  caption?: string;
  name?: string;
  properties: Record<string, string>;
  bounds?: Bounds;
  sourceStreams: string[];
  children: DesignerControl[];
};

export type VisualPreviewModel = {
  name: string;
  width: number;
  height: number;
  scale: number;
  controls: VisualPreviewControl[];
};

export type VisualPreviewControl = {
  id: string;
  path: string;
  kind: string;
  label: string;
  left: number;
  top: number;
  width: number;
  height: number;
  confidence: Bounds["confidence"];
  type?: string;
  progId?: string;
};

export type Bounds = {
  left: number;
  top: number;
  width: number;
  height: number;
  unit: "twips";
  confidence: "high" | "medium" | "low";
};

export type ApplicationModel = {
  generatedAt: string;
  sourceFiles: Array<{
    name: string;
    kind: ExtractedFile["kind"];
    path: string;
    size: number;
  }>;
  modules: ModuleModel[];
  projectReferences: ProjectReferenceModel[];
  forms: FormModel[];
  documentControls: ActiveXControlModel[];
  dependencies: DependencyModel[];
  assets: AssetModel[];
  migrationNotes: string[];
};

export type ModuleModel = {
  name: string;
  fileName: string;
  kind: "standard" | "class" | "form" | "document";
  sourcePath: string;
  procedures: ProcedureModel[];
  declarations: string[];
  variables: DeclarationModel[];
  constants: DeclarationModel[];
  events: EventModel[];
  references: string[];
  riskMarkers: RiskMarker[];
};

export type DeclarationModel = {
  name: string;
  scope: string;
  type?: string;
  line: number;
  statement: string;
};

export type RiskMarker = {
  category: string;
  line: number;
  text: string;
  reason: string;
};

export type ProcedureModel = {
  name: string;
  kind: "Sub" | "Function" | "Property";
  scope: "Public" | "Private" | "Friend" | "Implicit";
  signature: string;
  parameters: string[];
  returnType?: string;
  lineStart: number;
  lineEnd?: number;
  calls: string[];
  uses: string[];
};

export type EventModel = {
  procedure: string;
  controlName: string;
  eventName: string;
  linkedControlPath?: string;
  linkedControlType?: string;
};

export type FormModel = {
  name: string;
  properties: Record<string, string>;
  controls: FlatControlModel[];
};

export type FlatControlModel = {
  id: string;
  path: string;
  name?: string;
  caption?: string;
  type?: string;
  progId?: string;
  bounds?: Bounds;
  sourceStreams: string[];
  properties: Record<string, string>;
  parentPath?: string;
};

export type DependencyModel = {
  category: string;
  value: string;
  source: string;
  reason: string;
};

export type AssetModel = {
  name: string;
  mimeType: string;
  size: number;
  sourcePath: string;
};

export type ActiveXControlModel = {
  id: string;
  xmlPath: string;
  binPath?: string;
  classId?: string;
  label?: string;
  persistence?: string;
  properties: Record<string, string>;
  persistenceAnalysis?: ActiveXPersistenceAnalysis;
  sourceFiles: string[];
};

export type ActiveXPersistenceAnalysis = {
  format: "cfb-storage" | "raw-stream" | "empty" | "unknown";
  confidence: "high" | "medium" | "low";
  size: number;
  streams: ActiveXPersistenceStream[];
  compObj?: {
    clsid?: string;
    userType?: string;
    clipboardFormat?: string;
    progId?: string;
  };
  properties: Array<{
    name: string;
    value: string;
    source: string;
    confidence: "high" | "medium" | "low";
  }>;
  media: Array<{ name: string; mimeType: string; size: number; sourcePath: string }>;
  warnings: string[];
};

export type ActiveXPersistenceStream = {
  path: string;
  size: number;
  strings: string[];
  guids: string[];
  signatures: string[];
  oforms?: string;
};

export type MediaMatch = {
  extension: string;
  mimeType: string;
  length: number;
  bytes?: Uint8Array;
  label?: string;
};

export const CLSID_STD_PICTURE = [0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11, 0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51];

export type ZipEntry = {
  name: string;
  method: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
};

export type DirectoryEntry = {
  index: number;
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  startSector: number;
  size: number;
  path: string;
};

export const END_OF_CHAIN = -2;
export const FREE_SECTOR = -1;
export const FAT_SECTOR = -3;
export const DIFAT_SECTOR = -4;
