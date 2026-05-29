function renderLlmRebuildBrief(model: ApplicationModel) {
  const eventCount = model.modules.reduce((sum, module) => sum + module.events.length, 0);
  const linkedEventCount = model.modules.reduce((sum, module) => sum + module.events.filter((event) => event.linkedControlPath).length, 0);
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
    `- Controls: ${model.forms.reduce((sum, form) => sum + form.controls.length, 0)}`,
    `- Modules: ${model.modules.length}`,
    `- Procedures: ${model.modules.reduce((sum, module) => sum + module.procedures.length, 0)}`,
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
    "",
  ];

  for (const module of model.modules.filter((candidate) => candidate.events.length)) {
    lines.push(`### ${module.name}`, "");
    for (const event of module.events) {
      lines.push(`- ${event.procedure}: ${event.controlName}.${event.eventName}${event.linkedControlPath ? ` -> ${event.linkedControlPath}` : " -> unlinked"}`);
    }
    lines.push("");
  }

  lines.push("## Dependency Warnings", "");
  if (model.dependencies.length === 0) {
    lines.push("- No high-confidence external dependencies were detected.");
  } else {
    for (const dependency of model.dependencies) {
      lines.push(`- ${dependency.category}: ${dependency.value} from ${dependency.source}. Migration concern: ${dependency.reason}.`);
    }
  }

  lines.push(
    "",
    "## Implementation Advice",
    "",
    "- Start by recreating the UI shape from `forms` and `controls` in `application-model.json`.",
    "- Then wire linked event handlers using the original VBA code as behavior reference.",
    "- Replace global workbook state with explicit frontend state stores.",
    "- Treat dependency detections as integration requirements. Do not silently drop them.",
    "- Preserve original names in comments or metadata so reviewers can trace VBA behavior to the new code.",
    "",
  );

  return lines.join("\n");
}

function getDisplayPath(file: ExtractedFile) {
  const streamPath = extractInternalStreamPath(file.sourcePath);
  return streamPath ? printableStreamName(streamPath) : file.sourcePath;
}

function buildDesignerSummary(group: ResultGroup): DesignerSummary | undefined {
  const hasFormCode = group.code.some((file) => file.kind === "frm");
  if (!hasFormCode && group.resources.length === 0) return undefined;
  if (group.name === "Project metadata") return undefined;

  const frame = parseVbFrameProperties(group.resources);
  const controlsByPath = new Map<string, DesignerControl>();
  const rootControls: DesignerControl[] = [];

  for (const resource of group.resources) {
    const internalPath = extractInternalStreamPath(resource.sourcePath);
    if (!internalPath) continue;
    const parts = internalPath.split("/");
    if (parts.length < 2 || parts[0] !== group.name) continue;

    for (let depth = 2; depth <= parts.length - 1; depth += 1) {
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
          children: [],
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

  for (const control of controlsByPath.values()) {
    finalizeControlBounds(control);
  }

  const directFormProps = extractResourceProperties(group.resources.filter((resource) => {
    const internalPath = extractInternalStreamPath(resource.sourcePath);
    return internalPath === `${group.name}/f` || internalPath === `${group.name}/\x01CompObj`;
  }));

  for (const [key, value] of Object.entries(directFormProps)) {
    if (!frame[key]) frame[key] = value;
  }

  return {
    formName: group.name,
    frame,
    controls: rootControls,
  };
}

function parseVbFrameProperties(resources: ExtractedFile[]) {
  const frame: Record<string, string> = {};
  const vbFrame = resources.find((resource) => extractInternalStreamPath(resource.sourcePath)?.endsWith("/\x03VBFrame"));
  if (!vbFrame) return frame;

  const text = decodeText(vbFrame.bytes);
  const beginMatch = text.match(/Begin\s+\{[^}]+\}\s+([^\s\r\n]+)/i);
  if (beginMatch) frame.Name = beginMatch[1];

  for (const line of text.split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!match) continue;
    frame[match[1]] = match[2].replace(/^"|"$/g, "");
  }
  return frame;
}

function mergeControlResource(control: DesignerControl, resource: ExtractedFile) {
  const internalPath = extractInternalStreamPath(resource.sourcePath);
  const streamName = internalPath?.split("/").at(-1);
  if (!streamName) return;
  control.sourceStreams.push(internalPath ?? resource.sourcePath);

  const properties = extractResourceProperties([resource]);
  const bounds = inferBoundsFromResource(resource);
  if (bounds && !control.bounds) control.bounds = bounds;
  for (const [key, value] of Object.entries(properties)) {
    control.properties[key] = value;
  }

  if (streamName === "\x01CompObj") {
    const progId = Object.values(properties).find((value) => /^Forms\./i.test(value.replace(/\s+\(.+\)$/, "")));
    if (progId) {
      control.progId = progId.replace(/\s+\(.+\)$/, "");
      control.type = labelProgId(control.progId) ?? progId;
    }
    const microsoftFormsName = Object.values(properties).find((value) => /Microsoft Forms 2\.0/i.test(value));
    if (!control.type && microsoftFormsName) control.type = microsoftFormsName;
  }

  const decodedName = stripConfidence(properties["Decoded Name"]);
  const decodedCaption = stripConfidence(properties["Decoded Caption"]);
  if (decodedName && isHumanLabel(decodedName) && propertyConfidence(properties["Decoded Name"]) !== "low") control.name = decodedName;
  if (decodedCaption && isHumanLabel(decodedCaption)) control.caption = decodedCaption;

  const captionCandidate = Object.entries(properties).find(([key, value]) => key.startsWith("Label ") && isHumanLabel(value));
  if (!control.caption && captionCandidate) control.caption = captionCandidate[1].replace(/\s+\(.+\)$/, "");
  if (!control.name && control.caption) control.name = control.caption;
}

function stripConfidence(value: string | undefined) {
  return value?.replace(/\s+\((?:high|medium|low)\)$/i, "");
}

function propertyConfidence(value: string | undefined) {
  return value?.match(/\((high|medium|low)\)$/i)?.[1].toLowerCase();
}

function finalizeControlBounds(control: DesignerControl) {
  const size = parseSizeHint(control.properties["Decoded DisplayedSize"]) ?? parseSizeHint(control.properties["Decoded LogicalSize"]);
  const position = parsePositionHint(control.properties["Decoded SitePosition"]);
  if (size && position) {
    control.bounds = {
      left: position.left,
      top: position.top,
      width: size.width,
      height: size.height,
      unit: "twips",
      confidence: size.confidence === "medium" && position.confidence === "medium" ? "medium" : "low",
    };
  }
}

function parseSizeHint(value: string | undefined) {
  const match = value?.match(/(\d+)\s+x\s+(\d+)\s+twips.*\((medium|low)\)/i);
  if (!match) return undefined;
  return { width: Number(match[1]), height: Number(match[2]), confidence: match[3] as "medium" | "low" };
}

function parsePositionHint(value: string | undefined) {
  const match = value?.match(/(\d+),\s*(\d+)\s+twips.*\((medium|low)\)/i);
  if (!match) return undefined;
  return { left: Number(match[1]), top: Number(match[2]), confidence: match[3] as "medium" | "low" };
}

function inferBoundsFromResource(resource: ExtractedFile): Bounds | undefined {
  const internalPath = extractInternalStreamPath(resource.sourcePath);
  if (!internalPath?.endsWith("/f")) return undefined;
  const bytes = resource.bytes;
  if (bytes.length < 32) return undefined;

  const candidates: Bounds[] = [];
  for (let offset = 8; offset + 16 <= Math.min(bytes.length, 160); offset += 4) {
    const a = readU32(bytes, offset);
    const b = readU32(bytes, offset + 4);
    const c = readU32(bytes, offset + 8);
    const d = readU32(bytes, offset + 12);
    if ([a, b, c, d].every((value) => value >= 0 && value < 60_000) && c > 0 && d > 0) {
      const looksLikeSize = c >= 30 && d >= 30 && c < 32_000 && d < 32_000;
      const looksLikePosition = a < 25_000 && b < 25_000;
      if (looksLikeSize) {
        candidates.push({ left: a, top: b, width: c, height: d, unit: "twips", confidence: looksLikePosition && offset <= 64 ? "medium" : "low" });
      }
    }
  }

  const best = candidates.sort((a, b) => scoreBounds(b) - scoreBounds(a))[0];
  return best?.confidence === "low" && (best.left > 25_000 || best.top > 25_000) ? undefined : best;
}

function scoreBounds(bounds: Bounds) {
  let score = 0;
  if (bounds.width > bounds.height) score += 1;
  if (bounds.width > 500 && bounds.height > 300) score += 2;
  if (bounds.left < 20_000 && bounds.top < 20_000) score += 1;
  if (bounds.confidence === "medium") score += 2;
  return score;
}

function extractResourceProperties(resources: ExtractedFile[]) {
  const properties: Record<string, string> = {};
  for (const resource of resources) {
    const records = resource.analysis?.oforms?.records ?? [];
    for (const record of records) {
      for (const property of record.properties) {
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

function isHumanLabel(value: string) {
  const plain = value.replace(/\s+\(.+\)$/, "");
  if (/^Forms\./i.test(plain) || /Embedded Object/i.test(plain) || /Microsoft Forms/i.test(plain)) return false;
  if (/^Tahoma$/i.test(plain) || /^0x[0-9a-f]+$/i.test(plain) || /^\{[0-9A-F-]+\}$/i.test(plain)) return false;
  if (/^'/.test(plain) || /^[-=]{5,}$/.test(plain) || /https?:\/\//i.test(plain) || /copyright|permission is hereby/i.test(plain)) return false;
  return /^[\w .'-]{2,80}$/i.test(plain);
}

function renderDesignerSummary(summary: DesignerSummary) {
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
    for (const control of summary.controls) tree.append(renderControlNode(control));
    section.append(tree);
  }

  return section;
}

function renderKeyValueGrid(values: Record<string, string>) {
  const dl = document.createElement("dl");
  dl.className = "designer-grid";
  for (const [key, value] of Object.entries(values).slice(0, 24)) {
    const dt = document.createElement("dt");
    dt.textContent = key;
    const dd = document.createElement("dd");
    dd.textContent = labelPossibleIdentifier(value);
    dl.append(dt, dd);
  }
  return dl;
}

function renderLayoutPreview(summary: DesignerSummary) {
  const wrapper = document.createElement("div");
  wrapper.className = "layout-preview-wrap";

  const title = document.createElement("h4");
  title.textContent = "Best-effort layout preview";
  wrapper.append(title);

  const controls = flattenDesignerControls(summary.controls);
  const bounded = controls.filter((control) => control.bounds);
  if (bounded.length === 0) {
    const empty = document.createElement("p");
    empty.className = "analysis-empty";
    empty.textContent = "No usable bounds detected yet.";
    wrapper.append(empty);
    return wrapper;
  }

  const maxRight = Math.max(...bounded.map((control) => (control.bounds!.left + control.bounds!.width)));
  const maxBottom = Math.max(...bounded.map((control) => (control.bounds!.top + control.bounds!.height)));
  const scale = Math.min(1, 760 / Math.max(maxRight, 1), 420 / Math.max(maxBottom, 1));

  const canvas = document.createElement("div");
  canvas.className = "layout-preview";
  canvas.style.width = `${Math.max(320, maxRight * scale + 24)}px`;
  canvas.style.height = `${Math.max(180, maxBottom * scale + 24)}px`;

  for (const control of bounded) {
    const bounds = control.bounds!;
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

function flattenDesignerControls(controls: DesignerControl[]): DesignerControl[] {
  return controls.flatMap((control) => [control, ...flattenDesignerControls(control.children)]);
}

function renderVisualLayoutPreview(summary: DesignerSummary) {
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
  svgButton.addEventListener("click", () => downloadBytes(`${safeFileName(summary.formName)}-preview.svg`, "image/svg+xml", encodeText(renderVisualPreviewSvg(preview))));
  actions.append(svgButton);

  const pngButton = document.createElement("button");
  pngButton.type = "button";
  pngButton.textContent = "PNG";
  pngButton.addEventListener("click", async () => {
    const blob = await renderVisualPreviewPng(preview);
    downloadBlob(`${safeFileName(summary.formName)}-preview.png`, blob);
  });
  actions.append(pngButton);

  const canvas = document.createElement("div");
  canvas.className = "layout-preview";
  canvas.style.width = `${preview.width}px`;
  canvas.style.height = `${preview.height}px`;

  for (const control of preview.controls) {
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

function buildVisualPreviewModel(summary: DesignerSummary): VisualPreviewModel | undefined {
  const controls = flattenDesignerControls(summary.controls)
    .filter((control) => control.bounds)
    .map((control) => ({ control, bounds: control.bounds! }));
  if (controls.length === 0) return undefined;

  const maxRight = Math.max(...controls.map(({ bounds }) => bounds.left + bounds.width));
  const maxBottom = Math.max(...controls.map(({ bounds }) => bounds.top + bounds.height));
  const scale = Math.min(1, 900 / Math.max(maxRight, 1), 560 / Math.max(maxBottom, 1));

  return {
    name: summary.formName,
    width: Math.ceil(Math.max(360, maxRight * scale + 24)),
    height: Math.ceil(Math.max(220, maxBottom * scale + 24)),
    scale,
    controls: controls.map(({ control, bounds }) => ({
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
      progId: control.progId,
    })),
  };
}

function getVisualControlKind(control: Pick<DesignerControl | FlatControlModel | VisualPreviewControl, "type" | "progId"> & { label?: string; caption?: string; name?: string }) {
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

function renderVisualControlContents(control: VisualPreviewControl) {
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

function createVisualPreviewArtifacts(groups: ResultGroup[]) {
  const artifacts: Array<{ path: string; bytes: Uint8Array }> = [];
  for (const group of groups) {
    if (!group.designer) continue;
    const preview = buildVisualPreviewModel(group.designer);
    if (!preview) continue;
    const base = `visual-previews/${safeFileName(group.designer.formName)}`;
    artifacts.push({ path: `${base}.svg`, bytes: encodeText(renderVisualPreviewSvg(preview)) });
    artifacts.push({ path: `${base}.html`, bytes: encodeText(renderVisualPreviewHtml(preview)) });
  }
  return artifacts;
}

function renderVisualPreviewSvg(preview: VisualPreviewModel) {
  const controls = preview.controls.map((control) => renderVisualControlSvg(control)).join("\n");
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

function renderVisualControlSvg(control: VisualPreviewControl) {
  const label = truncateLabel(control.label, Math.max(6, Math.floor(control.width / 7)));
  const style = visualControlStyle(control.kind);
  const textY = control.top + Math.max(15, Math.min(control.height - 6, control.height / 2 + 4));
  const textX = control.left + (["checkbox", "radio"].includes(control.kind) ? 24 : 8);
  const parts = [
    `  <g data-path="${escapeXml(control.path)}" data-kind="${escapeXml(control.kind)}">`,
    `    <rect x="${control.left}" y="${control.top}" width="${control.width}" height="${control.height}" rx="${style.radius}" fill="${style.fill}" stroke="${style.stroke}" stroke-width="1"/>`,
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

function renderVisualPreviewHtml(preview: VisualPreviewModel) {
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
    ${preview.controls.map((control) => `<div class="control ${escapeHtml(control.kind)}" title="${escapeHtml(control.path)}" style="left:${control.left}px;top:${control.top}px;width:${control.width}px;height:${control.height}px">${escapeHtml(control.label)}</div>`).join("\n    ")}
  </div>
</body>
</html>
`;
}

async function renderVisualPreviewPng(preview: VisualPreviewModel) {
  const canvas = document.createElement("canvas");
  canvas.width = preview.width * 2;
  canvas.height = preview.height * 2;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Canvas is not available.");
  ctx.scale(2, 2);
  drawVisualPreviewCanvas(ctx, preview);
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error("Could not render PNG."))), "image/png");
  });
}

function drawVisualPreviewCanvas(ctx: CanvasRenderingContext2D, preview: VisualPreviewModel) {
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, preview.width, preview.height);
  ctx.strokeStyle = "#edf4f3";
  ctx.lineWidth = 1;
  for (let x = 0; x <= preview.width; x += 20) {
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, preview.height);
    ctx.stroke();
  }
  for (let y = 0; y <= preview.height; y += 20) {
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(preview.width, y);
    ctx.stroke();
  }

  ctx.font = "11px Segoe UI, Arial, sans-serif";
  ctx.fillStyle = "#607178";
  ctx.fillText(preview.name, 10, 18);

  for (const control of preview.controls) drawVisualControlCanvas(ctx, control);
}

function drawVisualControlCanvas(ctx: CanvasRenderingContext2D, control: VisualPreviewControl) {
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

function visualControlStyle(kind: string) {
  const styles: Record<string, { fill: string; stroke: string; text: string; radius: number }> = {
    button: { fill: "#eef3f4", stroke: "#8a9ca1", text: "#172026", radius: 4 },
    label: { fill: "rgba(255,255,255,0)", stroke: "rgba(255,255,255,0)", text: "#172026", radius: 0 },
    textbox: { fill: "#ffffff", stroke: "#8a9ca1", text: "#172026", radius: 2 },
    combo: { fill: "#ffffff", stroke: "#8a9ca1", text: "#172026", radius: 2 },
    listbox: { fill: "#ffffff", stroke: "#8a9ca1", text: "#172026", radius: 2 },
    frame: { fill: "rgba(255,255,255,0.52)", stroke: "#8fb9b4", text: "#526168", radius: 3 },
    tabs: { fill: "rgba(255,255,255,0.7)", stroke: "#8fb9b4", text: "#526168", radius: 3 },
    image: { fill: "#f4f7f8", stroke: "#8a9ca1", text: "#607178", radius: 3 },
    checkbox: { fill: "rgba(255,255,255,0)", stroke: "rgba(255,255,255,0)", text: "#172026", radius: 0 },
    radio: { fill: "rgba(255,255,255,0)", stroke: "rgba(255,255,255,0)", text: "#172026", radius: 0 },
  };
  return styles[kind] ?? { fill: "rgba(217,235,232,0.88)", stroke: "#1e6b66", text: "#173532", radius: 4 };
}

function roundedRect(ctx: CanvasRenderingContext2D, x: number, y: number, width: number, height: number, radius: number) {
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

function truncateLabel(value: string, maxLength: number) {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(1, maxLength - 3))}...`;
}

function escapeXml(value: string) {
  return value.replace(/[<>&"']/g, (char) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;", "'": "&apos;" })[char] ?? char);
}

function renderControlNode(control: DesignerControl) {
  const node = document.createElement("details");
  node.className = "control-node";
  node.open = true;

  const summary = document.createElement("summary");
  const label = control.name || control.caption || control.id;
  summary.textContent = `${label}${control.type ? ` · ${control.type}` : ""}`;
  node.append(summary);

  const props: Record<string, string> = {
    ID: control.id,
    Path: printableStreamName(control.path),
  };
  if (control.progId) props.ProgID = control.progId;
  if (control.caption) props.Caption = control.caption;
  if (control.bounds) props.Bounds = `${control.bounds.left}, ${control.bounds.top}, ${control.bounds.width}, ${control.bounds.height} ${control.bounds.unit} (${control.bounds.confidence})`;
  Object.assign(props, control.properties);
  node.append(renderKeyValueGrid(props));

  for (const child of control.children) {
    node.append(renderControlNode(child));
  }

  return node;
}

function countControls(controls: DesignerControl[]): number {
  return controls.reduce((sum, control) => sum + 1 + countControls(control.children), 0);
}
