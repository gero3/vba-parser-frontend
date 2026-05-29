function renderResults(files: ExtractedFile[]) {
  results.innerHTML = "";
  summaryPanel.classList.toggle("hidden", files.length === 0);
  moduleCount.textContent = String(files.filter((file) => file.kind === "vba").length);
  formCount.textContent = String(files.filter((file) => file.kind === "frm").length);
  frxCount.textContent = String(files.filter((file) => file.kind === "frx").length);

  for (const group of buildResultGroups(files)) {
    results.append(renderResultGroup(group));
  }
}

function buildResultGroups(files: ExtractedFile[]) {
  const groups = new Map<string, ResultGroup>();

  for (const file of files) {
    const owner = getResultOwner(file);
    let group = groups.get(owner);
    if (!group) {
      group = { name: owner, code: [], resources: [], media: [], other: [] };
      groups.set(owner, group);
    }

    if (file.kind === "vba" || file.kind === "frm") group.code.push(file);
    else if (file.kind === "frx") group.resources.push(file);
    else if (file.kind === "media") group.media.push(file);
    else group.other.push(file);
  }

  for (const group of groups.values()) {
    group.designer = buildDesignerSummary(group);
  }

  return [...groups.values()].sort((a, b) => {
    const aHasForm = a.code.some((file) => file.kind === "frm") || a.resources.length > 0;
    const bHasForm = b.code.some((file) => file.kind === "frm") || b.resources.length > 0;
    if (aHasForm !== bHasForm) return aHasForm ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
}

function getResultOwner(file: ExtractedFile) {
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

function extractInternalStreamPath(sourcePath: string) {
  const marker = "vbaProject.bin/";
  const markerIndex = sourcePath.indexOf(marker);
  if (markerIndex < 0) return undefined;
  return sourcePath.slice(markerIndex + marker.length).replace(/\s+@\s+0x[0-9a-f]+.*$/i, "").replace(/\s+\(recovered by scan\)$/i, "");
}

function renderResultGroup(group: ResultGroup) {
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

function describeGroup(group: ResultGroup) {
  const parts = [
    countLabel(group.code.length, "code file"),
    countLabel(group.resources.length, "resource stream"),
    countLabel(group.media.length, "media file"),
    countLabel(group.other.length, "other file"),
  ].filter(Boolean);
  return parts.join(" · ") || "No extracted files";
}

function countLabel(count: number, label: string) {
  if (!count) return "";
  return `${count} ${label}${count === 1 ? "" : "s"}`;
}

function appendGroupSection(parent: HTMLElement, title: string, files: ExtractedFile[]) {
  if (files.length === 0) return;

  const section = document.createElement("section");
  section.className = "group-subsection";
  const heading = document.createElement("h3");
  heading.textContent = title;
  section.append(heading);

  const list = document.createElement("div");
  list.className = "group-items";
  for (const file of files.sort((a, b) => kindRank(a.kind) - kindRank(b.kind) || getDisplayPath(a).localeCompare(getDisplayPath(b)))) {
    list.append(renderFileCard(file));
  }
  section.append(list);
  parent.append(section);
}

function renderFileCard(file: ExtractedFile) {
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
  downloadButton.addEventListener("click", () => downloadFile(file));
  header.append(downloadButton);
  card.append(header);

  if (file.text) {
    const pre = document.createElement("pre");
    pre.textContent = file.text;
    card.append(pre);
  } else if (file.analysis) {
    card.append(renderAnalysis(file.analysis, file.bytes));
  } else if (canPreviewImage(file.mimeType)) {
    const url = URL.createObjectURL(new Blob([file.bytes], { type: file.mimeType }));
    const image = document.createElement("img");
    image.src = url;
    image.alt = file.name;
    image.className = "preview-image";
    image.addEventListener("load", () => URL.revokeObjectURL(url), { once: true });
    card.append(image);
  } else {
    const binaryNote = document.createElement("p");
    binaryNote.className = "binary-note";
    binaryNote.textContent = "Binary stream extracted. Download it to inspect or recover additional form resources.";
    card.append(binaryNote);
  }

  return card;
}

function canPreviewImage(mimeType: string) {
  return ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/x-icon"].includes(mimeType);
}
