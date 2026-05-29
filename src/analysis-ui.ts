function renderAnalysis(analysis: BinaryAnalysis, bytes: Uint8Array) {
  const section = document.createElement("section");
  section.className = "analysis";

  const title = document.createElement("h3");
  title.textContent = analysis.title;
  section.append(title);

  const summary = document.createElement("dl");
  summary.className = "analysis-grid";
  for (const item of analysis.summary) {
    const label = document.createElement("dt");
    label.textContent = item.label;
    const value = document.createElement("dd");
    value.textContent = item.value;
    summary.append(label, value);
  }
  section.append(summary);

  section.append(renderAnalysisTable("Detected structures", ["Offset", "Type", "Details"], analysis.signatures.map((item) => [toHex(item.offset), item.label, item.detail])));
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
  details.addEventListener(
    "toggle",
    () => {
      if (!details.open || pre.dataset.loaded === "true") return;
      pre.textContent = bytes.length ? hexDump(bytes, 0, bytes.length) : "(empty stream)";
      pre.dataset.loaded = "true";
    },
    { once: false },
  );
  details.append(summaryText, pre);
  section.append(details);

  return section;
}

function renderOFormsAnalysis(oforms: OFormsAnalysis) {
  const wrapper = document.createElement("section");
  wrapper.className = "oforms-analysis";

  const heading = document.createElement("h4");
  heading.textContent = `MS-OFORMS parse: ${oforms.kind}`;
  wrapper.append(heading);

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  for (const header of ["Offset", "Record", "Size", "Properties"]) {
    const cell = document.createElement("th");
    cell.textContent = header;
    headerRow.append(cell);
  }
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const record of oforms.records) {
    const row = document.createElement("tr");
    for (const value of [toHex(record.offset), record.type, formatBytes(record.size)]) {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.append(cell);
    }

    const propertiesCell = document.createElement("td");
    const list = document.createElement("dl");
    list.className = "property-list";
    for (const property of record.properties) {
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
    for (const note of oforms.notes) {
      const item = document.createElement("li");
      item.textContent = note;
      notes.append(item);
    }
    wrapper.append(notes);
  }

  return wrapper;
}

function renderLazyGuids(analysis: BinaryAnalysis, bytes: Uint8Array) {
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

  details.addEventListener("toggle", () => {
    if (!details.open || container.dataset.loaded === "true") return;
    container.textContent = "";
    container.append(renderAnalysisTable("", ["Offset", "Value", "Label"], collectGuids(bytes).map((item) => [toHex(item.offset), item.value, labelGuid(item.value) ?? ""])));
    container.dataset.loaded = "true";
  });

  details.append(summary, container);
  wrapper.append(details);
  return wrapper;
}

function renderLazyStrings(analysis: BinaryAnalysis, bytes: Uint8Array) {
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

  details.addEventListener("toggle", () => {
    if (!details.open || container.dataset.loaded === "true") return;
    container.textContent = "";
    container.append(renderStringTable(collectStrings(bytes)));
    container.dataset.loaded = "true";
  });

  details.append(summary, container);
  wrapper.append(details);
  return wrapper;
}

function renderStringTable(strings: ExtractedString[]) {
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  for (const header of ["Offset", "Encoding", "Value"]) {
    const cell = document.createElement("th");
    cell.textContent = header;
    headerRow.append(cell);
  }
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const item of strings) {
    const row = document.createElement("tr");
    for (const value of [toHex(item.offset), item.encoding, item.value]) {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.append(cell);
    }
    tbody.append(row);
  }
  table.append(tbody);
  return table;
}

function renderAnalysisTable(title: string, headers: string[], rows: string[][]) {
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
  for (const header of headers) {
    const cell = document.createElement("th");
    cell.textContent = header;
    headerRow.append(cell);
  }
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const row of rows) {
    const tableRow = document.createElement("tr");
    for (const value of row) {
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
