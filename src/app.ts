const fileInput = document.querySelector<HTMLInputElement>("#fileInput")!;
const dropZone = document.querySelector<HTMLElement>("#dropZone")!;
const clearButton = document.querySelector<HTMLButtonElement>("#clearButton")!;
const downloadAllButton = document.querySelector<HTMLButtonElement>("#downloadAllButton")!;
const statusText = document.querySelector<HTMLElement>("#statusText")!;
const statusPanel = document.querySelector<HTMLElement>("#statusPanel")!;
const summaryPanel = document.querySelector<HTMLElement>("#summaryPanel")!;
const moduleCount = document.querySelector<HTMLElement>("#moduleCount")!;
const formCount = document.querySelector<HTMLElement>("#formCount")!;
const frxCount = document.querySelector<HTMLElement>("#frxCount")!;
const results = document.querySelector<HTMLElement>("#results")!;

let currentFiles: ExtractedFile[] = [];

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) void handleFile(file);
});

clearButton.addEventListener("click", resetUi);
downloadAllButton.addEventListener("click", () => downloadAll(currentFiles));

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  const file = event.dataTransfer?.files[0];
  if (file) void handleFile(file);
});

async function handleFile(file: File) {
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

    const media = extracted.filter((item) => item.kind === "media").length;
    setStatus(`Extracted ${extracted.length} item${extracted.length === 1 ? "" : "s"}${media ? `, including ${media} embedded media file${media === 1 ? "" : "s"}` : ""}.`, "ready");
  } catch (error) {
    console.error(error);
    setStatus(error instanceof Error ? error.message : "Could not parse this file.", "error");
  }
}
