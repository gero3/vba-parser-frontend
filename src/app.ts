import type { ExtractedFile } from "./types";
import { downloadBlob } from "./runtime";
import { extractOffice } from "./extract";
import { renderResults } from "./results-ui";
import { createZip } from "./zip-export";

export const fileInput = document.querySelector<HTMLInputElement>("#fileInput")!;
export const dropZone = document.querySelector<HTMLElement>("#dropZone")!;
export const clearButton = document.querySelector<HTMLButtonElement>("#clearButton")!;
export const downloadAllButton = document.querySelector<HTMLButtonElement>("#downloadAllButton")!;
export const statusText = document.querySelector<HTMLElement>("#statusText")!;
export const statusPanel = document.querySelector<HTMLElement>("#statusPanel")!;
export const summaryPanel = document.querySelector<HTMLElement>("#summaryPanel")!;
export const moduleCount = document.querySelector<HTMLElement>("#moduleCount")!;
export const formCount = document.querySelector<HTMLElement>("#formCount")!;
export const frxCount = document.querySelector<HTMLElement>("#frxCount")!;
export const results = document.querySelector<HTMLElement>("#results")!;

export let currentFiles: ExtractedFile[] = [];

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

export async function handleFile(file: File) {
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

export function downloadAll(files: ExtractedFile[]) {
  if (files.length === 0) return;
  downloadBlob("vba-extractor-output.zip", new Blob([createZip(files)], { type: "application/zip" }));
}

export function resetUi(clearInput = true) {
  currentFiles = [];
  results.innerHTML = "";
  summaryPanel.classList.add("hidden");
  setStatus("Waiting for a document.", "idle");
  if (clearInput) fileInput.value = "";
}

export function setStatus(message: string, state: "idle" | "busy" | "ready" | "warn" | "error") {
  statusText.textContent = message;
  statusPanel.dataset.state = state;
}
