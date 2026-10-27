"use strict";

/*
 * XLSX.UP runs entirely in the browser. JSZip reads and rewrites the ZIP
 * container; we deliberately edit only worksheet XML, never cell values.
 */
const ui = {
  input: document.getElementById("fileInput"),
  dropzone: document.getElementById("dropzone"),
  change: document.getElementById("changeFile"),
  details: document.getElementById("fileDetails"),
  filename: document.getElementById("fileName"),
  filesize: document.getElementById("fileSize"),
  panel: document.getElementById("analysisPanel"),
  sheets: document.getElementById("sheetList"),
  count: document.getElementById("sheetCount"),
  summary: document.getElementById("summary"),
  action: document.getElementById("unlockButton"),
  actionLabel: document.getElementById("actionLabel"),
  status: document.getElementById("status"),
  progress: document.getElementById("progress"),
  progressFill: document.getElementById("progressFill"),
  download: document.getElementById("downloadLink")
};

const state = {
  file: null,
  zip: null,
  sheets: [],
  outputUrl: null,
  busy: false
};

const worksheetPattern = /^xl\/worksheets\/[^/]+\.xml$/i;
const excelMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function setStatus(message, kind = "") {
  ui.status.textContent = message;
  ui.status.className = "message" + (kind ? " " + kind : "");
}

function formatSize(size) {
  return size >= 1048576
    ? (size / 1048576).toFixed(2) + " MB"
    : Math.max(1, Math.ceil(size / 1024)) + " KB";
}

function releaseOutput() {
  if (state.outputUrl) URL.revokeObjectURL(state.outputUrl);
  state.outputUrl = null;
  ui.download.removeAttribute("href");
  ui.download.hidden = true;
}

function reset() {
  releaseOutput();
  state.file = null;
  state.zip = null;
  state.sheets = [];
  state.busy = false;
  ui.details.hidden = true;
  ui.panel.hidden = true;
  ui.sheets.replaceChildren();
  ui.summary.textContent = "";
  ui.count.textContent = "";
  ui.action.disabled = true;
  ui.actionLabel.textContent = "Unlock & download .xlsx";
  ui.progress.hidden = true;
  ui.progressFill.style.width = "0%";
  setStatus("Select an .xlsx file to get started.");
}

// Return the end index of a valid XML opening tag, ignoring > inside quotes.
// This avoids serializing the rest of the worksheet (which can alter formatting).
function openingTagEnd(xml, start) {
  let quote = "";
  for (let i = start; i < xml.length; i++) {
    const char = xml[i];
    if (quote) {
      if (char === quote) quote = "";
    } else if (char === "'" || char === '"') {
      quote = char;
    } else if (char === ">") {
      return i;
    }
  }
  throw new Error("A worksheet contains an incomplete XML tag.");
}

function escapeRegex(value) {
  // Legal XML tag names only need literal dots escaped for RegExp construction.
  return value.replace(/\./g, "\\.");
}

/**
 * Remove all <sheetProtection/> or <sheetProtection>...</sheetProtection>
 * elements, including namespace-prefixed forms, without changing other XML.
 * Returns the changed XML and number of protection elements removed.
 */
function stripSheetProtection(xml) {
  const opening = /<((?:[A-Za-z_][\w.-]*:)?sheetProtection)(?=[\s/>])/g;
  let found;
  let last = 0;
  let result = "";
  let removed = 0;

  while ((found = opening.exec(xml)) !== null) {
    const start = found.index;
    const tagEnd = openingTagEnd(xml, start + 1);
    let end = tagEnd + 1;
    const openTag = xml.slice(start, end);
    const selfClosing = /\/\s*>$/.test(openTag);

    if (!selfClosing) {
      const closing = new RegExp("</\\s*" + escapeRegex(found[1]) + "\\s*>", "g");
      closing.lastIndex = end;
      const closeMatch = closing.exec(xml);
      if (!closeMatch) throw new Error("Malformed sheetProtection XML element.");
      end = closeMatch.index + closeMatch[0].length;
    }

    result += xml.slice(last, start);
    last = end;
    removed++;
    opening.lastIndex = end;
  }
  if (!removed) return { xml, removed: 0 };
  return { xml: result + xml.slice(last), removed };
}

function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("Invalid workbook XML.");
  }
  return doc;
}

function elementsNamed(documentNode, localName) {
  return Array.from(documentNode.getElementsByTagName("*"))
    .filter((node) => node.localName === localName);
}

async function getSheetNameMap(zip) {
  const names = new Map();
  const workbook = zip.file("xl/workbook.xml");
  const rels = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbook || !rels) return names;

  try {
    const [bookText, relsText] = await Promise.all([
      workbook.async("string"),
      rels.async("string")
    ]);
    const book = parseXml(bookText);
    const relationships = parseXml(relsText);
    const targets = new Map();

    for (const relation of elementsNamed(relationships, "Relationship")) {
      if (relation.getAttribute("TargetMode") === "External") continue;
      const id = relation.getAttribute("Id");
      const target = relation.getAttribute("Target");
      if (!id || !target) continue;
      const url = new URL(target, "https://xlsx-up.local/xl/");
      targets.set(id, decodeURIComponent(url.pathname.slice(1)));
    }

    for (const sheet of elementsNamed(book, "sheet")) {
      const id = sheet.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id"
      ) || sheet.getAttribute("r:id");
      const path = targets.get(id);
      if (path) names.set(path, sheet.getAttribute("name") || path.split("/").pop());
    }
  } catch (error) {
    // Sheet names are optional metadata; processing by path still works.
    console.warn("Could not read worksheet names:", error);
  }
  return names;
}

function showSheetRows(rows) {
  ui.sheets.replaceChildren();
  const protectedRows = rows.filter((sheet) => sheet.protected);
  const shown = rows.length <= 12
    ? rows
    : (protectedRows.length ? protectedRows.slice(0, 40) : rows.slice(0, 12));

  for (const sheet of shown) {
    const row = document.createElement("div");
    row.className = "sheet-row" + (sheet.protected ? " protected" : "");
    const symbol = document.createElement("span");
    symbol.className = "sheet-symbol";
    symbol.setAttribute("aria-hidden", "true");
    symbol.textContent = sheet.protected ? "↗" : "✓";
    const name = document.createElement("span");
    name.className = "sheet-name";
    name.title = sheet.name;
    name.textContent = sheet.name;
    const status = document.createElement("span");
    status.className = "sheet-state";
    status.textContent = sheet.protected ? "Protected" : "Editable";
    row.append(symbol, name, status);
    ui.sheets.append(row);
  }
  if (shown.length < rows.length) {
    const note = document.createElement("div");
    note.className = "summary";
    note.textContent = "Showing " + shown.length + " of " + rows.length + " worksheets.";
    ui.sheets.append(note);
  }

  ui.count.textContent = rows.length + (rows.length === 1 ? " sheet" : " sheets");
  ui.summary.replaceChildren();
  const count = document.createElement("strong");
  count.textContent = protectedRows.length + (protectedRows.length === 1 ? " protected sheet" : " protected sheets");
  ui.summary.append(count, document.createTextNode(
    protectedRows.length
      ? " detected. All protected worksheets will be made editable in a new copy."
      : " detected. This file has no worksheet protection to remove."
  ));
}

async function loadFile(file) {
  if (state.busy) return;
  reset();
  if (!file) return;

  if (!/\.xlsx$/i.test(file.name)) {
    setStatus("Please choose an .xlsx file. Legacy .xls, .xlsm and .xlsb are not supported.", "error");
    return;
  }
  if (typeof JSZip === "undefined") {
    setStatus("The ZIP library couldn't load. Please reconnect to the internet and refresh.", "error");
    return;
  }

  state.busy = true;
  state.file = file;
  ui.filename.textContent = file.name;
  ui.filesize.textContent = formatSize(file.size) + " · Original will remain unchanged";
  ui.details.hidden = false;
  setStatus("Reading workbook locally…");

  try {
    const bytes = await file.arrayBuffer();
    const header = new Uint8Array(bytes, 0, Math.min(bytes.byteLength, 4));
    if (header[0] !== 0x50 || header[1] !== 0x4b) {
      throw new Error("This is not a standard .xlsx ZIP archive. Password-to-open encrypted files are not supported.");
    }

    const zip = await JSZip.loadAsync(bytes);
    if (!zip.file("[Content_Types].xml") || !zip.file("xl/workbook.xml")) {
      throw new Error("The selected file is not a valid Excel .xlsx workbook.");
    }

    const paths = Object.keys(zip.files)
      .filter((path) => worksheetPattern.test(path) && !zip.files[path].dir)
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    if (!paths.length) throw new Error("No worksheet XML files were found inside this workbook.");
    const nameMap = await getSheetNameMap(zip);
    const sheets = [];

    for (let i = 0; i < paths.length; i++) {
      setStatus("Checking worksheets… " + (i + 1) + " / " + paths.length);
      const path = paths[i];
      const xml = await zip.file(path).async("string");
      const removed = stripSheetProtection(xml).removed;
      sheets.push({
        path,
        name: nameMap.get(path) || path.split("/").pop(),
        protected: removed > 0
      });
    }

    state.zip = zip;
    state.sheets = sheets;
    ui.panel.hidden = false;
    showSheetRows(sheets);
    const count = sheets.filter((sheet) => sheet.protected).length;
    ui.action.disabled = count === 0;

    if (count) {
      const signed = Object.keys(zip.files).some((p) => p.startsWith("_xmlsignatures/"));
      setStatus(
        count + (count === 1 ? " protected worksheet found." : " protected worksheets found.") +
        (signed ? " Warning: this file has a digital signature that may become invalid." : "")
      );
    } else {
      setStatus("No worksheet protection found. Nothing needs to be changed.");
    }
  } catch (error) {
    state.zip = null;
    state.sheets = [];
    ui.panel.hidden = true;
    ui.action.disabled = true;
    console.error("Workbook analysis failed:", error);
    const message = /encrypted|password/i.test(String(error))
      ? "This workbook is encrypted or password-to-open protected. This tool only removes worksheet editing protection."
      : (error && error.message) || "Unable to read this workbook.";
    setStatus(message, "error");
  } finally {
    state.busy = false;
  }
}

async function unlockAndDownload() {
  if (state.busy) return;
  if (state.outputUrl) {
    ui.download.click();
    return;
  }
  if (!state.zip || !state.file) return;
  state.busy = true;
  ui.action.disabled = true;
  ui.progress.hidden = false;
  ui.progressFill.style.width = "0%";
  ui.actionLabel.textContent = "Preparing workbook…";
  setStatus("Removing sheet protection…");

  try {
    for (const sheet of state.sheets.filter((item) => item.protected)) {
      const entry = state.zip.file(sheet.path);
      const xml = await entry.async("string");
      const transformed = stripSheetProtection(xml);
      if (!transformed.removed) {
        throw new Error("Protection could not be found again in " + sheet.name);
      }
      state.zip.file(sheet.path, transformed.xml, {
        date: entry.date,
        createFolders: false
      });
    }

    const blob = await state.zip.generateAsync(
      { type: "blob", mimeType: excelMime, compression: "DEFLATE", compressionOptions: { level: 6 } },
      (progress) => {
        ui.progressFill.style.width = Math.round(progress.percent) + "%";
        setStatus("Creating your .xlsx file… " + Math.round(progress.percent) + "%");
      }
    );

    const filename = state.file.name.replace(/\.xlsx$/i, "") + "-unprotected.xlsx";
    state.outputUrl = URL.createObjectURL(blob);
    ui.download.href = state.outputUrl;
    ui.download.download = filename;
    ui.download.hidden = false;
    ui.actionLabel.textContent = "Download .xlsx again";
    ui.progressFill.style.width = "100%";
    setStatus("Done! The editable copy is ready. Your original file was not changed.", "success");
    ui.download.click();
  } catch (error) {
    console.error("Workbook export failed:", error);
    ui.actionLabel.textContent = "Try again";
    setStatus((error && error.message) || "Could not prepare the output .xlsx.", "error");
  } finally {
    state.busy = false;
    ui.action.disabled = false;
  }
}

ui.dropzone.addEventListener("click", () => {
  if (!state.busy) ui.input.click();
});
ui.dropzone.addEventListener("keydown", (event) => {
  if ((event.key === "Enter" || event.key === " ") && !state.busy) {
    event.preventDefault();
    ui.input.click();
  }
});
ui.change.addEventListener("click", () => {
  if (!state.busy) ui.input.click();
});
ui.input.addEventListener("change", () => {
  const file = ui.input.files?.[0];
  ui.input.value = ""; // Allow uploading the same file again.
  if (file) loadFile(file);
});

for (const eventName of ["dragenter", "dragover"]) {
  ui.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (state.busy) return;
    ui.dropzone.classList.add("dragging");
    if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
  });
}
for (const eventName of ["dragleave", "dragend"]) {
  ui.dropzone.addEventListener(eventName, () => {
    ui.dropzone.classList.remove("dragging");
  });
}
ui.dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  ui.dropzone.classList.remove("dragging");
  if (!state.busy && event.dataTransfer?.files?.length) {
    loadFile(event.dataTransfer.files[0]);
  }
});
ui.action.addEventListener("click", unlockAndDownload);
window.addEventListener("beforeunload", releaseOutput);
