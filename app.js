const MAX_BYTES = 8 * 1024 * 1024;
const MAX_FILES = 8;

const els = {
  drop: document.getElementById("drop"),
  file: document.getElementById("file"),
  pick: document.getElementById("pick"),
  clear: document.getElementById("clear"),
  status: document.getElementById("status"),
  error: document.getElementById("error"),
  files: document.getElementById("files"),
  preview: document.getElementById("preview"),
  run: document.getElementById("run"),
};

let items = [];

function setError(msg) {
  els.error.hidden = !msg;
  els.error.textContent = msg || "";
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function stem(name) {
  return name.replace(/\.[^.]+$/, "") || "sheet";
}

function decodeText(buf) {
  const bytes = new Uint8Array(buf);
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes);
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be").decode(bytes);
  }
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}

function looksCsv(text) {
  const sample = text.slice(0, 4000);
  if (!sample.trim()) return false;
  const lines = sample.split(/\r?\n/).filter((l) => l.trim()).slice(0, 8);
  if (!lines.length) return false;
  const seps = [",", ";", "\t", "|"];
  return seps.some((sep) => {
    const counts = lines.map((l) => l.split(sep).length);
    return counts[0] >= 2 && counts.filter((c) => c === counts[0]).length >= Math.min(2, counts.length);
  });
}

function tsvToCsv(text) {
  const wb = XLSX.read(text, { type: "string", FS: "\t", raw: true });
  return sheetToCsv(wb.Sheets[wb.SheetNames[0]]);
}

function sheetToCsv(sheet) {
  return XLSX.utils.sheet_to_csv(sheet, { FS: ",", RS: "\n" });
}

function workbookToCsvs(name, wb) {
  const sheets = wb.SheetNames || [];
  return sheets.map((sheetName) => {
    const csv = sheetToCsv(wb.Sheets[sheetName]);
    const file =
      sheets.length === 1
        ? `${stem(name)}.csv`
        : `${stem(name)}-${sheetName.replace(/[^\w.-]+/g, "_")}.csv`;
    return { name: file, csv, note: sheetName };
  });
}

async function fromSpreadsheet(name, buf) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  return workbookToCsvs(name, wb);
}

function csvFromXmlTables(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) return [];
  const tables = [...doc.getElementsByTagName("*")].filter((el) => el.localName.toLowerCase() === "table");
  const out = [];
  tables.forEach((table, idx) => {
    const rows = [];
    for (const tr of table.getElementsByTagName("*")) {
      if (tr.localName.toLowerCase() !== "tr" && tr.localName.toLowerCase() !== "row") continue;
      const cells = [];
      for (const cell of tr.children) {
        if (!["td", "th", "cell"].includes(cell.localName.toLowerCase())) continue;
        const index = cell.getAttributeNS("urn:schemas-microsoft-com:office:spreadsheet", "Index");
        if (index !== null) {
          const column = Number(index);
          if (!Number.isInteger(column) || column <= cells.length || column > 16384) {
            throw new Error("Invalid XML column index.");
          }
          while (cells.length < column - 1) cells.push('""');
        }
        cells.push(`"${(cell.textContent || "").replaceAll('"', '""')}"`);
      }
      if (cells.length) rows.push(cells.join(","));
    }
    if (rows.length) out.push({ name: `table-${idx + 1}.csv`, csv: rows.join("\n") + "\n", note: "xml table" });
  });
  return out;
}

async function fromNumbersZip(name, buf) {
  const zip = await JSZip.loadAsync(buf);
  const paths = Object.keys(zip.files).filter((p) => !zip.files[p].dir);
  const found = [];

  for (const path of paths) {
    const lower = path.toLowerCase();
    if (lower.endsWith(".csv") || lower.endsWith(".tsv")) {
      const text = await zip.files[path].async("string");
      found.push({
        name: `${stem(name)}-${path.split("/").pop().replace(/\.[^.]+$/, "")}.csv`,
        csv: lower.endsWith(".tsv") ? tsvToCsv(text) : text,
        note: path,
      });
    }
  }

  if (!found.length) {
    const xmlish = paths.filter((p) => /\.(xml|html)$/i.test(p) || /index\.xml/i.test(p));
    for (const path of xmlish) {
      const text = await zip.files[path].async("string");
      const tables = csvFromXmlTables(text);
      tables.forEach((t, i) => {
        t.name = `${stem(name)}-${i + 1}.csv`;
        found.push(t);
      });
    }
  }

  if (!found.length) {
    const xlsxPath = paths.find((p) => /\.(xlsx|xls|ods)$/i.test(p) || p.includes("xl/workbook"));
    if (xlsxPath && /\.(xlsx|xls|ods)$/i.test(xlsxPath)) {
      const inner = await zip.files[xlsxPath].async("arraybuffer");
      return fromSpreadsheet(name, inner);
    }
    if (paths.some((p) => p.startsWith("xl/"))) {
      return fromSpreadsheet(name, buf);
    }
  }

  if (!found.length) {
    throw new Error(
      `${name}: this Numbers file has no readable table. Export as Excel from Numbers once, then drop the .xlsx here.`
    );
  }
  return found;
}

async function convertFile(file) {
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf.slice(0, 4));
  const isZip = head[0] === 0x50 && head[1] === 0x4b;
  const lower = file.name.toLowerCase();

  if (isZip && (lower.endsWith(".numbers") || file.type.includes("numbers"))) {
    return fromNumbersZip(file.name, buf);
  }
  if (isZip || /\.(xlsx|xls|ods)$/i.test(file.name)) {
    return fromSpreadsheet(file.name, buf);
  }

  const text = decodeText(buf);
  if (looksCsv(text) || /\.(csv|tsv|txt)$/i.test(file.name)) {
    const csv = file.name.toLowerCase().endsWith(".tsv") ? tsvToCsv(text) : text;
    return [{ name: `${stem(file.name)}.csv`, csv, note: "already text" }];
  }

  throw new Error(`${file.name}: not a spreadsheet I can turn into CSV.`);
}

function render() {
  const ok = items.filter((i) => i.csvs && i.csvs.length);
  els.status.textContent = items.length
    ? `${items.length} file${items.length === 1 ? "" : "s"} · ${ok.reduce((n, i) => n + i.csvs.length, 0)} csv${ok.length === 1 ? "" : "s"} · files stayed in this tab`
    : "Waiting for a spreadsheet dump.";
  els.files.innerHTML = items
    .map((i) => {
      if (i.error) {
        return `<article><strong>${escapeHtml(i.file.name)}</strong><div class="warn">${escapeHtml(i.error)}</div></article>`;
      }
      const names = i.csvs.map((c) => c.name).join(", ");
      return `<article><strong>${escapeHtml(i.file.name)}</strong><div class="meta">${escapeHtml(names)}</div></article>`;
    })
    .join("");
  const first = ok[0]?.csvs[0]?.csv || "";
  els.preview.hidden = !first;
  els.preview.textContent = first ? first.split("\n").slice(0, 12).join("\n") : "";
  els.run.disabled = ok.length === 0;
}

async function addFiles(list) {
  setError("");
  for (const file of list) {
    if (items.length >= MAX_FILES) {
      setError(`Cap is ${MAX_FILES} files.`);
      break;
    }
    if (file.size > MAX_BYTES) {
      setError(`${file.name} is over 8 MB.`);
      continue;
    }
    const entry = { file, csvs: [], error: "" };
    try {
      entry.csvs = await convertFile(file);
    } catch (err) {
      entry.error = err && err.message ? err.message : String(err);
    }
    items.push(entry);
  }
  render();
}

async function downloadZip() {
  const ok = items.filter((i) => i.csvs && i.csvs.length);
  if (!ok.length) return;
  const zip = new JSZip();
  const used = new Set();
  for (const item of ok) {
    for (const out of item.csvs) {
      let name = out.name;
      let n = 2;
      while (used.has(name.toLowerCase())) {
        name = out.name.replace(/\.csv$/i, `-${n}.csv`);
        n += 1;
      }
      used.add(name.toLowerCase());
      zip.file(name, out.csv);
    }
  }
  const blob = await zip.generateAsync({ type: "blob" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "csvback.csv.zip";
  a.click();
  URL.revokeObjectURL(a.href);
}

els.pick.addEventListener("click", () => els.file.click());
els.file.addEventListener("change", () => {
  addFiles([...els.file.files]);
  els.file.value = "";
});
els.clear.addEventListener("click", () => {
  items = [];
  setError("");
  render();
});
els.run.addEventListener("click", () => downloadZip());

["dragenter", "dragover"].forEach((ev) => {
  els.drop.addEventListener(ev, (e) => {
    e.preventDefault();
    els.drop.classList.add("over");
  });
});
els.drop.addEventListener("dragleave", () => els.drop.classList.remove("over"));
els.drop.addEventListener("drop", (e) => {
  e.preventDefault();
  els.drop.classList.remove("over");
  addFiles([...e.dataTransfer.files]);
});

render();
