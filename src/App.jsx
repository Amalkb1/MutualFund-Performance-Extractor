import { useMemo, useRef, useState } from "react";
import seedCsv from "../data/mf_master_dataset.csv?raw";

const STORAGE_KEY = "fundfolio-master-v1";
const REQUIRED_COLUMNS = [
  "Scheme Name",
  "Benchmark",
  "Return 1 Year (%) Regular",
  "Return 1 Year (%) Benchmark",
  "Return 3 Year (%) Regular",
  "Return 3 Year (%) Benchmark",
  "Return 5 Year (%) Regular",
  "Return 5 Year (%) Benchmark",
  "Daily AUM (Cr.)",
];
const MASTER_COLUMNS = [
  ...REQUIRED_COLUMNS,
  "Source File",
];
const NUMERIC_COLUMNS = REQUIRED_COLUMNS.filter((column) =>
  column.startsWith("Return ") ||
  column.endsWith("(Cr.)"),
);
const AUM_COLUMN = "Daily AUM (Cr.)";
const AUM_SOURCE_COLUMNS = [AUM_COLUMN, "Previous Month Average AUM (Cr.)"];
const PAGE_SIZE = 10;

function cleanHeader(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function parseCsv(text) {
  const records = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"' && field === "") {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n" || character === "\r") {
      row.push(field);
      records.push(row);
      row = [];
      field = "";
      if (character === "\r" && text[index + 1] === "\n") index += 1;
    } else {
      field += character;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    records.push(row);
  }

  const headers = records.shift() ?? [];
  return records.filter((record) => record.some((value) => value !== ""))
    .map((record) => Object.fromEntries(
      headers.map((header, index) => [header, record[index] ?? ""]),
    ));
}

function getInitialMaster() {
  const seeded = parseCsv(seedCsv).map((row) => normalizeMasterRow(row));
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (!saved) return { rows: seeded, warning: "" };

    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed)) throw new Error("Saved data is not a row list.");
    return { rows: parsed.map((row) => normalizeMasterRow(row)), warning: "" };
  } catch (error) {
    console.error("Could not load saved fund data.", error);
    return {
      rows: seeded,
      warning:
        "Saved browser data could not be loaded. The bundled dataset is shown instead.",
    };
  }
}

function normalizeMasterRow(row) {
  const normalized = Object.fromEntries(REQUIRED_COLUMNS.map((column) => {
    let value = row[column] ?? "";
    if (column === "Daily AUM (Cr.)" && !value && row["AUM Type"] === "Daily") {
      value = row["Previous Month Average AUM (Cr.)"] ?? "";
    }
    return [column, value];
  }));
  normalized["Source File"] = row["Source File"] ?? "";
  return normalized;
}

function parseNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") return "";
  const cleaned = String(value).trim().replace(/[\^*†‡]+$/, "").replace(/,/g, "");
  return Number.isFinite(Number(cleaned)) ? cleaned : "";
}

async function extractExcel(file) {
  const XLSX = await import("@e965/xlsx");
  const workbook = XLSX.read(new Uint8Array(await file.arrayBuffer()), {
    type: "array",
    cellDates: true,
    cellText: false,
  });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The workbook does not contain a worksheet.");

  const sheetRange = sheet["!ref"];
  if (!sheetRange) throw new Error("The first worksheet is empty.");
  const range = XLSX.utils.decode_range(sheetRange);
  const getRawValue = (rowIndex, columnIndex) => {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
    return cell?.v ?? "";
  };
  const headerIndex = Array.from(
    { length: range.e.r - range.s.r + 1 },
    (_, index) => range.s.r + index,
  ).find((rowIndex) => cleanHeader(getRawValue(rowIndex, 0)) === "Scheme Name");
  if (headerIndex === undefined) {
    throw new Error("No 'Scheme Name' header row found. Check that this is the right file.");
  }

  const headers = Array.from(
    { length: range.e.c + 1 },
    (_, columnIndex) => cleanHeader(getRawValue(headerIndex, columnIndex)),
  );
  const headerMap = new Map();
  headers.forEach((header, index) => {
    if (header && !headerMap.has(header)) headerMap.set(header, index);
  });

  const aumSource = AUM_SOURCE_COLUMNS.find((column) => headerMap.has(column));
  const warnings = [];
  const missing = REQUIRED_COLUMNS.filter((column) =>
    column === AUM_COLUMN ? !aumSource : !headerMap.has(column),
  );
  if (missing.length) warnings.push(`Missing columns (left blank): ${missing.join(", ")}`);
  if (aumSource && aumSource !== AUM_COLUMN) {
    warnings.push(`Using "${aumSource}" for the "${AUM_COLUMN}" master column.`);
  }

  const sourceFile = file.name.replace(/\.xlsx?$/i, "");
  const rows = Array.from(
    { length: range.e.r - headerIndex },
    (_, index) => headerIndex + index + 1,
  )
    .filter((rowIndex) => {
      const scheme = getRawValue(rowIndex, headerMap.get("Scheme Name"));
      const benchmark = getRawValue(rowIndex, headerMap.get("Benchmark"));
      return String(scheme ?? "").trim() && String(benchmark ?? "").trim();
    })
    .map((rowIndex) => {
      const result = Object.fromEntries(MASTER_COLUMNS.map((column) => [column, ""]));
      for (const column of REQUIRED_COLUMNS) {
        const sourceColumn = column === AUM_COLUMN ? aumSource : column;
        const value = sourceColumn ? getRawValue(rowIndex, headerMap.get(sourceColumn)) : "";
        if (column === "Scheme Name" || column === "Benchmark") {
          result[column] = String(value ?? "").trim();
        } else {
          result[column] = NUMERIC_COLUMNS.includes(column) ? parseNumber(value) : String(value ?? "");
        }
      }
      result["Source File"] = sourceFile;
      return result;
    });

  if (!rows.length) warnings.push("No fund rows with both a scheme name and benchmark were found.");
  return { rows, warnings };
}

function mergeRows(master, incoming) {
  const byScheme = new Map();
  for (const row of [...master, ...incoming]) {
    byScheme.set(String(row["Scheme Name"]).trim().toLowerCase(), normalizeMasterRow(row));
  }
  return [...byScheme.values()].sort((a, b) =>
    String(a["Scheme Name"]).localeCompare(String(b["Scheme Name"])),
  );
}

function filterRows(rows, query, categories) {
  const selectedCategories = new Set(categories);
  const categoryRows = selectedCategories.size
    ? rows.filter((row) => selectedCategories.has(row["Source File"]))
    : rows;
  const terms = query.split(",").map((term) => term.trim().toLowerCase()).filter(Boolean);
  if (!terms.length) return categoryRows;

  return categoryRows.filter((row) => {
    const name = String(row["Scheme Name"]).trim().toLowerCase();
    return terms.some((term) => {
      if (categoryRows.some((candidate) =>
        String(candidate["Scheme Name"]).trim().toLowerCase() === term,
      )) {
        return name === term;
      }
      return term.split(/\s+/).every((word) => name.includes(word));
    });
  });
}

function formatValue(value) {
  if (value === "" || value === null || value === undefined) return "—";
  return String(value);
}

function csvForRows(rows, delimiter = ",") {
  const escape = (value) => {
    const text = value === null || value === undefined ? "" : String(value);
    return /[",\r\n\t]/.test(text)
      ? `"${text.replace(/"/g, '""')}"`
      : text;
  };
  return [
    REQUIRED_COLUMNS.map(escape).join(delimiter),
    ...rows.map((row) => REQUIRED_COLUMNS.map((column) => escape(row[column])).join(delimiter)),
  ].join("\r\n");
}

function downloadRows(rows, filename) {
  const blob = new Blob(["\uFEFF", csvForRows(rows)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function Icon({ name, size = 18 }) {
  const paths = {
    grid: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    search: <><circle cx="10.8" cy="10.8" r="6.8" /><path d="m16 16 4.5 4.5" /></>,
    upload: <><path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5" /><path d="M5 14v5a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-5" /></>,
    download: <><path d="M12 4v12m0 0 4.5-4.5M12 16l-4.5-4.5" /><path d="M5 15v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4" /></>,
    copy: <><rect x="8" y="8" width="12" height="12" rx="2" /><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" /></>,
    chart: <><path d="M4 19V5m0 14h17" /><path d="m7 14 4-4 3 2 5-6" /></>,
    file: <><path d="M13 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V10Z" /><path d="M13 3v7h7M8 15h8m-8 3h8" /></>,
    arrow: <><path d="M5 12h14m-6-6 6 6-6 6" /></>,
  };
  return (
    <svg aria-hidden="true" width={size} height={size} viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      {paths[name]}
    </svg>
  );
}

function MetricCard({ label, value, detail, icon }) {
  return (
    <article className="metric-card">
      <div className={`metric-icon metric-icon-${icon}`}><Icon name={icon} /></div>
      <div className="metric-label">{label}</div>
      <div className="metric-value">{value}</div>
      <div className="metric-detail">{detail}</div>
    </article>
  );
}

function FundTable({ rows, selected, onToggle }) {
  return (
    <div className="table-scroll">
      <table>
        <thead>
          <tr>
            {onToggle && <th className="check-cell"><span className="sr-only">Select</span></th>}
            {REQUIRED_COLUMNS.map((column) => <th key={column}>{column}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const name = row["Scheme Name"];
            return (
              <tr key={`${name}-${row["Source File"]}`}>
                {onToggle && (
                  <td className="check-cell">
                    <input type="checkbox" aria-label={`Select ${name}`}
                      checked={selected.has(name)} onChange={() => onToggle(name)} />
                  </td>
                )}
                {REQUIRED_COLUMNS.map((column) => (
                  <td
                    className={column === "Scheme Name" ? "scheme-name" : column === "Benchmark" ? "benchmark" : "number-cell"}
                    key={column}
                  >
                    {formatValue(row[column])}
                  </td>
                ))}
              </tr>
            );
          })}
          {!rows.length && (
          <tr><td className="empty-cell" colSpan={REQUIRED_COLUMNS.length + (onToggle ? 1 : 0)}>
              No funds match your filters.
            </td></tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function Pagination({ page, pageCount, onChange }) {
  if (pageCount <= 1) return null;
  return (
    <div className="pagination">
      <span>Page {page} of {pageCount}</span>
      <div className="pagination-actions">
        <button className="button button-quiet button-small" disabled={page <= 1}
          onClick={() => onChange(page - 1)}>Previous</button>
        <button className="button button-quiet button-small" disabled={page >= pageCount}
          onClick={() => onChange(page + 1)}>Next</button>
      </div>
    </div>
  );
}

export default function App() {
  const [initialData] = useState(getInitialMaster);
  const [master, setMaster] = useState(initialData.rows);
  const [activeView, setActiveView] = useState("master");
  const [searchText, setSearchText] = useState("");
  const [selectedCategories, setSelectedCategories] = useState([]);
  const [selectedSchemes, setSelectedSchemes] = useState(new Set());
  const [files, setFiles] = useState([]);
  const [importMessages, setImportMessages] = useState([]);
  const [mode, setMode] = useState("add");
  const [notice, setNotice] = useState(null);
  const [storageWarning, setStorageWarning] = useState(initialData.warning);
  const [masterPage, setMasterPage] = useState(1);
  const [resultPage, setResultPage] = useState(1);
  const [isBuilding, setIsBuilding] = useState(false);
  const fileInputRef = useRef(null);

  const categories = useMemo(
    () => [...new Set(master.map((row) => row["Source File"]).filter(Boolean))].sort(),
    [master],
  );
  const matches = useMemo(
    () => filterRows(master, searchText, selectedCategories),
    [master, searchText, selectedCategories],
  );
  const results = useMemo(
    () => selectedSchemes.size
      ? matches.filter((row) => selectedSchemes.has(row["Scheme Name"]))
      : matches,
    [matches, selectedSchemes],
  );
  const requestedNames = searchText.split(",").map((name) => name.trim()).filter(Boolean);
  const notFound = requestedNames.filter((term) =>
    !filterRows(master, term, selectedCategories).length,
  );
  const masterPages = Math.max(1, Math.ceil(master.length / PAGE_SIZE));
  const resultPages = Math.max(1, Math.ceil(results.length / PAGE_SIZE));
  const masterPageRows = master.slice((masterPage - 1) * PAGE_SIZE, masterPage * PAGE_SIZE);
  const resultPageRows = results.slice((resultPage - 1) * PAGE_SIZE, resultPage * PAGE_SIZE);

  function persistRows(rows, successMessage) {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
      setMaster(rows);
      setStorageWarning("");
      setNotice({ type: "success", text: successMessage });
      setMasterPage(1);
    } catch (error) {
      console.error("Could not save fund data in this browser.", error);
      setStorageWarning(
        "Your changes were not saved in this browser. Export the master CSV before leaving this page.",
      );
      setNotice({ type: "error", text: "Could not save the updated master dataset." });
    }
  }

  async function buildMaster() {
    if (!files.length || isBuilding) return;
    setIsBuilding(true);
    setNotice(null);
    const parsedFrames = [];
    const messages = [];

    for (const file of files) {
      try {
        const { rows, warnings } = await extractExcel(file);
        if (!rows.length) {
          messages.push({ type: "error", text: `${file.name}: no usable fund rows were found.` });
          continue;
        }
        parsedFrames.push(rows);
        messages.push({ type: "success", text: `${file.name}: ${rows.length} schemes` });
        warnings.forEach((warning) => messages.push({
          type: "warning",
          text: `${file.name}: ${warning}`,
        }));
      } catch (error) {
        console.error(`Could not read ${file.name}.`, error);
        messages.push({
          type: "error",
          text: `${file.name}: could not read the workbook (${error.message || "unknown error"}).`,
        });
      }
    }

    if (parsedFrames.length) {
      const incoming = parsedFrames.flat();
      const base = mode === "replace" ? [] : master;
      const merged = mergeRows(base, incoming);
      persistRows(merged, `${incoming.length} imported scheme rows processed. Master now has ${merged.length} schemes.`);
      if (messages.some((message) => message.type !== "success")) {
        setNotice({
          type: "warning",
          text: "The master was updated with readable files. Review the import messages below.",
        });
      }
    } else {
      setNotice({ type: "error", text: "No files could be imported. Check the messages below." });
    }
    setFiles([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
    setImportMessages(messages);
    setIsBuilding(false);
  }

  function toggleCategory(category) {
    setSelectedCategories((current) => current.includes(category)
      ? current.filter((item) => item !== category)
      : [...current, category]);
    setSelectedSchemes(new Set());
    setResultPage(1);
  }

  function toggleScheme(name) {
    setSelectedSchemes((current) => {
      const next = new Set(current);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
    setResultPage(1);
  }

  async function copyResults() {
    try {
      await navigator.clipboard.writeText(csvForRows(results, "\t"));
      setNotice({ type: "success", text: `${results.length} rows copied for Excel.` });
    } catch (error) {
      console.error("Could not copy fund data to the clipboard.", error);
      const textarea = document.createElement("textarea");
      textarea.value = csvForRows(results, "\t");
      textarea.setAttribute("readonly", "");
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      try {
        textarea.select();
        const copied = document.execCommand("copy");
        setNotice(copied
          ? { type: "success", text: `${results.length} rows copied for Excel.` }
          : { type: "error", text: "Clipboard access was blocked. Use Download CSV instead." });
      } catch (fallbackError) {
        console.error("Clipboard fallback failed.", fallbackError);
        setNotice({ type: "error", text: "Clipboard access was blocked. Use Download CSV instead." });
      } finally {
        textarea.remove();
      }
    }
  }

  const totalSourceLabel = categories.length === 1 ? "source file" : "source files";

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#" onClick={(event) => {
          event.preventDefault();
          setActiveView("master");
        }}>
          <span className="brand-mark"><span /></span>
          <span className="brand-name">fundfolio<span>.</span></span>
        </a>

        <div className="sidebar-label">WORKSPACE</div>
        <nav className="side-nav" aria-label="Main navigation">
          <button className={activeView === "master" ? "nav-item active" : "nav-item"}
            onClick={() => setActiveView("master")}>
            <Icon name="grid" /><span>Master data</span>
          </button>
          <button className={activeView === "explore" ? "nav-item active" : "nav-item"}
            onClick={() => setActiveView("explore")}>
            <Icon name="search" /><span>Fund explorer</span>
          </button>
        </nav>

        <div className="sidebar-bottom">
          <div className="privacy-card">
            <span className="privacy-dot" />
            <div><strong>Private by design</strong><span>Your data stays in this browser</span></div>
          </div>
          <div className="sidebar-footer">MUTUAL FUND PERFORMANCE<br />PERSONAL WORKSPACE</div>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div className="breadcrumb"><span>Workspace</span><span className="breadcrumb-slash">/</span>
            <strong>{activeView === "master" ? "Master data" : "Fund explorer"}</strong>
          </div>
          <button className="button button-primary topbar-button" onClick={() => {
            setActiveView("master");
            document.getElementById("upload-panel")?.scrollIntoView({ behavior: "smooth" });
          }}>
            <Icon name="upload" size={16} /> Import data
          </button>
        </header>

        <div className="page-content">
          <section className="page-heading">
            <div>
              <div className="eyebrow">YOUR FUND DATA, AT A GLANCE</div>
              <h1>{activeView === "master" ? "Mutual fund performance" : "Explore funds"}</h1>
              <p>{activeView === "master"
                ? "Build your performance library and keep every scheme in one place."
                : "Find, compare, and export the fund performance you need."}</p>
            </div>
            <div className="dataset-status"><span className="status-dot" /> Dataset ready</div>
          </section>

          {storageWarning && <div className="notice notice-warning" role="alert">{storageWarning}</div>}
          {notice && <div className={`notice notice-${notice.type}`} role="status">
            <span>{notice.text}</span><button aria-label="Dismiss message" onClick={() => setNotice(null)}>×</button>
          </div>}

          <section className="metrics-grid" aria-label="Dataset overview">
            <MetricCard label="Total schemes" value={master.length.toLocaleString("en-IN")}
              detail="In your performance library" icon="grid" />
            <MetricCard label="Categories" value={categories.length.toLocaleString("en-IN")}
              detail="Distinct source files" icon="chart" />
            <MetricCard label="Last updated" value={master.length ? "Ready" : "—"}
              detail="Data is stored in this browser" icon="file" />
          </section>

          <div className="view-tabs" role="tablist" aria-label="Dataset views">
            <button role="tab" aria-selected={activeView === "master"}
              className={activeView === "master" ? "view-tab selected" : "view-tab"}
              onClick={() => setActiveView("master")}><Icon name="grid" size={16} /> Master data</button>
            <button role="tab" aria-selected={activeView === "explore"}
              className={activeView === "explore" ? "view-tab selected" : "view-tab"}
              onClick={() => setActiveView("explore")}><Icon name="search" size={16} /> Fund explorer</button>
          </div>

          {activeView === "master" ? (
            <div className="content-stack">
              <section className="panel upload-panel" id="upload-panel">
                <div className="panel-heading">
                  <div className="panel-title-wrap">
                    <div className="panel-icon"><Icon name="upload" /></div>
                    <div><h2>Import performance data</h2>
                      <p>Upload Excel reports from Mutual Sahi Hai to build your master.</p></div>
                  </div>
                </div>
                <label className="dropzone" htmlFor="excel-files"
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={(event) => {
                    event.preventDefault();
                    setFiles(Array.from(event.dataTransfer.files ?? []));
                    setImportMessages([]);
                    setNotice(null);
                  }}>
                  <span className="drop-icon"><Icon name="upload" size={22} /></span>
                  <span className="drop-main">Drop Excel files here or <strong>browse files</strong></span>
                  <span className="drop-help">Supports .xlsx and .xls spreadsheets · Select multiple files</span>
                  <input ref={fileInputRef} id="excel-files" type="file" accept=".xlsx,.xls"
                    multiple onChange={(event) => {
                      setFiles(Array.from(event.target.files ?? []));
                      setImportMessages([]);
                      setNotice(null);
                    }} />
                </label>

                {files.length > 0 && (
                  <div className="selected-files">
                    {files.map((file) => <span className="file-pill" key={`${file.name}-${file.size}`}>
                      <Icon name="file" size={14} />{file.name}
                      <button aria-label={`Remove ${file.name}`} onClick={() =>
                        setFiles((current) => current.filter((item) => item !== file))
                      }>×</button>
                    </span>)}
                  </div>
                )}

                <div className="import-controls">
                  <fieldset className="mode-fieldset">
                    <legend>When importing</legend>
                    <label className="radio-option">
                      <input type="radio" name="import-mode" checked={mode === "add"}
                        onChange={() => setMode("add")} />
                      <span><strong>Add / update</strong><small>Keep existing schemes and update matches</small></span>
                    </label>
                    <label className="radio-option">
                      <input type="radio" name="import-mode" checked={mode === "replace"}
                        onChange={() => setMode("replace")} />
                      <span><strong>Replace master</strong><small>Start fresh with the uploaded files</small></span>
                    </label>
                  </fieldset>
                  <button className="button button-primary build-button" disabled={!files.length || isBuilding}
                    onClick={buildMaster}>
                    <Icon name="upload" size={16} />{isBuilding ? "Processing files…" : "Build master"}
                  </button>
                </div>
                {importMessages.length > 0 && <div className="import-messages" aria-live="polite">
                  {importMessages.map((item, index) => (
                    <p className={`import-message ${item.type}`} key={`${index}-${item.text}`}>{item.text}</p>
                  ))}
                </div>}
              </section>

              <section className="panel table-panel">
                <div className="table-panel-heading">
                  <div><div className="table-title-row"><h2>Master dataset</h2>
                    <span className="count-badge">{master.length} schemes</span></div>
                    <p>{categories.length} {totalSourceLabel} in your library</p></div>
                  <button className="button button-outline" disabled={!master.length}
                    onClick={() => downloadRows(master, "mf_master_dataset.csv")}>
                    <Icon name="download" size={16} /> Export master
                  </button>
                </div>
                {master.length ? (
                  <>
                    <FundTable rows={masterPageRows} />
                    <Pagination page={masterPage} pageCount={masterPages} onChange={setMasterPage} />
                  </>
                ) : <div className="empty-state">
                  <span className="empty-icon"><Icon name="file" size={24} /></span>
                  <strong>Your master is empty</strong>
                  <span>Upload an Excel report above to get started.</span>
                </div>}
              </section>
            </div>
          ) : (
            <div className="content-stack">
              <section className="panel search-panel">
                <div className="panel-heading">
                  <div className="panel-title-wrap">
                    <div className="panel-icon"><Icon name="search" /></div>
                    <div><h2>Find the funds you need</h2>
                      <p>Search scheme names, filter by source, then export your results.</p></div>
                  </div>
                </div>
                <div className="search-controls">
                  <label className="field search-field">
                    <span>Scheme name</span>
                    <span className="input-with-icon"><Icon name="search" size={17} />
                      <input value={searchText} onChange={(event) => {
                        setSearchText(event.target.value);
                        setSelectedSchemes(new Set());
                        setResultPage(1);
                      }} placeholder="e.g. HDFC Liquid, SBI Liquid" />
                    </span>
                    <small>Separate multiple names with commas. Partial names work too.</small>
                  </label>
                  <fieldset className="field category-field">
                    <legend>Category / source file</legend>
                    <div className="category-list">
                      {categories.length ? categories.map((category) => (
                        <label className="category-option" key={category}>
                          <input type="checkbox" checked={selectedCategories.includes(category)}
                            onChange={() => toggleCategory(category)} />
                          <span>{category}</span>
                        </label>
                      )) : <span className="field-help">Import data to filter by source.</span>}
                    </div>
                    {selectedCategories.length > 0 && <button className="clear-filters"
                      onClick={() => {
                        setSelectedCategories([]);
                        setSelectedSchemes(new Set());
                        setResultPage(1);
                      }}>Clear source filters</button>}
                  </fieldset>
                </div>
              </section>

              {notFound.length > 0 && <div className="notice notice-warning" role="status">
                Not found in master: {notFound.join(", ")}
              </div>}

              <section className="panel table-panel">
                <div className="table-panel-heading result-heading">
                  <div><div className="table-title-row"><h2>Matching schemes</h2>
                    <span className="count-badge">{results.length} {results.length === 1 ? "scheme" : "schemes"}</span></div>
                    <p>{selectedSchemes.size
                      ? `${selectedSchemes.size} selected · Clear selections to use all matches`
                      : "Select rows to export a subset, or leave all matches selected."}</p></div>
                  <div className="result-actions">
                    {selectedSchemes.size > 0 && <button className="button button-outline"
                      onClick={() => setSelectedSchemes(new Set())}>Use all matches</button>}
                    <button className="button button-outline" disabled={!results.length} onClick={copyResults}>
                      <Icon name="copy" size={16} /> Copy for Excel
                    </button>
                    <button className="button button-primary" disabled={!results.length}
                      onClick={() => downloadRows(results, "filtered_funds_data.csv")}>
                      <Icon name="download" size={16} /> Download CSV
                    </button>
                  </div>
                </div>
                <FundTable rows={resultPageRows} onToggle={toggleScheme}
                  selected={selectedSchemes} />
                <Pagination page={resultPage} pageCount={resultPages} onChange={setResultPage} />
              </section>
            </div>
          )}

          <footer className="page-footer">
            <span>Performance data is provided by your uploaded spreadsheets.</span>
            <span>Stored locally in this browser <Icon name="arrow" size={14} /></span>
          </footer>
        </div>
      </main>
    </div>
  );
}
