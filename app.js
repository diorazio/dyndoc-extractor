(() => {
  // ---------- Elements ----------
  const manifestInput = document.getElementById('manifestInput');
  const searchInput = document.getElementById('searchInput');
  // Dedicated filter controls
  const filterTypeEl = document.getElementById('filterType');
  const filterManifestEl = document.getElementById('filterManifest');
  const filterFileEl = document.getElementById('filterFile');
  const filterValueEl = document.getElementById('filterValue');
  const filterNameEl = document.getElementById('filterName');
  const clearFiltersBtn = document.getElementById('clearFiltersBtn');
  const folderInput = document.getElementById('folderInput');
  const filesInput = document.getElementById('filesInput');
  const targetFolderNameEl = document.getElementById('targetFolderName');

  const includeCommentsEl = document.getElementById('includeComments');
  const usePathEl = document.getElementById('usePath');
  const avoidDupesEl = document.getElementById('avoidDupes');

  // NEW: extraction toggles + templates config
  const extractFiltersEl = document.getElementById('extractFilters');
  const extractTemplatesEl = document.getElementById('extractTemplates');
  const extractUuidsEl = document.getElementById('extractUuids');

  const templatesRootNameEl = document.getElementById('templatesRootName');
  const templatesLangEl = document.getElementById('templatesLang');
  const standardTemplatesNameEl = document.getElementById('standardTemplatesName');
  const specialtyTemplatesNameEl = document.getElementById('specialtyTemplatesName');

  const addBtn = document.getElementById('addBtn');
  const compareBtn = document.getElementById('compareBtn');
  const csvBtn = document.getElementById('csvBtn');
  const clearBtn = document.getElementById('clearBtn');

  // NEW: session export/import
  const exportSessionBtn = document.getElementById('exportSessionBtn');
  const importSessionBtn = document.getElementById('importSessionBtn');
  const importSessionFile = document.getElementById('importSessionFile');

  const statusEl = document.getElementById('status');
  const tbody = document.getElementById('resultsBody');

  const rowsCountEl = document.getElementById('rowsCount');
  const visibleCountEl = document.getElementById('visibleCount');

  const helpBtn = document.getElementById('helpBtn');
  const helpOverlay = document.getElementById('helpOverlay');
  const compareOverlay = document.getElementById('compareOverlay');

  // Compare modal controls
  const domainPaste = document.getElementById('domainPaste');
  const smartTemplatePaste = document.getElementById('smartTemplatePaste');

  // Smart Template paste preview
  const tplPasteMeta = document.getElementById('tplPasteMeta');
  const tplPastePreviewBody = document.getElementById('tplPastePreviewBody');
  const normalizeWhitespaceEl = document.getElementById('normalizeWhitespace');
  const caseInsensitiveEl = document.getElementById('caseInsensitive');
  if (caseInsensitiveEl) caseInsensitiveEl.checked = true;
  const runCompareBtn = document.getElementById('runCompareBtn');
  const clearPasteBtn = document.getElementById('clearPasteBtn');
  const compareStatus = document.getElementById('compareStatus');
  
  

  const dlFullWorkbookBtn = document.getElementById('dlFullWorkbookBtn');

  // Paste preview
  const pasteMeta = document.getElementById('pasteMeta');
  const pastePreviewBody = document.getElementById('pastePreviewBody');

  // Compare outputs - counts
  const dCkiNonEmpty = document.getElementById('dCkiNonEmpty');
  const dCkiUnique   = document.getElementById('dCkiUnique');
  const dCkiDupes    = document.getElementById('dCkiDupes');
  const eCkiRows      = document.getElementById('eCkiRows');
  const eCkiUnique    = document.getElementById('eCkiUnique');
  const eCkiDupes     = document.getElementById('eCkiDupes');

  const dNameNonEmpty = document.getElementById('dNameNonEmpty');
  const dNameUnique   = document.getElementById('dNameUnique');
  const dNameDupes    = document.getElementById('dNameDupes');
  const eNameRows     = document.getElementById('eNameRows');
  const eNameUnique   = document.getElementById('eNameUnique');
  const eNameDupes    = document.getElementById('eNameDupes');

  // Missing directions
  const missingInDomainCkiCount = document.getElementById('missingInDomainCkiCount');
  const missingInExtractCkiCount = document.getElementById('missingInExtractCkiCount');
  const missingInDomainNameCount = document.getElementById('missingInDomainNameCount');
  const missingInExtractNameCount = document.getElementById('missingInExtractNameCount');

  // Previews
  const missingInDomainCkiPreview = document.getElementById('missingInDomainCkiPreview');
  const missingInDomainNamePreview = document.getElementById('missingInDomainNamePreview');
  const missingInDomainCkiPreviewCount = document.getElementById('missingInDomainCkiPreviewCount');
  const missingInDomainNamePreviewCount = document.getElementById('missingInDomainNamePreviewCount');

  // ---------- State ----------
  const state = {
  rows: [],
  keySet: new Set(),
  lastCompare: null,

  // NEW: manifest behavior
  manifestMode: "auto", // "auto" follows folder root; "manual" is user-typed
};

  // ---------- Helpers ----------
  function escapeHtml(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function escapeCsv(val) {
    const s = String(val ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function download(filename, text, mime="text/csv;charset=utf-8") {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function normalized(s) { return String(s ?? "").toLowerCase(); }
  function rowKey(r) { return [r.manifest, r.file, r.type, r.value, r.comment].join("||"); }

  function rebuildManifestFilterOptions() {
    if (!filterManifestEl) return;
    const prev = filterManifestEl.value || "";

    const values = Array.from(
      new Set(state.rows.map(r => String(r.manifest || "").trim()).filter(Boolean))
    );
    values.sort((a,b) => String(a).localeCompare(String(b)));

    filterManifestEl.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = "All";
    filterManifestEl.appendChild(all);

    for (const v of values) {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      filterManifestEl.appendChild(opt);
    }

    // Restore selection if still present
    if (prev && values.includes(prev)) {
      filterManifestEl.value = prev;
    } else {
      filterManifestEl.value = "";
    }
  }

  function getFilteredRows() {
    const q = normalized(searchInput?.value).trim();
    const typeFilter = String(filterTypeEl?.value || "").trim();
    const manifestFilter = String(filterManifestEl?.value || "").trim();
    const fileQ = normalized(filterFileEl?.value).trim();
    const valueQ = normalized(filterValueEl?.value).trim();
    const nameQ = normalized(filterNameEl?.value).trim();

    return state.rows.filter(r => {
      if (typeFilter && r.type !== typeFilter) return false;
      if (manifestFilter && r.manifest !== manifestFilter) return false;

      if (fileQ && !normalized(r.file).includes(fileQ)) return false;
      if (valueQ && !normalized(r.value).includes(valueQ)) return false;
      if (nameQ && !normalized(r.comment).includes(nameQ)) return false;

      if (q) {
        const hay = [r.manifest, r.file, r.type, r.value, r.comment].map(normalized).join(" ");
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function updateCounts() {
    rowsCountEl.textContent = String(state.rows.length);
    visibleCountEl.textContent = String(getFilteredRows().length);

    csvBtn.disabled = getFilteredRows().length === 0;
    clearBtn.disabled = state.rows.length === 0;
    compareBtn.disabled = state.rows.length === 0;

    // NEW
    exportSessionBtn.disabled = (state.rows.length === 0);
  }

  function render() {
    const rows = getFilteredRows();
    tbody.innerHTML = "";

    for (const r of rows) {
      const tr = document.createElement("tr");

	const pillClass =
	  r.type === "concept_cki" ? "ck" :
	  r.type === "eventset_name" ? "name" :
	  r.type === "smart_template" ? "tmpl" :
	  r.type === "uuid" ? "uuid" : "";

      tr.innerHTML = `
        <td>${escapeHtml(r.manifest)}</td>
        <td class="mono">${escapeHtml(r.file)}</td>
        <td><span class="pill ${pillClass}">${escapeHtml(r.type)}</span></td>
        <td class="mono">${escapeHtml(r.value)}</td>
        <td>${escapeHtml(r.comment)}</td>
      `;

      tbody.appendChild(tr);
    }

    rebuildManifestFilterOptions();
    updateCounts();
  }

  function textOrEmpty(node) { return node ? (node.textContent ?? "").trim() : ""; }
  function nextNonWhitespaceSibling(node) {
    let s = node.nextSibling;
    while (s && s.nodeType === Node.TEXT_NODE && !s.textContent.trim()) s = s.nextSibling;
    return s;
  }

  // sanitize "almost XML"
  function sanitizeXml(input) {
    let s = String(input ?? "");
    s = s.replace(/^\uFEFF/, "");
    s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
    s = s.replace(/^\s*[-+]\s+(?=<)/gm, "");
    const firstLt = s.search(/</);
    if (firstLt > 0) s = s.slice(firstLt);
    const idx = s.indexOf("<?xml");
    if (idx > 0) s = s.slice(idx);
    return s.trimStart();
  }

  function parseEventsets(xmlText, includeComments) {
    xmlText = sanitizeXml(xmlText);

    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "application/xml");

    const parseError = doc.querySelector("parsererror");
    if (parseError) throw new Error(parseError.textContent.trim());

    const eventsetNodes = Array.from(
      doc.querySelectorAll("patcaremeasurementfilter > eventsets > eventset")
    );

    return eventsetNodes.map(es => {
      const conceptCKI = textOrEmpty(es.querySelector("concept_cki"));
      const eventsetName = textOrEmpty(es.querySelector("eventset_name"));

      const value = conceptCKI || eventsetName;
      const type = conceptCKI ? "concept_cki" : (eventsetName ? "eventset_name" : "");

      let comment = "";
      if (includeComments) {
        const sib = nextNonWhitespaceSibling(es);
        if (sib && sib.nodeType === Node.COMMENT_NODE) {
          comment = (sib.textContent ?? "").trim();
        }
      }

      return { type, value, comment };
    }).filter(x => x.type && x.value);
  }
  

  // ---------- Templates (HTML) extraction ----------
  function sanitizeHtml(input) {
    let s = String(input ?? "");
    s = s.replace(/^\uFEFF/, "");
    s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
    // Keep everything; DOMParser("text/html") is fairly forgiving.
    return s;
  }

  function collapseWhitespace(s) {
    return String(s ?? "").replace(/\s+/g, " ").trim();
  }

  // Finds <... dd:template_cki="..."> and associates it with the nearest section label if present.
  function parseSmartTemplates(htmlText) {
    htmlText = sanitizeHtml(htmlText);

    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlText, "text/html");

    const nodes = Array.from(doc.querySelectorAll('[dd\\:template_cki], [template_cki]'));
    const out = [];

    for (const node of nodes) {
      const cki = (node.getAttribute("dd:template_cki") || node.getAttribute("template_cki") || "").trim();
      if (!cki) continue;

      let name = "";
      const container = node.closest(".ddsubsection") || node.closest(".ddsection") || node.parentElement;
      if (container) {
        const display = container.querySelector(".ddsectiondisplay");
        if (display) {
          // Prefer the underlined inner span (matches the common "title" markup),
          // otherwise fallback to the ...
          const inner =
            display.querySelector('span[style*="underline"]') ||
            display.querySelector("span") ||
            display;
          name = collapseWhitespace(inner.textContent || "");
        }
      }

      out.push({ type: "smart_template", value: cki, comment: name });
    }

    return out;
  }
  
  function cleanLabel(s) {
    return collapseWhitespace(s)
      .replace(/:\s*$/, "")      // drop trailing colon
      .replace(/\u00A0/g, " ")   // nbsp
      .trim();
  }

  function classToLabel(cls) {
    return String(cls || "")
      .replace(/[-_]+/g, " ")
      .replace(/([a-z])([A-Z])/g, "$1 $2")
      .trim();
  }

  function findDisplayTitle(container) {
    if (!container) return "";

    // Prefer the container’s own heading (if present)
    const display =
      container.querySelector(":scope > .ddsectiondisplay") ||
      container.querySelector(".ddsectiondisplay");

    if (!display) return "";

    // If there’s a nested styled span, use it; otherwise use the whole display
    const inner =
      display.querySelector('span[style*="underline"]') ||
      display.querySelector("span") ||
      display;

    return cleanLabel(inner.textContent || "");
  }

  function findRowLabel(node) {
    // Look for a local label in the same line/row first
    const row = node.closest("div, p, li, td, th, tr");
    if (!row) return "";

    // Prefer a strong/b/label that occurs BEFORE the uuid node
    const candidates = Array.from(row.querySelectorAll("strong, b, label"));
    for (const el of candidates) {
      // el is before node if node is FOLLOWING el
      const rel = el.compareDocumentPosition(node);
      if (rel & Node.DOCUMENT_POSITION_FOLLOWING) {
        const t = cleanLabel(el.textContent || "");
        if (t) return t;
      }
    }

    return "";
  }

  function deriveFromClasses(node) {
    const bad = new Set(["ddemrcontent", "ddsectiondisplay", "ddrefreshable", "ddinsertfreetext", "ddremovable"]);
    const cls = Array.from(node.classList || []).find(c => c && !bad.has(c) && !c.startsWith("dd"));
    return cls ? cleanLabel(classToLabel(cls)) : "";
  }

  function parseReferenceUuids(htmlText) {
    htmlText = sanitizeHtml(htmlText);

    const doc = new DOMParser().parseFromString(htmlText, "text/html");
    const nodes = Array.from(doc.querySelectorAll('[dd\\:referenceuuid], [referenceuuid]'));

    const out = [];

    for (const node of nodes) {
      const uuid = (node.getAttribute("dd:referenceuuid") || node.getAttribute("referenceuuid") || "").trim();
      if (!uuid) continue;

      let name = "";

      // 1) Prefer nearest ddsubsection title (Your Care Team)
      const sub = node.closest(".ddsubsection");
      name = findDisplayTitle(sub);

      // 2) Else ddsection title (My Summary) — only if there was no subsection title
      if (!name) {
        const sec = node.closest(".ddsection");
        name = findDisplayTitle(sec);
      }

      // 3) Else row-level label (Patient Name)
      if (!name) {
        name = findRowLabel(node);
      }

      // 4) Else derive from class (patientname -> "patientname"/"patient name")
      if (!name) {
        name = deriveFromClasses(node);
      }

      out.push({ type: "uuid", value: uuid, comment: name });
    }

    return out;
  }

  // ---------- Folder selection (nested) ----------
  // Normalize segments so we can match "en_gb" vs "en-gb" vs "en gb", etc.
  function normalizeSegment(s) {
    return String(s ?? "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ");
  }

  function splitPathSegments(path) {
    return String(path || "").split("/").filter(Boolean);
  }

  function pathHasSegment(path, seg) {
    if (!seg) return false;
    const wanted = normalizeSegment(seg);
    const segs = splitPathSegments(path).map(normalizeSegment);
    return segs.includes(wanted);
  }

  function pathContainsOrderedSegments(path, orderedSegs) {
    const segs = splitPathSegments(path).map(normalizeSegment);
    let idx = -1;
    for (const raw of (orderedSegs || [])) {
      const wanted = normalizeSegment(raw);
      idx = segs.indexOf(wanted, idx + 1);
      if (idx < 0) return false;
    }
    return true;
  }

  function fileExtLower(name) {
    const s = String(name || "");
    const i = s.lastIndexOf(".");
    return i >= 0 ? s.slice(i).toLowerCase() : "";
  }

  function selectFilterXmlFolderFiles(folderXmlFiles) {
    const targetSeg = normalizeSegment(targetFolderNameEl.value || "Filters");
    let used = folderXmlFiles;
    let note = "";

    if (folderXmlFiles.length && targetSeg) {
      const subset = folderXmlFiles.filter(f => pathHasSegment(f.webkitRelativePath || "", targetSeg));
      if (subset.length) {
        used = subset;
        note = `Filters: folder contains "${targetSeg}" — using ${subset.length} XML file(s) under that subfolder (ignored ${folderXmlFiles.length - subset.length} other XMLs).`;
      } else {
        note = `Filters: folder upload — using ${folderXmlFiles.length} XML file(s) (no "${targetSeg}" subfolder detected).`;
      }
    } else if (folderXmlFiles.length) {
      note = `Filters: folder upload — using ${folderXmlFiles.length} XML file(s).`;
    }

    return { files: used, note };
  }

  function selectTemplateHtmlFolderFiles(folderHtmlFiles) {
    const templatesRoot = (templatesRootNameEl?.value || "Templates").trim() || "Templates";
    const standardName = (standardTemplatesNameEl?.value || "Standard Templates").trim() || "Standard Templates";
    const specialtyName = (specialtyTemplatesNameEl?.value || "Specialty Templates").trim() || "Specialty Templates";
    const lang = (templatesLangEl?.value || "en").trim() || "en";

    const rootNorm = normalizeSegment(templatesRoot);
    const langNorm = normalizeSegment(lang);

    // First, scope to "Templates" folder if it exists anywhere in the upload.
    let underTemplates = folderHtmlFiles.filter(f => pathHasSegment(f.webkitRelativePath || "", rootNorm));
    let scoped = underTemplates.length ? underTemplates : folderHtmlFiles;

    // Preferred: .../Templates/Standard Templates/<lang>/... and .../Templates/Specialty Templates/<lang>/...
    const std = scoped.filter(f => pathContainsOrderedSegments(f.webkitRelativePath || "", [rootNorm, standardName, langNorm]));
    const spec = scoped.filter(f => pathContainsOrderedSegments(f.webkitRelativePath || "", [rootNorm, specialtyName, langNorm]));

    let used = [];
    let note = "";
    if (std.length || spec.length) {
      used = [...std, ...spec];
      note = `Templates: using ${used.length} HTML file(s) in "${lang}" under Standard/Specialty templates.`;
    } else {
      // Fallback: any .../Templates/<lang>/... (covers cases where there aren't Standard/Specialty folders)
      const anyLang = scoped.filter(f => pathContainsOrderedSegments(f.webkitRelativePath || "", [rootNorm, langNorm]));
      if (anyLang.length) {
        used = anyLang;
        note = `Templates: using ${used.length} HTML file(s) in "${lang}" under "${templatesRoot}".`;
      } else {
        used = scoped;
        note = underTemplates.length
          ? `Templates: using ${used.length} HTML file(s) under "${templatesRoot}" (no "${lang}" folder detected).`
          : `Templates: using ${used.length} HTML file(s) (no "${templatesRoot}" folder detected).`;
      }
    }

    return { files: used, note };
  }

  function collectSelectedFilesByType() {
    const folderAll = Array.from(folderInput.files || []);
    const looseAll  = Array.from(filesInput.files || []);

    const folderXml = folderAll.filter(f => fileExtLower(f.name) === ".xml");
    const looseXml  = looseAll.filter(f => fileExtLower(f.name) === ".xml");

    const folderHtml = folderAll.filter(f => [".html", ".htm"].includes(fileExtLower(f.name)));
    const looseHtml  = looseAll.filter(f => [".html", ".htm"].includes(fileExtLower(f.name)));

    const filtersSel = selectFilterXmlFolderFiles(folderXml);
    const templatesSel = selectTemplateHtmlFolderFiles(folderHtml);

    // Loose files: include as-is (no subfolder heuristics available)
    const notes = [];
    if (filtersSel.note) notes.push(filtersSel.note);
    if (templatesSel.note) notes.push(templatesSel.note);

    return {
      xmlFiles: [...filtersSel.files, ...looseXml],
      htmlFiles: [...templatesSel.files, ...looseHtml],
      note: notes.filter(Boolean).join("\n")
    };
  }

  function collectSelectedFilesForActiveExtractors(doFilters, doTemplates, doUuids) {
    const sel = collectSelectedFilesByType();
    return {
      xmlFiles: doFilters ? sel.xmlFiles : [],
      htmlFiles: (doTemplates || doUuids) ? sel.htmlFiles : [],
      note: sel.note
    };
  }

  // If Manifest is blank, use TOP-LEVEL folder name (first segment of webkitRelativePath)
  function inferRootFolderNameFromFolderSelection() {
    const all = Array.from(folderInput.files || []);
    if (!all.length) return { name: "", note: "" };

    const roots = new Set();
    for (const f of all) {
      const p = f.webkitRelativePath || "";
      const root = p.split("/")[0] || "";
      if (root) roots.add(root);
    }
    if (!roots.size) return { name: "", note: "" };

    const name = Array.from(roots)[0];
    const note = roots.size > 1
      ? `Manifest was blank — using "${name}" (note: multiple root folders detected: ${Array.from(roots).join(", ")}).`
      : `Manifest was blank — using "${name}" from the selected folder.`;
    return { name, note };
  }

  // ---------- Modals ----------
  function openOverlay(overlayEl){ overlayEl.style.display = "flex"; }
  function closeOverlay(overlayEl){ overlayEl.style.display = "none"; }

  helpBtn.addEventListener("click", () => openOverlay(helpOverlay));
  compareBtn.addEventListener("click", () => openOverlay(compareOverlay));

  document.querySelectorAll("[data-close]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-close");
      const el = document.getElementById(id);
      if (el) closeOverlay(el);
    });
  });

  [helpOverlay, compareOverlay].forEach(ov => {
    ov.addEventListener("click", (e) => { if (e.target === ov) closeOverlay(ov); });
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (helpOverlay.style.display === "flex") closeOverlay(helpOverlay);
    if (compareOverlay.style.display === "flex") closeOverlay(compareOverlay);
  });

  // ---------- Compare logic ----------
  function normalizeValue(s, doWhitespace, doCase) {
    let v = String(s ?? "");
    // Normalize common invisible/odd whitespace that Excel/pastes can introduce
    v = v.replace(/\u00A0/g, " ")   // non-breaking space
         .replace(/\u200B/g, "");  // zero-width space

    if (doWhitespace) v = v.trim().replace(/\s+/g, " ");
    else v = v.trim();
    if (doCase) v = v.toLowerCase();
    return v;
  }

  function headerKey(s){
    return String(s ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  }

  function detectDelimiterFromSample(lines) {
    const sample = lines.slice(0, Math.min(10, lines.length)).join("\n");
    if (sample.includes("\t")) return "\t";
    if (sample.includes(",")) return ",";
    if (sample.includes("|")) return "|";
    return "\t";
  }

  function parseTwoColumnPaste(text, doWhitespace, doCase) {
    const raw = String(text ?? "");
    const lines = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n")
      .split("\n")
      .filter(l => l.length > 0);

    if (!lines.length) {
      return {
        conceptSet: new Set(), nameSet: new Set(),
        rowCount: 0,
        conceptNonEmpty: 0, nameNonEmpty: 0,
        conceptDupes: 0, nameDupes: 0,
        oneColRows: 0, multiColRows: 0,
        delim: "\t",
        headers: ["CONCEPT_CKI","EVENTSET_NAME"],
        preview: [],
        previewWithCki: []
      };
    }

    const delim = detectDelimiterFromSample(lines);

    // Header handling: many Excel pastes include headers, but some do not.
    // If the first line doesn't look like a header, treat it as data.
    let headerParts = lines[0].split(delim);
    let keys = headerParts.map(headerKey);
    const hasHeader = keys.some(k => k.includes("CONCEPTCKI") || k.includes("EVENTSETNAME"));
    const startRow = hasHeader ? 1 : 0;

    if (!hasHeader) {
      headerParts = ["CONCEPT_CKI", "EVENTSET_NAME"];
      keys = headerParts.map(headerKey);
    }

    let conceptIdx = hasHeader ? keys.findIndex(k => k.includes("CONCEPTCKI")) : 0;
    let nameIdx = hasHeader ? keys.findIndex(k => k.includes("EVENTSETNAME")) : 1;

    if (conceptIdx < 0) conceptIdx = 0;
    if (nameIdx < 0) nameIdx = 1;

    const conceptSet = new Set();
    const nameSet = new Set();

    let conceptNonEmpty = 0;
    let nameNonEmpty = 0;
    let oneColRows = 0;
    let multiColRows = 0;

    const preview = [];
    const previewWithCki = [];
	const rowsRaw = [];

    for (let i = startRow; i < lines.length; i++) {
      const parts = lines[i].split(delim);

      if (parts.length === 1) oneColRows++;
      if (parts.length > 2) multiColRows++;

      const cRaw = parts[conceptIdx] ?? "";
      let nRaw = parts[nameIdx] ?? "";

      // TSV safety: if extra tabs exist and nameIdx is 1, join remaining into name
      if (delim === "\t" && nameIdx === 1 && parts.length > 2) {
        nRaw = parts.slice(1).join("\t");
      }
	  
	  const cOut = String(cRaw ?? "").trim();
	  const nOut = String(nRaw ?? "").trim();
	  if (cOut || nOut) rowsRaw.push([cOut, nOut]);

      const c = normalizeValue(cRaw, doWhitespace, doCase);
      const n = normalizeValue(nRaw, doWhitespace, doCase);

      if (c) { conceptNonEmpty++; conceptSet.add(c); }
      if (n) { nameNonEmpty++; nameSet.add(n); }

      if (preview.length < 6) preview.push({ concept_cki: cRaw, eventset_name: nRaw });
      if ((String(cRaw || "").trim()) && previewWithCki.length < 6) {
        previewWithCki.push({ concept_cki: cRaw, eventset_name: nRaw });
      }
    }

    return {
      conceptSet, nameSet,
      rowCount: Math.max(0, lines.length - startRow),
      conceptNonEmpty,
      nameNonEmpty,
      conceptDupes: conceptNonEmpty - conceptSet.size,
      nameDupes: nameNonEmpty - nameSet.size,
      oneColRows,
      multiColRows,
      delim,
      headers: headerParts,
      preview,
      previewWithCki,
	  rowsRaw
    };
  }

  function parseSmartTemplatePaste(text, doWhitespace, doCase) {
    const raw = String(text ?? "");
    const lines = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n")
      .split("\n")
      .filter(l => l.length > 0);

    if (!lines.length) {
      return {
        templateSet: new Set(),
        templateNameByCki: new Map(),
        rowCount: 0,
        templateNonEmpty: 0,
        templateDupes: 0,
        oneColRows: 0,
        multiColRows: 0,
        delim: "\t",
        headers: ["TEMPLATE_CKI", "NAME"],
        preview: [],
        previewWithTemplate: [],
        rowsRaw: [] // ✅ added
      };
    }

    const delim = detectDelimiterFromSample(lines);
    let headerParts = lines[0].split(delim);
    let keys = headerParts.map(headerKey);
    const hasHeader = keys.some(k => k.includes("TEMPLATECKI") || k.includes("SMARTTEMPLATE") || k.includes("DDTEMPLATECKI"));
    const startRow = hasHeader ? 1 : 0;

    if (!hasHeader) {
      headerParts = ["TEMPLATE_CKI", "NAME"];
      keys = headerParts.map(headerKey);
    }

    let tplIdx = keys.findIndex(k => k.includes("TEMPLATECKI") || k.includes("SMARTTEMPLATE") || k.includes("DDTEMPLATECKI"));
    if (tplIdx < 0) tplIdx = 0;
    const nameIdx = (tplIdx === 0 ? 1 : 0);

    const templateSet = new Set();
    const templateNameByCki = new Map(); // normalized cki -> name

    let templateNonEmpty = 0;
    let oneColRows = 0;
    let multiColRows = 0;

    const preview = [];
    const previewWithTemplate = [];
    const rowsRaw = []; // ✅ added (paste-mirror, exact 2 columns)

    for (let i = startRow; i < lines.length; i++) {
      const parts = lines[i].split(delim);
      if (parts.length === 1) oneColRows++;
      if (parts.length > 2) multiColRows++;

      const tplRaw = parts[tplIdx] ?? "";
      let nameRaw = parts[nameIdx] ?? "";

      // TSV safety: if extra tabs exist and the name is col 1, join remaining into name
      if (delim === "\t" && nameIdx === 1 && parts.length > 2) {
        nameRaw = parts.slice(1).join("\t");
      }

      const tplOut = String(tplRaw ?? "").trim();
      const nameOut = String(nameRaw ?? "").trim();
      if (tplOut || nameOut) rowsRaw.push([tplOut, nameOut]); // ✅ added

      const tpl = normalizeValue(tplRaw, doWhitespace, doCase);
      const nameTrim = String(nameRaw ?? "").trim();

      if (tpl) {
        templateNonEmpty++;
        templateSet.add(tpl);
        if (nameTrim && !templateNameByCki.has(tpl)) templateNameByCki.set(tpl, nameTrim);
      }

      if (preview.length < 6) preview.push({ template_cki: tplRaw, name: nameRaw });
      if (String(tplRaw || "").trim() && previewWithTemplate.length < 6) {
        previewWithTemplate.push({ template_cki: tplRaw, name: nameRaw });
      }
    }

    return {
      templateSet,
      templateNameByCki,
      rowCount: startRow === 0 ? lines.length : Math.max(0, lines.length - 1),
      templateNonEmpty,
      templateDupes: templateNonEmpty - templateSet.size,
      oneColRows,
      multiColRows,
      delim,
      headers: headerParts,
      preview,
      previewWithTemplate,
      rowsRaw // ✅ added
    };
  }

  function buildExtractSets(doWhitespace, doCase) {
    const conceptSet = new Set();
    const nameSet = new Set();
    const templateSet = new Set();

    const templateNameByCki = new Map();     // normalized smart_template cki -> name (from comment)
    const conceptCommentByCki = new Map();   // normalized concept_cki -> comment/name

    let conceptRows = 0;
    let nameRows = 0;
    let templateRows = 0;
	
	function extractEventsetNameOnly(comment) {
	  const s = String(comment ?? "");
	  // Split on the exact separator we used: " — "
	  const idx = s.indexOf(" — ");
	  return (idx >= 0 ? s.slice(0, idx) : s).trim();
	}

	for (const r of state.rows) {
	  if (r.type === "concept_cki") {
	    conceptRows++;
	    const v = normalizeValue(r.value, doWhitespace, doCase);
	    if (v) conceptSet.add(v);

	    const cm = String(r.comment ?? "").trim().replace(/\s+/g, " ");
	    if (v && cm && !conceptCommentByCki.has(v)) conceptCommentByCki.set(v, cm);

	  } else if (r.type === "eventset_name") {
	    nameRows++;
	    const v = normalizeValue(r.value, doWhitespace, doCase);
	    if (v) nameSet.add(v);

	  } else if (r.type === "domain_value") {
	    // ✅ paired filter row: value = concept_cki, comment = eventset_name (maybe with " — xml comment")
	    const ckiNorm = normalizeValue(r.value, doWhitespace, doCase);
	    const nameOnly = extractEventsetNameOnly(r.comment);
	    const nameNorm = normalizeValue(nameOnly, doWhitespace, doCase);

	    if (ckiNorm) {
	      conceptRows++;
	      conceptSet.add(ckiNorm);
	      // tie concept -> eventset_name (first non-empty)
	      if (nameOnly && !conceptCommentByCki.has(ckiNorm)) {
	        conceptCommentByCki.set(ckiNorm, nameOnly);
	      }
	    }

	    if (nameNorm) {
	      nameRows++;
	      nameSet.add(nameNorm);
	    }

	  } else if (r.type === "smart_template") {
	    templateRows++;
	    const v = normalizeValue(r.value, doWhitespace, doCase);
	    if (v) templateSet.add(v);

	    const nm = String(r.comment ?? "").trim().replace(/\s+/g, " ");
	    if (v && nm && !templateNameByCki.has(v)) templateNameByCki.set(v, nm);
	  }
	}

    return {
      conceptSet,
      nameSet,
      templateSet,

      templateNameByCki,
      conceptCommentByCki, // ✅ IMPORTANT: return this

      conceptRows,
      nameRows,
      templateRows,

      conceptDupes: conceptRows - conceptSet.size,
      nameDupes: nameRows - nameSet.size,
      templateDupes: templateRows - templateSet.size
    };
  }

  function setDiff(a, b) {
    const out = [];
    for (const v of a) if (!b.has(v)) out.push(v);
    return out;
  }

  function previewList(arr, limit = 50) {
    if (!arr.length) return `None.`;
    return arr.slice(0, limit).join("\n");
  }

  function renderPastePreview(parsed) {
    if (!parsed || parsed.rowCount === 0) {
      pasteMeta.textContent = "No paste detected";
      pastePreviewBody.innerHTML = `<div class="status" style="margin:0;">Paste data above to see a 2-column preview with visible borders.</div>`;
      return;
    }

    const delimName = parsed.delim === "\t" ? "TAB" : (parsed.delim === "," ? "CSV" : (parsed.delim === "|" ? "PIPE" : "TAB"));
    const oneColWarn = parsed.oneColRows > 0
      ? ` • <span class="warn">${parsed.oneColRows.toLocaleString()}</span> row(s) appear 1-column (missing delimiter)`
      : "";
    const multiColWarn = parsed.multiColRows > 0
      ? ` • ${parsed.multiColRows.toLocaleString()} row(s) had extra columns`
      : "";

    pasteMeta.innerHTML =
      `Detected: <b>${delimName}</b> • Rows: <b>${parsed.rowCount.toLocaleString()}</b>${oneColWarn}${multiColWarn}`;

    const combined = [];
    const seen = new Set();
    for (const r of (parsed.preview || [])) {
      const k = (r.concept_cki || "") + "\t" + (r.eventset_name || "");
      if (!seen.has(k)) { seen.add(k); combined.push(r); }
    }
    for (const r of (parsed.previewWithCki || [])) {
      const k = (r.concept_cki || "") + "\t" + (r.eventset_name || "");
      if (!seen.has(k)) { seen.add(k); combined.push(r); }
    }

    const rowsHtml = combined.slice(0, 12).map(r => `
      <tr>
        <td>${escapeHtml(r.concept_cki)}</td>
        <td>${escapeHtml(r.eventset_name)}</td>
      </tr>
    `).join("");

    pastePreviewBody.innerHTML = `
      <table class="gridTable" aria-label="Paste preview table">
        <colgroup>
          <col style="width:45%">
          <col style="width:55%">
        </colgroup>
        <thead>
          <tr>
            <th>CONCEPT_CKI</th>
            <th>EVENTSET_NAME</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="2" style="color:#6B7280;">No rows detected after header.</td></tr>`}
        </tbody>
      </table>
    `;
  }

  function renderTemplatePastePreview(parsed) {
    if (!parsed || parsed.rowCount === 0) {
      tplPasteMeta.textContent = "No paste detected";
      tplPastePreviewBody.innerHTML = `<div class="status" style="margin:0;">Paste Smart Template data above to see a 2-column preview with visible borders.</div>`;
      return;
    }

    const delimName = parsed.delim === "\t" ? "TAB" : (parsed.delim === "," ? "CSV" : (parsed.delim === "|" ? "PIPE" : "TAB"));
    const oneColWarn = parsed.oneColRows > 0
      ? ` • <span class="warn">${parsed.oneColRows.toLocaleString()}</span> row(s) appear 1-column (missing delimiter)`
      : "";
    const multiColWarn = parsed.multiColRows > 0
      ? ` • ${parsed.multiColRows.toLocaleString()} row(s) had extra columns`
      : "";

    tplPasteMeta.innerHTML =
      `Detected: <b>${delimName}</b> • Rows: <b>${parsed.rowCount.toLocaleString()}</b>${oneColWarn}${multiColWarn}`;

    const combined = [];
    const seen = new Set();
    for (const r of (parsed.preview || [])) {
      const k = (r.template_cki || "") + "\t" + (r.name || "");
      if (!seen.has(k)) { seen.add(k); combined.push(r); }
    }
    for (const r of (parsed.previewWithTemplate || [])) {
      const k = (r.template_cki || "") + "\t" + (r.name || "");
      if (!seen.has(k)) { seen.add(k); combined.push(r); }
    }

    const rowsHtml = combined.slice(0, 12).map(r => `
      <tr>
        <td>${escapeHtml(r.template_cki)}</td>
        <td>${escapeHtml(r.name)}</td>
      </tr>
    `).join("");

    tplPastePreviewBody.innerHTML = `
      <table class="gridTable" aria-label="Smart template paste preview table">
        <colgroup>
          <col style="width:55%">
          <col style="width:45%">
        </colgroup>
        <thead>
          <tr>
            <th>TEMPLATE_CKI</th>
            <th>NAME</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="2" style="color:#6B7280;">No rows detected after header.</td></tr>`}
        </tbody>
      </table>
    `;
  }

  function recomputePastePreviews() {
    const doWhitespace = normalizeWhitespaceEl.checked;
    const doCase = caseInsensitiveEl.checked;

    const pasted = domainPaste.value || "";
    if (!pasted.trim()) {
      renderPastePreview(null);
    } else {
      const parsed = parseTwoColumnPaste(pasted, doWhitespace, doCase);
      renderPastePreview(parsed);
    }

    const pastedTpl = smartTemplatePaste.value || "";
    if (!pastedTpl.trim()) {
      renderTemplatePastePreview(null);
    } else {
      const parsedTpl = parseSmartTemplatePaste(pastedTpl, doWhitespace, doCase);
      renderTemplatePastePreview(parsedTpl);
    }
  }

  domainPaste.addEventListener("input", recomputePastePreviews);
  smartTemplatePaste.addEventListener("input", recomputePastePreviews);
  normalizeWhitespaceEl.addEventListener("change", recomputePastePreviews);
  caseInsensitiveEl.addEventListener("change", recomputePastePreviews);

  // ---------- XLSX export ----------
  function ensureXlsxAvailable() {
    if (typeof XLSX === "undefined" || !XLSX.utils) {
      compareStatus.innerHTML = `<span class="bad">XLSX library not loaded.</span> If you are offline, download xlsx.full.min.js and reference it locally.`;
      return false;
    }
    return true;
  }

  function setToSortedArray(set) {
    const arr = Array.from(set);
    arr.sort((a,b) => String(a).localeCompare(String(b)));
    return arr;
  }

  function makeTwoColumnAoA(col1Header, col1Values, col2Header, col2Values) {
    const a = col1Values || [];
    const b = col2Values || [];
    const n = Math.max(a.length, b.length);
    const aoa = [[col1Header, col2Header]];
    for (let i = 0; i < n; i++) {
      aoa.push([a[i] ?? "", b[i] ?? ""]);
    }
    return aoa;
  }

  function addSheet(wb, name, aoa) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  function downloadWorkbookXlsx(filename, wb) {
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function buildFullCompareWorkbook() {
    const lc = state.lastCompare;
    const wb = XLSX.utils.book_new();

    // Safety
    const missingCki = Array.isArray(lc?.missingInDomainCki) ? lc.missingInDomainCki : [];
    const missingCkiComments = Array.isArray(lc?.missingInDomainCkiComments) ? lc.missingInDomainCkiComments : [];
    const missingName = Array.isArray(lc?.missingInDomainName) ? lc.missingInDomainName : [];

    const missingTpl = Array.isArray(lc?.missingInDomainTpl) ? lc.missingInDomainTpl : [];
    const missingTplNames = Array.isArray(lc?.missingInDomainTplNames) ? lc.missingInDomainTplNames : [];

    // Keep your existing sheets (organized)
    // Extract
    const extractAoA = [["TYPE", "VALUE", "COMMENT/NAME"]];
    for (let i = 0; i < (lc.extractConceptUnique || []).length; i++) {
      extractAoA.push([
        "concept_cki",
        lc.extractConceptUnique[i] ?? "",
        (lc.extractConceptUniqueComments?.[i] ?? "")
      ]);
    }
    extractAoA.push(["", "", ""]);
    for (const v of (lc.extractNameUnique || [])) extractAoA.push(["eventset_name", v ?? "", ""]);
    extractAoA.push(["", "", ""]);
    for (let i = 0; i < (lc.extractTplUnique || []).length; i++) {
      extractAoA.push(["smart_template", lc.extractTplUnique[i] ?? "", (lc.extractTplUniqueNames?.[i] ?? "")]);
    }
    addSheet(wb, "Package values", extractAoA);

	// Domain values (paste-mirror, EXACT 2 columns)
	const domainAoA = [["CONCEPT_CKI", "EVENTSET_NAME"]];

	const domainRows = Array.isArray(lc?.domainPasteRows) ? lc.domainPasteRows : [];

	if (domainRows.length) {
	  for (const row of domainRows) {
	    domainAoA.push([row?.[0] ?? "", row?.[1] ?? ""]);
	  }
	} else {
	  // Fallback for older sessions
	  for (const v of (lc.domainConceptUnique || [])) domainAoA.push([v ?? "", ""]);
	  for (const v of (lc.domainNameUnique || [])) domainAoA.push(["", v ?? ""]);
	}

	addSheet(wb, "Domain values", domainAoA);
	
    // ---- Full Compare sheet (metrics + missing in domain sections) ----
    const summary = [];
    summary.push(["Normalization", lc.normalizationLabel || ""]);
    summary.push(["Paste diagnostics", lc.pasteDiagnostics || ""]);
    summary.push([""]);
    summary.push(["Metric", "Domain", "Package"]);

    // Concept metrics
    summary.push(["CONCEPT_CKI non-empty cells", lc.domainConceptNonEmpty ?? 0, lc.extractConceptRows ?? 0]);
    summary.push(["CONCEPT_CKI unique values", lc.domainConceptUniqueCount ?? 0, lc.extractConceptUniqueCount ?? 0]);
    summary.push(["CONCEPT_CKI duplicates (cells-unique)", lc.domainConceptDupes ?? 0, lc.extractConceptDupes ?? 0]);
    summary.push(["CONCEPT_CKI missing in Domain (Package\\Domain)", "", missingCki.length]);
    summary.push([""]);

    // Eventset metrics
    summary.push(["EVENTSET_NAME non-empty cells", lc.domainNameNonEmpty ?? 0, lc.extractNameRows ?? 0]);
    summary.push(["EVENTSET_NAME unique values", lc.domainNameUniqueCount ?? 0, lc.extractNameUniqueCount ?? 0]);
    summary.push(["EVENTSET_NAME duplicates (cells-unique)", lc.domainNameDupes ?? 0, lc.extractNameDupes ?? 0]);
    summary.push(["EVENTSET_NAME missing in Domain (Package\\Domain)", "", missingName.length]);
    summary.push([""]);

    // Smart template metrics
    summary.push(["SMART_TEMPLATE non-empty cells", lc.domainTplNonEmpty ?? 0, lc.extractTplRows ?? 0]);
    summary.push(["SMART_TEMPLATE unique values", lc.domainTplUniqueCount ?? 0, lc.extractTplUniqueCount ?? 0]);
    summary.push(["SMART_TEMPLATE duplicates (cells-unique)", lc.domainTplDupes ?? 0, lc.extractTplDupes ?? 0]);
    summary.push(["SMART_TEMPLATE missing in Domain (Package\\Domain)", "", missingTpl.length]);
    summary.push([""]);

    // Missing-in-domain sections (ONLY)
    summary.push(["MISSING IN DOMAIN (Package\\Domain)"]);
    summary.push([""]);

    summary.push(["Missing CONCEPT_CKI in Domain"]);
    summary.push(["CONCEPT_CKI", "Comment/Name (from Package)"]);
    for (let i = 0; i < missingCki.length; i++) {
      summary.push([missingCki[i] ?? "", missingCkiComments[i] ?? ""]);
    }

    summary.push([""]);
    summary.push(["Missing EVENTSET_NAME in Domain"]);
    summary.push(["EVENTSET_NAME"]);
    for (const v of missingName) summary.push([v ?? ""]);

    summary.push([""]);
    summary.push(["Missing SMART_TEMPLATE in Domain"]);
    summary.push(["TEMPLATE_CKI", "NAME (from Package)"]);
    for (let i = 0; i < missingTpl.length; i++) {
      summary.push([missingTpl[i] ?? "", missingTplNames[i] ?? ""]);
    }

    addSheet(wb, "Full compare", summary);
    return wb;
  }

  dlFullWorkbookBtn.addEventListener("click", () => {
    if (!state.lastCompare) return;
    if (!ensureXlsxAvailable()) return;
    const wb = buildFullCompareWorkbook();
    const stamp = new Date().toISOString().slice(0,10);
    downloadWorkbookXlsx(`full_compare_${stamp}.xlsx`, wb);
  });

  // ---------- Compare run ----------
  runCompareBtn.addEventListener("click", () => {
    const doWhitespace = normalizeWhitespaceEl.checked;
    const doCase = caseInsensitiveEl.checked;

    if (state.rows.length === 0) {
      compareStatus.innerHTML = `<span class="bad">No extracted rows.</span> Load XML/HTML files first, then compare.`;
      return;
    }

    const pastedDomain = domainPaste.value || "";
    const pastedTpl = smartTemplatePaste.value || "";

    const haveDomain = !!pastedDomain.trim();
    const haveTpl = !!pastedTpl.trim();

    if (!haveDomain && !haveTpl) {
      compareStatus.innerHTML = `<span class="bad">No pasted data.</span> Paste Domain values (Concept/Eventset) and/or Smart Template values, then try again.`;
      return;
    }

    compareStatus.textContent = "Parsing paste and computing differences…";

    const parsedDomain = haveDomain ? parseTwoColumnPaste(pastedDomain, doWhitespace, doCase) : parseTwoColumnPaste("", doWhitespace, doCase);
    const parsedTpl = haveTpl ? parseSmartTemplatePaste(pastedTpl, doWhitespace, doCase) : parseSmartTemplatePaste("", doWhitespace, doCase);

    renderPastePreview(haveDomain ? parsedDomain : null);
    renderTemplatePastePreview(haveTpl ? parsedTpl : null);

    const extracted = buildExtractSets(doWhitespace, doCase);

    // ----- Concept CKI + Eventset Name (ONLY Missing in Domain) -----
    let domainCkiSet = new Set();
    let domainNameSet = new Set();
    let missingInDomainCki = [];
    let missingInDomainName = [];

    if (haveDomain) {
      domainCkiSet = parsedDomain.conceptSet;
      domainNameSet = parsedDomain.nameSet;

      missingInDomainCki = setDiff(extracted.conceptSet, domainCkiSet);
      missingInDomainName = setDiff(extracted.nameSet, domainNameSet);

      missingInDomainCki.sort((a,b) => String(a).localeCompare(String(b)));
      missingInDomainName.sort((a,b) => String(a).localeCompare(String(b)));

      dCkiNonEmpty.textContent = String(parsedDomain.conceptNonEmpty);
      dCkiUnique.textContent   = String(domainCkiSet.size);
      dCkiDupes.textContent    = String(Math.max(0, parsedDomain.conceptDupes));

      eCkiRows.textContent     = String(extracted.conceptRows);
      eCkiUnique.textContent   = String(extracted.conceptSet.size);
      eCkiDupes.textContent    = String(Math.max(0, extracted.conceptDupes));

      dNameNonEmpty.textContent = String(parsedDomain.nameNonEmpty);
      dNameUnique.textContent   = String(domainNameSet.size);
      dNameDupes.textContent    = String(Math.max(0, parsedDomain.nameDupes));

      eNameRows.textContent     = String(extracted.nameRows);
      eNameUnique.textContent   = String(extracted.nameSet.size);
      eNameDupes.textContent    = String(Math.max(0, extracted.nameDupes));

      // Missing in domain counts
      missingInDomainCkiCount.textContent = String(missingInDomainCki.length);
      missingInDomainNameCount.textContent = String(missingInDomainName.length);

      // We no longer compute "Missing in Extract"
      if (missingInExtractCkiCount) missingInExtractCkiCount.textContent = "0";
      if (missingInExtractNameCount) missingInExtractNameCount.textContent = "0";

      // Previews
      missingInDomainCkiPreview.textContent = previewList(missingInDomainCki, 50);
      missingInDomainNamePreview.textContent = previewList(missingInDomainName, 50);
      missingInDomainCkiPreviewCount.textContent = `${Math.min(50, missingInDomainCki.length)} / ${missingInDomainCki.length}`;
      missingInDomainNamePreviewCount.textContent = `${Math.min(50, missingInDomainName.length)} / ${missingInDomainName.length}`;
    } else {
      // Clear concept/name UI if that paste is not provided
      [
        dCkiNonEmpty, dCkiUnique, dCkiDupes, eCkiRows, eCkiUnique, eCkiDupes,
        dNameNonEmpty, dNameUnique, dNameDupes, eNameRows, eNameUnique, eNameDupes,
        missingInDomainCkiCount, missingInDomainNameCount
      ].forEach(el => el.textContent = "0");

      if (missingInExtractCkiCount) missingInExtractCkiCount.textContent = "0";
      if (missingInExtractNameCount) missingInExtractNameCount.textContent = "0";

      missingInDomainCkiPreview.textContent = "";
      missingInDomainNamePreview.textContent = "";
      missingInDomainCkiPreviewCount.textContent = "0 / 0";
      missingInDomainNamePreviewCount.textContent = "0 / 0";
    }

    // ----- Smart Templates (ONLY Missing in Domain) -----
    let domainTplSet = new Set();
    let missingInDomainTpl = [];

    if (haveTpl) {
      domainTplSet = parsedTpl.templateSet;

      missingInDomainTpl = setDiff(extracted.templateSet, domainTplSet);
      missingInDomainTpl.sort((a,b) => String(a).localeCompare(String(b)));

      dTplNonEmpty.textContent = String(parsedTpl.templateNonEmpty);
      dTplUnique.textContent   = String(domainTplSet.size);
      dTplDupes.textContent    = String(Math.max(0, parsedTpl.templateDupes));

      eTplRows.textContent     = String(extracted.templateRows);
      eTplUnique.textContent   = String(extracted.templateSet.size);
      eTplDupes.textContent    = String(Math.max(0, extracted.templateDupes));

      missingInDomainTplCount.textContent = String(missingInDomainTpl.length);

      // no longer computed
      if (missingInExtractTplCount) missingInExtractTplCount.textContent = "0";

      missingInDomainTplPreview.textContent = previewList(missingInDomainTpl, 50);
      missingInDomainTplPreviewCount.textContent = `${Math.min(50, missingInDomainTpl.length)} / ${missingInDomainTpl.length}`;
    } else {
      [dTplNonEmpty, dTplUnique, dTplDupes, eTplRows, eTplUnique, eTplDupes, missingInDomainTplCount]
        .forEach(el => el.textContent = "0");

      if (missingInExtractTplCount) missingInExtractTplCount.textContent = "0";

      missingInDomainTplPreview.textContent = "";
      missingInDomainTplPreviewCount.textContent = "0 / 0";
    }

    // Diagnostics
    const diagDomain1 = haveDomain && parsedDomain.oneColRows ? ` • ${parsedDomain.oneColRows.toLocaleString()} 1-column row(s)` : "";
    const diagDomain2 = haveDomain && parsedDomain.multiColRows ? ` • ${parsedDomain.multiColRows.toLocaleString()} extra-column row(s)` : "";
    const diagTpl1 = haveTpl && parsedTpl.oneColRows ? ` • ${parsedTpl.oneColRows.toLocaleString()} 1-column row(s)` : "";
    const diagTpl2 = haveTpl && parsedTpl.multiColRows ? ` • ${parsedTpl.multiColRows.toLocaleString()} extra-column row(s)` : "";

    const pasteDiagnostics = [
      haveDomain ? `Concept/Eventset paste: ${(diagDomain1 || diagDomain2) ? (diagDomain1 + diagDomain2).trim() : "ok"}` : "Concept/Eventset paste: (skipped)",
      haveTpl ? `Smart template paste: ${(diagTpl1 || diagTpl2) ? (diagTpl1 + diagTpl2).trim() : "ok"}` : "Smart template paste: (skipped)"
    ].join(" • ");

    const parsedDomainRows = haveDomain ? parsedDomain.rowCount.toLocaleString() : "0";
    const parsedTplRows = haveTpl ? parsedTpl.rowCount.toLocaleString() : "0";

    compareStatus.innerHTML =
      `<span class="ok">Compare complete.</span> ` +
      `Parsed <b>${parsedDomainRows}</b> Concept/Eventset row(s) and <b>${parsedTplRows}</b> Smart Template row(s). ` +
      `Differences are computed on <b>unique</b> values.` +
      `<br/>Normalization: <b>${doWhitespace ? "whitespace" : "none"}</b>, <b>${doCase ? "case-insensitive" : "case-sensitive"}</b>.` +
      `<br/><b>Main goal:</b> Missing in Domain (Package \\ Domain).` +
      `<br/>${escapeHtml(pasteDiagnostics)}`;

    // Build name arrays for templates
    const domainTplUnique = setToSortedArray(domainTplSet);
    const domainTplUniqueNames = domainTplUnique.map(v => parsedTpl.templateNameByCki?.get?.(v) || "");

    const extractTplUnique = setToSortedArray(extracted.templateSet);
    const extractTplUniqueNames = extractTplUnique.map(v => extracted.templateNameByCki?.get?.(v) || "");

    // Concept comments for all extract concept unique (helps workbook organization)
    const extractConceptUnique = setToSortedArray(extracted.conceptSet);
    const extractConceptUniqueComments = extractConceptUnique.map(v => extracted.conceptCommentByCki.get(v) || "");

    // Missing concept comments (for missing list)
    const missingInDomainCkiComments = missingInDomainCki.map(v => extracted.conceptCommentByCki.get(v) || "");

    const missingInDomainTplNames = missingInDomainTpl.map(v => extracted.templateNameByCki?.get?.(v) || "");
	
	const domainPasteRows = haveDomain ? (parsedDomain.rowsRaw || []) : [];
	
	const domainTplPasteRows = haveTpl ? (parsedTpl.rowsRaw || []) : [];

    state.lastCompare = {
      normalizationLabel: `${doWhitespace ? "whitespace" : "none"}, ${doCase ? "case-insensitive" : "case-sensitive"}`,
      pasteDiagnostics,

      // Domain
      domainConceptUnique: setToSortedArray(domainCkiSet),
      domainNameUnique: setToSortedArray(domainNameSet),
      domainConceptUniqueCount: domainCkiSet.size,
      domainNameUniqueCount: domainNameSet.size,
      domainConceptNonEmpty: haveDomain ? parsedDomain.conceptNonEmpty : 0,
      domainNameNonEmpty: haveDomain ? parsedDomain.nameNonEmpty : 0,
      domainConceptDupes: haveDomain ? Math.max(0, parsedDomain.conceptDupes) : 0,
      domainNameDupes: haveDomain ? Math.max(0, parsedDomain.nameDupes) : 0,
		
	  domainPasteRows,
	  domainTplPasteRows,
		
      // Extract
      extractConceptUnique,
      extractConceptUniqueComments,
      extractNameUnique: setToSortedArray(extracted.nameSet),
      extractConceptUniqueCount: extracted.conceptSet.size,
      extractNameUniqueCount: extracted.nameSet.size,
      extractConceptRows: extracted.conceptRows,
      extractNameRows: extracted.nameRows,
      extractConceptDupes: Math.max(0, extracted.conceptDupes),
      extractNameDupes: Math.max(0, extracted.nameDupes),

      // Smart Templates (Domain)
      domainTplUnique,
      domainTplUniqueNames,
      domainTplUniqueCount: domainTplSet.size,
      domainTplNonEmpty: haveTpl ? parsedTpl.templateNonEmpty : 0,
      domainTplDupes: haveTpl ? Math.max(0, parsedTpl.templateDupes) : 0,

      // Smart Templates (Extract)
      extractTplUnique,
      extractTplUniqueNames,
      extractTplUniqueCount: extracted.templateSet.size,
      extractTplRows: extracted.templateRows,
      extractTplDupes: Math.max(0, extracted.templateDupes),

      // Missing in Domain ONLY
      missingInDomainCki,
      missingInDomainName,
      missingInDomainCkiComments,
      missingInDomainTpl,
      missingInDomainTplNames
    };

    dlFullWorkbookBtn.disabled = false;

    document.querySelectorAll("#compareOverlay details").forEach(d => d.open = true);
  });

  clearPasteBtn.addEventListener("click", () => {
    domainPaste.value = "";
    smartTemplatePaste.value = "";
    compareStatus.textContent = "Paste your data, then click Run Compare.";
    renderPastePreview(null);
    renderTemplatePastePreview(null);

    state.lastCompare = null;
    dlFullWorkbookBtn.disabled = true;

    [
      dCkiNonEmpty, dCkiUnique, dCkiDupes, eCkiRows, eCkiUnique, eCkiDupes,
      dNameNonEmpty, dNameUnique, dNameDupes, eNameRows, eNameUnique, eNameDupes,
      missingInDomainCkiCount, missingInExtractCkiCount, missingInDomainNameCount, missingInExtractNameCount,

      dTplNonEmpty, dTplUnique, dTplDupes, eTplRows, eTplUnique, eTplDupes,
      missingInDomainTplCount, missingInExtractTplCount
    ].forEach(el => el.textContent = "0");

    missingInDomainCkiPreview.textContent = "";
    missingInDomainNamePreview.textContent = "";
    missingInDomainCkiPreviewCount.textContent = "0 / 0";
    missingInDomainNamePreviewCount.textContent = "0 / 0";

    missingInDomainTplPreview.textContent = "";
    missingInDomainTplPreviewCount.textContent = "0 / 0";

    document.querySelectorAll("#compareOverlay details").forEach(d => d.open = false);
  });

  // ---------- Session export/import ----------
    function buildSessionObject() {
    return {
      app: "Oracle XML Manifest Extractor",
      version: 2,
      exportedAt: new Date().toISOString(),

      inputs: {
        manifest: manifestInput.value || "",
        search: searchInput.value || "",

        // Filter section selections
        filterManifest: filterManifestEl ? (filterManifestEl.value || "") : "",
        filterFile: filterFileEl ? (filterFileEl.value || "") : "",
        filterType: filterTypeEl ? (filterTypeEl.value || "") : "",

        // Advanced
        targetFolderName: targetFolderNameEl.value || "Filters",
        templatesRootName: templatesRootNameEl ? (templatesRootNameEl.value || "Templates") : "Templates",
        standardTemplatesName: standardTemplatesNameEl ? (standardTemplatesNameEl.value || "Standard Templates") : "Standard Templates",
        specialtyTemplatesName: specialtyTemplatesNameEl ? (specialtyTemplatesNameEl.value || "Specialty Templates") : "Specialty Templates",
        templatesLang: templatesLangEl ? (templatesLangEl.value || "en") : "en",
      },

      toggles: {
        // Extractors
        extractFilters: !!(extractFiltersEl && extractFiltersEl.checked),
        extractTemplates: !!(extractTemplatesEl && extractTemplatesEl.checked),
		extractUuids: !!(extractUuidsEl && extractUuidsEl.checked),

        // Parse / output
        includeComments: !!includeCommentsEl.checked,
        usePath: !!usePathEl.checked,
        avoidDupes: !!avoidDupesEl.checked,

        // Compare normalization
        normalizeWhitespace: !!normalizeWhitespaceEl.checked,
        caseInsensitive: !!caseInsensitiveEl.checked,
      },

      compare: {
        domainPaste: domainPaste.value || "",
        smartTemplatePaste: smartTemplatePaste ? (smartTemplatePaste.value || "") : "",
        lastCompare: state.lastCompare || null,
      },

      rows: Array.isArray(state.rows) ? state.rows : []
    };
  }

  function applyLastCompareToUI(lc) {
    if (!lc) {
      dlFullWorkbookBtn.disabled = true;

      [
        dCkiNonEmpty, dCkiUnique, dCkiDupes, eCkiRows, eCkiUnique, eCkiDupes,
        dNameNonEmpty, dNameUnique, dNameDupes, eNameRows, eNameUnique, eNameDupes,
        dTplNonEmpty, dTplUnique, dTplDupes, eTplRows, eTplUnique, eTplDupes,
        missingInDomainCkiCount, missingInDomainNameCount, missingInDomainTplCount
      ].forEach(el => el.textContent = "0");

      // “Missing in extract” is no longer used
      if (missingInExtractCkiCount) missingInExtractCkiCount.textContent = "0";
      if (missingInExtractNameCount) missingInExtractNameCount.textContent = "0";
      if (missingInExtractTplCount) missingInExtractTplCount.textContent = "0";

      missingInDomainCkiPreview.textContent = "Run compare to see preview.";
      missingInDomainNamePreview.textContent = "Run compare to see preview.";
      missingInDomainTplPreview.textContent = "Run compare to see preview.";
      missingInDomainCkiPreviewCount.textContent = "0 / 0";
      missingInDomainNamePreviewCount.textContent = "0 / 0";
      missingInDomainTplPreviewCount.textContent = "0 / 0";

      compareStatus.textContent = "Paste your data, then click Run Compare.";
      document.querySelectorAll("#compareOverlay details").forEach(d => d.open = false);
      return;
    }

    dCkiNonEmpty.textContent = String(lc.domainConceptNonEmpty ?? 0);
    dCkiUnique.textContent   = String(lc.domainConceptUniqueCount ?? (lc.domainConceptUnique?.length ?? 0));
    dCkiDupes.textContent    = String(lc.domainConceptDupes ?? 0);

    eCkiRows.textContent     = String(lc.extractConceptRows ?? 0);
    eCkiUnique.textContent   = String(lc.extractConceptUniqueCount ?? (lc.extractConceptUnique?.length ?? 0));
    eCkiDupes.textContent    = String(lc.extractConceptDupes ?? 0);

    dNameNonEmpty.textContent = String(lc.domainNameNonEmpty ?? 0);
    dNameUnique.textContent   = String(lc.domainNameUniqueCount ?? (lc.domainNameUnique?.length ?? 0));
    dNameDupes.textContent    = String(lc.domainNameDupes ?? 0);

    eNameRows.textContent     = String(lc.extractNameRows ?? 0);
    eNameUnique.textContent   = String(lc.extractNameUniqueCount ?? (lc.extractNameUnique?.length ?? 0));
    eNameDupes.textContent    = String(lc.extractNameDupes ?? 0);

    // Smart template metrics
    dTplNonEmpty.textContent = String(lc.domainTplNonEmpty ?? 0);
    dTplUnique.textContent   = String(lc.domainTplUniqueCount ?? (lc.domainTplUnique?.length ?? 0));
    dTplDupes.textContent    = String(lc.domainTplDupes ?? 0);

    eTplRows.textContent     = String(lc.extractTplRows ?? 0);
    eTplUnique.textContent   = String(lc.extractTplUniqueCount ?? (lc.extractTplUnique?.length ?? 0));
    eTplDupes.textContent    = String(lc.extractTplDupes ?? 0);

    const midCki = Array.isArray(lc.missingInDomainCki) ? lc.missingInDomainCki : [];
    const midName = Array.isArray(lc.missingInDomainName) ? lc.missingInDomainName : [];
    const midTpl = Array.isArray(lc.missingInDomainTpl) ? lc.missingInDomainTpl : [];

    missingInDomainCkiCount.textContent = String(midCki.length);
    missingInDomainNameCount.textContent = String(midName.length);
    missingInDomainTplCount.textContent = String(midTpl.length);

    // “Missing in extract” no longer used
    if (missingInExtractCkiCount) missingInExtractCkiCount.textContent = "0";
    if (missingInExtractNameCount) missingInExtractNameCount.textContent = "0";
    if (missingInExtractTplCount) missingInExtractTplCount.textContent = "0";

    missingInDomainCkiPreview.textContent = previewList(midCki, 50);
    missingInDomainNamePreview.textContent = previewList(midName, 50);
    missingInDomainTplPreview.textContent = previewList(midTpl, 50);

    missingInDomainCkiPreviewCount.textContent = `${Math.min(50, midCki.length)} / ${midCki.length}`;
    missingInDomainNamePreviewCount.textContent = `${Math.min(50, midName.length)} / ${midName.length}`;
    missingInDomainTplPreviewCount.textContent = `${Math.min(50, midTpl.length)} / ${midTpl.length}`;

    compareStatus.innerHTML =
      `<span class="ok">Compare restored from imported session.</span> ` +
      `Normalization: <b>${escapeHtml(lc.normalizationLabel || "unknown")}</b>.<br/>` +
      `${escapeHtml(lc.pasteDiagnostics || "")}`;

    dlFullWorkbookBtn.disabled = false;

    document.querySelectorAll("#compareOverlay details").forEach(d => d.open = true);
  }

    function applySessionObject(obj) {
    if (!obj || !Array.isArray(obj.rows)) {
      throw new Error("Invalid session file (expected JSON with rows[])." );
    }

    if (obj.inputs) {
      manifestInput.value = obj.inputs.manifest ?? "";
      searchInput.value = obj.inputs.search ?? "";

      // Advanced
      targetFolderNameEl.value = obj.inputs.targetFolderName ?? "Filters";
      if (templatesRootNameEl) templatesRootNameEl.value = obj.inputs.templatesRootName ?? "Templates";
      if (standardTemplatesNameEl) standardTemplatesNameEl.value = obj.inputs.standardTemplatesName ?? "Standard Templates";
      if (specialtyTemplatesNameEl) specialtyTemplatesNameEl.value = obj.inputs.specialtyTemplatesName ?? "Specialty Templates";
      if (templatesLangEl) templatesLangEl.value = obj.inputs.templatesLang ?? "en";
    }

    if (obj.toggles) {
      // Extractors
      if (extractFiltersEl) extractFiltersEl.checked = (obj.toggles.extractFilters !== undefined) ? !!obj.toggles.extractFilters : true;
      if (extractTemplatesEl) extractTemplatesEl.checked = (obj.toggles.extractTemplates !== undefined) ? !!obj.toggles.extractTemplates : true;
	  if (extractUuidsEl) extractUuidsEl.checked = (obj.toggles.extractUuids !== undefined) ? !!obj.toggles.extractUuids : true;

      // Parse / output
      includeCommentsEl.checked = !!obj.toggles.includeComments;
      usePathEl.checked = !!obj.toggles.usePath;
      avoidDupesEl.checked = !!obj.toggles.avoidDupes;

      // Compare normalization
      normalizeWhitespaceEl.checked = !!obj.toggles.normalizeWhitespace;
      caseInsensitiveEl.checked = !!obj.toggles.caseInsensitive;
    }

    // Reset rows
    state.rows = [];
    state.keySet.clear();

    for (const r of obj.rows) {
      const row = {
        manifest: String(r.manifest ?? ""),
        file: String(r.file ?? ""),
        type: String(r.type ?? ""),
        value: String(r.value ?? ""),
        comment: String(r.comment ?? "")
      };
      const k = rowKey(row);
      if (!state.keySet.has(k)) {
        state.keySet.add(k);
        state.rows.push(row);
      }
    }

    // Compare inputs
    domainPaste.value = obj.compare?.domainPaste ?? "";
    if (smartTemplatePaste) smartTemplatePaste.value = obj.compare?.smartTemplatePaste ?? "";
    state.lastCompare = obj.compare?.lastCompare ?? null;

    // Restore filter selections AFTER we have rows (so options exist)
    rebuildManifestFilterOptions();

    if (obj.inputs) {
      if (filterManifestEl && (obj.inputs.filterManifest ?? "") !== "") filterManifestEl.value = obj.inputs.filterManifest;
      if (filterFileEl) filterFileEl.value = obj.inputs.filterFile ?? "";
      if (filterTypeEl && (obj.inputs.filterType ?? "") !== "") filterTypeEl.value = obj.inputs.filterType;
    }

	// Render + restore compare UI
	render();
	recomputePastePreviews(); // ✅ updates BOTH paste previews
	applyLastCompareToUI(state.lastCompare);
  }

  exportSessionBtn.addEventListener("click", () => {
    const session = buildSessionObject();
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    download(`oracle_manifest_session_${stamp}.json`, JSON.stringify(session, null, 2), "application/json;charset=utf-8");
  });

  importSessionBtn.addEventListener("click", () => importSessionFile.click());

  importSessionFile.addEventListener("change", async () => {
    const file = importSessionFile.files && importSessionFile.files[0];
    if (!file) return;

    try {
      const text = await file.text();
      const obj = JSON.parse(text);
      applySessionObject(obj);
      statusEl.innerHTML = `<span class="ok">Imported session.</span> Rows restored: <b>${state.rows.length}</b>.`;
    } catch (e) {
      statusEl.innerHTML = `<span class="bad">Import failed.</span> ${escapeHtml(String(e.message || e))}`;
    } finally {
      importSessionFile.value = "";
    }
  });

  // ---------- Upload actions ----------
  addBtn.addEventListener("click", async () => {
    const typedManifest = (manifestInput.value || "").trim();
const inferred = inferRootFolderNameFromFolderSelection();

let manifest = "";
let manifestNote = "";

// AUTO mode: if a folder is selected, always follow its root folder name.
// This is what makes folder1 -> folder2 automatically update the manifest.
if (state.manifestMode === "auto" && inferred.name) {
  manifest = inferred.name;

  const prev = typedManifest;
  if (!prev) {
    manifestNote = inferred.note;
  } else if (prev !== manifest) {
    manifestNote = `Manifest auto-updated to "${manifest}" from the selected folder.`;
  }

  manifestInput.value = manifest; // keep UI in sync
} else if (typedManifest) {
  // MANUAL mode (or no folder root available): use whatever the user typed
  manifest = typedManifest;
} else if (inferred.name) {
  // Fallback if blank but we can infer anyway
  manifest = inferred.name;
  manifestNote = inferred.note;
  manifestInput.value = manifest;
  state.manifestMode = "auto";
} else {
  // No folder selection and no manifest typed
  manifest = "Untitled Manifest";
  manifestNote = `Manifest was blank — using "${manifest}".`;
  manifestInput.value = manifest;
  state.manifestMode = "auto";
}


    const includeComments = includeCommentsEl.checked;
    const usePath = usePathEl.checked;
    const avoidDupes = avoidDupesEl.checked;

    const doFilters = extractFiltersEl ? !!extractFiltersEl.checked : true;
    const doTemplates = extractTemplatesEl ? !!extractTemplatesEl.checked : false;
	const doUuids = extractUuidsEl ? !!extractUuidsEl.checked : false;

	if (!doFilters && !doTemplates && !doUuids) {
	  statusEl.innerHTML = `<span class="bad">Nothing to extract.</span> Enable an extractor, then try again.`;
	  return;
	}

	const sel = collectSelectedFilesForActiveExtractors(doFilters, doTemplates, doUuids);
    const xmlFiles = sel.xmlFiles;
    const htmlFiles = sel.htmlFiles;
    const selectionNote = sel.note;

  const totalFiles =
    (doFilters ? xmlFiles.length : 0) +
    ((doTemplates || doUuids) ? htmlFiles.length : 0);
    if (totalFiles === 0) {
      const want = doFilters && doTemplates ? "XML or HTML" : (doFilters ? "XML" : "HTML");
      statusEl.innerHTML = `<span class="bad">No ${want} files selected.</span> Choose a folder and/or files, then try again.`;
      return;
    }

    addBtn.disabled = true;

    const preface = [manifestNote, selectionNote].filter(Boolean).join("\n\n");
    statusEl.textContent = `${preface ? preface + "\n\n" : ""}Adding ${totalFiles} file(s)…`;

    let addedRows = 0;
    let skippedDupes = 0;
    const errs = [];

	const work = [];
	if (doFilters) {
	  for (const f of xmlFiles) work.push({ kind: "xml", file: f });
	}
	if (doTemplates || doUuids) {
	  for (const f of htmlFiles) work.push({ kind: "html", file: f });
	}

    for (let i = 0; i < work.length; i++) {
      const job = work[i];
      const f = job.file;
      try {
        const raw = await f.text();
        
		const items = job.kind === "xml"
		  ? parseEventsets(raw, includeComments)
		  : [
		      ...(doTemplates ? parseSmartTemplates(raw) : []),
		      ...(doUuids ? parseReferenceUuids(raw) : [])
		    ];
		

        const fileDisplay = usePath ? (f.webkitRelativePath || f.name) : f.name;

        for (const it of items) {
			const row = {
			  manifest,
			  file: fileDisplay,
			  type: String(it?.type ?? ""),
			  value: String(it?.value ?? ""),
			  comment: String(it?.comment ?? "")
			};

          if (avoidDupes) {
            const k = rowKey(row);
            if (state.keySet.has(k)) { skippedDupes++; continue; }
            state.keySet.add(k);
          }

          state.rows.push(row);
          addedRows++;
        }
      } catch (e) {
        errs.push(`❌ ${f.name}\n${String(e.message || e)}\n`);
      }
    }

    folderInput.value = "";
    filesInput.value = "";

    render();

    let msg = `${preface ? preface + "\n\n" : ""}Added ${addedRows} row(s) from ${totalFiles} file(s).`;
    if (avoidDupes && skippedDupes) msg += ` Skipped ${skippedDupes} duplicate row(s).`;
    if (errs.length) msg += `\n\nErrors (${errs.length} file[s]):\n` + errs.join("\n");

    statusEl.textContent = msg;
    addBtn.disabled = false;
  });

  // IMPORTANT: Clear Results clears manifest field + everything else
  clearBtn.addEventListener("click", () => {
    state.rows = [];
    state.keySet.clear();
    state.lastCompare = null;
    state.manifestMode = "auto";

    // Clear main inputs
    manifestInput.value = "";

    // Clear filter inputs
    searchInput.value = "";
    if (filterTypeEl) filterTypeEl.value = "";
    if (filterManifestEl) filterManifestEl.value = "";
    if (filterFileEl) filterFileEl.value = "";
    if (filterValueEl) filterValueEl.value = "";
    if (filterNameEl) filterNameEl.value = "";
    targetFolderNameEl.value = "Filters";
    templatesRootNameEl.value = "Templates";
    templatesLangEl.value = "en";
    standardTemplatesNameEl.value = "Standard Templates";
    specialtyTemplatesNameEl.value = "Specialty Templates";

    extractFiltersEl.checked = true;
    extractTemplatesEl.checked = true;
	if (extractUuidsEl) extractUuidsEl.checked = true;
    folderInput.value = "";
    filesInput.value = "";

    // Clear compare inputs + reset compare UI
	domainPaste.value = "";
	smartTemplatePaste.value = "";
	recomputePastePreviews(); // ✅ updates both preview tables
	applyLastCompareToUI(null);

    dlFullWorkbookBtn.disabled = true;

    render();
    statusEl.textContent = "Cleared all results.";
  });

  csvBtn.addEventListener("click", () => {
    const rows = getFilteredRows();
    if (!rows.length) return;

    const header = ["manifest name","file","type","value","comment"];
    const lines = [header.join(",")];

    for (const r of rows) {
      lines.push([
        escapeCsv(r.manifest),
        escapeCsv(r.file),
        escapeCsv(r.type),
        escapeCsv(r.value),
        escapeCsv(r.comment)
      ].join(","));
    }

    const stamp = new Date().toISOString().slice(0,10);
    download(`oracle_manifest_extract_${stamp}.csv`, lines.join("\n"));
  });

// NEW: If user types a manifest, lock it (manual). If they clear it, return to auto.
manifestInput.addEventListener("input", () => {
  state.manifestMode = manifestInput.value.trim() ? "manual" : "auto";
});

  // ---------- Filters ----------
  function clearFilters() {
    searchInput.value = "";
    if (filterTypeEl) filterTypeEl.value = "";
    if (filterManifestEl) filterManifestEl.value = "";
    if (filterFileEl) filterFileEl.value = "";
    if (filterValueEl) filterValueEl.value = "";
    if (filterNameEl) filterNameEl.value = "";
    render();
  }

  if (clearFiltersBtn) clearFiltersBtn.addEventListener("click", clearFilters);

  // Live updates
  searchInput.addEventListener("input", render);
  if (filterFileEl) filterFileEl.addEventListener("input", render);
  if (filterValueEl) filterValueEl.addEventListener("input", render);
  if (filterNameEl) filterNameEl.addEventListener("input", render);
  if (filterTypeEl) filterTypeEl.addEventListener("change", render);
  if (filterManifestEl) filterManifestEl.addEventListener("change", render);

  // ---------- Column resizing (results table) ----------
  function initResizableResultsTable() {
    const table = document.getElementById("resultsTable");
    if (!table) return;

    const colgroup = table.querySelector("colgroup");
    if (!colgroup) return;

    const cols = Array.from(colgroup.querySelectorAll("col"));
    const ths = Array.from(table.querySelectorAll("thead th"));
    if (!cols.length || cols.length !== ths.length) return;

    // Prevent double-init
    if (table.dataset.colResizeInit === "1") return;
    table.dataset.colResizeInit = "1";

    // Restore widths if present
    try {
      const raw = localStorage.getItem("oracleManifestColWidths");
      const saved = raw ? JSON.parse(raw) : null;
      if (Array.isArray(saved) && saved.length === cols.length) {
        for (let i = 0; i < saved.length; i++) {
          const w = Number(saved[i]);
          if (Number.isFinite(w) && w > 40) cols[i].style.width = `${Math.round(w)}px`;
        }
      }
    } catch (_) {}

    function saveWidths() {
      try {
        const widths = cols.map(c => Math.round(parseFloat(getComputedStyle(c).width) || 0));
        localStorage.setItem("oracleManifestColWidths", JSON.stringify(widths));
      } catch (_) {}
    }

    function startResize(clientX, colIndex) {
      const startX = clientX;
      const startW = parseFloat(getComputedStyle(cols[colIndex]).width) || 0;
      const minW = 90;

      function onMove(e) {
        const x = (e.touches && e.touches[0]) ? e.touches[0].clientX : e.clientX;
        const dx = x - startX;
        const nextW = Math.max(minW, startW + dx);
        cols[colIndex].style.width = `${Math.round(nextW)}px`;
        e.preventDefault();
      }

      function onUp() {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        document.removeEventListener("touchmove", onMove);
        document.removeEventListener("touchend", onUp);
        saveWidths();
      }

      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
      document.addEventListener("touchmove", onMove, { passive: false });
      document.addEventListener("touchend", onUp);
    }

    ths.forEach((th, i) => {
      const handle = document.createElement("div");
      handle.className = "colResizer";
      handle.setAttribute("role", "separator");
      handle.setAttribute("aria-orientation", "vertical");
      handle.setAttribute("aria-label", "Resize column");
      handle.title = "Drag to resize";

      handle.addEventListener("mousedown", (e) => {
        e.preventDefault();
        startResize(e.clientX, i);
      });

      handle.addEventListener("touchstart", (e) => {
        if (!(e.touches && e.touches[0])) return;
        e.preventDefault();
        startResize(e.touches[0].clientX, i);
      }, { passive: false });

      th.appendChild(handle);
    });
  }

  // Initial render
  render();
  renderPastePreview(null);
  initResizableResultsTable();
})();
