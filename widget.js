/* Sectioned Hierarchy Table - SAC custom widget (main web component)
 * Reads a data binding ("tableBinding") and groups hierarchy members into
 * sections at the level chosen by the user.
 * NOTE: the exact field names for hierarchy info on a member (parentId, level)
 * must be confirmed in a live tenant - enable "showDiagnostics" to see raw data.
 */
(function () {
  const template = document.createElement("template");
  template.innerHTML = `
  <style>
    :host { display:block; font-family:"72","72full",Arial,sans-serif; font-size:13px; color:#1d2d3e; }
    .wrap { height:100%; display:flex; flex-direction:column; }
    .bar { display:flex; align-items:center; gap:12px; padding:4px 8px; }
    .title { font-weight:700; font-size:15px; flex:1; }
    .bar label { font-size:12px; }
    select { font:inherit; padding:2px 4px; }
    .export { font:inherit; padding:2px 10px; cursor:pointer; border:1px solid #0a6ed1; background:#fff; color:#0a6ed1; border-radius:4px; }
    .export:hover { background:#ebf5fe; }
    .pager { display:inline-flex; align-items:center; gap:6px; font-size:12px; }
    .pager[hidden] { display:none; }
    .pg { font:inherit; padding:1px 8px; cursor:pointer; border:1px solid #0a6ed1; background:#fff; color:#0a6ed1; border-radius:4px; }
    .pg:disabled { opacity:.35; cursor:default; }
    .scroll { overflow:auto; flex:1; }
    table { border-collapse:collapse; width:100%; }
    th, td { padding:4px 8px; white-space:nowrap; border-bottom:1px solid #d9d9d9; }
    th { position:sticky; top:0; background:var(--hdr,#1d2d3e); color:#fff; text-align:right; font-weight:600; z-index:1; }
    th.l, td.l { text-align:left; }
    th.grp { text-align:center; border-bottom:1px solid #8aa0b8; border-left:2px solid #fff; }
    thead tr:nth-child(2) th { top:26px; }
    td.n { text-align:right; font-variant-numeric:tabular-nums; }
    tr.grand td { font-weight:700; }
    td.gc { background:var(--gshade,#eef3f8); }
    tbody.collapsible tr.sec td { cursor:pointer; user-select:none; }
    tbody.collapsible tr.sec td::before { content:"\\25BE\\00a0"; }
    tbody.collapsible.collapsed tr.sec td::before { content:"\\25B8\\00a0"; }
    tbody.collapsed tr.m { display:none; }
    tr.nh { display:none; }
    td.has { cursor:pointer; user-select:none; }
    td.has::before { content:"\\25BE\\00a0"; color:#6a6d70; }
    td.has.closed::before { content:"\\25B8\\00a0"; }
    tr.sec td { background:var(--sec,#e8eef5); font-weight:700; border-top:2px solid var(--hdr,#1d2d3e); }
    tr.tot td { font-weight:700; border-top:1px solid #1d2d3e; }
    .neg { color:var(--neg,#bb0000); }
    .pos { color:var(--pos,#107e3e); }
    .msg { padding:12px; color:#666; }
    .warn { padding:6px 8px; background:#fff4d6; color:#7a5a00; font-size:12px; }
    pre { margin:0; padding:8px; font-size:11px; background:#f7f7f7; max-height:160px; overflow:auto; }
    /* full-height mode: no inner scrolling, widget grows with its content (better for PDF export) */
    :host([data-full]) .wrap { height:auto; }
    :host([data-full]) .scroll { overflow:visible; }
    :host([data-full]) th { position:static; }
    /* print / PDF: repeat header on each page, keep sections and TOTAL/EXISTING pairs together */
    @media print {
      thead { display:table-header-group; }
      tbody.section { break-inside:avoid; page-break-inside:avoid; }
      tbody.collapsed tr.m { display:table-row; }
      tr.nh { display:table-row; }
      tr { break-inside:avoid; page-break-inside:avoid; }
      tr.keep { break-after:avoid; page-break-after:avoid; }
      tr.sec { break-after:avoid; page-break-after:avoid; }
      .bar select, .pick, .export { display:none; }
      * { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    }
  </style>
  <div class="wrap">
    <div class="bar"><span class="title"></span><span class="pager" hidden><button class="pg prev" type="button" aria-label="Previous page">&#9664;</button><span class="pglbl"></span><button class="pg next" type="button" aria-label="Next page">&#9654;</button></span><label class="pick">Section level: <select></select></label><button class="export" type="button">Export to Excel</button></div>
    <div class="warn" hidden></div>
    <div class="scroll"></div>
    <pre hidden></pre>
  </div>`;

  const val = (v) => (v && typeof v === "object") ? v : { raw: v, formatted: v == null ? "" : String(v) };

  /* ---------- minimal dependency-free .xlsx writer (stored zip, inline strings) ---------- */
  const enc = new TextEncoder();
  const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const zipStore = (files) => {
    const parts = [], central = []; let off = 0;
    for (const f of files) {
      const nm = enc.encode(f.name), d = f.data, crc = crc32(d);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, d.length, true); lh.setUint32(22, d.length, true); lh.setUint16(26, nm.length, true);
      parts.push(new Uint8Array(lh.buffer), nm, d);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, d.length, true); ch.setUint32(24, d.length, true); ch.setUint16(28, nm.length, true);
      ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), nm);
      off += 30 + nm.length + d.length;
    }
    const cdSize = central.reduce((a, p) => a + p.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
    return new Blob([...parts, ...central, new Uint8Array(e.buffer)]);
  };
  const xe = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colName = (i) => { let s = ""; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s; return s; };

  // style ids (see styles.xml below): 0 text, 1 header, 2 bold text, 3 num, 4 bold num, 5 pct(fraction), 6 bold pct, 7 pct(literal), 8 bold pct, 9 section band
  const buildXlsx = (table) => {
    const rowsXml = [], merges = [];
    let maxCol = 0, headRows = table.tHead ? table.tHead.rows.length : 0;
    Array.from(table.rows).forEach((tr, ri) => {
      const isHead = tr.parentElement.tagName === "THEAD";
      const bold = /\b(grand|tot)\b/.test(tr.className);
      const band = /\bsec\b/.test(tr.className);
      let ci = 0, cx = "";
      for (const cell of Array.from(tr.cells)) {
        const ref = colName(ci) + (ri + 1), span = cell.colSpan || 1, text = cell.textContent;
        const raw = cell.dataset ? cell.dataset.raw : undefined;
        const shaded = /\bgc\b/.test(cell.className || "");
        if (span > 1) merges.push(ref + ":" + colName(ci + span - 1) + (ri + 1));
        if (raw !== undefined && !isHead) {
          const n = Number(raw), shown = parseFloat(String(text).replace(/,/g, ""));
          let style = bold ? 4 : 3, val = n;
          if (/%\s*$/.test(text)) {
            if (Math.abs(shown / 100 - n) < 1e-6 + Math.abs(n) * 1e-3) style = bold ? 6 : 5;
            else style = bold ? 8 : 7;
          }
          if (shaded) style = style >= 5 ? style + 5 : style + 11;   // 5-8 -> 10-13, 3-4 -> 14-15
          cx += `<c r="${ref}" s="${style}"><v>${val}</v></c>`;
        } else {
          const s = isHead ? 1 : band ? 9 : shaded ? 16 : bold ? 2 : 0;
          cx += text === "" ? `<c r="${ref}" s="${s}"/>` : `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xe(text)}</t></is></c>`;
        }
        for (let k = 1; k < span; k++) cx += `<c r="${colName(ci + k) + (ri + 1)}" s="${isHead ? 1 : band ? 9 : 0}"/>`;
        ci += span;
      }
      maxCol = Math.max(maxCol, ci);
      rowsXml.push(`<row r="${ri + 1}">${cx}</row>`);
    });
    const cols = `<cols><col min="1" max="1" width="38" customWidth="1"/><col min="2" max="${Math.max(maxCol, 2)}" width="13" customWidth="1"/></cols>`;
    const pane = headRows ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${headRows}" topLeftCell="A${headRows + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : "";
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}${cols}<sheetData>${rowsXml.join("")}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : ""}</worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.0"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="0.0&quot;%&quot;"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1D2D3E"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF5"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="17">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="165" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="165" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="166" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="166" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="164" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
</cellXfs></styleSheet>`;
    const files = [
      ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
      ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Export" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ["xl/styles.xml", styles],
      ["xl/worksheets/sheet1.xml", sheet],
    ].map(([name, xml]) => ({ name, data: enc.encode(xml) }));
    return zipStore(files);
  };
  /* ---------- end xlsx writer ---------- */
  const rawOf = (v) => v == null ? null : (v.raw !== undefined ? v.raw : v.rawValue);
  const fmtOf = (v) => {
    if (v == null) return "";
    if (typeof v !== "object") return String(v);
    if (v.formatted !== undefined) return String(v.formatted);
    if (v.formattedValue !== undefined) return String(v.formattedValue);
    return v.raw !== undefined ? String(v.raw) : "";
  };
  const memberId = (m) => (m && typeof m === "object") ? (m.id !== undefined ? m.id : m.label) : m;
  const memberLabel = (m) => (m && typeof m === "object") ? (m.label !== undefined ? m.label : m.id) : m;

  class SectionTable extends HTMLElement {
    constructor() {
      super();
      this._shadow = this.attachShadow({ mode: "open" });
      this._shadow.appendChild(template.content.cloneNode(true));
      this._props = { sectionLevel: 1, showSectionTotals: true, showLevelPicker: true,
        showDiagnostics: false, groupColumnsByPeriod: true, collapsibleSections: true, startCollapsed: false, shadeGrowthColumns: true, growthShade: "#eef3f8", totalsPosition: "top", showExportButton: true, showPager: true, fullHeight: false, pageSize: 0, pageNumber: 1, sectionFilter: "", growthColors: true, growthKeyword: "%", positiveColor: "#107e3e", periodKeywords: "DAY,WEEK,MONTH,YEAR", columnOrder: "LY,TY,%", title: "", headerColor: "#1d2d3e", sectionColor: "#e8eef5", negativeColor: "#bb0000" };
      this._binding = null;
      this._collapsed = new Map();   // section id -> collapsed? (user toggles)
      this._nodeCollapsed = new Set(); // member rows (below section level) the user collapsed
      this._allDefault = null;       // set by collapseAll()/expandAll()
      this._sectionIds = [];
      this._model = null;
      this._$ = (s) => this._shadow.querySelector(s);
      this._$(".export").addEventListener("click", () => this.exportToExcel());
      // view-time paging: moves only the on-screen page, never changes the saved pageNumber
      this._viewPage = null;
      const step = (d) => { const n = this.getPageCount(); this._viewPage = Math.min(n, Math.max(1, this._curPage() + d)); this._render(); };
      this._$(".prev").addEventListener("click", () => step(-1));
      this._$(".next").addEventListener("click", () => step(1));
      this._$("select").addEventListener("change", (e) => {
        const lvl = parseInt(e.target.value, 10);
        this._props.sectionLevel = lvl;
        this.dispatchEvent(new CustomEvent("propertiesChanged", { detail: { properties: { sectionLevel: lvl } } }));
        this.dispatchEvent(new Event("onSectionLevelChanged"));
        this._render();
      });
    }

    /* ---- SAC lifecycle ---- */
    onCustomWidgetBeforeUpdate(changed) {
      if (changed && (changed.pageNumber !== undefined || changed.pageSize !== undefined)) this._viewPage = null;
      Object.assign(this._props, changed);
    }
    onCustomWidgetAfterUpdate(changed) {
      if (changed && changed.tableBinding !== undefined) this._binding = changed.tableBinding;
      this._buildModel();
      this._render();
    }
    onCustomWidgetResize() { /* layout is CSS driven */ }
    onCustomWidgetDestroy() {}

    // Some SAC versions deliver the binding through a property setter/getter.
    set tableBinding(b) { this._binding = b; this._buildModel(); this._render(); }
    get tableBinding() { return this._binding; }

    /* ---- scripting API ---- */
    setSectionLevel(level) {
      this._props.sectionLevel = Math.max(1, parseInt(level, 10) || 1);
      this.dispatchEvent(new CustomEvent("propertiesChanged", { detail: { properties: { sectionLevel: this._props.sectionLevel } } }));
      this._render();
    }
    getSectionLevel() { return this._props.sectionLevel; }
    _collapsedDefault() { return this._allDefault === null ? !!this._props.startCollapsed : this._allDefault; }
    collapseAll() { this._collapsed.clear(); this._nodeCollapsed.clear(); this._allDefault = true; this._render(); }
    expandAll() { this._collapsed.clear(); this._nodeCollapsed.clear(); this._allDefault = false; this._render(); }
    exportToExcel(fileName) {
      const table = this._$("table");
      if (!table) return;
      const name = (fileName || this._props.title || "Section table").replace(/[\\/:*?"<>|]/g, "_");
      const blob = new Blob([buildXlsx(table)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = name + " " + new Date().toISOString().slice(0, 10) + ".xlsx";
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    }
    setSectionFilter(text) {
      this._props.sectionFilter = String(text || "");
      this.dispatchEvent(new CustomEvent("propertiesChanged", { detail: { properties: { sectionFilter: this._props.sectionFilter } } }));
      this._render();
    }
    getMaxLevel() { return this._model ? this._model.maxDepth : 1; }

    /* ---- data -> model ---- */
    _buildModel() {
      const b = this._binding;
      this._model = null;
      if (!b || b.state !== "success" || !Array.isArray(b.data)) return;

      const rows = b.data;
      const keys = rows.length ? Object.keys(rows[0]) : [];
      const dimKeys = keys.filter((k) => /^dimensions_\d+$/.test(k)).sort();
      const measKeys = keys.filter((k) => /^measures_\d+$/.test(k)).sort();
      const hasMD = keys.includes("@MeasureDimension");
      const hierKey = dimKeys[0];
      const bizKey = dimKeys[1];
      if (!hierKey) return;

      const colOrder = [];           // [{id,label}]
      const colSeen = new Map();
      const addCol = (id, label) => { if (!colSeen.has(id)) { colSeen.set(id, true); colOrder.push({ id, label }); } };
      const metaMeas = (b.metadata && (b.metadata.mainStructureMembers || b.metadata.measures)) || {};

      const nodes = new Map();       // id -> node
      const order = [];
      let sawParent = false;

      for (const r of rows) {
        const m = r[hierKey];
        const id = String(memberId(m));
        let n = nodes.get(id);
        if (!n) {
          const p = (m && typeof m === "object") ? (m.parentId !== undefined ? m.parentId : m.parent) : undefined;
          if (p !== undefined && p !== null && p !== "") sawParent = true;
          n = { id, label: String(memberLabel(m)), parentId: (p === undefined || p === null || p === "") ? null : String(p),
                children: [], depth: 1, biz: new Map(), bizOrder: [] };
          nodes.set(id, n); order.push(n);
        }
        const bizLabel = bizKey ? String(memberLabel(r[bizKey])) : "";
        if (!n.biz.has(bizLabel)) { n.biz.set(bizLabel, new Map()); n.bizOrder.push(bizLabel); }
        const cells = n.biz.get(bizLabel);
        if (hasMD) {
          const md = r["@MeasureDimension"];
          const cid = String(memberId(md));
          addCol(cid, String(memberLabel(md)));
          cells.set(cid, r[measKeys[0] || "measures_0"]);
        } else {
          for (const mk of measKeys) {
            const meta = metaMeas[mk] || {};
            addCol(mk, meta.label || meta.id || mk);
            cells.set(mk, r[mk]);
          }
        }
      }

      // parent links + depth
      for (const n of order) {
        const p = n.parentId && nodes.get(n.parentId);
        if (p && p !== n) p.children.push(n); else n.parentId = null;
      }
      const setDepth = (n, d) => { n.depth = d; n.children.forEach((c) => setDepth(c, d + 1)); };
      const roots = order.filter((n) => !n.parentId);
      roots.forEach((n) => setDepth(n, 1));
      const maxDepth = order.reduce((a, n) => Math.max(a, n.depth), 1);

      this._model = { roots, order, cols: colOrder, maxDepth, sawParent, bizKey: !!bizKey,
        dimLabel: (b.metadata && b.metadata.dimensions && b.metadata.dimensions[hierKey] && b.metadata.dimensions[hierKey].description) || "" };
    }

    /* ---- rendering ---- */
    _curPage() { return this._viewPage || Math.max(1, parseInt(this._props.pageNumber, 10) || 1); }
    _render() {
      this._$(".pager").hidden = true;
      this._renderInner();
      // SAC developer guide: tells the export framework the widget has finished rendering
      this.dispatchEvent(new Event("customWidgetRenderComplete"));
    }
    _renderInner() {
      const root = this._shadow.host;
      root.style.setProperty("--hdr", this._props.headerColor);
      root.style.setProperty("--sec", this._props.sectionColor);
      root.style.setProperty("--neg", this._props.negativeColor);
      root.style.setProperty("--pos", this._props.positiveColor);
      root.style.setProperty("--gshade", this._props.growthShade);
      this._$(".title").textContent = this._props.title || "";
      this._sectionIds = [];
      // paged mode: this instance shows one page of the table (for PDF export / publications)
      const pageSize = Math.max(0, parseInt(this._props.pageSize, 10) || 0);
      const paged = pageSize > 0;
      if (this._props.fullHeight || paged) root.setAttribute("data-full", ""); else root.removeAttribute("data-full");
      const host = this._$(".scroll");
      host.textContent = "";
      const warn = this._$(".warn");
      warn.hidden = true;
      this._$(".pick").hidden = !this._props.showLevelPicker || paged;
      this._$(".pick").style.display = this._$(".pick").hidden ? "none" : "";
      this._$(".export").hidden = !this._props.showExportButton || paged;
      const collapsibleOn = !!this._props.collapsibleSections && !paged;

      const diag = this._$("pre");
      diag.hidden = !this._props.showDiagnostics;
      if (this._props.showDiagnostics) {
        const b = this._binding;
        diag.textContent = b ? JSON.stringify({ state: b.state, metadata: b.metadata, firstRows: (b.data || []).slice(0, 3) }, null, 1) : "no binding";
      }

      let M = this._model;
      if (!M) {
        const d = document.createElement("div");
        d.className = "msg";
        d.textContent = (this._binding && this._binding.state === "loading") ? "Loading..." : "Add a data source and assign dimensions and measures in the Builder panel.";
        host.appendChild(d);
        return;
      }

      // level picker (left empty when hidden: SAC's PDF export prints a hidden <select>'s options)
      const sel = this._$("select");
      sel.textContent = "";
      const pickerShown = !!this._props.showLevelPicker && !paged;
      for (let i = 1; pickerShown && i <= M.maxDepth; i++) {
        const o = document.createElement("option");
        o.value = String(i); o.textContent = "Level " + i;
        sel.appendChild(o);
      }
      const L = Math.min(Math.max(1, this._props.sectionLevel | 0), M.maxDepth);
      if (pickerShown) sel.value = String(L);

      if (!M.sawParent) {
        warn.hidden = false;
        warn.textContent = "No parent information found on the row dimension - is a hierarchy active on it? Showing a flat list (one level). Turn on 'Show diagnostics' to inspect the data.";
      }

      const table = document.createElement("table");
      const head = table.createTHead();
      const mkTh = (row, t, left, span, cls) => {
        const th = document.createElement("th");
        th.textContent = t;
        th.className = (left ? "l " : "") + (cls || "");
        if (span > 1) th.colSpan = span;
        row.appendChild(th);
        return th;
      };

      // Optional period grouping (Day / Week / Month / Year) detected from measure labels
      const kws = String(this._props.periodKeywords || "DAY,WEEK,MONTH,YEAR")
        .split(",").map((s) => s.trim()).filter(Boolean);
      const grouped = !!this._props.groupColumnsByPeriod && kws.length > 0;
      let cols = M.cols.map((c) => ({ id: c.id, label: c.label, group: -1, sub: c.label }));
      if (grouped) {
        const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        cols.forEach((c) => {
          const i = kws.findIndex((k) => new RegExp("(^|[^A-Za-z])" + esc(k) + "([^A-Za-z]|$)", "i").test(c.label));
          if (i >= 0) {
            c.group = i;
            const s = c.label.replace(new RegExp("(^|[^A-Za-z])" + esc(kws[i]) + "(?=[^A-Za-z]|$)", "i"), "$1").replace(/\s+/g, " ").trim();
            c.sub = s || c.label;
          }
        });
        // keep each period together; unmatched measures stay first.
        // Within a period, order by "Column order" (e.g. LY,TY,%); unlisted columns keep SAC's order, after the listed ones.
        const ord = String(this._props.columnOrder || "").split(",").map((s) => s.trim()).filter(Boolean);
        const rank = (c) => {
          const s = (c.sub || c.label).toUpperCase();
          const i = ord.findIndex((k) => /^[A-Za-z0-9]+$/.test(k)
            ? new RegExp("(^|[^A-Za-z0-9])" + esc(k.toUpperCase()) + "([^A-Za-z0-9]|$)").test(s)
            : s.includes(k.toUpperCase()));
          return i < 0 ? ord.length : i;
        };
        cols = cols.map((c, idx) => ({ c, idx, r: c.group >= 0 ? rank(c) : 0 }))
          .sort((a, b) => (a.c.group - b.c.group) || (a.r - b.r) || (a.idx - b.idx)).map((x) => x.c);
      }
      M = Object.assign({}, M, { cols });

      const lead = 1 + (M.bizKey ? 1 : 0);
      if (grouped) {
        const r1 = head.insertRow();
        mkTh(r1, "", true, lead);
        for (let i = 0; i < cols.length;) {
          let j = i;
          while (j < cols.length && cols[j].group === cols[i].group) j++;
          const g = cols[i].group;
          mkTh(r1, g >= 0 ? kws[g].charAt(0).toUpperCase() + kws[g].slice(1).toLowerCase() : "", false, j - i, g >= 0 ? "grp" : "");
          i = j;
        }
      }
      const thead = head.insertRow();
      mkTh(thead, M.dimLabel || "", true, 1);
      if (M.bizKey) mkTh(thead, "", true, 1);
      cols.forEach((c) => mkTh(thead, grouped ? c.sub : c.label, false, 1));
      const sectionBody = () => table.createTBody();
      let tbody = sectionBody();
      const filt = String(this._props.sectionFilter || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
      const passes = (n) => !filt.length || filt.some((f) => n.label.toLowerCase().includes(f) || n.id.toLowerCase() === f);

      const growthKw = String(this._props.growthKeyword || "%").trim();
      const useGrowth = !!this._props.growthColors && growthKw.length > 0;
      const shadePct = !!this._props.shadeGrowthColumns && growthKw.length > 0;
      const isGrowth = (c) => useGrowth && (c.label.toUpperCase().includes(growthKw.toUpperCase()));
      const cellsFor = (tr, node, biz, cls) => {
        const cells = node.biz.get(biz) || new Map();
        for (const c of M.cols) {
          const v = cells.get(c.id);
          const td = tr.insertCell();
          const r = rawOf(v);
          const num = (r === null || r === undefined || r === "") ? NaN : Number(r);
          let k = "n";
          if (!isNaN(num)) {
            if (isGrowth(c)) { if (num > 0) k += " pos"; else if (num < 0) k += " neg"; }
            else if (num < 0) k += " neg";
          }
          if (shadePct && c.label.toUpperCase().includes(growthKw.toUpperCase())) k += " gc";
          td.className = k;
          if (!isNaN(num)) td.dataset.raw = String(num);
          td.textContent = fmtOf(v);
        }
      };
      // one table row per business member (e.g. TOTAL / EXISTING) for a node
      const emitNode = (node, rowClass, indent, labelOverride, anc, expandable) => {
        const bizList = node.bizOrder.length ? node.bizOrder : [""];
        bizList.forEach((biz, i) => {
          const tr = tbody.insertRow();
          tr.className = rowClass + (i < bizList.length - 1 ? " keep" : "");
          tr._anc = anc || [];
          const l = tr.insertCell(); l.className = "l";
          if (i === 0) {
            l.textContent = labelOverride || node.label; l.style.paddingLeft = (8 + indent * 16) + "px";
            if (expandable && node.children.length && collapsibleOn) {
              l.classList.add("has");
              l.setAttribute("role", "button"); l.setAttribute("tabindex", "0");
              l._nid = node.id; tr._lbl = l;
              const tgl = () => {
                if (this._nodeCollapsed.has(node.id)) this._nodeCollapsed.delete(node.id); else this._nodeCollapsed.add(node.id);
                this._applyVis(table);
              };
              l.addEventListener("click", tgl);
              l.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); tgl(); } });
            }
          }
          if (M.bizKey) { const b = tr.insertCell(); b.className = "l"; b.textContent = biz; }
          cellsFor(tr, node, biz);
        });
      };
      const walk = (node, indent, anc) => {
        anc = anc || [];
        emitNode(node, "m", indent, undefined, anc, true);
        node.children.forEach((c) => walk(c, indent + 1, anc.concat(node.id)));
      };

      const visit = (node) => {
        if (node.depth < L) {            // above section level: grand-total style rows
          if (!filt.length) emitNode(node, "grand", node.depth - 1);
          node.children.forEach(visit);
        } else if (node.depth === L) {   // section
          if (!passes(node)) return;
          tbody = sectionBody();         // own <tbody> so print CSS can keep it on one page
          tbody.className = "section";
          this._sectionIds.push(node.id);
          const collapsible = collapsibleOn;
          const isCollapsed = collapsible && (this._collapsed.has(node.id) ? this._collapsed.get(node.id) : this._collapsedDefault());
          if (isCollapsed) tbody.classList.add("collapsed");
          if (collapsible) tbody.classList.add("collapsible");
          const secBody = tbody;
          const hdr = tbody.insertRow(); hdr.className = "sec";
          const c = hdr.insertCell(); c.textContent = node.label; c.colSpan = 1 + (M.bizKey ? 1 : 0) + M.cols.length;
          if (collapsible) {
            c.setAttribute("role", "button"); c.setAttribute("tabindex", "0");
            c.setAttribute("aria-expanded", String(!isCollapsed));
            const toggle = () => {
              const now = secBody.classList.toggle("collapsed");
              this._collapsed.set(node.id, now);
              c.setAttribute("aria-expanded", String(!now));
            };
            c.addEventListener("click", toggle);
            c.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); } });
          }
          const totalsTop = this._props.totalsPosition !== "bottom";
          if (this._props.showSectionTotals && totalsTop) emitNode(node, "tot", 0, "Totals");
          node.children.forEach((ch) => walk(ch, 1));
          if (this._props.showSectionTotals && !totalsTop) emitNode(node, "tot", 0, "Totals");
        }
      };
      M.roots.forEach(visit);

      if (paged) {
        const n = this._paginate(table, pageSize);
        const p = this._curPage();
        const pager = this._$(".pager");
        pager.hidden = !this._props.showPager;
        this._$(".pglbl").textContent = "Page " + Math.min(p, n) + " of " + n;
        this._$(".prev").disabled = p <= 1;
        this._$(".next").disabled = p >= n;
        if (p > n) {
          const d = document.createElement("div");
          d.className = "msg";
          d.textContent = "Page " + p + " is empty: the table has " + n + " page(s) at " + pageSize + " rows per page.";
          host.appendChild(d);
          return;
        }
        this._keepPage(table, p);
      } else {
        this._pages = null;
      }
      host.appendChild(table);
      if (!paged) this._applyVis(table);
    }

    /* ---- paging (PDF export draws only the widget frame, so each instance shows one page) ----
     * Rows are kept in groups: a row marked "keep" (e.g. TOTAL) stays with the next (EXISTING).
     * Whole sections are moved to the next page when they don't fit; a section longer than a
     * page is split and continues under a repeated section band marked "(cont.)". */
    _paginate(table, size) {
      const groupsOf = (rows) => {
        const out = []; let g = [];
        rows.forEach((r) => { g.push(r); if (!r.classList.contains("keep")) { out.push(g); g = []; } });
        if (g.length) out.push(g);
        return out;
      };
      const pages = [[]];
      let used = 0;
      const cur = () => pages[pages.length - 1];
      const newPage = () => { pages.push([]); used = 0; };
      for (const body of Array.from(table.tBodies)) {
        const rows = Array.from(body.rows);
        if (!rows.length) continue;
        const isSec = body.classList.contains("section");
        if (!isSec) {                          // grand-total rows above the sections
          for (const g of groupsOf(rows)) {
            if (used + g.length > size && used > 0) newPage();
            cur().push({ body, rows: g, cont: false }); used += g.length;
          }
          continue;
        }
        if (used + rows.length > size && used > 0 && rows.length <= size) newPage();
        if (used + rows.length <= size) { cur().push({ body, rows, cont: false }); used += rows.length; continue; }
        // section longer than the space left: split it, band first
        const band = rows[0];
        const groups = groupsOf(rows.slice(1));
        if (used > 0 && size - used < 1 + (groups[0] ? groups[0].length : 0)) newPage();
        let entry = { body, rows: [band], cont: false }; cur().push(entry); used += 1;
        for (const g of groups) {
          if (used + g.length > size && entry.rows.length > 1) {
            newPage();
            entry = { body, rows: [band], cont: true }; cur().push(entry); used += 1;
          }
          entry.rows.push(...g); used += g.length;
        }
      }
      this._pages = pages.filter((p) => p.length);
      return this._pages.length || 1;
    }
    _keepPage(table, p) {
      const page = (this._pages || [])[p - 1] || [];
      Array.from(table.tBodies).forEach((b) => b.remove());
      for (const e of page) {
        const tb = table.createTBody();
        tb.className = e.body.className;
        e.rows.forEach((r, i) => {
          if (i === 0 && e.cont) {
            const c = r.cloneNode(true);
            c.cells[0].textContent = r.cells[0].textContent + " (cont.)";
            tb.appendChild(c);
          } else tb.appendChild(r);
        });
      }
    }
    getPageCount() { return this._pages ? this._pages.length || 1 : 1; }
    setPageNumber(n) {
      this._props.pageNumber = Math.max(1, parseInt(n, 10) || 1);
      this._viewPage = null;
      this.dispatchEvent(new CustomEvent("propertiesChanged", { detail: { properties: { pageNumber: this._props.pageNumber } } }));
      this._render();
    }

    // show/hide member rows whose ancestor (below the section) is collapsed
    _applyVis(table) {
      for (const tr of Array.from(table.rows)) {
        if (!tr._anc) continue;
        tr.classList.toggle("nh", tr._anc.some((a) => this._nodeCollapsed.has(a)));
        if (tr._lbl) tr._lbl.classList.toggle("closed", this._nodeCollapsed.has(tr._lbl._nid));
      }
    }
  }

  customElements.define("com-shoprite-sectiontable", SectionTable);
})();
