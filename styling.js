/* Styling panel for the Sectioned Hierarchy Table */
(function () {
  const template = document.createElement("template");
  template.innerHTML = `
  <style>
    :host { font-family:"72",Arial,sans-serif; font-size:13px; }
    fieldset { border:0; margin:0; padding:8px; display:grid; grid-template-columns:130px 1fr; gap:8px; align-items:center; }
  </style>
  <form><fieldset>
    <label>Title</label><input id="title" type="text">
    <label>Section level</label><input id="sectionLevel" type="number" min="1" max="10">
    <label>Section totals</label><input id="showSectionTotals" type="checkbox">
    <label>Level picker</label><input id="showLevelPicker" type="checkbox">
    <label>Full height (PDF)</label><input id="fullHeight" type="checkbox">
    <label>Rows per page (PDF)</label><input id="pageSize" type="number" min="0" max="500" title="0 = off">
    <label>Page number</label><input id="pageNumber" type="number" min="1" max="999">
    <label>Section filter</label><input id="sectionFilter" type="text" placeholder="e.g. Eastern Cape Checkers">
    <label>Export button</label><input id="showExportButton" type="checkbox">
    <label>Totals position</label><select id="totalsPosition"><option value="top">Top of section</option><option value="bottom">Bottom of section</option></select>
    <label>Shade % columns</label><input id="shadeGrowthColumns" type="checkbox">
    <label>% column shade</label><input id="growthShade" type="color">
    <label>Collapsible sections</label><input id="collapsibleSections" type="checkbox">
    <label>Start collapsed</label><input id="startCollapsed" type="checkbox">
    <label>Group by period</label><input id="groupColumnsByPeriod" type="checkbox">
    <label>Period keywords</label><input id="periodKeywords" type="text">
    <label>Column order</label><input id="columnOrder" type="text" placeholder="e.g. LY,TY,%">
    <label>Diagnostics</label><input id="showDiagnostics" type="checkbox">
    <label>Header colour</label><input id="headerColor" type="color">
    <label>Section colour</label><input id="sectionColor" type="color">
    <label>Growth colours</label><input id="growthColors" type="checkbox">
    <label>Growth keyword</label><input id="growthKeyword" type="text">
    <label>Positive colour</label><input id="positiveColor" type="color">
    <label>Negative colour</label><input id="negativeColor" type="color">
  </fieldset></form>`;

  class SectionTableStyling extends HTMLElement {
    constructor() {
      super();
      this._shadow = this.attachShadow({ mode: "open" });
      this._shadow.appendChild(template.content.cloneNode(true));
      this._shadow.querySelectorAll("input, select").forEach((el) => {
        el.addEventListener("change", () => {
          const v = el.type === "checkbox" ? el.checked : (el.type === "number" ? (Number.isNaN(parseInt(el.value, 10)) ? (parseInt(el.min, 10) || 0) : parseInt(el.value, 10)) : el.value);
          this.dispatchEvent(new CustomEvent("propertiesChanged", { detail: { properties: { [el.id]: v } } }));
        });
      });
    }
    onCustomWidgetBeforeUpdate(changed) { this._p = Object.assign(this._p || {}, changed); }
    onCustomWidgetAfterUpdate() {
      const p = this._p || {};
      this._shadow.querySelectorAll("input, select").forEach((el) => {
        if (p[el.id] === undefined) return;
        if (el.type === "checkbox") el.checked = !!p[el.id]; else el.value = p[el.id];
      });
    }
  }
  customElements.define("com-shoprite-sectiontable-styling", SectionTableStyling);
})();
