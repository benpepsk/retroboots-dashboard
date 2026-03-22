(function () {
  const $ = (id) => document.getElementById(id);
  const els = {
    dataSourceLabel: $("data-source-label"),
    generatedAt: $("generated-at"),
    refreshDashboard: $("refresh-dashboard"),
    summaryNote: $("summary-note"),
    workbenchCounts: $("workbench-counts"),
    statusFilter: $("status-filter"),
    brandFilter: $("brand-filter"),
    channelFilter: $("channel-filter"),
    periodFilter: $("period-filter"),
    resetFilters: $("reset-filters"),
    activeFilterChips: $("active-filter-chips"),
    priorityList: $("priority-list"),
    overallKpiGrid: $("overall-kpi-grid"),
    yearKpiGrid: $("year-kpi-grid"),
    trendChart: $("trend-chart"),
    trendMeta: $("trend-meta"),
    businessHealth: $("business-health"),
    forecastOutlook: $("forecast-outlook"),
    topModels: $("top-models"),
    ageBuckets: $("age-buckets"),
    brandPerformance: $("brand-performance"),
    channelPerformance: $("channel-performance"),
    marginDistribution: $("margin-distribution"),
    priceRealization: $("price-realization"),
    countryPerformance: $("country-performance"),
    sizeDistribution: $("size-distribution"),
    purchaseSources: $("purchase-sources"),
    inventoryTurnover: $("inventory-turnover"),
    opportunities: $("opportunities"),
    risks: $("risks"),
    inventoryActions: $("inventory-actions"),
    marginLeaks: $("margin-leaks"),
    opportunitiesCount: $("opportunities-count"),
    risksCount: $("risks-count"),
    inventoryCount: $("inventory-count"),
    marginLeaksCount: $("margin-leaks-count"),
    uiTooltip: $("ui-tooltip"),
    tabButtons: Array.from(document.querySelectorAll("[data-table]")),
    tabPanels: Array.from(document.querySelectorAll("[data-table-panel]"))
  };

  const FILTER_META = [
    { key: "status", label: "Status", el: els.statusFilter },
    { key: "brand", label: "Brand", el: els.brandFilter },
    { key: "channel", label: "Channel", el: els.channelFilter },
    { key: "period", label: "Time Range", el: els.periodFilter }
  ];

  let state = { records: [], summary: {}, dataSource: "local", activeTable: "opportunities", loadIssue: null };

  const euro = (v, d = 2) => new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: d }).format(Number(v) || 0);
  const num = (v, d = 0) => new Intl.NumberFormat("de-DE", { maximumFractionDigits: d }).format(Number(v) || 0);
  const pct = (v) => `${num(v, 1)} %`;
  const ratio = (v) => num(v, 2);
  const compactEuro = (v) => new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR", notation: "compact", maximumFractionDigits: 1 }).format(Number(v) || 0);
  const titleCase = (value) => String(value ?? "")
    .split(/(\s+|\/|-)/)
    .map((part) => (/^\s+$|^\/$|^-$/.test(part) ? part : (part ? part.charAt(0).toUpperCase() + part.slice(1).toLowerCase() : part)))
    .join("");
  const displayModel = (value) => titleCase(String(value ?? "").replace(/\+/g, "+"));
  const displayProductName = (brand, model, fallback = "Unknown", includeBrand = false) => {
    const cleanBrand = brand ? titleCase(brand) : "";
    const cleanModel = model ? displayModel(model) : "";
    if (includeBrand && cleanBrand && cleanModel) return `${cleanBrand} ${cleanModel}`;
    if (cleanBrand) return cleanBrand;
    if (cleanModel) return cleanModel;
    return fallback;
  };
  const displayStatus = (value) => {
    if (value === "verkauft") return "Sold";
    if (value === "lager") return "In Stock";
    return titleCase(value);
  };
  const date = (v) => { const d = new Date(v); return Number.isNaN(d.getTime()) ? null : d; };
  const daysAgo = (n) => { const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() - n); return d; };
  const sum = (rows, fn) => rows.reduce((t, r) => t + (Number(fn(r)) || 0), 0);
  const avg = (rows, fn) => rows.length ? sum(rows, fn) / rows.length : 0;
  const badge = (t, tone) => `<span class="pill ${tone}">${t}</span>`;
  const cell = (a, b) => `<div class="cell-primary"><strong>${a}</strong><span>${b}</span></div>`;
  const uniq = (vals) => [...new Set(vals.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "de"));

  function showTooltip(html, event) {
    if (!els.uiTooltip) return;
    els.uiTooltip.innerHTML = html;
    els.uiTooltip.hidden = false;
    moveTooltip(event);
  }

  function moveTooltip(event) {
    if (!els.uiTooltip || els.uiTooltip.hidden) return;
    const pad = 18;
    const width = els.uiTooltip.offsetWidth || 220;
    const height = els.uiTooltip.offsetHeight || 80;
    const left = Math.min(event.clientX + pad, window.innerWidth - width - 12);
    const top = Math.min(event.clientY + pad, window.innerHeight - height - 12);
    els.uiTooltip.style.left = `${left}px`;
    els.uiTooltip.style.top = `${top}px`;
  }

  function hideTooltip() {
    if (!els.uiTooltip) return;
    els.uiTooltip.hidden = true;
  }

  function parseCsvLine(line) {
    const out = []; let cur = ""; let quoted = false;
    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];
      if (ch === '"') { if (quoted && line[i + 1] === '"') { cur += '"'; i += 1; } else quoted = !quoted; }
      else if (ch === "," && !quoted) { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur);
    return out;
  }

  function nOrNull(v) {
    if (v == null || v === "") return null;
    const n = Number(String(v).replace(",", "."));
    return Number.isFinite(n) ? n : null;
  }

  function excelDateToIso(v) {
    const n = nOrNull(v);
    if (n == null) return v || null;
    const days = Math.floor(n - 25569);
    return new Date(days * 86400 * 1000).toISOString().slice(0, 10);
  }

  function norm(v) {
    const t = String(v ?? "").trim();
    if (!t || t === "System.Xml.XmlElement") return null;
    return t.toLowerCase() === "shopfiy" ? "shopify" : t;
  }

  function normalizeHeaderKey(value) {
    return String(value || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function buildAliasRow(row) {
    const aliased = {};
    Object.keys(row).forEach((key) => {
      aliased[normalizeHeaderKey(key)] = row[key];
    });
    return aliased;
  }

  function pickField(row, keys) {
    for (const key of keys) {
      const normalized = normalizeHeaderKey(key);
      if (row[normalized] != null && row[normalized] !== "") return row[normalized];
    }
    return "";
  }

  function normalizeStatusValue(value) {
    const status = norm(value);
    if (!status) return null;
    const normalized = normalizeHeaderKey(status);
    if (["verkauft", "sold", "sale", "verkauf"].includes(normalized)) return "verkauft";
    if (["lager", "bestand", "instock", "stock", "inventory", "available"].includes(normalized)) return "lager";
    return status;
  }

  function parseFlexibleDate(value) {
    if (!value) return null;
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    const text = String(value).trim();
    const gvizMatch = text.match(/^Date\((\d{4}),(\d{1,2}),(\d{1,2})/);
    if (gvizMatch) {
      const year = Number(gvizMatch[1]);
      const month = Number(gvizMatch[2]);
      const day = Number(gvizMatch[3]);
      return new Date(Date.UTC(year, month, day)).toISOString().slice(0, 10);
    }
    const deMatch = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (deMatch) {
      const day = Number(deMatch[1]);
      const month = Number(deMatch[2]) - 1;
      const year = Number(deMatch[3]);
      return new Date(Date.UTC(year, month, day)).toISOString().slice(0, 10);
    }
    const excelIso = excelDateToIso(text);
    return excelIso || text;
  }

  function buildRecord(row) {
    const ek = nOrNull(pickField(row, ["einkaufspreis_netto", "purchase_price_net", "unit_cost", "cost_net"]));
    const vk = nOrNull(pickField(row, ["verkaufspreis_netto", "sale_price_net", "selling_price_net", "revenue_net"]));
    const status = normalizeStatusValue(pickField(row, ["status", "inventory_status", "state"]));
    const gp = ek != null && vk != null ? vk - ek : null;
    return {
      article_id: nOrNull(pickField(row, ["artikel_id", "article_id", "product_id", "id"])),
      brand: norm(pickField(row, ["marke", "brand"])),
      series: norm(pickField(row, ["serie", "series"])),
      sub_series: norm(pickField(row, ["sub_serie", "sub_series"])),
      variant_name: norm(pickField(row, ["sub_sub_serie", "variant_name", "variant"])),
      generation: norm(pickField(row, ["generation"])),
      color: norm(pickField(row, ["farbe", "color"])),
      surface: norm(pickField(row, ["untergrund", "surface", "sole_type"])),
      size: norm(pickField(row, ["groesse", "size"])),
      condition: norm(pickField(row, ["zustand", "condition"])),
      status,
      purchase_date: parseFlexibleDate(pickField(row, ["kaufdatum", "purchase_date", "buy_date"])),
      purchase_price_net: ek,
      target_sale_price_net: nOrNull(pickField(row, ["ziel_vk_netto", "target_sale_price_net", "target_net"])),
      sale_price_net: vk,
      sale_date: parseFlexibleDate(pickField(row, ["verkaufsdatum", "sale_date", "sold_at"])),
      purchase_platform: norm(pickField(row, ["plattform_einkauf", "purchase_platform", "buy_channel"])),
      sales_channel: norm(pickField(row, ["plattform_verkauf", "sales_channel", "channel"])),
      sales_country: norm(pickField(row, ["land_verkauf", "sales_country", "country"])),
      gross_profit: gp,
      margin_pct: gp != null && vk ? (gp / vk) * 100 : null,
      inventory_units: status === "lager" ? 1 : 0,
      sold_units: status === "verkauft" ? 1 : 0,
      inventory_value_cost: status === "lager" ? ek || 0 : 0,
      days_in_stock: ((Date.now() - (date(parseFlexibleDate(pickField(row, ["kaufdatum", "purchase_date", "buy_date"]))) || new Date()).getTime()) / 86400000) | 0,
      model: [
        pickField(row, ["serie", "series"]),
        pickField(row, ["sub_serie", "sub_series"]),
        pickField(row, ["sub_sub_serie", "variant_name", "variant"])
      ].map(norm).filter(Boolean).join(" ")
    };
  }

  function csvToRecords(text) {
    const lines = text.trim().split(/\r?\n/).filter(Boolean);
    const headers = parseCsvLine(lines[0]);
    return lines.slice(1).map((line) => {
      const vals = parseCsvLine(line);
      const rawRow = {};
      headers.forEach((h, i) => { rawRow[h] = vals[i] || ""; });
      return buildRecord(buildAliasRow(rawRow));
    }).filter((row) => row.article_id || row.brand || row.status);
  }

  function gvizRowsToRecords(table) {
    const headers = (table.cols || []).map((col) => col.label || col.id || "");
    return (table.rows || []).map((row) => {
      const rawRow = {};
      headers.forEach((header, index) => {
        const c = (row.c || [])[index];
        if (!c || c.v == null) { rawRow[header] = ""; return; }
        rawRow[header] = c.v instanceof Date ? c.v.toISOString().slice(0, 10) : c.v;
      });
      return buildRecord(buildAliasRow(rawRow));
    }).filter((row) => row.article_id || row.brand || row.status);
  }

  function loadGoogleSheetViaGviz(sheetId, gid) {
    return new Promise((resolve, reject) => {
      const callbackName = `codexSheetCb_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
      const script = document.createElement("script");
      const timeout = window.setTimeout(() => { cleanup(); reject(new Error("Google Sheet request timed out")); }, 12000);
      function cleanup() {
        window.clearTimeout(timeout);
        if (script.parentNode) script.parentNode.removeChild(script);
        try { delete window[callbackName]; } catch (error) { window[callbackName] = undefined; }
      }
      window[callbackName] = (response) => {
        cleanup();
        if (!response || response.status !== "ok") { reject(new Error("Google Sheet could not be loaded")); return; }
        try { resolve(gvizRowsToRecords(response.table || {})); } catch (error) { reject(error); }
      };
      script.onerror = () => { cleanup(); reject(new Error("Google Sheet script load failed")); };
      script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${gid}&headers=1&tqx=responseHandler:${callbackName}`;
      document.head.appendChild(script);
    });
  }

  async function loadGoogleSheetViaCsv(url) {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`Google Sheet CSV failed (${res.status})`);
    return csvToRecords(await res.text());
  }

  async function loadData() {
    if (window.dashboardSettings && window.dashboardSettings.googleSheetId) {
      let liveError = null;
      if (window.dashboardSettings.googleSheetCsvUrl) {
        try {
          const records = await loadGoogleSheetViaCsv(window.dashboardSettings.googleSheetCsvUrl);
          if (records.length) return { records, summary: { generated_at: new Date().toISOString(), row_count: records.length, live_ok: true, live_mode: "csv" }, dataSource: "google-sheet" };
          if (!liveError) liveError = new Error("Google Sheet CSV returned no rows");
        } catch (e) { if (!liveError) liveError = e; }
      }
      try {
        const records = await loadGoogleSheetViaGviz(window.dashboardSettings.googleSheetId, window.dashboardSettings.googleSheetGid || "0");
        if (records.length) return { records, summary: { generated_at: new Date().toISOString(), row_count: records.length, live_ok: true, live_mode: "gviz" }, dataSource: "google-sheet" };
        if (!liveError) liveError = new Error("Google Sheet returned no rows");
      } catch (e) { if (!liveError) liveError = e; }
      if (window.dashboardLocalData && window.dashboardLocalData.records) {
        return { records: window.dashboardLocalData.records || [], summary: { ...(window.dashboardLocalData.summary || {}), live_ok: false, live_error: liveError ? liveError.message : "Unbekannter Live-Fehler" }, dataSource: "local" };
      }
    }
    if (window.dashboardLocalData && window.dashboardLocalData.records) {
      return { records: window.dashboardLocalData.records || [], summary: window.dashboardLocalData.summary || {}, dataSource: "local" };
    }
    const res = await fetch(window.dashboardSettings.localDataUrl);
    const json = await res.json();
    return { records: json.records || [], summary: json.summary || {}, dataSource: "local" };
  }

  async function refreshDashboardData() {
    els.refreshDashboard.disabled = true;
    els.refreshDashboard.textContent = "Refreshing...";
    try {
      const payload = await loadData();
      state = { ...payload, activeTable: state.activeTable || "opportunities" };
      setupFilters(state.records);
      setActiveTable(state.activeTable);
      render();
    } catch (error) {
      els.priorityList.innerHTML = `<div class="empty-state">Refresh failed: ${error.message}</div>`;
    } finally {
      els.refreshDashboard.disabled = false;
      els.refreshDashboard.innerHTML = '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M13.65 2.35A7.96 7.96 0 0 0 8 0C3.58 0 0 3.58 0 8s3.58 8 8 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0 1 8 14 6 6 0 1 1 8 2c1.66 0 3.14.69 4.22 1.78L9 7h7V0l-2.35 2.35z" fill="currentColor"/></svg> Refresh';
    }
  }

  function setActiveTable(name) {
    state.activeTable = name;
    els.tabButtons.forEach((b) => b.classList.toggle("is-active", b.dataset.table === name));
    els.tabPanels.forEach((p) => p.classList.toggle("is-active", p.dataset.tablePanel === name));
  }

  function setupFilters(records) {
    const fill = (el, vals, formatter = (v) => v) => { el.innerHTML = ""; const o = document.createElement("option"); o.value = "all"; o.textContent = "All"; el.appendChild(o); vals.forEach((v) => { const x = document.createElement("option"); x.value = v; x.textContent = formatter(v); el.appendChild(x); }); };
    fill(els.statusFilter, uniq(records.map((r) => r.status)), displayStatus);
    fill(els.brandFilter, uniq(records.map((r) => r.brand)), titleCase);
    fill(els.channelFilter, uniq(records.map((r) => r.sales_channel)));
    if (!setupFilters.bound) {
      FILTER_META.forEach(({ el }) => el.addEventListener("change", render));
      els.resetFilters.addEventListener("click", () => { FILTER_META.forEach(({ el }) => { el.value = "all"; }); render(); });
      els.refreshDashboard.addEventListener("click", refreshDashboardData);
      els.tabButtons.forEach((b) => b.addEventListener("click", () => setActiveTable(b.dataset.table)));
      // Nav link active state
      document.querySelectorAll(".nav-link").forEach((link) => {
        link.addEventListener("click", () => {
          document.querySelectorAll(".nav-link").forEach((l) => l.classList.remove("active"));
          link.classList.add("active");
        });
      });
      setupFilters.bound = true;
    }
  }

  function filteredRecords() {
    const cutoff = els.periodFilter.value === "all" ? null : daysAgo(Number(els.periodFilter.value));
    return state.records.filter((r) => {
      if (els.statusFilter.value !== "all" && r.status !== els.statusFilter.value) return false;
      if (els.brandFilter.value !== "all" && r.brand !== els.brandFilter.value) return false;
      if (els.channelFilter.value !== "all" && r.sales_channel !== els.channelFilter.value) return false;
      if (cutoff) {
        const d = date(r.sale_date || r.purchase_date);
        if (!d || d < cutoff) return false;
      }
      return true;
    });
  }

  function modelGroups(rows) {
    const map = new Map();
    rows.forEach((r) => {
      const model = displayModel(r.model || r.series || "Unknown");
      const key = `${r.brand || "-"}::${model}`;
      if (!map.has(key)) map.set(key, { brand: r.brand || "-", model, rows: [] });
      map.get(key).rows.push(r);
    });
    return [...map.values()];
  }

  function findOpportunities(rows) {
    return modelGroups(rows).map((g) => {
      const soldRows = g.rows.filter((r) => r.sold_units);
      const sold = soldRows.length;
      const stock = g.rows.filter((r) => r.inventory_units).length;
      const revenue = sum(soldRows, (r) => r.sale_price_net);
      const profit = sum(soldRows, (r) => r.gross_profit);
      const margin = revenue ? (profit / revenue) * 100 : 0;
      let signal = "", tone = "good";
      if (sold >= 2 && stock <= 1) { signal = "Restock review"; tone = "warn"; }
      else if (margin >= 30 && sold >= 2) signal = "Scale the winner";
      else if (revenue >= 1000 && sold >= 1) signal = "Expand demand";
      return {
        model: cell(g.model, `${titleCase(g.brand)} | ${sold} sold | ${euro(revenue, 0)}`),
        brand: titleCase(g.brand), sold_units: num(sold), inventory_units: num(stock),
        margin_badge: badge(pct(margin), margin >= 30 ? "good" : "warn"),
        signal_badge: signal ? badge(signal, tone) : ""
      };
    }).filter((r) => r.signal_badge).sort((a, b) => Number(b.sold_units.replace(/\./g, "")) - Number(a.sold_units.replace(/\./g, ""))).slice(0, 10);
  }

  function findRisks(rows) {
    return modelGroups(rows).map((g) => {
      const inv = g.rows.filter((r) => r.inventory_units);
      const stock = inv.length;
      const sold = g.rows.filter((r) => r.sold_units).length;
      const age = avg(inv, (r) => r.days_in_stock);
      const value = sum(inv, (r) => r.inventory_value_cost);
      const noTarget = inv.filter((r) => !r.target_sale_price_net).length;
      let signal = "", tone = "warn";
      if (stock >= 2 && sold === 0) { signal = "Slow mover"; tone = "bad"; }
      else if (age >= 180) { signal = "Dead stock"; tone = "bad"; }
      else if (noTarget > 0) { signal = "Missing target price"; tone = "warn"; }
      else if (value >= 500 && sold <= 1) signal = "Capital tied up";
      return {
        model: cell(g.model, `${titleCase(g.brand)} | ${stock} in stock | ${euro(value, 0)} cost`),
        brand: titleCase(g.brand), inventory_units: num(stock), days_in_stock: num(age),
        inventory_value_cost: euro(value, 0), signal_badge: signal ? badge(signal, tone) : ""
      };
    }).filter((r) => r.signal_badge).sort((a, b) => Number(b.inventory_units.replace(/\./g, "")) - Number(a.inventory_units.replace(/\./g, ""))).slice(0, 10);
  }

  function inventoryActions(rows) {
    return rows.filter((r) => r.inventory_units).map((r) => {
      let action = "Monitor", tone = "warn";
      if ((r.days_in_stock || 0) > 180) { action = "De-risk now"; tone = "bad"; }
      else if (!r.target_sale_price_net) { action = "Set target price"; tone = "warn"; }
      else if ((r.days_in_stock || 0) > 120) { action = "Review pricing"; tone = "warn"; }
      else if (r.target_sale_price_net && r.purchase_price_net && ((r.target_sale_price_net - r.purchase_price_net) / r.target_sale_price_net) * 100 >= 30) { action = "Hold price"; tone = "good"; }
      return {
        article: cell(displayProductName(r.brand, r.model, `${titleCase(r.brand || "")} #${r.article_id}`.trim(), false), `${r.size || "-"} | ${r.surface || "-"}`),
        brand: r.brand ? titleCase(r.brand) : "-",
        status_badge: badge(displayStatus(r.status || "-"), r.status === "lager" ? "warn" : "good"),
        purchase_price_net: euro(r.purchase_price_net, 0),
        target_sale_price_net: r.target_sale_price_net ? euro(r.target_sale_price_net, 0) : "-",
        days_in_stock: num(r.days_in_stock),
        action_badge: badge(action, tone)
      };
    }).sort((a, b) => Number(b.days_in_stock.replace(/\./g, "")) - Number(a.days_in_stock.replace(/\./g, ""))).slice(0, 14);
  }

  function marginLeaks(rows) {
    return rows.filter((r) => r.sold_units && r.sale_price_net != null).map((r) => {
      const margin = Number(r.margin_pct) || 0;
      let tone = "warn", signal = "Monitor margin";
      if (margin < 0) { tone = "bad"; signal = "Negative sale"; }
      else if (margin < 15) { tone = "bad"; signal = "Clear profit leak"; }
      else if (margin < 25) { tone = "warn"; signal = "Below target margin"; }
      return {
        article: cell(displayProductName(r.brand, r.model, `${titleCase(r.brand || "")} #${r.article_id}`.trim(), false), `${r.sales_channel || "-"} | ${r.sale_date || "-"}`),
        brand: r.brand ? titleCase(r.brand) : "-",
        sale_price_net: euro(r.sale_price_net, 0), profit_net: euro(r.gross_profit, 0),
        margin_badge: badge(pct(margin), tone), signal_badge: badge(signal, tone), _sort: margin
      };
    }).sort((a, b) => a._sort - b._sort).slice(0, 12);
  }

  function renderRankList(container, items, fmt) {
    if (!container) return;
    if (!items.length) { container.innerHTML = `<div class="empty-state">No data for this selection.</div>`; return; }
    const max = Math.max(...items.map((item) => Number(item.value) || 0), 1);
    container.innerHTML = items.map((item) => `<div class="rank-row" data-tooltip-title="${item.label}" data-tooltip-body="${item.tooltip || fmt(item.value)}"><div class="rank-header"><strong>${item.label}</strong><span>${fmt(item.value)}</span></div><div class="rank-bar"><div class="rank-fill" style="width:${(item.value / max) * 100}%"></div></div></div>`).join("");
    container.querySelectorAll(".rank-row").forEach((row) => {
      row.addEventListener("mouseenter", (event) => showTooltip(`<strong>${row.dataset.tooltipTitle}</strong>${row.dataset.tooltipBody}`, event));
      row.addEventListener("mousemove", moveTooltip);
      row.addEventListener("mouseleave", hideTooltip);
    });
  }

  function renderTable(container, rows, cols) {
    if (!container) return;
    if (!rows.length) { container.innerHTML = `<div class="empty-state">No matches for the current selection.</div>`; return; }
    const head = cols.map(([l]) => `<th>${l}</th>`).join("");
    const body = rows.map((row) => `<tr>${cols.map(([, k]) => `<td>${row[k] == null ? "-" : row[k]}</td>`).join("")}</tr>`).join("");
    container.innerHTML = `<div class="table-scroll"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  }

  function renderBusinessHealth(rows) {
    if (!els.businessHealth) return;
    const sold = rows.filter((r) => r.sold_units);
    const inv = rows.filter((r) => r.inventory_units);
    const dead = inv.filter((r) => (r.days_in_stock || 0) > 180);
    const lowMargin = sold.filter((r) => (r.margin_pct || 0) < 20);
    const capitalAtRisk = inv.filter((r) => (r.days_in_stock || 0) > 120);
    const items = [
      { label: "Sell-through", value: pct((sold.length / ((sold.length + inv.length) || 1)) * 100), note: "sold vs. total" },
      { label: "Dead Stock >180d", value: num(dead.length), note: `${euro(sum(dead, (r) => r.inventory_value_cost), 0)} tied up` },
      { label: "Profit Leaks", value: num(lowMargin.length), note: "sales below 20% margin" },
      { label: "Avg Inventory Age", value: `${num(avg(inv, (r) => r.days_in_stock))} days`, note: "active inventory" },
      { label: "No Target Price", value: num(inv.filter((r) => !r.target_sale_price_net).length), note: "pricing risk" },
      { label: "Capital >120d", value: euro(sum(capitalAtRisk, (r) => r.inventory_value_cost), 0), note: "losing mobility" },
      { label: "Active Brands", value: num(uniq(rows.map((r) => r.brand)).length), note: "current selection" }
    ];
    els.businessHealth.innerHTML = items.map((i) => `<div class="mini-metric"><span>${i.label}</span><strong>${i.value}</strong><em>${i.note}</em></div>`).join("");
  }

  function renderAgeBuckets(rows) {
    if (!els.ageBuckets) return;
    const inv = rows.filter((r) => r.inventory_units);
    const buckets = [
      { label: "0-60 Days", min: 0, max: 60 },
      { label: "61-120 Days", min: 61, max: 120 },
      { label: "121-180 Days", min: 121, max: 180 },
      { label: "180+ Days", min: 181, max: Infinity }
    ].map((b) => {
      const m = inv.filter((r) => (Number(r.days_in_stock) || 0) >= b.min && (Number(r.days_in_stock) || 0) <= b.max);
      return { label: `${b.label} | ${num(m.length)} Pairs`, value: sum(m, (r) => r.inventory_value_cost), tooltip: `${num(m.length)} pairs<br>${euro(sum(m, (r) => r.inventory_value_cost), 0)} cost` };
    });
    renderRankList(els.ageBuckets, buckets, (v) => euro(v, 0));
  }

  function renderChannelPerformance(rows) {
    if (!els.channelPerformance) return;
    const groups = new Map();
    rows.filter((r) => r.sold_units && r.sales_channel).forEach((r) => groups.set(r.sales_channel, (groups.get(r.sales_channel) || 0) + (r.sale_price_net || 0)));
    renderRankList(els.channelPerformance, [...groups.entries()].map(([label, value]) => ({ label: titleCase(label), value, tooltip: `Revenue: ${euro(value, 0)}` })).sort((a, b) => b.value - a.value).slice(0, 6), (v) => euro(v, 0));
  }

  function renderBrandPerformance(rows) {
    if (!els.brandPerformance) return;
    const table = uniq(rows.map((r) => r.brand)).map((brand) => {
      const brandRows = rows.filter((r) => r.brand === brand);
      const sold = brandRows.filter((r) => r.sold_units);
      const inv = brandRows.filter((r) => r.inventory_units);
      const rev = sum(sold, (r) => r.sale_price_net);
      const prof = sum(sold, (r) => r.gross_profit);
      const margin = rev ? (prof / rev) * 100 : 0;
      const stockCost = sum(inv, (r) => r.inventory_value_cost);
      const sellThrough = (sold.length / ((sold.length + inv.length) || 1)) * 100;
      return {
        brand: titleCase(brand), sold: num(sold.length), revenue: euro(rev, 0), profit: euro(prof, 0),
        margin_badge: badge(pct(margin), margin >= 30 ? "good" : margin < 15 ? "bad" : "warn"),
        sell_through_badge: badge(pct(sellThrough), sellThrough >= 40 ? "good" : sellThrough < 20 ? "bad" : "warn"),
        stock_cost: euro(stockCost, 0), _profit: prof
      };
    }).filter((r) => r.brand).sort((a, b) => b._profit - a._profit);
    renderTable(els.brandPerformance, table, [["Brand", "brand"], ["Sales", "sold"], ["Net Revenue", "revenue"], ["Net Profit", "profit"], ["Margin", "margin_badge"], ["Sell-through", "sell_through_badge"], ["Inv. Cost", "stock_cost"]]);
  }

  // ─── NEW: Margin Distribution ───
  function renderMarginDistribution(rows) {
    if (!els.marginDistribution) return;
    const sold = rows.filter((r) => r.sold_units && r.margin_pct != null);
    if (!sold.length) { els.marginDistribution.innerHTML = `<div class="empty-state">No margin data available.</div>`; return; }
    const brackets = [
      { label: "Negative", min: -Infinity, max: 0, tone: "bad" },
      { label: "0-15%", min: 0, max: 15, tone: "bad" },
      { label: "15-25%", min: 15, max: 25, tone: "warn" },
      { label: "25-35%", min: 25, max: 35, tone: "warn" },
      { label: "35-50%", min: 35, max: 50, tone: "good" },
      { label: "50%+", min: 50, max: Infinity, tone: "good" }
    ];
    const total = sold.length;
    const data = brackets.map((b) => {
      const count = sold.filter((r) => r.margin_pct >= b.min && r.margin_pct < b.max).length;
      const revenue = sum(sold.filter((r) => r.margin_pct >= b.min && r.margin_pct < b.max), (r) => r.sale_price_net);
      return { ...b, count, pct: (count / total) * 100, revenue };
    });
    els.marginDistribution.innerHTML = data.map((d) => `
      <div class="margin-bar-row">
        <span class="margin-bar-label">${d.label}</span>
        <div class="margin-bar-track"><div class="margin-bar-fill ${d.tone}" style="width:${d.pct}%"></div></div>
        <span class="margin-bar-value">${d.count}</span>
      </div>
    `).join("");
  }

  // ─── NEW: Price Realization ───
  function renderPriceRealization(rows) {
    if (!els.priceRealization) return;
    const sold = rows.filter((r) => r.sold_units && r.sale_price_net != null);
    const withTarget = sold.filter((r) => r.target_sale_price_net);
    const inv = rows.filter((r) => r.inventory_units);
    const invWithTarget = inv.filter((r) => r.target_sale_price_net && r.purchase_price_net);

    const avgSalePrice = avg(sold, (r) => r.sale_price_net);
    const avgTargetPrice = withTarget.length ? avg(withTarget, (r) => r.target_sale_price_net) : 0;
    const realizationRate = avgTargetPrice ? (avgSalePrice / avgTargetPrice) * 100 : 0;
    const aboveTarget = withTarget.filter((r) => r.sale_price_net >= r.target_sale_price_net).length;
    const belowTarget = withTarget.filter((r) => r.sale_price_net < r.target_sale_price_net).length;
    const avgMarkup = invWithTarget.length ? avg(invWithTarget, (r) => ((r.target_sale_price_net - r.purchase_price_net) / r.purchase_price_net) * 100) : 0;

    const items = [
      { label: "Avg Sale Price", value: euro(avgSalePrice, 0), note: `from ${num(sold.length)} sales` },
      { label: "Avg Target Price", value: avgTargetPrice ? euro(avgTargetPrice, 0) : "N/A", note: `${num(withTarget.length)} with target` },
      { label: "Realization Rate", value: realizationRate ? pct(realizationRate) : "N/A", note: "actual vs target" },
      { label: "Above Target", value: num(aboveTarget), note: `${num(belowTarget)} below target` },
      { label: "Avg Markup Plan", value: pct(avgMarkup), note: "target vs cost (inventory)" },
      { label: "Price Discipline", value: withTarget.length ? pct((aboveTarget / withTarget.length) * 100) : "N/A", note: "% sold at or above target" }
    ];
    els.priceRealization.innerHTML = items.map((i) => `<div class="mini-metric"><span>${i.label}</span><strong>${i.value}</strong><em>${i.note}</em></div>`).join("");
  }

  // ─── NEW: Country Performance ───
  function renderCountryPerformance(rows) {
    if (!els.countryPerformance) return;
    const groups = new Map();
    rows.filter((r) => r.sold_units && r.sales_country).forEach((r) => {
      const key = r.sales_country;
      if (!groups.has(key)) groups.set(key, { revenue: 0, count: 0, profit: 0 });
      const g = groups.get(key);
      g.revenue += r.sale_price_net || 0;
      g.count += 1;
      g.profit += r.gross_profit || 0;
    });
    const items = [...groups.entries()].map(([label, data]) => ({
      label: titleCase(label),
      value: data.revenue,
      tooltip: `${num(data.count)} sales<br>${euro(data.revenue, 0)} revenue<br>${euro(data.profit, 0)} profit`
    })).sort((a, b) => b.value - a.value).slice(0, 8);
    renderRankList(els.countryPerformance, items, (v) => euro(v, 0));
  }

  // ─── NEW: Size Distribution ───
  function renderSizeDistribution(rows) {
    if (!els.sizeDistribution) return;
    const groups = new Map();
    rows.filter((r) => r.size).forEach((r) => {
      const key = r.size;
      if (!groups.has(key)) groups.set(key, { sold: 0, stock: 0, revenue: 0 });
      const g = groups.get(key);
      if (r.sold_units) { g.sold += 1; g.revenue += r.sale_price_net || 0; }
      if (r.inventory_units) g.stock += 1;
    });
    const items = [...groups.entries()].map(([label, data]) => ({
      label: `Size ${label}`,
      value: data.sold,
      tooltip: `${num(data.sold)} sold | ${num(data.stock)} in stock<br>${euro(data.revenue, 0)} revenue`
    })).sort((a, b) => b.value - a.value).slice(0, 10);
    renderRankList(els.sizeDistribution, items, (v) => `${num(v)} sold`);
  }

  // ─── NEW: Purchase Sources ───
  function renderPurchaseSources(rows) {
    if (!els.purchaseSources) return;
    const groups = new Map();
    rows.filter((r) => r.purchase_platform).forEach((r) => {
      const key = r.purchase_platform;
      if (!groups.has(key)) groups.set(key, { count: 0, cost: 0, soldCount: 0, revenue: 0 });
      const g = groups.get(key);
      g.count += 1;
      g.cost += r.purchase_price_net || 0;
      if (r.sold_units) { g.soldCount += 1; g.revenue += r.sale_price_net || 0; }
    });
    const items = [...groups.entries()].map(([label, data]) => ({
      label: titleCase(label),
      value: data.count,
      tooltip: `${num(data.count)} purchased<br>${euro(data.cost, 0)} total cost<br>${num(data.soldCount)} sold for ${euro(data.revenue, 0)}`
    })).sort((a, b) => b.value - a.value).slice(0, 8);
    renderRankList(els.purchaseSources, items, (v) => `${num(v)} items`);
  }

  // ─── NEW: Inventory Turnover ───
  function renderInventoryTurnover(rows) {
    if (!els.inventoryTurnover) return;
    const sold = rows.filter((r) => r.sold_units);
    const inv = rows.filter((r) => r.inventory_units);
    const totalSold = sold.length;
    const totalInv = inv.length;

    // Calculate monthly velocity
    const soldDates = sold.map((r) => date(r.sale_date)).filter(Boolean).sort((a, b) => a - b);
    let monthlyVelocity = 0;
    if (soldDates.length >= 2) {
      const spanDays = (soldDates[soldDates.length - 1] - soldDates[0]) / 86400000;
      const spanMonths = Math.max(spanDays / 30, 1);
      monthlyVelocity = totalSold / spanMonths;
    }

    const daysOfStock = monthlyVelocity > 0 ? (totalInv / monthlyVelocity) * 30 : 0;
    const turnoverRatio = totalInv > 0 ? totalSold / totalInv : 0;
    const avgDaysToSell = avg(sold.filter((r) => r.purchase_date && r.sale_date), (r) => {
      const buy = date(r.purchase_date);
      const sell = date(r.sale_date);
      return buy && sell ? (sell - buy) / 86400000 : 0;
    });

    const items = [
      { label: "Monthly Velocity", value: `${num(monthlyVelocity, 1)} /mo`, note: "avg pairs sold per month" },
      { label: "Days of Stock", value: `${num(daysOfStock)} days`, note: "at current sell rate" },
      { label: "Turnover Ratio", value: ratio(turnoverRatio), note: "sold / current stock" },
      { label: "Avg Days to Sell", value: `${num(avgDaysToSell)} days`, note: "purchase to sale" },
      { label: "Stock Coverage", value: monthlyVelocity > 0 ? `${num(totalInv / monthlyVelocity, 1)} months` : "N/A", note: "stock / monthly demand" },
      { label: "Capital Efficiency", value: sold.length ? ratio(sum(sold, (r) => r.sale_price_net) / Math.max(sum(sold, (r) => r.purchase_price_net), 1)) : "N/A", note: "revenue per cost invested" }
    ];
    els.inventoryTurnover.innerHTML = items.map((i) => `<div class="mini-metric"><span>${i.label}</span><strong>${i.value}</strong><em>${i.note}</em></div>`).join("");
  }

  function addMonths(baseDate, offset) {
    return new Date(baseDate.getFullYear(), baseDate.getMonth() + offset, 1);
  }

  function monthKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }

  // ─── FULL-CYCLE FORECAST ENGINE v2 ───
  // Considers: purchase pipeline, model demand, size velocity, brand mix,
  // aging decay, fresh-stock boost, and dynamic stock evolution.

  function agingDecay(daysInStock) {
    if (daysInStock <= 30) return 1.0;
    if (daysInStock <= 60) return 0.95;
    if (daysInStock <= 90) return 0.82;
    if (daysInStock <= 120) return 0.65;
    if (daysInStock <= 180) return 0.40;
    if (daysInStock <= 270) return 0.22;
    if (daysInStock <= 365) return 0.10;
    return 0.04;
  }

  function agingPriceDiscount(daysInStock) {
    if (daysInStock <= 30) return 1.02;  // fresh stock can sell above target
    if (daysInStock <= 60) return 1.0;
    if (daysInStock <= 120) return 0.94;
    if (daysInStock <= 180) return 0.84;
    if (daysInStock <= 270) return 0.70;
    return 0.55;
  }

  function buildForecastIntelligence(rows) {
    const sold = rows.filter((r) => r.sold_units);
    const inventory = rows.filter((r) => r.inventory_units);
    const all = rows;

    // ─── BRAND INTELLIGENCE ───
    const brandMap = new Map();
    all.forEach((r) => {
      const b = r.brand || "_unknown";
      if (!brandMap.has(b)) brandMap.set(b, {
        sold: 0, stock: 0, total: 0, revenue: 0, profit: 0,
        soldDates: [], purchaseDates: [], purchaseCost: 0,
        models: new Map(), sizes: new Map()
      });
      const g = brandMap.get(b);
      g.total += 1;
      if (r.sold_units) {
        g.sold += 1;
        g.revenue += r.sale_price_net || 0;
        g.profit += r.gross_profit || 0;
        const d = date(r.sale_date);
        if (d) g.soldDates.push(d);
      }
      if (r.inventory_units) g.stock += 1;
      const pd = date(r.purchase_date);
      if (pd) g.purchaseDates.push(pd);
      g.purchaseCost += r.purchase_price_net || 0;
      // Model tracking
      const model = r.model || r.series || "_unknown";
      if (!g.models.has(model)) g.models.set(model, { sold: 0, stock: 0 });
      const mm = g.models.get(model);
      if (r.sold_units) mm.sold += 1;
      if (r.inventory_units) mm.stock += 1;
      // Size tracking
      const sz = r.size || "_unknown";
      if (!g.sizes.has(sz)) g.sizes.set(sz, { sold: 0, stock: 0 });
      const sm = g.sizes.get(sz);
      if (r.sold_units) sm.sold += 1;
      if (r.inventory_units) sm.stock += 1;
    });

    brandMap.forEach((g) => {
      g.sellThrough = g.total ? g.sold / g.total : 0;
      g.avgPrice = g.sold ? g.revenue / g.sold : 0;
      g.avgProfit = g.sold ? g.profit / g.sold : 0;
      g.avgMargin = g.revenue ? g.profit / g.revenue : 0;
      const sortedSales = g.soldDates.sort((a, b) => a - b);
      if (sortedSales.length >= 2) {
        const span = (sortedSales[sortedSales.length - 1] - sortedSales[0]) / 86400000;
        g.monthlyVelocity = g.sold / Math.max(span / 30, 1);
      } else {
        g.monthlyVelocity = g.sold > 0 ? g.sold / 6 : 0;
      }
      const sortedPurchases = g.purchaseDates.sort((a, b) => a - b);
      if (sortedPurchases.length >= 2) {
        const span = (sortedPurchases[sortedPurchases.length - 1] - sortedPurchases[0]) / 86400000;
        g.monthlyPurchaseRate = g.total / Math.max(span / 30, 1);
      } else {
        g.monthlyPurchaseRate = 0;
      }
    });

    // ─── MODEL INTELLIGENCE ───
    const modelMap = new Map();
    all.forEach((r) => {
      const key = `${r.brand || ""}::${r.model || r.series || ""}`;
      if (!modelMap.has(key)) modelMap.set(key, {
        brand: r.brand, model: r.model || r.series || "_unknown",
        sold: 0, stock: 0, revenue: 0, profit: 0, avgDaysToSell: []
      });
      const m = modelMap.get(key);
      if (r.sold_units) {
        m.sold += 1;
        m.revenue += r.sale_price_net || 0;
        m.profit += r.gross_profit || 0;
        const buy = date(r.purchase_date);
        const sell = date(r.sale_date);
        if (buy && sell) m.avgDaysToSell.push((sell - buy) / 86400000);
      }
      if (r.inventory_units) m.stock += 1;
    });

    modelMap.forEach((m) => {
      m.sellThrough = (m.sold + m.stock) ? m.sold / (m.sold + m.stock) : 0;
      m.avgPrice = m.sold ? m.revenue / m.sold : 0;
      m.velocity = m.avgDaysToSell.length
        ? m.avgDaysToSell.reduce((a, b) => a + b, 0) / m.avgDaysToSell.length
        : 999;
      m.demandScore = Math.min(1, (m.sellThrough * 0.5) + (Math.min(m.sold, 10) / 10 * 0.3) + (m.velocity < 90 ? 0.2 : m.velocity < 180 ? 0.1 : 0));
    });

    // ─── SIZE INTELLIGENCE ───
    const sizeMap = new Map();
    all.forEach((r) => {
      const sz = r.size;
      if (!sz) return;
      if (!sizeMap.has(sz)) sizeMap.set(sz, { sold: 0, stock: 0, revenue: 0 });
      const s = sizeMap.get(sz);
      if (r.sold_units) { s.sold += 1; s.revenue += r.sale_price_net || 0; }
      if (r.inventory_units) s.stock += 1;
    });

    sizeMap.forEach((s) => {
      s.sellThrough = (s.sold + s.stock) ? s.sold / (s.sold + s.stock) : 0;
      s.demandScore = Math.min(1, s.sellThrough * 1.3);
    });

    // ─── PURCHASE PIPELINE ───
    const now = new Date();
    const recent3mo = new Date(now.getFullYear(), now.getMonth() - 3, 1);
    const recentPurchases = all.filter((r) => {
      const d = date(r.purchase_date);
      return d && d >= recent3mo;
    });
    const monthlyPurchaseRate = recentPurchases.length / 3;
    const purchaseBrandMix = new Map();
    recentPurchases.forEach((r) => {
      const b = r.brand || "_unknown";
      purchaseBrandMix.set(b, (purchaseBrandMix.get(b) || 0) + 1);
    });
    const avgPurchaseCost = recentPurchases.length
      ? sum(recentPurchases, (r) => r.purchase_price_net) / recentPurchases.length : 0;

    // Purchase quality: are we buying brands that sell well?
    let purchaseQuality = 0;
    if (recentPurchases.length) {
      purchaseQuality = recentPurchases.reduce((total, r) => {
        const bd = brandMap.get(r.brand || "_unknown");
        return total + (bd ? bd.sellThrough : 0.15);
      }, 0) / recentPurchases.length;
    }

    return { brandMap, modelMap, sizeMap, monthlyPurchaseRate, purchaseBrandMix, avgPurchaseCost, purchaseQuality };
  }

  function scoreInventoryItem(item, intel, globalPriceRealization) {
    const brand = intel.brandMap.get(item.brand || "_unknown") || { sellThrough: 0.15, avgPrice: 0, avgMargin: 0, monthlyVelocity: 0 };
    const modelKey = `${item.brand || ""}::${item.model || item.series || ""}`;
    const model = intel.modelMap.get(modelKey) || { demandScore: 0.15, avgPrice: 0, velocity: 999 };
    const size = intel.sizeMap.get(item.size) || { demandScore: 0.2 };
    const days = item.days_in_stock || 0;

    // Layer 1: Brand velocity (0-1)
    const brandScore = Math.min(1, brand.sellThrough * 1.4);
    // Layer 2: Model demand (0-1)
    const modelScore = model.demandScore;
    // Layer 3: Size demand (0-1)
    const sizeScore = size.demandScore;
    // Layer 4: Aging decay (0-1)
    const ageScore = agingDecay(days);
    // Layer 5: Pricing readiness (0-1)
    const pricingScore = item.target_sale_price_net ? 1.0 : 0.55;

    // Weighted combination (model & brand matter most, then age, then size, then pricing)
    const sellProbability = Math.min(0.95, Math.max(0.02,
      Math.pow(brandScore, 0.25) *
      Math.pow(modelScore, 0.30) *
      Math.pow(sizeScore, 0.10) *
      Math.pow(ageScore, 0.25) *
      Math.pow(pricingScore, 0.10)
    ));

    // Expected sale price with aging discount
    const basePrice = item.target_sale_price_net
      ? item.target_sale_price_net * globalPriceRealization
      : brand.avgPrice || model.avgPrice || item.purchase_price_net * 1.3 || 100;
    const expectedPrice = basePrice * agingPriceDiscount(days);
    const cost = item.purchase_price_net || 0;

    return {
      item, sellProbability, expectedPrice,
      expectedProfit: expectedPrice - cost, cost,
      brandScore, modelScore, sizeScore, ageScore, pricingScore,
      daysInStock: days, brand: item.brand,
      model: item.model || item.series || "_unknown",
      size: item.size
    };
  }

  function buildForecast(rows) {
    const sold = rows.filter((r) => r.sold_units && date(r.sale_date));
    const inventory = rows.filter((r) => r.inventory_units);
    if (!sold.length || !inventory.length) {
      return { scenarios: null, method: "Not enough data." };
    }

    const intel = buildForecastIntelligence(rows);

    // Global price realization
    const withTarget = sold.filter((r) => r.target_sale_price_net);
    const globalPriceRealization = withTarget.length
      ? sum(withTarget, (r) => r.sale_price_net) / sum(withTarget, (r) => r.target_sale_price_net)
      : 0.92;

    // Score every inventory item
    const scoredItems = inventory.map((item) => scoreInventoryItem(item, intel, globalPriceRealization));
    scoredItems.sort((a, b) => b.sellProbability - a.sellProbability);

    // Demand history
    const lastSaleDate = sold.map((r) => date(r.sale_date)).sort((a, b) => a - b).pop();
    const anchor = new Date(lastSaleDate.getFullYear(), lastSaleDate.getMonth(), 1);
    const historyMonths = Array.from({ length: 6 }, (_, i) => addMonths(anchor, i - 5));
    const historyMap = new Map(historyMonths.map((d) => [monthKey(d), 0]));
    sold.forEach((r) => {
      const d = date(r.sale_date);
      if (!d) return;
      const key = monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
      if (historyMap.has(key)) historyMap.set(key, historyMap.get(key) + 1);
    });
    const historyArr = historyMonths.map((d) => historyMap.get(monthKey(d)));
    const weightedMonthly = historyArr.reduce((t, v, i) => t + v * (i + 1), 0) / historyArr.reduce((t, _, i) => t + (i + 1), 0);
    const recent3Avg = avg(historyArr.slice(-3).map((v) => ({ v })), (o) => o.v);
    const prior3Avg = avg(historyArr.slice(0, 3).map((v) => ({ v })), (o) => o.v);
    const momentum = prior3Avg > 0 ? Math.min(1.5, Math.max(0.6, recent3Avg / prior3Avg)) : 1;

    // ─── PURCHASE PIPELINE PROJECTION ───
    // How many new items arrive per month, and what quality?
    const purchaseRate = intel.monthlyPurchaseRate;
    const purchaseQuality = intel.purchaseQuality;
    const avgNewCost = intel.avgPurchaseCost || avg(inventory, (r) => r.purchase_price_net);

    // Build typical new-arrival profile from recent purchase mix
    function generateNewArrivals(count) {
      const arrivals = [];
      const brandWeights = [...intel.purchaseBrandMix.entries()].sort((a, b) => b[1] - a[1]);
      const totalWeight = brandWeights.reduce((t, [, w]) => t + w, 0) || 1;
      for (let i = 0; i < Math.round(count); i++) {
        // Pick brand based on recent purchase distribution
        let roll = Math.random() * totalWeight;
        let pickedBrand = brandWeights[0] ? brandWeights[0][0] : "_unknown";
        for (const [brand, weight] of brandWeights) {
          roll -= weight;
          if (roll <= 0) { pickedBrand = brand; break; }
        }
        const bd = intel.brandMap.get(pickedBrand) || { avgPrice: 200, avgMargin: 0.3, sellThrough: 0.3 };
        const estTargetPrice = avgNewCost * (1 + Math.max(bd.avgMargin || 0.3, 0.2));
        arrivals.push({
          sellProbability: Math.min(0.90, Math.max(0.3,
            Math.pow(Math.min(1, (bd.sellThrough || 0.3) * 1.4), 0.3) * 0.95 // fresh stock boost
          )),
          expectedPrice: estTargetPrice * globalPriceRealization * 1.02, // fresh premium
          cost: avgNewCost,
          expectedProfit: estTargetPrice * globalPriceRealization * 1.02 - avgNewCost,
          daysInStock: 0,
          brand: pickedBrand,
          isNewArrival: true
        });
      }
      return arrivals;
    }

    // ─── THREE SCENARIOS ───
    const scenarioConfig = {
      conservative: { demandFactor: 0.7, priceFactor: 0.90, purchaseFactor: 0.5 },
      base:         { demandFactor: 1.0, priceFactor: 1.0,  purchaseFactor: 1.0 },
      optimistic:   { demandFactor: 1.3, priceFactor: 1.06, purchaseFactor: 1.2 }
    };

    function runScenario(config) {
      const monthlyDemand = Math.max(1, weightedMonthly * momentum * config.demandFactor);
      const monthlyNewStock = purchaseRate * config.purchaseFactor;
      let pool = [...scoredItems];
      const months = [];
      let totalSold = 0, totalRevenue = 0, totalProfit = 0;
      let totalNewArrivals = 0, totalNewCost = 0;

      for (let m = 0; m < 3; m++) {
        const monthDate = addMonths(anchor, m + 1);
        const label = monthDate.toLocaleDateString("en-GB", { month: "short", year: "2-digit" });

        // STEP 1: New arrivals enter the pool
        const arrivals = generateNewArrivals(monthlyNewStock);
        pool = [...pool, ...arrivals].sort((a, b) => b.sellProbability - a.sellProbability);
        totalNewArrivals += arrivals.length;
        totalNewCost += arrivals.length * avgNewCost;

        // STEP 2: Sell top items by probability
        const budget = Math.min(Math.round(monthlyDemand), pool.length);
        const selling = pool.slice(0, budget);
        pool = pool.slice(budget);

        let monthRevenue = 0, monthProfit = 0;
        let newSold = 0, oldSold = 0;
        selling.forEach((s) => {
          const price = s.expectedPrice * config.priceFactor;
          monthRevenue += price;
          monthProfit += price - s.cost;
          if (s.isNewArrival) newSold++; else oldSold++;
        });

        // STEP 3: Age remaining pool items (+30 days per month)
        pool.forEach((s) => {
          if (!s.isNewArrival) {
            s.daysInStock += 30;
            s.sellProbability = Math.max(0.02, s.sellProbability * 0.92);
            s.expectedPrice *= 0.97;
          }
        });

        totalSold += selling.length;
        totalRevenue += monthRevenue;
        totalProfit += monthProfit;

        months.push({
          label, units: selling.length, revenue: monthRevenue, profit: monthProfit,
          remainingStock: pool.length, newArrivals: arrivals.length,
          newSold, oldSold,
          startStock: selling.length + pool.length
        });
      }

      return {
        totalUnits: totalSold, totalRevenue, totalProfit,
        totalMargin: totalRevenue ? (totalProfit / totalRevenue) * 100 : 0,
        monthly: months, remainingStock: pool.length,
        totalNewArrivals, totalNewCost,
        investmentRequired: totalNewCost
      };
    }

    const scenarios = {
      conservative: runScenario(scenarioConfig.conservative),
      base: runScenario(scenarioConfig.base),
      optimistic: runScenario(scenarioConfig.optimistic)
    };

    // ─── DRIVERS ───
    const avgSellProb = avg(scoredItems.map((s) => ({ p: s.sellProbability })), (o) => o.p);
    const highProbItems = scoredItems.filter((s) => s.sellProbability > 0.5).length;
    const lowProbItems = scoredItems.filter((s) => s.sellProbability < 0.2).length;
    const freshStock = scoredItems.filter((s) => s.daysInStock <= 60).length;
    const deadStock = scoredItems.filter((s) => s.daysInStock > 180).length;

    // Top models by expected revenue contribution
    const modelContrib = new Map();
    scoredItems.forEach((s) => {
      const key = s.model || "_unknown";
      if (!modelContrib.has(key)) modelContrib.set(key, { count: 0, expectedRev: 0, avgProb: 0 });
      const g = modelContrib.get(key);
      g.count += 1;
      g.expectedRev += s.expectedPrice * s.sellProbability;
      g.avgProb += s.sellProbability;
    });
    const topModels = [...modelContrib.entries()].map(([model, d]) => ({
      model: displayModel(model), count: d.count,
      expectedRevenue: d.expectedRev, avgProb: d.avgProb / d.count
    })).sort((a, b) => b.expectedRevenue - a.expectedRevenue).slice(0, 5);

    // Size analysis for forecast
    const sizeContrib = new Map();
    scoredItems.forEach((s) => {
      const sz = s.size || "_unknown";
      if (!sizeContrib.has(sz)) sizeContrib.set(sz, { count: 0, avgProb: 0 });
      const g = sizeContrib.get(sz);
      g.count += 1;
      g.avgProb += s.sellProbability;
    });
    const topSizes = [...sizeContrib.entries()].map(([size, d]) => ({
      size, count: d.count, avgProb: d.avgProb / d.count
    })).sort((a, b) => b.avgProb - a.avgProb).slice(0, 5);
    const worstSizes = [...sizeContrib.entries()].map(([size, d]) => ({
      size, count: d.count, avgProb: d.avgProb / d.count
    })).sort((a, b) => a.avgProb - b.avgProb).slice(0, 3);

    // Brand contribution
    const brandContrib = new Map();
    scoredItems.forEach((s) => {
      const b = s.brand || "_unknown";
      if (!brandContrib.has(b)) brandContrib.set(b, { count: 0, totalProb: 0, expectedRev: 0 });
      const g = brandContrib.get(b);
      g.count += 1;
      g.totalProb += s.sellProbability;
      g.expectedRev += s.expectedPrice * s.sellProbability;
    });
    const brandContribArr = [...brandContrib.entries()].map(([brand, d]) => ({
      brand: titleCase(brand), count: d.count, avgProb: d.totalProb / d.count,
      expectedRevenue: d.expectedRev
    })).sort((a, b) => b.expectedRevenue - a.expectedRevenue);

    const drivers = {
      avgSellProbability: avgSellProb, highProbItems, lowProbItems,
      momentum, monthlyDemandBase: weightedMonthly,
      priceRealization: globalPriceRealization,
      avgAge: avg(scoredItems.map((s) => ({ d: s.daysInStock })), (o) => o.d),
      freshStock, deadStock, totalInventory: inventory.length,
      purchaseRate, purchaseQuality, avgPurchaseCost: avgNewCost,
      brandContributions: brandContribArr,
      topModels, topSizes, worstSizes,
      purchaseBrandMix: [...intel.purchaseBrandMix.entries()]
        .map(([b, c]) => ({ brand: titleCase(b), count: c }))
        .sort((a, b) => b.count - a.count).slice(0, 5)
    };

    // Waterfall
    const base = scenarios.base;
    const waterfall = base.monthly.map((m) => ({
      label: m.label, sold: m.units, newArrivals: m.newArrivals,
      remainingStock: m.remainingStock, startStock: m.startStock
    }));

    const brandCount = brandContrib.size;
    const modelCount = modelContrib.size;
    return {
      scenarios, drivers, waterfall,
      method: `Full-cycle forecast scoring ${num(inventory.length)} items across ${num(modelCount)} models, ${num(brandCount)} brands and ${num(sizeContrib.size)} sizes. Layers: brand velocity, model demand, size demand, aging decay (sigmoid), pricing readiness. Purchase pipeline: ${num(purchaseRate, 1)} pairs/mo projected arrivals (quality score: ${pct(purchaseQuality * 100)}). Dynamic stock model adds new inventory each month and ages unsold stock. Momentum: ${ratio(momentum)}.`
    };
  }

  function renderForecast(rows) {
    if (!els.forecastOutlook) return;
    const forecast = buildForecast(rows);
    if (!forecast.scenarios) {
      els.forecastOutlook.innerHTML = `<div class="empty-state">${forecast.method}</div>`;
      return;
    }

    const s = forecast.scenarios;
    const d = forecast.drivers;
    const w = forecast.waterfall;
    const base = s.base;

    els.forecastOutlook.innerHTML = `
      <p class="forecast-section-label">3-Month Scenario Comparison</p>
      <div class="forecast-scenarios">
        <div class="scenario-card conservative">
          <span class="scenario-label">Conservative</span>
          <span class="scenario-value">${euro(s.conservative.totalRevenue, 0)}</span>
          <span class="scenario-sub">${num(s.conservative.totalUnits)} sales &middot; ${pct(s.conservative.totalMargin)} margin</span>
          <span class="scenario-detail">Profit: ${euro(s.conservative.totalProfit, 0)}<br>+${num(s.conservative.totalNewArrivals)} new arrivals<br>Investment: ${euro(s.conservative.investmentRequired, 0)}</span>
        </div>
        <div class="scenario-card base">
          <span class="scenario-label">Base Case</span>
          <span class="scenario-value">${euro(base.totalRevenue, 0)}</span>
          <span class="scenario-sub">${num(base.totalUnits)} sales &middot; ${pct(base.totalMargin)} margin</span>
          <span class="scenario-detail">Profit: ${euro(base.totalProfit, 0)}<br>+${num(base.totalNewArrivals)} new arrivals<br>Investment: ${euro(base.investmentRequired, 0)}</span>
        </div>
        <div class="scenario-card optimistic">
          <span class="scenario-label">Optimistic</span>
          <span class="scenario-value">${euro(s.optimistic.totalRevenue, 0)}</span>
          <span class="scenario-sub">${num(s.optimistic.totalUnits)} sales &middot; ${pct(s.optimistic.totalMargin)} margin</span>
          <span class="scenario-detail">Profit: ${euro(s.optimistic.totalProfit, 0)}<br>+${num(s.optimistic.totalNewArrivals)} new arrivals<br>Investment: ${euro(s.optimistic.investmentRequired, 0)}</span>
        </div>
      </div>

      <p class="forecast-section-label">Monthly Flow (Base Case): Stock + Arrivals &minus; Sales</p>
      <div class="forecast-months">
        ${base.monthly.map((m) => `
          <div class="forecast-row">
            <div><strong>${m.label}</strong><span>+${num(m.newArrivals)} new &middot; ${num(m.oldSold)} old sold &middot; ${num(m.newSold)} new sold</span></div>
            <div class="forecast-metric"><strong>${num(m.units)}</strong><span>sales</span></div>
            <div class="forecast-metric"><strong>${euro(m.revenue, 0)}</strong><span>revenue</span></div>
            <div class="forecast-metric"><strong>${euro(m.profit, 0)}</strong><span>profit</span></div>
          </div>`).join("")}
      </div>

      <p class="forecast-section-label">Stock Evolution</p>
      <div class="forecast-waterfall">
        ${w.map((m) => `
          <div class="waterfall-bar-row">
            <span class="waterfall-label">${m.label}</span>
            <div class="waterfall-track">
              <div class="waterfall-fill stock" style="width:${Math.min(100, (m.remainingStock / Math.max(d.totalInventory, 1)) * 100)}%">${num(m.remainingStock)}</div>
            </div>
            <span class="waterfall-value">+${num(m.newArrivals)} &minus;${num(m.sold)}</span>
          </div>`).join("")}
      </div>

      <p class="forecast-section-label">Purchase Pipeline</p>
      <div class="forecast-drivers">
        <div class="driver-card"><span class="driver-label">Purchase Rate</span><span class="driver-value">${num(d.purchaseRate, 1)} /mo</span><span class="driver-note">recent 3-month average</span></div>
        <div class="driver-card"><span class="driver-label">Purchase Quality</span><span class="driver-value">${pct(d.purchaseQuality * 100)}</span><span class="driver-note">buy brands that sell? ${d.purchaseQuality >= 0.35 ? "Good" : d.purchaseQuality >= 0.25 ? "Fair" : "Weak"}</span></div>
        <div class="driver-card"><span class="driver-label">Avg Purchase Cost</span><span class="driver-value">${euro(d.avgPurchaseCost, 0)}</span><span class="driver-note">per pair (recent buys)</span></div>
        <div class="driver-card"><span class="driver-label">Buying Mix</span><span class="driver-value">${d.purchaseBrandMix[0] ? d.purchaseBrandMix[0].brand : "N/A"}</span><span class="driver-note">${d.purchaseBrandMix.map((b) => `${b.brand}: ${b.count}`).join(" &middot; ")}</span></div>
      </div>

      <p class="forecast-section-label">Model Demand Ranking</p>
      <div class="forecast-drivers">
        ${d.topModels.map((m) => `
          <div class="driver-card"><span class="driver-label">${m.model}</span><span class="driver-value">${euro(m.expectedRevenue, 0)}</span><span class="driver-note">${num(m.count)} in stock &middot; ${pct(m.avgProb * 100)} sell prob</span></div>
        `).join("")}
      </div>

      <p class="forecast-section-label">Demand & Risk Drivers</p>
      <div class="forecast-drivers">
        <div class="driver-card"><span class="driver-label">Sell Probability</span><span class="driver-value">${pct(d.avgSellProbability * 100)}</span><span class="driver-note">${num(d.highProbItems)} high (&gt;50%) &middot; ${num(d.lowProbItems)} low (&lt;20%)</span></div>
        <div class="driver-card"><span class="driver-label">Momentum</span><span class="driver-value">${ratio(d.momentum)}x</span><span class="driver-note">${d.momentum >= 1 ? "Accelerating" : "Decelerating"} demand</span></div>
        <div class="driver-card"><span class="driver-label">Stock Freshness</span><span class="driver-value">${num(d.freshStock)} fresh</span><span class="driver-note">${num(d.deadStock)} dead (&gt;180d) &middot; avg ${num(d.avgAge)}d</span></div>
        <div class="driver-card"><span class="driver-label">Price Realization</span><span class="driver-value">${pct(d.priceRealization * 100)}</span><span class="driver-note">actual vs target price</span></div>
        <div class="driver-card"><span class="driver-label">Best Sizes</span><span class="driver-value">${d.topSizes[0] ? d.topSizes[0].size : "-"}</span><span class="driver-note">${d.topSizes.map((s) => `${s.size}: ${pct(s.avgProb * 100)}`).join(" &middot; ")}</span></div>
        <div class="driver-card"><span class="driver-label">Slowest Sizes</span><span class="driver-value">${d.worstSizes[0] ? d.worstSizes[0].size : "-"}</span><span class="driver-note">${d.worstSizes.map((s) => `${s.size}: ${pct(s.avgProb * 100)}`).join(" &middot; ")}</span></div>
      </div>

      <div class="forecast-method" style="margin-top:12px">${forecast.method}</div>
    `;
  }

  function renderTrendChart(rows) {
    if (!els.trendChart) return;
    if (!rows.length) { els.trendChart.innerHTML = `<div class="empty-state">No sales data in the selected time range.</div>`; return; }
    const monthly = new Map();
    rows.forEach((r) => {
      const d = date(r.sale_date);
      if (!d) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const cur = monthly.get(key) || { revenue: 0, sales: 0 };
      cur.revenue += r.sale_price_net || 0;
      cur.sales += r.sold_units || 0;
      monthly.set(key, cur);
    });
    const items = [...monthly.entries()].map(([label, v]) => ({ label, revenue: v.revenue, sales: v.sales })).sort((a, b) => a.label.localeCompare(b.label)).slice(-12);
    const w = 760, h = 320, leftPad = 64, rightPad = 44, topPad = 34, bottomPad = 34;
    const maxRevenue = Math.max(...items.map((i) => i.revenue), 1);
    const maxSales = Math.max(...items.map((i) => i.sales), 1);
    const makeTicks = (max, steps = 4) => {
      const rough = max / steps;
      const mag = 10 ** Math.floor(Math.log10(Math.max(rough, 1)));
      const norm = rough / mag;
      let step = mag;
      if (norm > 5) step = 10 * mag; else if (norm > 2) step = 5 * mag; else if (norm > 1) step = 2 * mag;
      const niceMax = Math.ceil(max / step) * step;
      return Array.from({ length: steps + 1 }, (_, i) => i * step).filter((t) => t <= niceMax);
    };
    const revTicks = makeTicks(maxRevenue, 4);
    const saleTicks = makeTicks(maxSales, 4);
    const revMax = revTicks[revTicks.length - 1] || maxRevenue;
    const saleMax = saleTicks[saleTicks.length - 1] || maxSales;
    const fmtMonth = (v) => { const [y, m] = v.split("-"); return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("en-GB", { month: "short", year: "2-digit" }); };
    const step = items.length > 1 ? (w - leftPad - rightPad) / (items.length - 1) : 0;
    const pts = items.map((i, x) => ({ ...i, x: leftPad + step * x, ry: h - bottomPad - ((i.revenue / revMax) * (h - topPad - bottomPad)), sy: h - bottomPad - ((i.sales / saleMax) * (h - topPad - bottomPad)) }));
    const revLine = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.ry}`).join(" ");
    const saleLine = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.sy}`).join(" ");
    const area = `${revLine} L ${pts[pts.length - 1].x} ${h - bottomPad} L ${pts[0].x} ${h - bottomPad} Z`;
    const gridTicks = revTicks.slice(1);
    const yRev = (t) => h - bottomPad - ((t / revMax) * (h - topPad - bottomPad));
    const ySale = (t) => h - bottomPad - ((t / saleMax) * (h - topPad - bottomPad));

    els.trendChart.innerHTML = `<div class="chart-legend"><span><i class="legend-swatch revenue"></i>Net Revenue</span><span><i class="legend-swatch sales"></i>Sales Count</span></div><div class="chart-hover-card" hidden></div><svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" aria-label="Trend"><line class="chart-hover-line" x1="${leftPad}" x2="${leftPad}" y1="${topPad}" y2="${h - bottomPad}" opacity="0"></line><text class="chart-axis-title" x="${leftPad - 28}" y="${topPad - 8}">Revenue</text><text class="chart-axis-title" x="${w - rightPad + 22}" y="${topPad - 8}" text-anchor="end">Sales</text>${gridTicks.map((t) => `<line class="chart-grid" x1="${leftPad}" x2="${w - rightPad}" y1="${yRev(t)}" y2="${yRev(t)}"></line>`).join("")}${revTicks.map((t) => `<text class="chart-axis-value" x="${leftPad - 10}" y="${yRev(t) + 4}" text-anchor="end">${t === 0 ? "0 €" : compactEuro(t)}</text>`).join("")}${saleTicks.map((t) => `<text class="chart-axis-value" x="${w - rightPad + 10}" y="${ySale(t) + 4}" text-anchor="start">${num(t)}</text>`).join("")}<path class="chart-area-fill" d="${area}"></path><path class="chart-line revenue-line" d="${revLine}"></path><path class="chart-line sales-line" d="${saleLine}"></path>${pts.map((p) => `<circle class="chart-dot revenue-dot" data-label="${fmtMonth(p.label)}" data-revenue="${euro(p.revenue, 0)}" data-sales="${num(p.sales)}" cx="${p.x}" cy="${p.ry}" r="5"></circle>`).join("")}${pts.map((p) => `<circle class="chart-dot sales-dot" data-label="${fmtMonth(p.label)}" data-revenue="${euro(p.revenue, 0)}" data-sales="${num(p.sales)}" cx="${p.x}" cy="${p.sy}" r="4"></circle>`).join("")}${pts.map((p) => `<text class="chart-label" x="${p.x}" y="${h - 10}" text-anchor="middle">${fmtMonth(p.label)}</text>`).join("")}</svg>`;
    els.trendMeta.textContent = `${items.length} months`;

    const hoverLine = els.trendChart.querySelector(".chart-hover-line");
    const hoverCard = els.trendChart.querySelector(".chart-hover-card");
    const posHover = (e) => { if (!hoverCard) return; const r = els.trendChart.getBoundingClientRect(); hoverCard.style.left = `${Math.min(e.clientX - r.left + 18, r.width - 170)}px`; hoverCard.style.top = `${Math.max(18, e.clientY - r.top - 18)}px`; };
    els.trendChart.querySelectorAll(".chart-dot").forEach((dot) => {
      dot.addEventListener("mouseenter", (e) => { dot.classList.add("is-active"); if (hoverLine) { hoverLine.setAttribute("x1", dot.getAttribute("cx")); hoverLine.setAttribute("x2", dot.getAttribute("cx")); hoverLine.setAttribute("opacity", "1"); } if (hoverCard) { hoverCard.hidden = false; hoverCard.innerHTML = `<strong>${dot.dataset.label}</strong><span>${dot.dataset.revenue} revenue</span><span>${dot.dataset.sales} sales</span>`; posHover(e); } });
      dot.addEventListener("mousemove", posHover);
      dot.addEventListener("mouseleave", () => { dot.classList.remove("is-active"); if (hoverLine) hoverLine.setAttribute("opacity", "0"); if (hoverCard) hoverCard.hidden = true; });
    });
    els.trendChart.querySelector("svg").addEventListener("mouseleave", () => { if (hoverLine) hoverLine.setAttribute("opacity", "0"); if (hoverCard) hoverCard.hidden = true; });
  }

  function renderFilterChips() {
    const active = FILTER_META.filter(({ el }) => el.value !== "all");
    if (!active.length) {
      els.activeFilterChips.innerHTML = `<span class="filter-chip">No active filters</span>`;
      els.summaryNote.textContent = "No filters active";
      return;
    }
    els.summaryNote.textContent = `${active.length} filters active`;
    els.activeFilterChips.innerHTML = active.map(({ key, label, el }) => `<span class="filter-chip">${label}: ${el.options[el.selectedIndex].text}<button type="button" data-clear-filter="${key}" aria-label="Clear ${label}">x</button></span>`).join("");
    els.activeFilterChips.querySelectorAll("[data-clear-filter]").forEach((btn) => btn.addEventListener("click", () => {
      const target = FILTER_META.find((item) => item.key === btn.dataset.clearFilter);
      if (target) { target.el.value = "all"; render(); }
    }));
  }

  function renderPriorities(rows, opps, risks, inv, leaks) {
    if (!els.priorityList) return;
    const sold = rows.filter((r) => r.sold_units);
    const missingTarget = inv.filter((r) => !r.target_sale_price_net).length;
    const aged = inv.filter((r) => (r.days_in_stock || 0) > 180).length;
    const capitalAtRisk = sum(inv.filter((r) => (r.days_in_stock || 0) > 120), (r) => r.inventory_value_cost);
    const lowMarginSales = sold.filter((r) => (r.margin_pct || 0) < 20).length;
    const items = [
      { title: "Protect Growth", copy: `${opps.length} models show momentum or restock pressure.`, tone: opps.length ? "good" : "warn", value: opps.length },
      { title: "Release Capital", copy: `${euro(capitalAtRisk, 0)} tied in inventory >120 days. ${aged} dead stock.`, tone: capitalAtRisk ? "bad" : "good", value: aged || 0 },
      { title: "Defend Margin", copy: `${lowMarginSales} sales below 20% margin. ${leaks.length} in leak queue.`, tone: lowMarginSales ? "warn" : "good", value: lowMarginSales || 0 },
      { title: "Restore Control", copy: `${missingTarget} without target price. ${risks.length} with risk.`, tone: missingTarget || risks.length ? "warn" : "good", value: missingTarget + risks.length }
    ];
    els.priorityList.innerHTML = items.map((i) => `<div class="priority-item"><div class="priority-copy"><strong>${i.title}</strong><span>${i.copy}</span></div>${badge(num(i.value), i.tone)}</div>`).join("");
  }

  function deriveMetricsFromRows(rows) {
    const sold = rows.filter((r) => r.sold_units);
    const inventory = rows.filter((r) => r.inventory_units);
    const sold2026 = sold.filter((r) => { const d = date(r.sale_date); return d && d.getFullYear() === 2026; });
    const bought2026 = rows.filter((r) => { const d = date(r.purchase_date); return d && d.getFullYear() === 2026; });

    const overallRevenue = sum(sold, (r) => r.sale_price_net);
    const overallProfit = sum(sold, (r) => r.gross_profit);
    const overallSoldPurchase = sum(sold, (r) => r.purchase_price_net);
    const inventoryCost = sum(inventory, (r) => r.inventory_value_cost);
    const deadStock = inventory.filter((r) => (r.days_in_stock || 0) > 180);
    const stockAtRisk = inventory.filter((r) => (r.days_in_stock || 0) > 120);
    const sellThrough = (sold.length / ((sold.length + inventory.length) || 1)) * 100;
    const yearRevenue = sum(sold2026, (r) => r.sale_price_net);
    const yearProfit = sum(sold2026, (r) => r.gross_profit);
    const yearPurchase = sum(bought2026, (r) => r.purchase_price_net);

    return {
      overall: {
        sold_pairs: sold.length, inventory_pairs: inventory.length,
        net_revenue: overallRevenue, profit_net: overallProfit,
        margin_pct: overallRevenue ? (overallProfit / overallRevenue) * 100 : 0,
        revenue_multiple: overallSoldPurchase ? overallRevenue / overallSoldPurchase : 0,
        inventory_cost: inventoryCost,
        dead_stock_count: deadStock.length, dead_stock_cost: sum(deadStock, (r) => r.inventory_value_cost),
        capital_at_risk_cost: sum(stockAtRisk, (r) => r.inventory_value_cost),
        sell_through_pct: sellThrough, avg_inventory_age: avg(inventory, (r) => r.days_in_stock),
        inventory_without_target: inventory.filter((r) => !r.target_sale_price_net).length,
        low_margin_sales: sold.filter((r) => (r.margin_pct || 0) < 20).length,
        average_purchase_price_inventory: avg(inventory, (r) => r.purchase_price_net),
        average_target_price_inventory: avg(inventory, (r) => r.target_sale_price_net),
        inventory_value_target: sum(inventory, (r) => r.target_sale_price_net),
        average_sale_price_net: avg(sold, (r) => r.sale_price_net),
        average_profit_per_pair: avg(sold, (r) => r.gross_profit),
        active_brands: uniq(rows.map((r) => r.brand)).length
      },
      y2026: {
        sold_pairs: sold2026.length, bought_pairs: bought2026.length,
        net_revenue: yearRevenue, profit_net: yearProfit, purchase_net: yearPurchase,
        profit_per_invested_euro: yearPurchase ? yearProfit / yearPurchase : 0,
        capital_turnover: yearPurchase ? yearRevenue / yearPurchase : 0,
        average_sale_price_net: avg(sold2026, (r) => r.sale_price_net),
        average_profit_per_pair: avg(sold2026, (r) => r.gross_profit),
        margin_pct: yearRevenue ? (yearProfit / yearRevenue) * 100 : 0
      }
    };
  }

  function renderSummaryCards(rows) {
    const metrics = deriveMetricsFromRows(rows);
    const o = metrics.overall || {};
    const y = metrics.y2026 || {};
    const renderCards = (container, cards, className = "kpi-card") => {
      if (!container) return;
      container.innerHTML = cards.map((c) => `<article class="${className} ${c.tone || ""}"><span class="${className === "focus-card" ? "focus-label" : "kpi-label"}">${c.label}</span><strong class="${className === "focus-card" ? "focus-value" : "kpi-value"}">${c.value}</strong><span class="${className === "focus-card" ? "focus-meta" : "kpi-delta"}">${c.meta || c.delta || ""}</span></article>`).join("");
    };
    renderCards(els.overallKpiGrid, [
      { label: "Net Revenue", value: euro(o.net_revenue, 0), meta: `${num(o.sold_pairs)} sold pairs`, tone: "highlight" },
      { label: "Net Profit", value: euro(o.profit_net, 0), meta: `${euro(o.average_profit_per_pair, 0)} /pair`, tone: "highlight" },
      { label: "Margin", value: pct(o.margin_pct), meta: `${num(o.low_margin_sales)} below 20%`, tone: o.margin_pct >= 30 ? "good" : o.margin_pct < 20 ? "bad" : "warn" },
      { label: "Inventory Cost", value: euro(o.inventory_cost, 0), meta: `${num(o.inventory_pairs)} pairs`, tone: "warn" },
      { label: "Capital >120d", value: euro(o.capital_at_risk_cost, 0), meta: `${num(o.dead_stock_count)} dead stock`, tone: o.capital_at_risk_cost ? "bad" : "good" },
      { label: "Sell-through", value: pct(o.sell_through_pct), meta: "sold vs. available", tone: o.sell_through_pct >= 40 ? "good" : o.sell_through_pct < 25 ? "bad" : "warn" },
      { label: "Avg Inv. Age", value: `${num(o.avg_inventory_age)} days`, meta: "active inventory", tone: o.avg_inventory_age > 120 ? "bad" : o.avg_inventory_age > 75 ? "warn" : "good" },
      { label: "No Target Price", value: num(o.inventory_without_target), meta: `${num(o.active_brands)} brands`, tone: o.inventory_without_target ? "warn" : "good" },
      { label: "Avg Cost In Stock", value: euro(o.average_purchase_price_inventory, 0), delta: "inventory only" },
      { label: "Avg Target Price", value: euro(o.average_target_price_inventory, 0), delta: "inventory only" },
      { label: "Inv. Value (Target)", value: euro(o.inventory_value_target, 0), delta: "based on target price" }
    ], "kpi-card");
    renderCards(els.yearKpiGrid, [
      { label: "Sold Pairs", value: num(y.sold_pairs), meta: "2026", tone: "highlight" },
      { label: "Bought Pairs", value: num(y.bought_pairs), meta: "2026", tone: "highlight" },
      { label: "Net Revenue", value: euro(y.net_revenue, 0), meta: `${euro(y.average_sale_price_net, 0)} avg price`, tone: "good" },
      { label: "Net Profit", value: euro(y.profit_net, 0), meta: `${euro(y.average_profit_per_pair, 0)} /pair`, tone: y.profit_net > 0 ? "good" : "warn" },
      { label: "Margin", value: pct(y.margin_pct), meta: "profit / revenue", tone: y.margin_pct >= 30 ? "good" : y.margin_pct < 20 ? "bad" : "warn" },
      { label: "ROI per Euro", value: ratio(y.profit_per_invested_euro), meta: `${ratio(y.capital_turnover)} turnover`, tone: y.profit_per_invested_euro >= 0.4 ? "good" : "warn" },
    ], "focus-card");
  }

  function render() {
    const rows = filteredRecords();
    const sold = rows.filter((r) => r.sold_units);
    const inv = rows.filter((r) => r.inventory_units);
    const opps = findOpportunities(rows);
    const risks = findRisks(rows);
    const actions = inventoryActions(rows);
    const leaks = marginLeaks(rows);

    renderSummaryCards(rows);
    renderFilterChips();
    renderPriorities(rows, opps, risks, inv, leaks);
    renderBusinessHealth(rows);
    renderForecast(rows);
    renderTrendChart(sold);
    renderAgeBuckets(rows);
    renderMarginDistribution(rows);
    renderPriceRealization(rows);
    renderCountryPerformance(rows);
    renderSizeDistribution(rows);
    renderPurchaseSources(rows);
    renderInventoryTurnover(rows);
    renderRankList(els.topModels, modelGroups(sold).map((g) => {
      const revenue = sum(g.rows, (r) => r.sale_price_net);
      const profit = sum(g.rows, (r) => r.gross_profit);
      const soldUnits = g.rows.filter((r) => r.sold_units).length;
      return { label: displayModel(g.model || "Unknown"), value: revenue, tooltip: `${titleCase(g.brand)}<br>${soldUnits} sold<br>${euro(profit, 0)} profit` };
    }).sort((a, b) => b.value - a.value).slice(0, 8), (v) => euro(v, 0));
    renderBrandPerformance(rows);
    renderChannelPerformance(rows);
    renderTable(els.opportunities, opps, [["Model", "model"], ["Brand", "brand"], ["Sales", "sold_units"], ["Stock", "inventory_units"], ["Margin", "margin_badge"], ["Signal", "signal_badge"]]);
    renderTable(els.risks, risks, [["Model", "model"], ["Brand", "brand"], ["Stock", "inventory_units"], ["Days", "days_in_stock"], ["Cost Value", "inventory_value_cost"], ["Signal", "signal_badge"]]);
    renderTable(els.inventoryActions, actions, [["Item", "article"], ["Brand", "brand"], ["Status", "status_badge"], ["Cost", "purchase_price_net"], ["Target", "target_sale_price_net"], ["Days", "days_in_stock"], ["Action", "action_badge"]]);
    renderTable(els.marginLeaks, leaks, [["Item", "article"], ["Brand", "brand"], ["Revenue", "sale_price_net"], ["Profit", "profit_net"], ["Margin", "margin_badge"], ["Signal", "signal_badge"]]);

    els.workbenchCounts.textContent = `${num(rows.length)} records`;
    els.opportunitiesCount.textContent = `${num(opps.length)} items`;
    els.risksCount.textContent = `${num(risks.length)} items`;
    els.inventoryCount.textContent = `${num(actions.length)} items`;
    els.marginLeaksCount.textContent = `${num(leaks.length)} items`;
    els.generatedAt.textContent = new Date(state.summary.generated_at || Date.now()).toLocaleString("de-DE");
    els.dataSourceLabel.textContent = state.dataSource === "google-sheet"
      ? (state.summary.live_mode === "csv" ? "Google Sheets (CSV)" : "Google Sheets (Live)")
      : (state.summary.live_error ? `Local (${state.summary.live_error})` : "Local Snapshot");
  }

  loadData().then((payload) => {
    state = payload;
    setupFilters(state.records);
    setActiveTable(state.activeTable);
    render();
  }).catch((error) => {
    const msg = `<div class="empty-state">Data could not be loaded: ${error.message}</div>`;
    if (els.overallKpiGrid) els.overallKpiGrid.innerHTML = msg;
    if (els.yearKpiGrid) els.yearKpiGrid.innerHTML = msg;
    if (els.priorityList) els.priorityList.innerHTML = msg;
  });
}());
