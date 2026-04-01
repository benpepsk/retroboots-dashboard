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
    flipCritical: $("flip-critical"),
    opportunitiesCount: $("opportunities-count"),
    risksCount: $("risks-count"),
    flipCriticalCount: $("flip-critical-count"),
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
    const ekNet = nOrNull(pickField(row, ["einkaufspreis_netto", "purchase_price_net", "unit_cost", "cost_net"]));
    const ekGross = nOrNull(pickField(row, ["einkaufspreis_brutto", "purchase_price_gross", "unit_cost_gross"]));
    const vkNet = nOrNull(pickField(row, ["verkaufspreis_netto", "sale_price_net", "selling_price_net", "revenue_net"]));
    const vkGross = nOrNull(pickField(row, ["verkaufspreis_brutto", "sale_price_gross", "selling_price_gross"]));
    const shippingIn = nOrNull(pickField(row, ["versandkosten_einkauf", "shipping_cost_purchase", "shipping_in"]));
    const shippingOut = nOrNull(pickField(row, ["versandkosten_verkauf", "shipping_cost_sale", "shipping_out"]));
    const status = normalizeStatusValue(pickField(row, ["status", "inventory_status", "state"]));
    const sheetProfit = nOrNull(pickField(row, ["netto_gewinn", "net_profit"]));
    const sheetPotProfit = nOrNull(pickField(row, ["pot_gewinn", "potential_profit"]));
    // Use sheet's netto_gewinn when available (includes shipping), otherwise compute
    const ek = ekNet;
    const vk = vkNet;
    const computedProfit = ek != null && vk != null ? vk - ek - (shippingIn || 0) - (shippingOut || 0) : null;
    const gp = sheetProfit != null ? sheetProfit : computedProfit;
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
      purchase_price_gross: ekGross,
      purchase_price_net: ek,
      shipping_cost_purchase: shippingIn,
      purchase_type: norm(pickField(row, ["kaufart", "purchase_type"])),
      target_sale_price_gross: nOrNull(pickField(row, ["ziel_vk_brutto", "target_sale_price_gross"])),
      target_sale_price_net: nOrNull(pickField(row, ["ziel_vk_netto", "target_sale_price_net", "target_net"])),
      sale_price_gross: vkGross,
      sale_price_net: vk,
      shipping_cost_sale: shippingOut,
      sale_date: parseFlexibleDate(pickField(row, ["verkaufsdatum", "sale_date", "sold_at"])),
      purchase_platform: norm(pickField(row, ["plattform_einkauf", "purchase_platform", "buy_channel"])),
      sales_channel: norm(pickField(row, ["plattform_verkauf", "sales_channel", "channel"])),
      sales_country: norm(pickField(row, ["land_verkauf", "sales_country", "country"])),
      potential_profit: sheetPotProfit,
      gross_profit: gp,
      margin_pct: gp != null && vk ? (gp / vk) * 100 : null,
      inventory_units: status === "lager" ? 1 : 0,
      sold_units: status === "verkauft" ? 1 : 0,
      inventory_value_cost: status === "lager" ? ek || 0 : 0,
      total_shipping: (shippingIn || 0) + (shippingOut || 0),
      days_in_stock: ((Date.now() - (date(parseFlexibleDate(pickField(row, ["kaufdatum", "purchase_date", "buy_date"]))) || new Date()).getTime()) / 86400000) | 0,
      strategie: (() => { const v = normalizeHeaderKey(norm(pickField(row, ["strategie", "strategy"])) || ""); return v === "hold" ? "hold" : v === "flip" ? "flip" : null; })(),
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
      const flipInv = inv.filter((r) => r.strategie !== "hold");
      const flipAge = avg(flipInv, (r) => r.days_in_stock);
      let signal = "", tone = "warn";
      if (stock >= 2 && sold === 0 && flipInv.length > 0) { signal = "Slow mover"; tone = "bad"; }
      else if (flipAge >= 180 && flipInv.length > 0) { signal = "Dead stock"; tone = "bad"; }
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
      if (r.strategie === "hold") { action = "Long-term hold"; tone = "good"; }
      else if ((r.days_in_stock || 0) > 180) { action = "De-risk now"; tone = "bad"; }
      else if (!r.target_sale_price_net) { action = "Set target price"; tone = "warn"; }
      else if ((r.days_in_stock || 0) > 120) { action = "Review pricing"; tone = "warn"; }
      else if (r.target_sale_price_net && r.purchase_price_net && ((r.target_sale_price_net - r.purchase_price_net) / r.target_sale_price_net) * 100 >= 30) { action = "Hold price"; tone = "good"; }
      const stratLabel = r.strategie === "hold" ? "Hold" : r.strategie === "flip" ? "Flip" : "-";
      const stratTone = r.strategie === "hold" ? "good" : r.strategie === "flip" ? "warn" : "warn";
      return {
        article: cell(displayProductName(r.brand, r.model, `${titleCase(r.brand || "")} #${r.article_id}`.trim(), false), `${r.size || "-"} | ${r.surface || "-"}`),
        brand: r.brand ? titleCase(r.brand) : "-",
        strategie_badge: badge(stratLabel, stratTone),
        purchase_price_net: euro(r.purchase_price_net, 0),
        target_sale_price_net: r.target_sale_price_net ? euro(r.target_sale_price_net, 0) : "-",
        days_in_stock: num(r.days_in_stock),
        action_badge: badge(action, tone)
      };
    }).sort((a, b) => Number(b.days_in_stock.replace(/\./g, "")) - Number(a.days_in_stock.replace(/\./g, ""))).slice(0, 14);
  }

  function flipCriticalCases(rows) {
    return rows
      .filter((r) => r.inventory_units && r.strategie === "flip" && (r.days_in_stock || 0) >= 90)
      .map((r) => {
        const days = r.days_in_stock || 0;
        const potProfit = r.target_sale_price_net != null && r.purchase_price_net != null
          ? r.target_sale_price_net - r.purchase_price_net : null;
        let urgency, tone;
        if (days > 180) { urgency = "Sofort verkaufen"; tone = "bad"; }
        else if (days > 120) { urgency = "Preischeck fällig"; tone = "warn"; }
        else { urgency = "Im Blick behalten"; tone = "warn"; }
        return {
          article: cell(
            displayProductName(r.brand, r.model, `${titleCase(r.brand || "")} #${r.article_id}`.trim(), true),
            [r.generation ? `Gen. ${r.generation}` : null, r.color ? titleCase(r.color) : null, r.surface ? r.surface.toUpperCase() : null, r.size ? `Gr. ${r.size}` : null].filter(Boolean).join(" · ")
          ),
          days_in_stock: num(days),
          purchase_price_net: euro(r.purchase_price_net, 0),
          target_sale_price_net: r.target_sale_price_net ? euro(r.target_sale_price_net, 0) : "-",
          pot_profit: potProfit != null ? euro(potProfit, 0) : "-",
          urgency_badge: badge(urgency, tone),
        };
      })
      .sort((a, b) => Number(b.days_in_stock.replace(/\./g, "")) - Number(a.days_in_stock.replace(/\./g, "")));
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
    const flipInv = inv.filter((r) => r.strategie !== "hold");
    const dead = flipInv.filter((r) => (r.days_in_stock || 0) > 180);
    const lowMargin = sold.filter((r) => (r.margin_pct || 0) < 20);
    const capitalAtRisk = flipInv.filter((r) => (r.days_in_stock || 0) > 120);
    const holdCount = inv.filter((r) => r.strategie === "hold").length;
    const items = [
      { label: "Sell-through", value: pct((sold.length / ((sold.length + inv.length) || 1)) * 100), note: "sold vs. total" },
      { label: "Flip >180d", value: num(dead.length), note: `${euro(sum(dead, (r) => r.inventory_value_cost), 0)} tied up` },
      { label: "Profit Leaks", value: num(lowMargin.length), note: "sales below 20% margin" },
      { label: "Avg Inventory Age", value: `${num(avg(inv, (r) => r.days_in_stock))} days`, note: "active inventory" },
      { label: "No Target Price", value: num(inv.filter((r) => !r.target_sale_price_net).length), note: "pricing risk" },
      { label: "Flip Capital >120d", value: euro(sum(capitalAtRisk, (r) => r.inventory_value_cost), 0), note: "flip – losing mobility" },
      { label: "Hold Inventory", value: num(holdCount), note: "bewusst gehalten" },
      { label: "Shipping Load", value: euro(sum(rows, (r) => r.total_shipping), 0), note: `${euro(avg(rows.filter((r) => r.total_shipping > 0), (r) => r.total_shipping), 0)} avg per item` }
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

  function renderPriorities(rows, opps, risks, inv) {
    if (!els.priorityList) return;
    const sold = rows.filter((r) => r.sold_units);
    const missingTarget = inv.filter((r) => !r.target_sale_price_net).length;
    const flipInv = inv.filter((r) => r.strategie !== "hold");
    const aged = flipInv.filter((r) => (r.days_in_stock || 0) > 180).length;
    const capitalAtRisk = sum(flipInv.filter((r) => (r.days_in_stock || 0) > 120), (r) => r.inventory_value_cost);
    const lowMarginSales = sold.filter((r) => (r.margin_pct || 0) < 20).length;
    const items = [
      { title: "Protect Growth", copy: `${opps.length} models show momentum or restock pressure.`, tone: opps.length ? "good" : "warn", value: opps.length },
      { title: "Release Capital", copy: `${euro(capitalAtRisk, 0)} tied in inventory >120 days. ${aged} flip >180d.`, tone: capitalAtRisk ? "bad" : "good", value: aged || 0 },
      { title: "Defend Margin", copy: `${lowMarginSales} sales below 20% margin.`, tone: lowMarginSales ? "warn" : "good", value: lowMarginSales || 0 },
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
    const flipInventory = inventory.filter((r) => r.strategie !== "hold");
    const deadStock = flipInventory.filter((r) => (r.days_in_stock || 0) > 180);
    const stockAtRisk = flipInventory.filter((r) => (r.days_in_stock || 0) > 120);
    const sellThrough = (sold.length / ((sold.length + inventory.length) || 1)) * 100;
    const yearRevenue = sum(sold2026, (r) => r.sale_price_net);
    const yearProfit = sum(sold2026, (r) => r.gross_profit);
    const yearPurchase = sum(bought2026, (r) => r.purchase_price_net);

    const totalShippingIn = sum(rows, (r) => r.shipping_cost_purchase);
    const totalShippingOut = sum(sold, (r) => r.shipping_cost_sale);
    const totalShipping = totalShippingIn + totalShippingOut;
    const privatePurchases = rows.filter((r) => r.purchase_type === "privat").length;
    const gewerblichPurchases = rows.filter((r) => r.purchase_type === "gewerblich").length;

    const yearShippingIn = sum(bought2026, (r) => r.shipping_cost_purchase);
    const yearShippingOut = sum(sold2026, (r) => r.shipping_cost_sale);

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
        active_brands: uniq(rows.map((r) => r.brand)).length,
        total_shipping: totalShipping, shipping_in: totalShippingIn, shipping_out: totalShippingOut,
        privat_count: privatePurchases, gewerblich_count: gewerblichPurchases
      },
      y2026: {
        sold_pairs: sold2026.length, bought_pairs: bought2026.length,
        net_revenue: yearRevenue, profit_net: yearProfit, purchase_net: yearPurchase,
        profit_per_invested_euro: yearPurchase ? yearProfit / yearPurchase : 0,
        capital_turnover: yearPurchase ? yearRevenue / yearPurchase : 0,
        average_sale_price_net: avg(sold2026, (r) => r.sale_price_net),
        average_profit_per_pair: avg(sold2026, (r) => r.gross_profit),
        margin_pct: yearRevenue ? (yearProfit / yearRevenue) * 100 : 0,
        shipping_total: yearShippingIn + yearShippingOut
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
      { label: "Flip Capital >120d", value: euro(o.capital_at_risk_cost, 0), meta: `${num(o.dead_stock_count)} flip >180d`, tone: o.capital_at_risk_cost ? "bad" : "good" },
      { label: "Sell-through", value: pct(o.sell_through_pct), meta: "sold vs. available", tone: o.sell_through_pct >= 40 ? "good" : o.sell_through_pct < 25 ? "bad" : "warn" },
      { label: "Avg Inv. Age", value: `${num(o.avg_inventory_age)} days`, meta: "active inventory", tone: o.avg_inventory_age > 120 ? "bad" : o.avg_inventory_age > 75 ? "warn" : "good" },
      { label: "No Target Price", value: num(o.inventory_without_target), meta: `${num(o.active_brands)} brands`, tone: o.inventory_without_target ? "warn" : "good" },
      { label: "Avg Cost In Stock", value: euro(o.average_purchase_price_inventory, 0), delta: "inventory only" },
      { label: "Avg Target Price", value: euro(o.average_target_price_inventory, 0), delta: "inventory only" },
      { label: "Inv. Value (Target)", value: euro(o.inventory_value_target, 0), delta: "based on target price" },
      { label: "Shipping Costs", value: euro(o.total_shipping, 0), delta: `${euro(o.shipping_in, 0)} in · ${euro(o.shipping_out, 0)} out`, tone: o.total_shipping > 0 ? "warn" : "" },
      { label: "Purchase Type", value: `${num(o.gewerblich_count)} B2B`, delta: `${num(o.privat_count)} private` }
    ], "kpi-card");
    renderCards(els.yearKpiGrid, [
      { label: "Sold Pairs", value: num(y.sold_pairs), meta: "2026", tone: "highlight" },
      { label: "Bought Pairs", value: num(y.bought_pairs), meta: "2026", tone: "highlight" },
      { label: "Net Revenue", value: euro(y.net_revenue, 0), meta: `${euro(y.average_sale_price_net, 0)} avg price`, tone: "good" },
      { label: "Net Profit", value: euro(y.profit_net, 0), meta: `${euro(y.average_profit_per_pair, 0)} /pair`, tone: y.profit_net > 0 ? "good" : "warn" },
      { label: "Margin", value: pct(y.margin_pct), meta: "profit / revenue", tone: y.margin_pct >= 30 ? "good" : y.margin_pct < 20 ? "bad" : "warn" },
      { label: "ROI per Euro", value: ratio(y.profit_per_invested_euro), meta: `${ratio(y.capital_turnover)} turnover`, tone: y.profit_per_invested_euro >= 0.4 ? "good" : "warn" },
      { label: "Shipping Costs", value: euro(y.shipping_total, 0), meta: "2026 total", tone: y.shipping_total > 0 ? "warn" : "" }
    ], "focus-card");
  }

  function render() {
    const rows = filteredRecords();
    const sold = rows.filter((r) => r.sold_units);
    const inv = rows.filter((r) => r.inventory_units);
    const opps = findOpportunities(rows);
    const risks = findRisks(rows);
    const flipCritical = flipCriticalCases(rows);

    renderSummaryCards(rows);
    renderFilterChips();
    renderPriorities(rows, opps, risks, inv);
    renderBusinessHealth(rows);
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
    renderTable(els.flipCritical, flipCritical, [["Schuh", "article"], ["Tage", "days_in_stock"], ["EK netto", "purchase_price_net"], ["Ziel VK", "target_sale_price_net"], ["Pot. Gewinn", "pot_profit"], ["Priorität", "urgency_badge"]]);

    els.workbenchCounts.textContent = `${num(rows.length)} records`;
    els.opportunitiesCount.textContent = `${num(opps.length)} items`;
    els.risksCount.textContent = `${num(risks.length)} items`;
    if (els.flipCriticalCount) els.flipCriticalCount.textContent = `${num(flipCritical.length)} Paare`;
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
