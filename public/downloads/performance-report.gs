/**
 * ============================================================
 *  PERFORMANCE REPORT (Search Console + GA4)
 * ============================================================
 *  Menu: SEO Tools > Performance Report (GSC + GA4) (menu created by onOpen() in seo-audit.gs).
 *
 *  Sidebar:
 *  - Period:  Last 7 days, Last week (Mon–Sun), Last 28 days, Last month, This month so far,
 *             Pick a week, Pick a month, or Custom dates. Optional comparison with the previous
 *             period or the same period last year (▲▼ change columns).
 *  - Sources: a Search Console property and/or a GA4 property (lists load from your Google account).
 *  - Pages:   all pages with data, the current cell selection, or pasted URLs.
 *  - Columns: top keywords per page, GA4 traffic source type (channel, source / medium, source),
 *             GA4 page match (landing page or any page view).
 *
 *  One request returns every page (instead of one request per page), so even big sites finish
 *  in one run, well under Apps Script's 6-minute limit.
 *
 *  Per page: Impressions | Clicks | CTR | Avg Position | Top Keywords | Sessions | Engaged Sessions |
 *            Engagement Rate | Top 5 Traffic Sources   (+ Δ columns when comparing)
 *  Above the table: a dashboard with key numbers, plain-language findings and short top lists
 *  (top pages, gainers, losers, quick wins, low CTR, low engagement, top queries, channels).
 *  Lists are on one line, comma-separated (like the audit's in-link columns). Non-HTML URLs are left out.
 *
 *  ONE-TIME SETUP for GA4 (the Search Console part works if Check Indexing (GSC) works):
 *  - Enable "Google Analytics Data API" (and "Google Analytics Admin API" for the property list)
 *    in this script's Google Cloud project.
 *  - appsscript.json must include https://www.googleapis.com/auth/analytics.readonly.
 * ============================================================
 */

const RPT = {
  STATUS_KEY: "RPT_STATUS",
  SETTINGS_KEY: "RPT_SETTINGS",
  NEW_SHEET: "__new__",
  GSC_ROW_LIMIT: 25000,          // Search Console's maximum rows per request
  GSC_MAX_PAGE_ROWS: 100000,     // pages read from Search Console
  GSC_MAX_KEYWORD_ROWS: 100000,  // page + keyword rows read for "all pages" reports
  PER_PAGE_KEYWORDS_UP_TO: 300,  // lists up to this size get exact keywords per page (one request each)
  GA_LIMIT: 100000,              // GA4 rows per request
  GA_MAX_ROWS: 250000,
  PARALLEL: 10,
  MAX_URLS: 5000,
  EMPTY: "-"
};

const RPT_TRAFFIC_DIMS = {
  channel: { dim: "sessionDefaultChannelGroup", label: "Channel" },
  sourceMedium: { dim: "sessionSourceMedium", label: "Source / medium" },
  source: { dim: "sessionSource", label: "Source" }
};

/**
 * Run this one from the Apps Script editor (▷ Run) to grant the Analytics permission and test the setup.
 * Results appear in the Execution log in plain words.
 */
function checkReportSetup() {
  const out = [];
  try { out.push("✅ Search Console: " + rptGscSites().length + " property/ies found"); }
  catch (e) { out.push("❌ Search Console: " + e.message + (e.enableUrl ? " → " + e.enableUrl : "")); }
  let props = [];
  try { props = rptGaProps(); out.push("✅ GA4 property list (Admin API): " + props.length + " property/ies" + (props.length ? ": " + props.slice(0, 5).map(p => p.name + " (" + p.id + ")").join(", ") : "")); }
  catch (e) { out.push("❌ GA4 property list (Admin API): " + e.message + (e.enableUrl ? " → " + e.enableUrl : "")); }
  if (props.length) {
    try { const t = rptGaTotals(props[0].id, rptIsoDaysAgo(30), rptIsoDaysAgo(1)); out.push("✅ GA4 report data (Data API): " + props[0].name + " had " + t.sessions + " sessions in the last 30 days"); }
    catch (e) { out.push("❌ GA4 report data (Data API): " + e.message + (e.enableUrl ? " → " + e.enableUrl : "")); }
  } else {
    out.push("ℹ GA4 report data (Data API): not tested (no property in the list)");
  }
  out.forEach(line => Logger.log(line));
  return out.join("\n");
}

function rptIsoDaysAgo(n) {
  return Utilities.formatDate(new Date(Date.now() - n * 864e5), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function openPerformanceReport() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(RPT_SIDEBAR_HTML).setTitle("Performance Report (GSC + GA4)"));
}

// ============================================================
//  DATA FOR THE SIDEBAR
// ============================================================

/** Properties from both services (each may fail on its own), tabs, the selection and the saved choices. */
function rptGetInfo() {
  const info = {
    gsc: { sites: [], error: "", enableUrl: "" },
    ga: { props: [], error: "", enableUrl: "" },
    tabs: SpreadsheetApp.getActive().getSheets().map(s => s.getName()),
    selection: rptSelectionInfo(),
    saved: rptSaved(),
    guide: RPT_SETUP_GUIDE
  };
  try { info.gsc.sites = rptGscSites(); } catch (e) { info.gsc.error = e.message; info.gsc.enableUrl = e.enableUrl || ""; }
  try { info.ga.props = rptGaProps(); } catch (e) { info.ga.error = e.message; info.ga.enableUrl = e.enableUrl || ""; }
  return info;
}

function rptSelectionInfo() {
  try {
    const r = SpreadsheetApp.getActiveRange();
    if (!r) return { count: 0, a1: "" };
    return { count: rptParseUrls(r.getValues().map(row => row.join("\n")).join("\n")).length, a1: r.getSheet().getName() + "!" + r.getA1Notation() };
  } catch (e) {
    return { count: 0, a1: "" };
  }
}

/** Last choices for this spreadsheet (one sheet per client keeps its own properties). */
function rptSaved() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(RPT.SETTINGS_KEY) || "{}"); } catch (e) { return {}; }
}

function rptStatus() {
  try { return JSON.parse(PropertiesService.getUserProperties().getProperty(RPT.STATUS_KEY) || "{}"); } catch (e) { return {}; }
}

function rptSet(patch, fresh) {
  const s = fresh ? {} : rptStatus();
  Object.keys(patch).forEach(k => { s[k] = patch[k]; });
  PropertiesService.getUserProperties().setProperty(RPT.STATUS_KEY, JSON.stringify(s));
}

/** Search Console properties this Google account can use. */
function rptGscSites() {
  const res = rptCall("https://searchconsole.googleapis.com/webmasters/v3/sites", null, "Google Search Console API");
  return (res.siteEntry || [])
    .filter(s => s.permissionLevel && s.permissionLevel !== "siteUnverifiedUser")
    .map(s => s.siteUrl)
    .sort((a, b) => rptLabel(a).localeCompare(rptLabel(b)));
}

/** GA4 properties this Google account can read (Analytics Admin API). */
function rptGaProps() {
  const out = [];
  let token = "";
  for (let i = 0; i < 10; i++) {
    const res = rptCall("https://analyticsadmin.googleapis.com/v1beta/accountSummaries?pageSize=200" + (token ? "&pageToken=" + encodeURIComponent(token) : ""), null, "Google Analytics Admin API");
    (res.accountSummaries || []).forEach(a => (a.propertySummaries || []).forEach(p => out.push({
      id: String(p.property).replace(/^properties\//, ""), name: p.displayName || "", account: a.displayName || ""
    })));
    token = res.nextPageToken || "";
    if (!token) break;
  }
  return out.sort((a, b) => (a.name + a.account).localeCompare(b.name + b.account));
}

// ============================================================
//  BUILD THE REPORT (called from the sidebar)
// ============================================================

/**
 * c = { start, end, compare, compareStart, compareEnd, compareLabel, gscSite, gaProp, gaName,
 *       source: "all" | "selection" | "paste", urls, maxPages, keywords, traffic, gaMatch, output }
 */
function rptRun(c) {
  try {
    const iso = /^\d{4}-\d{2}-\d{2}$/;
    if (!iso.test(c.start || "") || !iso.test(c.end || "") || c.start > c.end) throw new Error("Choose a valid period (start date before end date).");
    if (c.compare && (!iso.test(c.compareStart || "") || !iso.test(c.compareEnd || ""))) throw new Error("The comparison period isn't valid.");
    const site = String(c.gscSite || "");
    const gaProp = String(c.gaProp || "").replace(/^properties\//, "").trim();
    if (!site && !gaProp) throw new Error("Choose a Search Console property, a GA4 property, or both.");
    if (gaProp && !/^\d+$/.test(gaProp)) throw new Error("A GA4 property ID is a number like 336987406 (GA4 Admin → Property settings).");
    const kwCount = Math.max(0, Math.min(10, Number(c.keywords) || 0));
    const traffic = RPT_TRAFFIC_DIMS[c.traffic] || RPT_TRAFFIC_DIMS.channel;
    const gaDim = c.gaMatch === "pagePath" ? "pagePath" : "landingPage";

    // Your own URLs (selection / paste), or every page that has data.
    let urls = null;
    if (c.source === "selection") {
      const r = SpreadsheetApp.getActiveRange();
      if (!r) throw new Error("Select the cells that contain your URLs first.");
      urls = rptParseUrls(r.getValues().map(row => row.join("\n")).join("\n"));
      if (!urls.length) throw new Error("No URLs found in the selected cells (" + r.getA1Notation() + ").");
    } else if (c.source === "paste") {
      urls = rptParseUrls(c.urls);
      if (!urls.length) throw new Error("Paste at least one URL (one per line).");
    }
    let skippedFiles = 0;
    if (urls) {
      const pages = urls.filter(u => !rptIsFile(u));
      skippedFiles = urls.length - pages.length;
      if (!pages.length) throw new Error("All the URLs are files (PDFs, images…), not web pages: nothing to report.");
      urls = pages.slice(0, RPT.MAX_URLS);
    }

    // Remember the sidebar choices (a scheduled run keeps its own settings and doesn't overwrite these).
    if (!c.fromSchedule) PropertiesService.getDocumentProperties().setProperty(RPT.SETTINGS_KEY, JSON.stringify({
      gscSite: site, gaProp, gaName: c.gaName || "", keywords: kwCount, traffic: c.traffic, gaMatch: gaDim,
      maxPages: c.maxPages, compare: !!c.compare, compareMode: c.compareMode || "previous", preset: c.preset || ""
    }));

    const steps = [];
    if (site) steps.push("Search Console: pages", kwCount ? "Search Console: keywords" : null);
    if (gaProp) steps.push("GA4: pages", "GA4: traffic sources");
    if (c.compare) steps.push("Comparison period");
    steps.push("Writing the report");
    const list = steps.filter(Boolean);
    let stepNo = 0;
    const step = name => rptSet({ phase: "running", step: name, stepNo: ++stepNo, steps: list });
    rptSet({ phase: "running", step: "Starting…", stepNo: 0, steps: list, startedAt: Date.now() }, true);

    const notes = [];
    const data = { gsc: null, gscPrev: null, gscTotals: null, gscTotalsPrev: null, kw: {}, ga: null, gaPrev: null, gaTotals: null, gaTotalsPrev: null, src: {} };

    // ---- Search Console ----
    if (site) {
      step("Search Console: pages");
      data.gsc = rptGscPages(site, c.start, c.end);
      data.gscTotals = rptGscTotals(site, c.start, c.end);
      data.queries = rptGscTopQueries(site, c.start, c.end);
      if (kwCount) {
        step("Search Console: keywords");
        if (urls && urls.length <= RPT.PER_PAGE_KEYWORDS_UP_TO) {
          data.kw = rptGscKeywordsFor(site, urls.map(u => (data.gsc[rptKey(u)] || {}).url || u), c.start, c.end, kwCount);
        } else {
          const bulk = rptGscKeywordsBulk(site, c.start, c.end, kwCount);
          data.kw = bulk.map;
          if (bulk.capped) notes.push("Keywords come from the top " + RPT.GSC_MAX_KEYWORD_ROWS.toLocaleString() + " page–keyword rows (very large site).");
        }
      }
    }

    // ---- GA4 ----
    if (gaProp) {
      step("GA4: pages");
      data.ga = rptGaPages(gaProp, gaDim, c.start, c.end);
      data.gaTotals = rptGaTotals(gaProp, c.start, c.end);
      data.channels = rptGaChannels(gaProp, c.start, c.end);
      step("GA4: traffic sources");
      data.src = rptGaSources(gaProp, gaDim, traffic.dim, c.start, c.end);
    }

    // ---- Comparison ----
    if (c.compare) {
      step("Comparison period");
      if (site) { data.gscPrev = rptGscPages(site, c.compareStart, c.compareEnd); data.gscTotalsPrev = rptGscTotals(site, c.compareStart, c.compareEnd); }
      if (gaProp) {
        data.gaPrev = rptGaPages(gaProp, gaDim, c.compareStart, c.compareEnd);
        data.gaTotalsPrev = rptGaTotals(gaProp, c.compareStart, c.compareEnd);
        data.channelsPrev = rptGaChannels(gaProp, c.compareStart, c.compareEnd);
      }
    }

    // ---- Rows ----
    step("Writing the report");
    const rows = rptBuildRows(urls, data, c, kwCount);
    skippedFiles += rows.skippedFiles || 0;
    if (skippedFiles) notes.push(skippedFiles + " non-HTML URL" + (skippedFiles === 1 ? "" : "s") + " (PDFs, images, files) left out.");
    const domain = site ? rptLabel(site) : (c.gaName || "GA4 " + gaProp);
    const sh = rptResultSheet(c.output, domain, c.start, c.end);
    const ins = rptWrite(sh, rows, data, c, { site, gaProp, domain, kwCount, traffic, gaDim, notes, source: c.source });
    try { SpreadsheetApp.getActive().setActiveSheet(sh); } catch (e) {} // no active sheet when a schedule runs

    const summary = {
      sheetName: sh.getName(), sheetId: sh.getSheetId(), domain, pages: rows.length, notes, steps: list,
      kpis: ins.kpis, findings: ins.findings.slice(0, 8), compare: !!c.compare,
      gsc: site ? rptGscSummary(data.gscTotals) : null, gscPrev: c.compare && site ? rptGscSummary(data.gscTotalsPrev) : null,
      ga: gaProp ? data.gaTotals : null, gaPrev: c.compare && gaProp ? data.gaTotalsPrev : null
    };
    rptSet({ phase: "done", step: "Done", stepNo: list.length, steps: list });
    return summary;
  } catch (e) {
    rptSet({ phase: "error", step: e.message });
    throw e;
  }
}

/** Brings a results tab to the front (Open results link). */
function rptOpenTab(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(String(name || ""));
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

// ============================================================
//  SCHEDULE: weekly / monthly reports built in the background (one schedule per spreadsheet)
// ============================================================

const RPT_SCHEDULE_KEY = "RPT_SCHEDULE";
const RPT_SCHEDULE_HANDLER = "rptScheduledRun";
const RPT_URL_TAB = "Scheduled report URLs";
const RPT_WEEKDAYS = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
const RPT_DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

function rptSchedule() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(RPT_SCHEDULE_KEY) || "null"); } catch (e) { return null; }
}

function rptStoreSchedule(s) {
  PropertiesService.getDocumentProperties().setProperty(RPT_SCHEDULE_KEY, JSON.stringify(s));
}

/** What the sidebar shows: the saved schedule, whether it is active, the next and the last run. */
function rptScheduleInfo() {
  let me = "";
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  const s = rptSchedule();
  if (!s) return { exists: false, defaultEmail: me };
  const active = !!rptFindTrigger(s.triggerId);
  return {
    exists: true, active, defaultEmail: me,
    form: { freq: s.freq, weekday: s.weekday, monthDay: s.monthDay, hour: s.hour, keep: s.keep, email: s.email, to: (s.to || []).join(", ") },
    summary: rptScheduleText(s), next: active ? rptNextRunText(s) : "", last: s.last || null, reports: (s.tabs || []).length
  };
}

/** Saves (or replaces) the schedule. form = { freq, weekday, monthDay, hour, keep, email, to }; choice = the sidebar's settings. */
function rptSaveSchedule(form, choice) {
  const freq = form.freq === "monthly" ? "monthly" : "weekly";
  const s = {
    freq,
    weekday: Math.max(0, Math.min(6, Number(form.weekday) || 0)),
    monthDay: Math.max(1, Math.min(28, Number(form.monthDay) || 4)),
    hour: Math.max(0, Math.min(23, Number(form.hour) || 0)),
    keep: Math.max(0, Number(form.keep) || 0),
    email: !!form.email,
    to: String(form.to || "").split(/[\s,;]+/).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)).slice(0, 10)
  };
  if (s.email && !s.to.length) throw new Error("Enter at least one email address, or untick the email option.");
  const c = Object.assign({}, choice);
  if (!c.gscSite && !c.gaProp) throw new Error("Choose a Search Console property, a GA4 property, or both (card 2) before saving the schedule.");
  ["start", "end", "compareStart", "compareEnd", "compareLabel", "preset"].forEach(k => delete c[k]);
  // Your own URLs are saved in a hidden tab, so every run reports on the same list.
  if (c.source === "selection" || c.source === "paste") {
    let urls;
    if (c.source === "paste") urls = rptParseUrls(c.urls);
    else {
      const r = SpreadsheetApp.getActiveRange();
      urls = r ? rptParseUrls(r.getValues().map(row => row.join("\n")).join("\n")) : [];
    }
    if (!urls.length) throw new Error(c.source === "paste" ? "Paste at least one URL (card 3)." : "Select the cells with your URLs first (card 3).");
    rptWriteUrlTab(urls.slice(0, RPT.MAX_URLS));
    c.source = "paste"; c.urls = ""; c.urlTab = true;
  }
  c.output = RPT.NEW_SHEET;
  const old = rptSchedule();
  if (old) rptDeleteTrigger(old.triggerId);
  s.choice = c;
  s.tabs = old ? old.tabs || [] : [];
  s.last = old ? old.last || null : null;
  s.triggerId = rptCreateTrigger(s);
  rptStoreSchedule(s);
  return rptScheduleInfo();
}

/** Pause (removes the timer, keeps the settings) or resume. */
function rptPauseSchedule(pause) {
  const s = rptSchedule();
  if (!s) throw new Error("No schedule saved yet.");
  rptDeleteTrigger(s.triggerId);
  s.triggerId = pause ? "" : rptCreateTrigger(s);
  rptStoreSchedule(s);
  return rptScheduleInfo();
}

/** Deletes the schedule (reports already created stay). */
function rptDeleteSchedule() {
  const s = rptSchedule();
  if (s) rptDeleteTrigger(s.triggerId);
  PropertiesService.getDocumentProperties().deleteProperty(RPT_SCHEDULE_KEY);
  const tab = SpreadsheetApp.getActive().getSheetByName(RPT_URL_TAB);
  if (tab && SpreadsheetApp.getActive().getSheets().length > 1) SpreadsheetApp.getActive().deleteSheet(tab);
  return rptScheduleInfo();
}

/** "Run now" button: one scheduled run straight away (same period, tab, keep and email rules). */
function rptRunScheduleNow() {
  rptScheduledRun();
  return rptScheduleInfo();
}

/** The timer calls this: builds the report for the last full week / month, tidies old tabs, sends the email. */
function rptScheduledRun() {
  const s = rptSchedule();
  if (!s) return null;
  const p = rptSchedulePeriod(s, new Date(), Session.getScriptTimeZone());
  const c = Object.assign({}, s.choice, p, { fromSchedule: true, output: RPT.NEW_SHEET });
  if (c.urlTab) c.urls = rptReadUrlTab();
  let result = null, error = "";
  try { result = rptRun(c); } catch (e) { error = e.message; }
  s.last = { at: Date.now(), ok: !error, error, sheet: result ? result.sheetName : "", pages: result ? result.pages : 0, period: rptRangeLabel(p.start, p.end), deleted: 0 };
  if (result) {
    s.tabs = (s.tabs || []).concat([result.sheetId]);
    s.last.deleted = rptPruneTabs(s);
  }
  if (s.email && s.to && s.to.length) {
    try { rptSendScheduleEmail(s, c, result, error); s.last.emailed = s.to.join(", "); }
    catch (e) { s.last.emailError = e.message; }
  }
  rptStoreSchedule(s);
  return s.last;
}

/** Last full week (Monday–Sunday) or last full month, plus the comparison period when Compare is on. */
function rptSchedulePeriod(s, now, tz) {
  const t = Utilities.formatDate(now, tz, "yyyy-MM-dd").split("-").map(Number);
  const DAY = 864e5, iso = ms => new Date(ms).toISOString().slice(0, 10);
  const today = Date.UTC(t[0], t[1] - 1, t[2]);
  let start, end, ps, pe, label;
  if (s.freq === "monthly") {
    start = Date.UTC(t[0], t[1] - 2, 1); end = Date.UTC(t[0], t[1] - 1, 0);
    ps = Date.UTC(t[0], t[1] - 3, 1); pe = Date.UTC(t[0], t[1] - 2, 0); label = "previous month";
  } else {
    const thisMonday = today - ((new Date(today).getUTCDay() + 6) % 7) * DAY;
    start = thisMonday - 7 * DAY; end = thisMonday - DAY;
    ps = start - 7 * DAY; pe = start - DAY; label = "previous week";
  }
  if (s.choice.compareMode === "year") {
    const back = ms => { const d = new Date(ms); return Date.UTC(d.getUTCFullYear() - 1, d.getUTCMonth(), d.getUTCDate()); };
    ps = back(start); pe = back(end); label = "same period last year";
  }
  const out = { start: iso(start), end: iso(end), preset: s.freq === "monthly" ? "lastMonth" : "lastWeek" };
  if (s.choice.compare) Object.assign(out, { compareStart: iso(ps), compareEnd: iso(pe), compareLabel: label });
  return out;
}

/** Keep only the newest N scheduled report tabs (tabs built by hand are never touched). Returns how many were deleted. */
function rptPruneTabs(s) {
  if (!s.keep) return 0;
  const ss = SpreadsheetApp.getActive();
  let deleted = 0;
  while (s.tabs.length > s.keep) {
    const id = s.tabs.shift();
    const sh = ss.getSheets().find(x => x.getSheetId() === id);
    if (sh && ss.getSheets().length > 1) { ss.deleteSheet(sh); deleted++; }
  }
  return deleted;
}

function rptCreateTrigger(s) {
  let b = ScriptApp.newTrigger(RPT_SCHEDULE_HANDLER).timeBased();
  b = s.freq === "monthly" ? b.onMonthDay(s.monthDay) : b.onWeekDay(ScriptApp.WeekDay[RPT_WEEKDAYS[s.weekday]]);
  return b.atHour(s.hour).create().getUniqueId();
}

function rptFindTrigger(id) {
  if (!id) return null;
  return ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === id) || null;
}

function rptDeleteTrigger(id) {
  const t = rptFindTrigger(id);
  if (t) ScriptApp.deleteTrigger(t);
}

function rptWriteUrlTab(urls) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RPT_URL_TAB) || ss.insertSheet(RPT_URL_TAB);
  sh.clear();
  if (sh.getMaxRows() < urls.length + 1) sh.insertRowsAfter(sh.getMaxRows(), urls.length + 1 - sh.getMaxRows());
  sh.getRange(1, 1).setValue("URLs used by the scheduled Performance Report (edit this list to change the pages)").setFontWeight("bold");
  sh.getRange(2, 1, urls.length, 1).setValues(urls.map(u => [u]));
  try { sh.hideSheet(); } catch (e) {}
}

function rptReadUrlTab() {
  const sh = SpreadsheetApp.getActive().getSheetByName(RPT_URL_TAB);
  if (!sh || sh.getLastRow() < 2) return "";
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]).join("\n");
}

function rptHourLabel(h) {
  const f = x => (x % 12 || 12) + (x < 12 ? " AM" : " PM");
  return f(h) + "–" + f((h + 1) % 24);
}

function rptScheduleText(s) {
  return (s.freq === "monthly" ? "Every month on day " + s.monthDay + ", " + rptHourLabel(s.hour) + " · report on last month"
    : "Every " + RPT_DAY_NAMES[s.weekday] + ", " + rptHourLabel(s.hour) + " · report on last week (Mon–Sun)") +
    (s.choice.compare ? " · compared with the " + (s.choice.compareMode === "year" ? "same period last year" : s.freq === "monthly" ? "previous month" : "previous week") : "") +
    (s.email ? " · email to " + s.to.join(", ") : "") + (s.keep ? " · keeps the last " + s.keep + " reports" : " · keeps all reports");
}

/** Next run (Google starts the timer at some point within the chosen hour). */
function rptNextRunText(s) {
  const tz = Session.getScriptTimeZone(), now = new Date();
  const hourNow = Number(Utilities.formatDate(now, tz, "H"));
  for (let i = 0; i < 40; i++) {
    const d = new Date(now.getTime() + i * 864e5);
    const dow = (Number(Utilities.formatDate(d, tz, "u")) + 6) % 7; // u: 1 = Monday … 7 = Sunday
    const dom = Number(Utilities.formatDate(d, tz, "d"));
    const match = s.freq === "monthly" ? dom === s.monthDay : dow === s.weekday;
    if (match && (i > 0 || hourNow < s.hour)) return Utilities.formatDate(d, tz, "EEE d MMM yyyy") + ", " + rptHourLabel(s.hour);
  }
  return "";
}

/** The email: key numbers, findings and a button to the report tab. Errors are emailed too. */
function rptSendScheduleEmail(s, c, r, error) {
  const ss = SpreadsheetApp.getActive();
  const domain = r ? r.domain : (c.gscSite ? rptLabel(c.gscSite) : c.gaName || "GA4 " + c.gaProp);
  const period = rptRangeLabel(c.start, c.end);
  const e = t => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  let subject, html;
  if (!r) {
    subject = "⚠ Scheduled report failed · " + domain;
    html = "<div style='font-family:Arial,sans-serif;font-size:14px;color:#202124'><h2 style='color:#c5221f'>The scheduled Performance Report could not be built</h2>" +
      "<p><b>" + e(domain) + "</b> · " + e(period) + "</p><p style='background:#fce8e6;padding:10px;border-radius:8px'>" + e(error) + "</p>" +
      "<p>Open the sheet → SEO Tools → Performance Report (GSC + GA4) and press <b>Run now</b> in the Schedule card to try again.</p>" +
      "<p><a href='" + ss.getUrl() + "'>Open the spreadsheet</a></p></div>";
  } else {
    subject = "📈 Performance Report · " + domain + " · " + period;
    const link = ss.getUrl() + "#gid=" + r.sheetId;
    const kpi = (r.kpis || []).map(k => "<tr><td style='padding:6px 10px;border-bottom:1px solid #eee'>" + e(k.label) + "</td><td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:right'><b>" + e(k.cur) + "</b></td>" +
      (r.compare ? "<td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:right;color:#5f6368'>" + e(k.prev) + "</td><td style='padding:6px 10px;border-bottom:1px solid #eee;font-weight:bold;color:" +
        (/▲/.test(k.change) ? "#137333" : /▼/.test(k.change) ? "#c5221f" : "#5f6368") + "'>" + e(k.change) + "</td>" : "") + "</tr>").join("");
    const head = "<tr style='background:#f1f3f4'><th style='padding:6px 10px;text-align:left'>Metric</th><th style='padding:6px 10px;text-align:right'>This period</th>" +
      (r.compare ? "<th style='padding:6px 10px;text-align:right'>Previous</th><th style='padding:6px 10px;text-align:left'>Change</th>" : "") + "</tr>";
    const finds = (r.findings || []).map(f => "<li style='margin:6px 0'>" + e(f[0] + " " + f[1]) + "</li>").join("");
    html = "<div style='font-family:Arial,sans-serif;font-size:14px;color:#202124;max-width:680px'>" +
      "<div style='background:#1a73e8;color:#fff;padding:16px;border-radius:10px 10px 0 0'><div style='font-size:20px;font-weight:bold'>📈 Performance Report · " + e(domain) + "</div>" +
      "<div style='opacity:.9;margin-top:4px'>📅 " + e(period) + (c.compare ? " · vs " + e(rptRangeLabel(c.compareStart, c.compareEnd)) : "") + " · " + r.pages + " pages</div></div>" +
      "<div style='border:1px solid #dadce0;border-top:0;padding:16px;border-radius:0 0 10px 10px'>" +
      "<h3 style='margin:0 0 8px'>Key numbers · whole site</h3><table style='border-collapse:collapse;width:100%'>" + head + kpi + "</table>" +
      (finds ? "<h3 style='margin:18px 0 6px'>Key findings · what to do first</h3><ul style='padding-left:18px;margin:0'>" + finds + "</ul>" : "") +
      "<p style='margin:20px 0'><a href='" + link + "' style='background:#188038;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font-weight:bold'>Open the full report ↗</a></p>" +
      "<p style='color:#5f6368;font-size:12px'>Automatic " + (s.freq === "monthly" ? "monthly" : "weekly") + " report · Prepared by Supriyo's Test Lab. " +
      "To change or stop it: SEO Tools → Performance Report (GSC + GA4) → Schedule.</p></div></div>";
  }
  MailApp.sendEmail({ to: s.to.join(","), subject, htmlBody: html, name: "Supriyo's Test Lab · SEO Tools" });
}

// ============================================================
//  SEARCH CONSOLE
// ============================================================

function rptGscUrl(site) {
  return "https://searchconsole.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(site) + "/searchAnalytics/query";
}

/** Every page with impressions in the period: key → { url, clicks, impressions, position }. */
function rptGscPages(site, start, end) {
  const out = {};
  for (let startRow = 0; startRow < RPT.GSC_MAX_PAGE_ROWS; startRow += RPT.GSC_ROW_LIMIT) {
    const res = rptCall(rptGscUrl(site), { startDate: start, endDate: end, dimensions: ["page"], rowLimit: RPT.GSC_ROW_LIMIT, startRow }, "Google Search Console API");
    const rows = res.rows || [];
    rows.forEach(r => { out[rptKey(r.keys[0])] = { url: r.keys[0], clicks: r.clicks, impressions: r.impressions, position: r.position }; });
    if (rows.length < RPT.GSC_ROW_LIMIT) break;
  }
  return out;
}

/** Whole-property totals (Search Console counts these per property, so they differ from the sum of the pages). */
function rptGscTotals(site, start, end) {
  const r = (rptCall(rptGscUrl(site), { startDate: start, endDate: end }, "Google Search Console API").rows || [])[0];
  return r ? { clicks: r.clicks, impressions: r.impressions, position: r.position } : { clicks: 0, impressions: 0, position: 0 };
}

function rptGscSummary(t) {
  return t ? { clicks: t.clicks, impressions: t.impressions, ctr: t.impressions ? t.clicks / t.impressions : 0, position: t.impressions ? t.position : 0 } : null;
}

/** The site's top 10 search queries by clicks (whole property). */
function rptGscTopQueries(site, start, end) {
  return (rptCall(rptGscUrl(site), { startDate: start, endDate: end, dimensions: ["query"], rowLimit: 10 }, "Google Search Console API").rows || [])
    .map(r => ({ q: r.keys[0], clicks: r.clicks, impressions: r.impressions, position: r.position }));
}

/** Exact top keywords for each listed page (one request per page, 10 at a time). */
function rptGscKeywordsFor(site, pages, start, end, n) {
  const out = {};
  const token = ScriptApp.getOAuthToken();
  for (let i = 0; i < pages.length; i += RPT.PARALLEL) {
    const batch = pages.slice(i, i + RPT.PARALLEL);
    const resps = UrlFetchApp.fetchAll(batch.map(p => ({
      url: rptGscUrl(site), method: "post", contentType: "application/json", headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true,
      payload: JSON.stringify({ startDate: start, endDate: end, dimensions: ["query"], rowLimit: 500,
        dimensionFilterGroups: [{ filters: [{ dimension: "page", operator: "equals", expression: p }] }] })
    })));
    resps.forEach((r, k) => {
      if (r.getResponseCode() !== 200) return;
      const rows = JSON.parse(r.getContentText()).rows || [];
      out[rptKey(batch[k])] = rows.sort((a, b) => b.impressions - a.impressions).slice(0, n).map(x => x.keys[0]);
    });
    if (i + RPT.PARALLEL < pages.length) Utilities.sleep(200);
  }
  return out;
}

/** Top keywords for every page from page + keyword rows (sorted by impressions per page). */
function rptGscKeywordsBulk(site, start, end, n) {
  const byPage = {};
  let read = 0, capped = false;
  for (let startRow = 0; ; startRow += RPT.GSC_ROW_LIMIT) {
    if (startRow >= RPT.GSC_MAX_KEYWORD_ROWS) { capped = true; break; }
    const rows = rptCall(rptGscUrl(site), { startDate: start, endDate: end, dimensions: ["page", "query"], rowLimit: RPT.GSC_ROW_LIMIT, startRow }, "Google Search Console API").rows || [];
    rows.forEach(r => { const k = rptKey(r.keys[0]); (byPage[k] = byPage[k] || []).push({ q: r.keys[1], i: r.impressions }); });
    read += rows.length;
    if (rows.length < RPT.GSC_ROW_LIMIT) break;
  }
  const map = {};
  Object.keys(byPage).forEach(k => { map[k] = byPage[k].sort((a, b) => b.i - a.i).slice(0, n).map(x => x.q); });
  return { map, capped, read };
}

// ============================================================
//  GA4
// ============================================================

function rptGaReport(prop, body) {
  const rows = [];
  for (let offset = 0; offset < RPT.GA_MAX_ROWS; offset += RPT.GA_LIMIT) {
    const res = rptCall("https://analyticsdata.googleapis.com/v1beta/properties/" + prop + ":runReport",
      Object.assign({}, body, { limit: RPT.GA_LIMIT, offset }), "Google Analytics Data API");
    (res.rows || []).forEach(r => rows.push({ d: (r.dimensionValues || []).map(x => x.value), m: (r.metricValues || []).map(x => Number(x.value) || 0) }));
    if (!res.rowCount || offset + RPT.GA_LIMIT >= res.rowCount) break;
  }
  return rows;
}

/** GA4 pages: key (host + path) → { url, sessions, engaged }. */
function rptGaPages(prop, dim, start, end) {
  const out = {};
  rptGaReport(prop, { dateRanges: [{ startDate: start, endDate: end }], dimensions: [{ name: "hostName" }, { name: dim }],
    metrics: [{ name: "sessions" }, { name: "engagedSessions" }] }).forEach(r => {
    const path = r.d[1];
    if (!path || path === "(not set)" || path.charAt(0) !== "/") return;
    const url = "https://" + r.d[0] + path, k = rptKey(url);
    const o = out[k] || (out[k] = { url, sessions: 0, engaged: 0 });
    o.sessions += r.m[0]; o.engaged += r.m[1];
  });
  return out;
}

function rptGaTotals(prop, start, end) {
  const r = rptGaReport(prop, { dateRanges: [{ startDate: start, endDate: end }], metrics: [{ name: "sessions" }, { name: "engagedSessions" }] })[0];
  const sessions = r ? r.m[0] : 0, engaged = r ? r.m[1] : 0;
  return { sessions, engaged, rate: sessions ? engaged / sessions : 0 };
}

/** Whole-site sessions per channel (Organic Search, Direct…), most first. */
function rptGaChannels(prop, start, end) {
  return rptGaReport(prop, { dateRanges: [{ startDate: start, endDate: end }], dimensions: [{ name: "sessionDefaultChannelGroup" }], metrics: [{ name: "sessions" }] })
    .map(r => ({ name: r.d[0] || "(not set)", sessions: r.m[0] })).sort((a, b) => b.sessions - a.sessions);
}

/** Top 5 traffic sources per page: key → "Organic Search 62%, Direct 20%, Referral 9%, …". */
function rptGaSources(prop, dim, trafficDim, start, end) {
  const per = {};
  rptGaReport(prop, { dateRanges: [{ startDate: start, endDate: end }], dimensions: [{ name: "hostName" }, { name: dim }, { name: trafficDim }],
    metrics: [{ name: "sessions" }] }).forEach(r => {
    if (!r.d[1] || r.d[1].charAt(0) !== "/") return;
    const k = rptKey("https://" + r.d[0] + r.d[1]);
    const p = per[k] || (per[k] = { total: 0, by: {} });
    p.total += r.m[0];
    p.by[r.d[2] || "(not set)"] = (p.by[r.d[2] || "(not set)"] || 0) + r.m[0];
  });
  const out = {};
  Object.keys(per).forEach(k => {
    const p = per[k];
    out[k] = Object.keys(p.by).sort((a, b) => p.by[b] - p.by[a]).slice(0, 5)
      .map(s => s + " " + (p.total ? Math.round(p.by[s] / p.total * 100) : 0) + "%").join(", ");
  });
  return out;
}

// ============================================================
//  ROWS + SHEET
// ============================================================

/** One row per page: your list in your order, or every page with data (most clicks, then sessions, first). */
function rptBuildRows(urls, data, c, kwCount) {
  const zeroG = { clicks: 0, impressions: 0, position: 0 }, zeroA = { sessions: 0, engaged: 0 };
  const gaKey = k => k.replace(/\?.*$/, "");
  let pages;
  if (urls) {
    pages = urls.map(u => ({ key: rptKey(u), url: u }));
  } else {
    const seen = {};
    pages = [];
    Object.keys(data.gsc || {}).forEach(k => { seen[k] = 1; pages.push({ key: k, url: data.gsc[k].url }); });
    Object.keys(data.ga || {}).forEach(k => { if (!seen[k]) { seen[k] = 1; pages.push({ key: k, url: data.ga[k].url }); } });
  }
  // Web pages only: PDFs, images and other files that get impressions or views are left out.
  const total = pages.length;
  pages = pages.filter(p => !rptIsFile(p.url));
  const skippedFiles = total - pages.length;
  const rows = pages.map(p => ({
    url: p.url,
    g: (data.gsc && data.gsc[p.key]) || zeroG,
    gp: (data.gscPrev && data.gscPrev[p.key]) || zeroG,
    a: (data.ga && (data.ga[p.key] || data.ga[gaKey(p.key)])) || zeroA,
    ap: (data.gaPrev && (data.gaPrev[p.key] || data.gaPrev[gaKey(p.key)])) || zeroA,
    kw: data.kw[p.key] || [],
    src: data.src[p.key] || data.src[gaKey(p.key)] || ""
  }));
  if (!urls) {
    rows.sort((x, y) => y.g.clicks - x.g.clicks || y.g.impressions - x.g.impressions || y.a.sessions - x.a.sessions);
    const max = Number(c.maxPages) || 0;
    if (max > 0) rows.length = Math.min(rows.length, max);
  }
  rows.skippedFiles = skippedFiles;
  return rows;
}

/** A file rather than a web page (by its extension): PDF, image, video, audio, document, archive, feed, script, font. */
function rptIsFile(url) {
  const path = String(url).replace(/^[a-z]+:\/\/[^\/]+/i, "").replace(/[?#].*$/, "");
  return /\.(pdf|jpe?g|png|gif|webp|avif|svg|ico|bmp|tiff?|heic|mp4|m4v|mov|avi|webm|mkv|mp3|wav|ogg|m4a|zip|rar|7z|gz|tar|docx?|xlsx?|pptx?|odt|ods|csv|txt|rtf|xml|json|rss|atom|js|css|map|woff2?|ttf|otf|eot|exe|dmg|apk|msi)$/i.test(path);
}

function rptPct(cur, prev) {
  if (!prev) return cur ? "New" : "";
  return (cur - prev) / prev;
}

function rptResultSheet(output, domain, start, end) {
  const ss = SpreadsheetApp.getActive();
  if (output && output !== RPT.NEW_SHEET) {
    const sh = ss.getSheetByName(String(output));
    if (!sh) throw new Error("Tab '" + output + "' was not found.");
    sh.clear();
    sh.clearConditionalFormatRules();
    return sh;
  }
  let name = ("Report · " + domain + " · " + rptRangeLabel(start, end, true)).replace(/[\[\]\*\?\/\\:]/g, " ").slice(0, 90);
  if (ss.getSheetByName(name)) { let i = 2; while (ss.getSheetByName(name + " (" + i + ")")) i++; name += " (" + i + ")"; }
  return ss.insertSheet(name);
}

/** "1 Oct – 31 Oct 2026" (short: "1–31 Oct 2026"). */
function rptRangeLabel(start, end, short) {
  const M = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const a = start.split("-").map(Number), b = end.split("-").map(Number);
  if (short && a[0] === b[0] && a[1] === b[1]) return a[2] + "–" + b[2] + " " + M[b[1] - 1] + " " + b[0];
  return a[2] + " " + M[a[1] - 1] + (a[0] !== b[0] ? " " + a[0] : "") + " – " + b[2] + " " + M[b[1] - 1] + " " + b[0];
}

function rptWrite(sh, rows, data, c, o) {
  const cmp = !!c.compare;
  const ctr = g => g.impressions ? g.clicks / g.impressions : 0;
  const rate = a => a.sessions ? a.engaged / a.sessions : 0;
  const r1 = n => Math.round(n * 10) / 10;
  const cols = [{ h: "Page", w: 380, v: r => r.url }];
  const add = (h, w, fmt, v, delta) => cols.push({ h, w, fmt, v, delta });
  if (o.site) {
    add("Impressions", 105, "#,##0", r => r.g.impressions);
    if (cmp) add("Impressions Δ", 100, "+0%;-0%;0%", r => rptPct(r.g.impressions, r.gp.impressions), "up");
    add("Clicks", 80, "#,##0", r => r.g.clicks);
    if (cmp) add("Clicks Δ", 80, "+0%;-0%;0%", r => rptPct(r.g.clicks, r.gp.clicks), "up");
    add("CTR", 75, "0.00%", r => ctr(r.g));
    if (cmp) add("CTR Δ (pts)", 90, "+0.00%;-0.00%;0.00%", r => r.g.impressions || r.gp.impressions ? ctr(r.g) - ctr(r.gp) : "", "up");
    add("Avg Position", 95, "0.0", r => r.g.impressions ? r1(r.g.position) : RPT.EMPTY);
    if (cmp) add("Position Δ", 90, "+0.0;-0.0;0.0", r => r.g.impressions && r.gp.impressions ? r1(r.g.position - r.gp.position) : "", "down");
    if (o.kwCount) add("Top " + o.kwCount + " Keywords", 260, null, r => r.kw.length ? r.kw.join(", ") : RPT.EMPTY);
  }
  if (o.gaProp) {
    add("Sessions", 90, "#,##0", r => r.a.sessions);
    if (cmp) add("Sessions Δ", 90, "+0%;-0%;0%", r => rptPct(r.a.sessions, r.ap.sessions), "up");
    add("Engaged Sessions", 110, "#,##0", r => r.a.engaged);
    add("Engagement Rate", 110, "0.00%", r => rate(r.a));
    if (cmp) add("Engagement Rate Δ (pts)", 120, "+0.00%;-0.00%;0.00%", r => r.a.sessions || r.ap.sessions ? rate(r.a) - rate(r.ap) : "", "up");
    add("Top 5 Traffic Sources (" + o.traffic.label + ")", 320, null, r => r.src || RPT.EMPTY);
  }

  // ---------- Layout: dashboard on top, page table below (same tab) ----------
  const W = Math.max(cols.length, 6);
  const C = { navy: "#1f3864", blue: "#1a73e8", light: "#e8f0fe", grey: "#f1f3f4", line: "#dadce0", text: "#3c4043", muted: "#5f6368", green: "#137333", red: "#c5221f" };
  const ins = rptInsights(rows, data, c, o);
  // A reused tab may still have the old layout: unfreeze, unmerge, remove banding.
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  try { sh.getBandings().forEach(b => b.remove()); } catch (e) {}
  const need = 160 + rows.length;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
  sh.setHiddenGridlines(true);
  cols.forEach((x, i) => sh.setColumnWidth(i + 1, x.w));
  for (let i = cols.length; i < W; i++) sh.setColumnWidth(i + 1, 110);
  const fullWidth = cols.reduce((a, x) => a + x.w, 0) + (W - cols.length) * 110;

  let row = 1;
  // One merged line across the dashboard width; tall enough for wrapped text.
  const line = (text, fontSize, color, bold, bg) => {
    const r = sh.getRange(row, 1, 1, W).merge();
    sh.getRange(row, 1).setValue(text);
    r.setFontSize(fontSize || 10).setFontColor(color || C.text).setFontWeight(bold ? "bold" : "normal").setWrap(true).setVerticalAlignment("middle");
    if (bg) r.setBackground(bg);
    const lines = Math.max(1, Math.ceil(String(text).length * (fontSize || 10) * 0.55 / fullWidth));
    sh.setRowHeight(row, Math.max(22, lines * (fontSize || 10) * 1.6 + 8));
    row++;
    return r;
  };
  const section = title => { line(title, 11, "#ffffff", true, C.navy); sh.setRowHeight(row - 1, 30); };
  const gap = () => { sh.setRowHeight(row, 10); row++; };
  // Small table: headers + rows; fmts[j] = number format; colors(value, j) = font colour or null.
  const mini = (headers, values, fmts, colors) => {
    const nc = headers.length;
    sh.getRange(row, 1, 1, nc).setValues([headers]).setFontWeight("bold").setBackground(C.grey).setFontColor(C.text)
      .setBorder(null, null, true, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
    row++;
    if (values.length) {
      const rg = sh.getRange(row, 1, values.length, nc);
      rg.setValues(values).setFontColor(C.text).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("middle")
        .setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
      (fmts || []).forEach((f, j) => { if (f) sh.getRange(row, j + 1, values.length, 1).setNumberFormat(f); });
      if (colors) values.forEach((v, i) => v.forEach((x, j) => { const col = colors(x, j, v); if (col) sh.getRange(row + i, j + 1).setFontColor(col).setFontWeight("bold"); }));
      row += values.length;
    }
  };

  // ---- Title ----
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");
  line("📈 Performance Report · " + o.domain, 18, C.blue, true);
  line("📅 " + rptRangeLabel(c.start, c.end) + (cmp ? "   ·   vs " + rptRangeLabel(c.compareStart, c.compareEnd) + (c.compareLabel ? " (" + c.compareLabel + ")" : "") : "") +
    "   ·   " + rows.length.toLocaleString() + " page" + (rows.length === 1 ? "" : "s") + (o.source === "all" ? " (all pages with data)" : " (your list)"), 11, C.text);
  line([o.site ? "Search Console: " + o.site : "", o.gaProp ? "GA4: " + (c.gaName ? c.gaName + " (" + o.gaProp + ")" : o.gaProp) +
    ", sessions by " + (o.gaDim === "landingPage" ? "landing page" : "any page view") : ""].filter(Boolean).join("   ·   ") +
    (o.notes.length ? "   ·   " + o.notes.join(" ") : "") + "   ·   Prepared by Supriyo's Test Lab · " + stamp, 9, C.muted);
  const jumpRow = row++;
  gap();

  // ---- Key numbers ----
  section("KEY NUMBERS · WHOLE SITE");
  mini(cmp ? ["Metric", "This period", "Previous period", "Change"] : ["Metric", "This period"], ins.kpis.map(k => cmp ? [k.label, k.cur, k.prev, k.change] : [k.label, k.cur]),
    [null, "@", "@", null], (x, j, v) => j === 3 && /▲/.test(String(x)) ? C.green : j === 3 && /▼/.test(String(x)) ? C.red : null);
  gap();

  // ---- Key findings ----
  if (ins.findings.length) {
    section("KEY FINDINGS · WHAT TO DO FIRST");
    ins.findings.forEach((f, i) => { line(f[0] + "  " + f[1], 10, C.text, false, i % 2 ? null : "#f8f9fa"); });
    gap();
  }

  // ---- Top lists ----
  ins.lists.forEach(l => {
    line(l.title, 11, C.blue, true, C.light);
    if (l.note) line(l.note, 9, C.muted);
    mini(l.headers, l.rows, l.fmts, l.colors);
    gap();
  });

  // ---- Page table ----
  section("PAGE DATA · " + rows.length.toLocaleString() + " PAGE" + (rows.length === 1 ? "" : "S") + (o.source === "all" ? (o.site ? " · SORTED BY CLICKS" : " · SORTED BY SESSIONS") : " · IN YOUR ORDER"));
  const headRow = row;
  sh.getRange(headRow, 1, 1, cols.length).setValues([cols.map(x => x.h)]).setFontWeight("bold").setBackground(C.navy).setFontColor("#ffffff")
    .setWrap(true).setVerticalAlignment("middle");
  sh.setRowHeight(headRow, 36);
  const first = headRow + 1;
  if (rows.length) {
    const rg = sh.getRange(first, 1, rows.length, cols.length);
    rg.setValues(rows.map(r => cols.map(x => { const v = x.v(r); return v === undefined || v === null ? "" : v; })))
      .setVerticalAlignment("top").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontColor(C.text);
    // Stripes on the data rows only: banding on the header row would paint over its navy background (white text on white).
    try { sh.getRange(first, 1, rows.length, cols.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch (e) {}
  }
  const n = Math.max(1, rows.length);
  const rules = [];
  cols.forEach((x, i) => {
    const rg = sh.getRange(first, i + 1, n, 1);
    if (x.fmt) rg.setNumberFormat(x.fmt);
    if (x.delta) {
      const good = x.delta === "up" ? "whenNumberGreaterThan" : "whenNumberLessThan";
      const bad = x.delta === "up" ? "whenNumberLessThan" : "whenNumberGreaterThan";
      rules.push(SpreadsheetApp.newConditionalFormatRule()[good](0).setFontColor(C.green).setBold(true).setRanges([rg]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()[bad](0).setFontColor(C.red).setBold(true).setRanges([rg]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("New").setFontColor(C.blue).setBold(true).setRanges([rg]).build());
    }
  });
  sh.setConditionalFormatRules(rules);

  // Link at the top that jumps straight to the table.
  sh.getRange(jumpRow, 1).setFormula('=HYPERLINK("#gid=' + sh.getSheetId() + '&range=A' + headRow + '","⬇  Jump to page data (' + rows.length.toLocaleString() + ' pages)")')
    .setFontColor(C.blue).setFontWeight("bold");
  return ins;
}

/**
 * Plain-language findings and short lists from the report rows, for people who won't read every page.
 * Lists use up to 5 rows (top pages: 10). Thresholds are written into each finding so it is clear why a page is listed.
 */
function rptInsights(rows, data, c, o) {
  const cmp = !!c.compare;
  const F = [], L = [], K = [];
  const n0 = x => Math.round(x || 0).toLocaleString();
  const pct = (a, b) => b ? (a - b) / b : null;
  const sp = x => (x > 0 ? "+" : "") + Math.round(x * 100) + "%";
  const pts = x => (x > 0 ? "+" : "") + (x * 100).toFixed(2) + " pts";
  const arrow = (d, goodWhenUp, txt) => d == null ? "" : Math.abs(d) < 0.0005 ? "= no change" : ((d > 0) === goodWhenUp ? "▲ " : "▼ ") + txt;
  const short = u => String(u).replace(/^https?:\/\/[^\/]+/i, "") || "/";
  const ctr = g => g.impressions ? g.clicks / g.impressions : 0;
  const rate = a => a.sessions ? a.engaged / a.sessions : 0;
  const gt = o.site ? rptGscSummary(data.gscTotals) : null, gp = o.site && cmp ? rptGscSummary(data.gscTotalsPrev) : null;
  const at = o.gaProp ? data.gaTotals : null, ap = o.gaProp && cmp ? data.gaTotalsPrev : null;
  const pageCols = (extra) => ["Page"].concat(extra);

  // ---- Key numbers ----
  if (gt) {
    K.push({ label: "🖱 Clicks", cur: n0(gt.clicks), prev: gp ? n0(gp.clicks) : "", change: gp ? arrow(pct(gt.clicks, gp.clicks), true, sp(pct(gt.clicks, gp.clicks) || 0)) : "" });
    K.push({ label: "👁 Impressions", cur: n0(gt.impressions), prev: gp ? n0(gp.impressions) : "", change: gp ? arrow(pct(gt.impressions, gp.impressions), true, sp(pct(gt.impressions, gp.impressions) || 0)) : "" });
    K.push({ label: "🎯 CTR", cur: (gt.ctr * 100).toFixed(2) + "%", prev: gp ? (gp.ctr * 100).toFixed(2) + "%" : "", change: gp ? arrow(gt.ctr - gp.ctr, true, pts(gt.ctr - gp.ctr)) : "" });
    const dpos = gp && gp.position ? Math.round((gt.position - gp.position) * 10) / 10 : null;
    K.push({ label: "📍 Avg position", cur: gt.position.toFixed(1), prev: gp ? gp.position.toFixed(1) : "", change: dpos == null ? "" : dpos === 0 ? "= no change" : (dpos < 0 ? "▲ " : "▼ ") + Math.abs(dpos) + (dpos < 0 ? " places better" : " places worse") });
  }
  if (at) {
    K.push({ label: "👥 Sessions", cur: n0(at.sessions), prev: ap ? n0(ap.sessions) : "", change: ap ? arrow(pct(at.sessions, ap.sessions), true, sp(pct(at.sessions, ap.sessions) || 0)) : "" });
    K.push({ label: "✅ Engaged sessions", cur: n0(at.engaged), prev: ap ? n0(ap.engaged) : "", change: ap ? arrow(pct(at.engaged, ap.engaged), true, sp(pct(at.engaged, ap.engaged) || 0)) : "" });
    K.push({ label: "⏱ Engagement rate", cur: (at.rate * 100).toFixed(2) + "%", prev: ap ? (ap.rate * 100).toFixed(2) + "%" : "", change: ap ? arrow(at.rate - ap.rate, true, pts(at.rate - ap.rate)) : "" });
    const ch = data.channels || [], tot = ch.reduce((a, x) => a + x.sessions, 0), org = ch.find(x => x.name === "Organic Search");
    if (tot) {
      const share = org ? org.sessions / tot : 0;
      const chp = data.channelsPrev || [], totp = chp.reduce((a, x) => a + x.sessions, 0), orgp = chp.find(x => x.name === "Organic Search");
      const sharep = cmp && totp ? (orgp ? orgp.sessions / totp : 0) : null;
      K.push({ label: "🌱 Organic Search share", cur: (share * 100).toFixed(1) + "%", prev: sharep == null ? "" : (sharep * 100).toFixed(1) + "%", change: sharep == null ? "" : arrow(share - sharep, true, pts(share - sharep)) });
    }
  }

  // ---- Findings ----
  if (gt && gp) {
    const dc = pct(gt.clicks, gp.clicks), di = pct(gt.impressions, gp.impressions);
    if (dc != null && di != null) {
      const why = dc < 0 && di > 0 ? "more people see the site in Google but fewer click (CTR " + pts(gt.ctr - gp.ctr) + "). Review titles and meta descriptions of the pages losing clicks, and check whether new impressions come from low positions."
        : dc < 0 ? "search visibility fell. Check rankings, indexing and recent changes on the pages losing clicks below."
        : di < 0 ? "fewer impressions but more clicks: the pages that rank are converting searchers better."
        : "visibility and clicks are both growing. Keep updating the pages that drive the growth.";
      F.push(["📊", "Search: clicks " + sp(dc) + ", impressions " + sp(di) + ", avg position " + gt.position.toFixed(1) + " (was " + gp.position.toFixed(1) + "): " + why]);
    }
  } else if (gt) {
    F.push(["📊", "Search: " + n0(gt.clicks) + " clicks from " + n0(gt.impressions) + " impressions (CTR " + (gt.ctr * 100).toFixed(2) + "%, avg position " + gt.position.toFixed(1) + "). Tick 'Compare' in the sidebar to see what changed."]);
  }
  if (at) {
    const ch = data.channels || [], tot = ch.reduce((a, x) => a + x.sessions, 0);
    let s = "Visitors: " + n0(at.sessions) + " sessions, " + (at.rate * 100).toFixed(1) + "% engaged";
    if (ap && ap.sessions) s += " (sessions " + sp(pct(at.sessions, ap.sessions)) + ", engagement " + pts(at.rate - ap.rate) + ")";
    if (tot) s += ". Main channels: " + ch.slice(0, 3).map(x => x.name + " " + Math.round(x.sessions / tot * 100) + "%").join(", ");
    F.push(["👥", s + "."]);
  }
  if (o.site) {
    const sum = rows.reduce((a, r) => a + r.g.clicks, 0);
    if (sum) {
      const top = rows.slice().sort((a, b) => b.g.clicks - a.g.clicks).slice(0, 10), share = top.reduce((a, r) => a + r.g.clicks, 0) / sum;
      F.push(["🏆", "The top " + top.length + " pages bring " + Math.round(share * 100) + "% of the clicks in this report" +
        (share >= 0.7 ? ": search traffic depends on a few pages. Protect them (no URL changes) and refresh their content first." : ": traffic is spread across many pages.")]);
    }
  }
  if (o.site && cmp) {
    const moves = rows.map(r => ({ r, d: r.g.clicks - r.gp.clicks }));
    const losers = moves.filter(m => m.d < 0).sort((a, b) => a.d - b.d), gainers = moves.filter(m => m.d > 0).sort((a, b) => b.d - a.d);
    const lost = -losers.reduce((a, m) => a + m.d, 0), won = gainers.reduce((a, m) => a + m.d, 0);
    if (losers.length) {
      const top3 = losers.slice(0, 3), share = lost ? -top3.reduce((a, m) => a + m.d, 0) / lost : 0;
      F.push(["📉", losers.length + (losers.length === 1 ? " page lost " : " pages lost ") + n0(lost) + " clicks; the top 3 (" + top3.map(m => short(m.r.url)).join(", ") + ") explain " + Math.round(share * 100) + "% of the loss. Start there: compare their positions and titles with the previous period."]);
    }
    if (gainers.length) F.push(["📈", gainers.length + (gainers.length === 1 ? " page gained " : " pages gained ") + n0(won) + " clicks; the biggest: " + gainers.slice(0, 3).map(m => short(m.r.url) + " (+" + n0(m.d) + ")").join(", ") + ". See what changed on them and repeat it on similar pages."]);
    const gone = rows.filter(r => r.gp.impressions >= 100 && !r.g.impressions);
    if (gone.length) F.push(["🚫", gone.length + " page" + (gone.length === 1 ? "" : "s") + " had 100+ impressions before and none now (e.g. " + gone.slice(0, 3).map(r => short(r.url)).join(", ") + (gone.length === 1 ? "). Check it still exists, returns 200 and is indexed" : "). Check they still exist, return 200 and are indexed") + " (Check Indexing tool)."]);
    const fresh = rows.filter(r => !r.gp.impressions && r.g.impressions >= 50);
    if (fresh.length) F.push(["🆕", fresh.length + " page" + (fresh.length === 1 ? "" : "s") + " started getting impressions this period (e.g. " + fresh.slice(0, 3).map(r => short(r.url)).join(", ") + ")."]);
    const fmt = [null, "#,##0", "#,##0", "+#,##0;-#,##0;0"];
    if (gainers.length) L.push({ title: "📈 Biggest gainers (clicks)", headers: pageCols(["Clicks", "Previous", "Change"]), rows: gainers.slice(0, 5).map(m => [m.r.url, m.r.g.clicks, m.r.gp.clicks, m.d]), fmts: fmt, colors: (x, j) => j === 3 ? RPT_GREEN : null });
    if (losers.length) L.push({ title: "📉 Biggest losers (clicks)", headers: pageCols(["Clicks", "Previous", "Change"]), rows: losers.slice(0, 5).map(m => [m.r.url, m.r.g.clicks, m.r.gp.clicks, m.d]), fmts: fmt, colors: (x, j) => j === 3 ? RPT_RED : null });
  }
  if (o.site) {
    const qw = rows.filter(r => r.g.impressions >= 100 && r.g.position >= 4 && r.g.position <= 20).sort((a, b) => b.g.impressions - a.g.impressions);
    if (qw.length) {
      F.push(["🚀", qw.length + " page" + (qw.length === 1 ? "" : "s") + (qw.length === 1 ? " ranks" : " rank") + " in positions 4–20 with 100+ impressions. Moving them into the top 3 brings the most new clicks for the least work: add internal links, refresh content, match search intent."]);
      L.push({ title: "🚀 Quick wins · positions 4–20, most impressions", headers: pageCols(["Impressions", "Clicks", "CTR", "Avg Position"]), rows: qw.slice(0, 5).map(r => [r.url, r.g.impressions, r.g.clicks, ctr(r.g), Math.round(r.g.position * 10) / 10]), fmts: [null, "#,##0", "#,##0", "0.00%", "0.0"] });
    }
    const siteCtr = gt ? gt.ctr : 0;
    const low = rows.filter(r => r.g.impressions >= 200 && r.g.position <= 10 && ctr(r.g) < siteCtr * 0.5).sort((a, b) => b.g.impressions - a.g.impressions);
    if (low.length) {
      F.push(["✏️", low.length + " page" + (low.length === 1 ? "" : "s") + (low.length === 1 ? " ranks" : " rank") + " on page 1 (position ≤ 10, 200+ impressions) but " + (low.length === 1 ? "gets" : "get") + " less than half the site's CTR (" + (siteCtr * 100).toFixed(2) + "%). Rewrite their titles and meta descriptions to earn the click."]);
      L.push({ title: "✏️ Low CTR on page 1 · rewrite title & description", headers: pageCols(["Impressions", "Clicks", "CTR", "Avg Position"]), rows: low.slice(0, 5).map(r => [r.url, r.g.impressions, r.g.clicks, ctr(r.g), Math.round(r.g.position * 10) / 10]), fmts: [null, "#,##0", "#,##0", "0.00%", "0.0"], colors: (x, j) => j === 3 ? RPT_RED : null });
    }
  }
  if (at && at.sessions) {
    const minS = Math.max(20, Math.round(at.sessions * 0.002));
    const weak = rows.filter(r => r.a.sessions >= minS && rate(r.a) < at.rate - 0.15).sort((a, b) => b.a.sessions - a.a.sessions);
    if (weak.length) {
      F.push(["⚠️", weak.length + " page" + (weak.length === 1 ? "" : "s") + " with " + minS + "+ sessions " + (weak.length === 1 ? "has" : "have") + " engagement 15+ pts below the site average (" + (at.rate * 100).toFixed(0) + "%). Check that the content matches what visitors searched for, page speed and the next step (CTA)."]);
      L.push({ title: "⚠️ Low engagement · many visitors leave quickly", headers: pageCols(["Sessions", "Engagement Rate", "Top Source"]), rows: weak.slice(0, 5).map(r => [r.url, r.a.sessions, rate(r.a), (r.src || "-").split(", ")[0]]), fmts: [null, "#,##0", "0.00%", null], colors: (x, j) => j === 2 ? RPT_RED : null });
    }
  }
  if (o.site && o.gaProp && o.gaDim === "landingPage") {
    const noGa = rows.filter(r => r.g.clicks >= 20 && !r.a.sessions);
    if (noGa.length) F.push(["🔧", noGa.length + " page" + (noGa.length === 1 ? "" : "s") + (noGa.length === 1 ? " gets" : " get") + " 20+ search clicks but no GA4 sessions as a landing page (e.g. " + noGa.slice(0, 3).map(r => short(r.url)).join(", ") + "). Check the GA4 tag, redirects and the consent banner on them."]);
  }

  // ---- Lists that are always useful ----
  const byClicks = o.site ? rows.slice().sort((a, b) => b.g.clicks - a.g.clicks) : rows.slice().sort((a, b) => b.a.sessions - a.a.sessions);
  if (o.site) {
    L.unshift({ title: "🏆 Top 10 pages by clicks", headers: pageCols(cmp ? ["Clicks", "Clicks Δ", "Impressions", "Avg Position"] : ["Clicks", "Impressions", "CTR", "Avg Position"]),
      rows: byClicks.slice(0, 10).filter(r => r.g.clicks || r.g.impressions).map(r => cmp ? [r.url, r.g.clicks, rptPct(r.g.clicks, r.gp.clicks), r.g.impressions, Math.round(r.g.position * 10) / 10]
        : [r.url, r.g.clicks, r.g.impressions, ctr(r.g), Math.round(r.g.position * 10) / 10]),
      fmts: cmp ? [null, "#,##0", "+0%;-0%;0%", "#,##0", "0.0"] : [null, "#,##0", "#,##0", "0.00%", "0.0"],
      colors: cmp ? (x, j) => j === 2 && typeof x === "number" ? (x > 0 ? RPT_GREEN : x < 0 ? RPT_RED : null) : null : null });
  } else {
    L.unshift({ title: "🏆 Top 10 pages by sessions", headers: pageCols(["Sessions", "Engaged Sessions", "Engagement Rate"]),
      rows: byClicks.slice(0, 10).filter(r => r.a.sessions).map(r => [r.url, r.a.sessions, r.a.engaged, rate(r.a)]), fmts: [null, "#,##0", "#,##0", "0.00%"] });
  }
  if (o.site && data.queries && data.queries.length) {
    L.push({ title: "🔎 Top search queries · whole site", headers: ["Query", "Clicks", "Impressions", "CTR", "Avg Position"],
      rows: data.queries.slice(0, 10).map(q => [q.q, q.clicks, q.impressions, q.impressions ? q.clicks / q.impressions : 0, Math.round(q.position * 10) / 10]),
      fmts: [null, "#,##0", "#,##0", "0.00%", "0.0"], note: "Queries Google shares (anonymised ones are hidden by Google)." });
  }
  if (at && data.channels && data.channels.length) {
    const tot = data.channels.reduce((a, x) => a + x.sessions, 0);
    const prev = {}; (data.channelsPrev || []).forEach(x => { prev[x.name] = x.sessions; });
    L.push({ title: "🌐 Traffic channels · whole site (GA4)", headers: cmp ? ["Channel", "Sessions", "Share", "Change"] : ["Channel", "Sessions", "Share"],
      rows: data.channels.slice(0, 8).map(x => cmp ? [x.name, x.sessions, tot ? x.sessions / tot : 0, rptPct(x.sessions, prev[x.name] || 0)] : [x.name, x.sessions, tot ? x.sessions / tot : 0]),
      fmts: [null, "#,##0", "0.0%", "+0%;-0%;0%"], colors: (x, j) => j === 3 && typeof x === "number" ? (x > 0 ? RPT_GREEN : x < 0 ? RPT_RED : null) : null });
  }
  if (!F.length) F.push(["ℹ️", "Not enough data in this period for findings. Try a longer period or tick 'Compare'."]);
  return { kpis: K, findings: F, lists: L };
}

const RPT_GREEN = "#137333", RPT_RED = "#c5221f";

// ============================================================
//  HELPERS
// ============================================================

/** One API call (GET when payload is null) with a retry for "slow down" / server errors; friendly errors. */
function rptCall(url, payload, api) {
  const opt = { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  if (payload) { opt.method = "post"; opt.contentType = "application/json"; opt.payload = JSON.stringify(payload); }
  let r;
  for (let attempt = 0; attempt < 3; attempt++) {
    r = UrlFetchApp.fetch(url, opt);
    const code = r.getResponseCode();
    if (code === 200) return JSON.parse(r.getContentText());
    if (code !== 429 && code < 500) break;
    Utilities.sleep(2000 * (attempt + 1));
  }
  throw rptError(r, api);
}

function rptError(r, api) {
  const text = r.getContentText(), code = r.getResponseCode();
  let msg = "";
  try { msg = (JSON.parse(text).error || {}).message || ""; } catch (e) {}
  let err;
  if (/has not been used|is disabled|SERVICE_DISABLED|accessNotConfigured/i.test(text)) {
    err = new Error("The " + api + " isn't enabled in this sheet's Google Cloud project yet (one-time setup).");
    err.enableUrl = (text.match(/https:\/\/console\.(?:developers|cloud)\.google\.com\/apis\/(?:api|library)\/[a-z.]+googleapis\.com[^"\s\\]*/) || [""])[0];
  } else if (/insufficient.*scope|ACCESS_TOKEN_SCOPE_INSUFFICIENT/i.test(text)) {
    err = new Error("Missing permission for the " + api + ": check appsscript.json (analytics.readonly), save, reload the sheet and authorise again.");
  } else if (code === 403) {
    err = new Error("Your Google account has no access to this " + (/Analytics/.test(api) ? "GA4 property" : "Search Console property") + (msg ? " (" + msg + ")" : "") + ".");
  } else if (code === 401) {
    err = new Error("Google sign-in expired. Reload the sheet and try again.");
  } else if (code === 429) {
    err = new Error(api + ": quota reached for now. Try again in a few minutes.");
  } else {
    err = new Error(api + ": " + (msg || "HTTP " + code));
  }
  return err;
}

/** Page key for matching Search Console and GA4: host without www + path without trailing slash (+ query). */
function rptKey(u) {
  const m = String(u).trim().match(/^(?:https?:\/\/)?([^\/?#]+)([^?#]*)(\?[^#]*)?/i);
  if (!m) return String(u).trim().toLowerCase();
  let path = m[2] || "/";
  if (path.length > 1) path = path.replace(/\/+$/, "");
  return m[1].toLowerCase().replace(/^www\./, "").replace(/:\d+$/, "") + (path || "/") + (m[3] || "");
}

function rptLabel(site) {
  return String(site).replace(/^sc-domain:/, "").replace(/^https?:\/\//, "").replace(/\/$/, "");
}

/** Unique http(s) URLs from text (one per line, or separated by spaces/commas). Bare domains get https://. */
function rptParseUrls(text) {
  const out = [], seen = {};
  String(text || "").split(/[\s,;]+/).forEach(t => {
    t = t.trim().replace(/^["'<(]+|[>)"',;]+$/g, "");
    if (!t) return;
    if (!/^https?:\/\//i.test(t)) {
      if (!/^[a-z0-9-]+(\.[a-z0-9-]+)+(\/|$)/i.test(t)) return;
      t = "https://" + t;
    }
    if (!/^https?:\/\/[^\s\/]+\.[^\s\/]+/i.test(t) || seen[t.toLowerCase()]) return;
    seen[t.toLowerCase()] = 1;
    out.push(t);
  });
  return out;
}

// The one-time GA4 setup, shown in the sidebar's guide (opens by itself when a property list can't load).
const RPT_SETUP_GUIDE = [
  { stage: "Same Google Cloud project as Check Indexing", steps: [
    { title: "Open the Cloud project linked to this sheet",
      text: "Extensions → Apps Script → ⚙ Project Settings shows the project number. Select that project at the top of Google Cloud.",
      link: "https://console.cloud.google.com/home/dashboard", linkText: "Open Google Cloud" }
  ] },
  { stage: "Enable the GA4 APIs", steps: [
    { title: "Enable Google Analytics Data API",
      text: "☰ → APIs & Services → Library → search \"Google Analytics Data API\" → Enable. This one reads the report data.",
      link: "https://console.cloud.google.com/apis/library/analyticsdata.googleapis.com", linkText: "Open Data API" },
    { title: "Enable Google Analytics Admin API",
      text: "Same place → \"Google Analytics Admin API\" → Enable. It fills the GA4 property list (without it, type the property ID).",
      link: "https://console.cloud.google.com/apis/library/analyticsadmin.googleapis.com", linkText: "Open Admin API" }
  ] },
  { stage: "Permission + access", steps: [
    { title: "Check appsscript.json",
      text: "Apps Script → ⚙ Project Settings → tick \"Show appsscript.json manifest file in editor\". Open appsscript.json, check it lists https://www.googleapis.com/auth/analytics.readonly, then Ctrl+S." },
    { title: "Authorise Analytics access from the editor",
      text: "Open performance-report.gs → in the function list at the top choose checkReportSetup (not openPerformanceReport) → Run. If asked: Review permissions → your Gmail → Advanced → Go to (app name) → tick \"See and download your Google Analytics data\" (or Select all) → Continue. The Execution log then shows ✅ / ❌ for each part." },
    { title: "Log says \"Missing permission\"?",
      text: "Google is reusing your old approval. Remove it: Google Account → Security → Your connections to third-party apps → your app name → Delete all connections. Then Run checkReportSetup again.",
      link: "https://myaccount.google.com/connections", linkText: "Open your connections" },
    { title: "Check your GA4 access",
      text: "Your Gmail needs at least Viewer access to the GA4 property (GA4 → Admin → Property access management)." },
    { title: "Find a property ID (if typing it)",
      text: "GA4 → Admin → Property settings → Property ID, a number like 336987406." }
  ] }
];

// ============================================================
//  SIDEBAR
// ============================================================

const RPT_SIDEBAR_HTML = `<!doctype html><html><head><meta charset="utf-8"><style>
  *{box-sizing:border-box}
  body{font-family:Arial,sans-serif;font-size:13px;color:#202124;margin:0;padding:0 12px 14px;background:#f8f9fa}
  .hero{display:flex;align-items:center;gap:10px;margin:0 -12px 10px;padding:12px;background:linear-gradient(135deg,#1a73e8,#1557b0);color:#fff}
  .hero .ic{width:34px;height:34px;border-radius:50%;background:rgba(255,255,255,.18);display:flex;align-items:center;justify-content:center;font-size:17px}
  .hero b{display:block;font-size:15px} .hero span{font-size:11px;opacity:.9}
  .card{background:#fff;border:1px solid #dadce0;border-radius:10px;padding:10px 12px;margin:10px 0}
  .sec{display:flex;align-items:center;gap:8px;font-size:14px;font-weight:bold;margin:0 0 8px}
  .num{flex:0 0 20px;width:20px;height:20px;border-radius:50%;background:#e8f0fe;color:#1a73e8;font-size:11px;display:flex;align-items:center;justify-content:center}
  .muted{color:#5f6368;font-size:12px;line-height:1.45} .hidden{display:none!important}
  .lbl{display:flex;justify-content:space-between;align-items:center;color:#5f6368;font-size:12px;margin:10px 0 5px}
  .q{display:inline-block;width:14px;height:14px;border-radius:50%;background:#9aa0a6;color:#fff;font-size:10px;font-weight:bold;text-align:center;line-height:14px;cursor:help}
  select,textarea,input[type=text],input[type=date]{width:100%;border:1px solid #dadce0;border-radius:6px;padding:8px;font-size:13px;background:#fff;color:#202124;font-family:inherit}
  select:focus,textarea:focus,input:focus{outline:none;border-color:#1a73e8;box-shadow:0 0 0 2px #e8f0fe}
  textarea{height:120px;font-family:monospace;font-size:11px}
  .chips{display:grid;grid-template-columns:1fr 1fr;gap:6px}
  .chips button{border:1px solid #dadce0;background:#fff;border-radius:16px;padding:7px 4px;font-size:12px;cursor:pointer;color:#3c4043}
  .chips button:hover{border-color:#aecbfa} .chips button.on{background:#1a73e8;border-color:#1a73e8;color:#fff;font-weight:bold}
  .dates{display:flex;gap:6px} .dates div{flex:1}
  .range{margin-top:10px;background:#e8f0fe;color:#174ea6;border-radius:8px;padding:8px 10px;font-size:12.5px;line-height:1.5}
  .range b{font-size:13px} .range .vs{color:#5f6368}
  .chk{display:flex;gap:8px;align-items:center;margin:10px 0 4px;padding:8px;border-radius:8px;background:#f8f9fa;cursor:pointer;font-size:12.5px;font-weight:bold}
  .chk input{accent-color:#1a73e8;width:15px;height:15px;margin:0}
  .warn{color:#b06000;font-size:12px;margin-top:6px} .warn:empty{display:none}
  .bad{background:#fce8e6;border:1px solid #f4b4ae;border-radius:8px;padding:8px 10px;color:#a50e0e;font-size:12px;margin-top:6px;line-height:1.45}
  .bad a.enable{display:inline-block;margin-top:5px;background:#d93025;color:#fff;border-radius:6px;padding:4px 9px;font-weight:bold}
  .src{display:flex;align-items:center;gap:6px;font-weight:bold;font-size:12.5px;margin:10px 0 5px} .src i{font-style:normal;font-size:14px}
  .guide{background:#fff;border:1px solid #c6dafc;border-radius:10px;margin:10px 0;overflow:hidden}
  .guide summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:8px;padding:10px 12px;background:#e8f0fe;color:#174ea6;font-weight:bold}
  .guide summary::-webkit-details-marker{display:none}
  .guide summary small{display:block;font-weight:normal;color:#5f6368;font-size:11px;margin-top:2px}
  .guide .chev{display:inline-block;margin-left:auto;font-size:11px;transition:transform .2s} .guide[open] .chev{transform:rotate(90deg)}
  .gbody{padding:4px 12px 10px} .stage{font-size:11px;font-weight:bold;letter-spacing:.4px;text-transform:uppercase;color:#1a73e8;margin:12px 0 6px}
  .gstep{display:flex;gap:9px;position:relative;padding-bottom:10px}
  .gstep:before{content:"";position:absolute;left:10px;top:22px;bottom:0;width:2px;background:#e8eaed} .gstep.last:before{display:none}
  .gnum{flex:0 0 22px;height:22px;border-radius:50%;background:#1a73e8;color:#fff;font-size:11px;font-weight:bold;display:flex;align-items:center;justify-content:center;z-index:1}
  .gtxt b{display:block;font-size:12.5px;margin:2px 0} .gtxt div{font-size:12px;color:#3c4043;line-height:1.45}
  .gtxt a{display:inline-block;margin-top:4px;font-size:11.5px;border:1px solid #c6dafc;border-radius:12px;padding:2px 9px}
  #go{width:100%;border:0;border-radius:8px;padding:12px;font-size:14px;font-weight:bold;cursor:pointer;background:linear-gradient(135deg,#188038,#137333);color:#fff;box-shadow:0 1px 3px rgba(0,0,0,.2);margin-top:4px}
  #go:hover{filter:brightness(1.07)} #go:disabled{background:#9aa0a6;box-shadow:none;cursor:default}
  .err{background:#fce8e6;border:1px solid #f4b4ae;border-radius:8px;padding:8px 10px;color:#a50e0e;font-size:12px;margin-top:10px;line-height:1.45}
  .bar{position:relative;background:#e8eaed;border-radius:4px;height:20px;overflow:hidden;margin:10px 0 6px}
  .fill{background:#1a73e8;height:100%;width:0;transition:width .5s} .fill.done{background:#188038}
  .pct{position:absolute;inset:0;text-align:center;line-height:20px;font-size:12px;font-weight:bold;color:#202124}
  .phead{display:flex;justify-content:space-between;align-items:center}
  .chip{font-size:11px;font-weight:bold;border-radius:10px;padding:3px 9px;background:#e8f0fe;color:#1a73e8} .chip.done{background:#e6f4ea;color:#188038}
  .steps{margin-top:6px} .st{display:flex;gap:8px;align-items:center;font-size:12px;color:#9aa0a6;margin:5px 0}
  .st .d{flex:0 0 18px;height:18px;border-radius:50%;border:2px solid #dadce0;font-size:10px;display:flex;align-items:center;justify-content:center}
  .st.ok{color:#202124} .st.ok .d{background:#188038;border-color:#188038;color:#fff} .st.now{color:#1a73e8;font-weight:bold} .st.now .d{border-color:#1a73e8;background:#1a73e8;color:#fff}
  .tiles{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:10px 0}
  .tile{border-radius:8px;padding:8px 10px;background:#f8f9fa;border-left:4px solid #1a73e8}
  .tile.ga{border-color:#f9ab00}
  .tile b{display:block;font-size:18px;color:#202124} .tile span{font-size:11.5px;color:#5f6368}
  .tile em{font-style:normal;font-size:11px;font-weight:bold;margin-left:4px} .up{color:#188038} .dn{color:#d93025}
  a{color:#1a73e8;cursor:pointer;text-decoration:none;font-weight:bold}
  .sig{text-align:center;color:#5f6368;font-size:12px;margin:16px 0 0}
  .sched{background:#fff;border:1px solid #dadce0;border-radius:10px;margin:10px 0;overflow:hidden}
  .sched summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:8px;padding:10px 12px;font-weight:bold;font-size:14px}
  .sched summary::-webkit-details-marker{display:none}
  .sched summary small{display:block;font-weight:normal;color:#5f6368;font-size:11px;margin-top:2px}
  .sched summary small.on{color:#137333;font-weight:bold}
  .sched .chev{display:inline-block;margin-left:auto;font-size:11px;color:#5f6368;transition:transform .2s} .sched[open] .chev{transform:rotate(90deg)}
  .sched[open] summary{border-bottom:1px solid #e8eaed} .sbody{padding:2px 12px 12px}
  .sbtns{display:flex;gap:8px;margin-top:12px}
  .sbtns button{flex:1;border-radius:8px;padding:10px 0;font-size:13px;font-weight:bold;cursor:pointer}
  .sbtns .blue{background:#1a73e8;color:#fff;border:0} .sbtns .ghost{background:#fff;color:#1a73e8;border:1px solid #dadce0}
  .sbtns button:disabled{opacity:.6;cursor:default}
  .sstat{border-radius:8px;padding:9px 10px;margin-top:10px;font-size:12px;line-height:1.55;background:#e6f4ea;color:#0d652d}
  .sstat.off{background:#f1f3f4;color:#3c4043} .sstat .bad{color:#c5221f;font-weight:bold}
  .slinks{display:flex;gap:16px;margin-top:8px;font-size:12.5px} .slinks a.del{color:#c5221f}
</style></head><body>
<div class="hero"><div class="ic">📈</div><div><b>Performance Report</b><span>Search Console + GA4 · pages, keywords, sessions</span></div></div>

<div class="card">
  <div class="sec"><span class="num">1</span>Period</div>
  <div class="chips" id="chips">
    <button data-p="last7">Last 7 days</button><button data-p="lastWeek">Last week</button>
    <button data-p="last28">Last 28 days</button><button data-p="lastMonth">Last month</button>
    <button data-p="thisMonth">This month so far</button><button data-p="custom">Custom dates</button>
    <button data-p="week">Pick a week</button><button data-p="month">Pick a month</button>
  </div>
  <div id="weekBox" class="hidden"><div class="lbl"><span>Week (Monday – Sunday)</span></div><select id="week"></select></div>
  <div id="monthBox" class="hidden"><div class="lbl"><span>Month</span></div><select id="month"></select></div>
  <div id="customBox" class="hidden"><div class="dates"><div><div class="lbl"><span>Start</span></div><input type="date" id="dStart"></div><div><div class="lbl"><span>End</span></div><input type="date" id="dEnd"></div></div></div>
  <label class="chk"><input type="checkbox" id="compare"><span>Compare with another period</span></label>
  <div id="compareBox" class="hidden"><select id="compareMode"><option value="previous">Previous period</option><option value="year">Same period last year</option></select></div>
  <div class="range" id="range"></div>
  <div class="warn" id="lagWarn"></div>
</div>

<div class="card">
  <div class="sec"><span class="num">2</span>Data sources</div>
  <div class="src"><i>🔍</i>Search Console property</div>
  <select id="gsc"><option value="">Loading…</option></select>
  <div class="bad hidden" id="gscErr"></div>
  <div class="src"><i>📊</i>GA4 property</div>
  <select id="ga"><option value="">Loading…</option></select>
  <input type="text" id="gaId" class="hidden" placeholder="GA4 property ID, e.g. 336987406" style="margin-top:6px">
  <div class="bad hidden" id="gaErr"></div>
  <div class="muted" style="margin-top:6px">Choose one or both. With both, each page shows search data and on-site behaviour side by side.</div>
</div>

<details class="guide" id="guide">
  <summary><span>🚀</span><span>GA4 setup guide<small>Analytics APIs + permission · about 3 minutes</small></span><span class="chev">▶</span></summary>
  <div class="gbody"><div id="guideSteps"></div></div>
</details>

<div class="card">
  <div class="sec"><span class="num">3</span>Pages</div>
  <select id="source"><option value="all">All pages with data</option><option value="selection">Current cell selection</option><option value="paste">Paste my URLs</option></select>
  <div id="allBox"><div class="lbl"><span>Most pages in the report</span><span class="q" title="Pages are sorted by clicks, then impressions, then sessions.">?</span></div>
    <select id="maxPages"><option value="100">Top 100</option><option value="500" selected>Top 500</option><option value="1000">Top 1,000</option><option value="5000">Top 5,000</option><option value="0">All pages</option></select></div>
  <div class="muted hidden" id="selInfo">Select the cells with your URLs in the sheet.</div>
  <div id="pasteBox" class="hidden"><textarea id="urls" placeholder="https://example.com/page-1/&#10;https://example.com/page-2/"></textarea>
    <div class="muted" id="pasteCount">0 URLs</div></div>
  <div class="warn" id="hostWarn"></div>
</div>

<div class="card">
  <div class="sec"><span class="num">4</span>Columns</div>
  <div id="kwBox"><div class="lbl"><span>Top keywords per page (Search Console)</span></div>
    <select id="keywords"><option value="0">None</option><option value="3">Top 3</option><option value="5">Top 5</option><option value="10" selected>Top 10</option></select></div>
  <div id="trafficBox"><div class="lbl"><span>Top 5 traffic sources show (GA4)</span><span class="q" title="Channel: Organic Search, Direct, Referral… · Source / medium: google / organic… · Source: google, facebook.com…">?</span></div>
    <select id="traffic"><option value="channel">Channel (Organic Search, Direct…)</option><option value="sourceMedium">Source / medium (google / organic…)</option><option value="source">Source (google, facebook.com…)</option></select>
    <div class="lbl"><span>Match GA4 sessions to pages by</span><span class="q" title="Landing page: sessions that started on the page (best next to Search Console clicks). Any page view: sessions that viewed the page at any point.">?</span></div>
    <select id="gaMatch"><option value="landingPage">Landing page (recommended)</option><option value="pagePath">Any page view</option></select></div>
  <div class="lbl"><span>Results sheet</span></div>
  <select id="output"><option value="__new__">Create New Sheet</option></select>
  <div class="warn" id="outWarn"></div>
</div>

<button id="go">Build report</button>
<div class="err hidden" id="err"></div>

<div class="card hidden" id="prog">
  <div class="phead"><b id="headline">Building…</b><span class="chip" id="chip">Running</span></div>
  <div class="bar"><div class="fill" id="fill"></div><div class="pct" id="pct">0%</div></div>
  <div class="steps" id="steps"></div>
  <div id="result" class="hidden">
    <div class="tiles" id="tiles"></div>
    <div class="muted" id="resNote"></div>
    <div style="margin-top:8px"><a id="openRes">Open report ↗</a></div>
  </div>
</div>

<details class="sched" id="schedCard">
  <summary><span class="num">5</span><span>Schedule<small id="schedSub">Off · build this report automatically</small></span><span class="chev">▶</span></summary>
  <div class="sbody">
    <div class="lbl"><span>Repeat</span></div>
    <select id="sFreq"><option value="weekly">Weekly · report on last week (Mon–Sun)</option><option value="monthly">Monthly · report on last month</option></select>
    <div id="sWeekBox"><div class="lbl"><span>Run every</span><span class="q" title="Search Console data takes 2–3 days, so Wednesday or later gives complete data for last week.">?</span></div>
      <select id="sWeekday"></select></div>
    <div id="sMonthBox" class="hidden"><div class="lbl"><span>Run on day</span><span class="q" title="Search Console data takes 2–3 days, so day 4 or later gives complete data for last month.">?</span></div>
      <select id="sMonthDay"></select></div>
    <div class="lbl"><span>Time (the sheet's time zone)</span></div>
    <select id="sHour"></select>
    <label class="chk"><input type="checkbox" id="sEmail" checked><span>Email when each report is ready</span></label>
    <input type="text" id="sTo" placeholder="you@gmail.com, client@example.com">
    <div class="lbl"><span>Keep</span><span class="q" title="Each run creates a new tab. With a limit, the oldest scheduled report tabs are deleted so the sheet stays tidy. Tabs you build yourself are never deleted.">?</span></div>
    <select id="sKeep"><option value="0">All scheduled reports</option><option value="6">Last 6 reports</option><option value="12">Last 12 reports</option><option value="24">Last 24 reports</option><option value="52">Last 52 reports</option></select>
    <div class="muted" id="keepHelp" style="margin-top:5px"></div>
    <div class="muted" style="margin-top:8px">Each run uses the sources, pages, columns and Compare setting chosen above (cards 1–4). Change them, then press Save schedule again.</div>
    <div class="sbtns"><button class="blue" id="sSave">Save schedule</button><button class="ghost" id="sNow">Run now</button></div>
    <div class="sstat off hidden" id="sStat"></div>
    <div class="slinks hidden" id="sLinks"><a id="sPause">Pause</a><a id="sDel" class="del">Delete schedule</a></div>
    <div class="err hidden" id="sErr"></div>
  </div>
</details>

<div class="sig">Made with ❤️ by Supriyo's Test Lab</div>

<script>
  var $ = function (id) { return document.getElementById(id); };
  var info = null, preset = "lastMonth", running = false, lastSheet = "", selTimer = null, gaTouched = false;
  var MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var LAG = 3; // Search Console data takes 2–3 days
  function show(id, on) { $(id).classList.toggle("hidden", !on); }
  function esc(t) { return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;"); }
  function nf(n) { return Number(n || 0).toLocaleString(); }
  function label(s) { return String(s).replace(/^sc-domain:/, "").replace(/^https?:\\/\\//, "").replace(/\\/$/, ""); }

  // ---- Dates (local calendar days) ----
  function day(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function add(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function iso(d) { return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function parse(s) { var p = String(s).split("-"); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function fmt(d) { return d.getDate() + " " + MON[d.getMonth()] + " " + d.getFullYear(); }
  function span(a, b) { return Math.round((b - a) / 864e5) + 1; }
  function monday(d) { return add(d, -((d.getDay() + 6) % 7)); }
  function lastOfMonth(y, m) { return new Date(y, m + 1, 0); }
  var today = day(new Date()), lagEnd = add(today, -LAG);

  function fillPickers() {
    var w = $("week"), m = $("month"), thisMon = monday(today), o = "";
    for (var i = 1; i <= 16; i++) { var s = add(thisMon, -7 * i), e = add(s, 6); o += '<option value="' + iso(s) + '">' + s.getDate() + " " + MON[s.getMonth()] + " – " + fmt(e) + (i === 1 ? " (last week)" : "") + "</option>"; }
    w.innerHTML = o; o = "";
    for (var j = 0; j < 16; j++) { var d = new Date(today.getFullYear(), today.getMonth() - j, 1); o += '<option value="' + iso(d) + '">' + MON[d.getMonth()] + " " + d.getFullYear() + (j === 0 ? " (so far)" : "") + "</option>"; }
    m.innerHTML = o; m.selectedIndex = 1;
    $("dEnd").value = iso(lagEnd); $("dStart").value = iso(add(lagEnd, -27)); $("dEnd").max = $("dStart").max = iso(add(today, -1));
  }
  /** The chosen period: { start, end, kind } (kind "month" / "monthToDate" / "other" decides the comparison). */
  function period() {
    var s, e, kind = "other";
    if (preset === "last7") { e = lagEnd; s = add(e, -6); }
    else if (preset === "last28") { e = lagEnd; s = add(e, -27); }
    else if (preset === "lastWeek") { s = add(monday(today), -7); e = add(s, 6); }
    else if (preset === "lastMonth") { s = new Date(today.getFullYear(), today.getMonth() - 1, 1); e = lastOfMonth(s.getFullYear(), s.getMonth()); kind = "month"; }
    else if (preset === "thisMonth") { s = new Date(today.getFullYear(), today.getMonth(), 1); e = lagEnd; kind = "monthToDate"; }
    else if (preset === "week") { s = parse($("week").value); e = add(s, 6); }
    else if (preset === "month") { s = parse($("month").value); e = lastOfMonth(s.getFullYear(), s.getMonth()); kind = "month"; if (e >= today) { e = lagEnd; kind = "monthToDate"; } }
    else { s = parse($("dStart").value || iso(lagEnd)); e = parse($("dEnd").value || iso(lagEnd)); }
    return { start: s, end: e, kind: kind };
  }
  function comparison(p) {
    if ($("compareMode").value === "year") {
      return { start: new Date(p.start.getFullYear() - 1, p.start.getMonth(), p.start.getDate()), end: new Date(p.end.getFullYear() - 1, p.end.getMonth(), p.end.getDate()), label: "same period last year" };
    }
    if (p.kind === "month") { var s = new Date(p.start.getFullYear(), p.start.getMonth() - 1, 1); return { start: s, end: lastOfMonth(s.getFullYear(), s.getMonth()), label: "previous month" }; }
    if (p.kind === "monthToDate") {
      var s2 = new Date(p.start.getFullYear(), p.start.getMonth() - 1, 1), last = lastOfMonth(s2.getFullYear(), s2.getMonth());
      return { start: s2, end: new Date(s2.getFullYear(), s2.getMonth(), Math.min(p.end.getDate(), last.getDate())), label: "same days last month" };
    }
    var n = span(p.start, p.end); return { start: add(p.start, -n), end: add(p.start, -1), label: "previous " + n + " days" };
  }
  function renderPeriod() {
    Array.prototype.forEach.call(document.querySelectorAll("#chips button"), function (b) { b.className = b.getAttribute("data-p") === preset ? "on" : ""; });
    show("weekBox", preset === "week"); show("monthBox", preset === "month"); show("customBox", preset === "custom");
    show("compareBox", $("compare").checked);
    var p = period(), html = "";
    if (p.end < p.start) { $("range").innerHTML = "⚠ Not enough data yet for this period (Search Console needs 2–3 days)."; return null; }
    html = "📅 <b>" + fmt(p.start) + " – " + fmt(p.end) + "</b> · " + span(p.start, p.end) + " days";
    if ($("compare").checked) { var c = comparison(p); html += '<br><span class="vs">vs ' + fmt(c.start) + " – " + fmt(c.end) + " (" + c.label + ")</span>"; }
    $("range").innerHTML = html;
    $("lagWarn").textContent = p.end > lagEnd ? "Search Console data for the last 2–3 days may still be incomplete." : "";
    return p;
  }
  document.querySelectorAll("#chips button").forEach(function (b) { b.onclick = function () { preset = b.getAttribute("data-p"); renderPeriod(); }; });
  ["week", "month", "dStart", "dEnd", "compareMode"].forEach(function (id) { $(id).onchange = renderPeriod; });
  $("compare").onchange = renderPeriod;

  // ---- Sources / pages / columns ----
  function gaValue() { return $("ga").value === "__manual__" ? $("gaId").value.trim() : $("ga").value; }
  function refresh() {
    var src = $("source").value, hasG = !!$("gsc").value, hasA = !!gaValue();
    show("allBox", src === "all"); show("selInfo", src === "selection"); show("pasteBox", src === "paste");
    show("gaId", $("ga").value === "__manual__");
    show("kwBox", hasG || !hasA); show("trafficBox", hasA || !hasG);
    $("outWarn").textContent = $("output").value !== "__new__" ? "All data in this tab will be replaced." : "";
    var warn = "";
    if (src === "paste") {
      var urls = pasted(); $("pasteCount").textContent = urls.length + " URL" + (urls.length === 1 ? "" : "s");
      var site = $("gsc").value;
      if (site && urls.length) {
        var host = label(site).split("/")[0].replace(/^www\\./, "");
        var off = urls.filter(function (u) { var h = (u.match(/^https?:\\/\\/([^\\/]+)/i) || [])[1] || ""; h = h.toLowerCase().replace(/^www\\./, ""); return h !== host && h.slice(-(host.length + 1)) !== "." + host; }).length;
        if (off) warn = "⚠ " + off + (off === 1 ? " URL isn't" : " URLs aren't") + " on " + host + ": " + (off === 1 ? "it" : "they") + " will show no Search Console data.";
      }
    }
    $("hostWarn").textContent = warn;
  }
  function pasted() {
    var seen = {}, out = [];
    $("urls").value.split(/[\\s,;]+/).forEach(function (t) {
      t = t.trim(); if (!t) return;
      if (!/^https?:\\/\\//i.test(t)) { if (!/^[a-z0-9-]+(\\.[a-z0-9-]+)+(\\/|$)/i.test(t)) return; t = "https://" + t; }
      if (seen[t.toLowerCase()]) return; seen[t.toLowerCase()] = 1; out.push(t);
    });
    return out;
  }
  function pollSelection() {
    if (selTimer) clearTimeout(selTimer);
    if ($("source").value !== "selection") return;
    google.script.run.withSuccessHandler(function (s) {
      $("selInfo").textContent = s && s.count ? s.count + " URL" + (s.count === 1 ? "" : "s") + " in the selection (" + s.a1 + ")" : "Select the cells with your URLs in the sheet (no URLs in the current selection).";
      selTimer = setTimeout(pollSelection, 3000);
    }).withFailureHandler(function () { selTimer = setTimeout(pollSelection, 5000); }).rptSelectionInfo();
  }
  // Suggest the GA4 property whose name matches the Search Console domain (until you pick one yourself).
  function suggestGa() {
    if (gaTouched || !info || !$("gsc").value) return;
    var word = label($("gsc").value).split("/")[0].replace(/^www\\./, "").split(".")[0].toLowerCase();
    var hit = info.ga.props.find(function (p) { return (p.name + " " + p.account).toLowerCase().replace(/[^a-z0-9]/g, "").indexOf(word.replace(/[^a-z0-9]/g, "")) !== -1; });
    if (hit) $("ga").value = hit.id;
  }
  $("source").onchange = function () { refresh(); pollSelection(); };
  $("gsc").onchange = function () { suggestGa(); refresh(); };
  $("ga").onchange = function () { gaTouched = true; refresh(); };
  $("gaId").oninput = refresh; $("urls").oninput = refresh; $("output").onchange = refresh;

  // ---- Run ----
  function bar(pct, done) { $("fill").style.width = pct + "%"; $("fill").className = "fill" + (done ? " done" : ""); $("pct").textContent = pct + "%"; $("pct").style.color = pct >= 50 ? "#fff" : "#202124"; }
  function renderSteps(s) {
    var list = s.steps || [];
    $("steps").innerHTML = list.map(function (name, i) {
      var cls = s.phase === "done" || i < (s.stepNo || 0) - 1 ? "ok" : i === (s.stepNo || 0) - 1 ? "now" : "";
      return "<div class='st " + cls + "'><span class='d'>" + (cls === "ok" ? "✓" : i + 1) + "</span>" + esc(name) + "</div>";
    }).join("");
    if (list.length) bar(s.phase === "done" ? 100 : Math.round(Math.max(0, (s.stepNo || 0) - 1) / list.length * 100), s.phase === "done");
  }
  function poll() {
    if (!running) return;
    google.script.run.withSuccessHandler(function (s) { if (running && s && s.phase === "running") renderSteps(s); setTimeout(poll, 1200); })
      .withFailureHandler(function () { setTimeout(poll, 2000); }).rptStatus();
  }
  $("go").onclick = function () {
    if (running) return;
    show("err", false);
    var p = renderPeriod(), g = $("gsc").value, a = gaValue(), src = $("source").value;
    var fail = function (m) { $("err").textContent = m; show("err", true); };
    if (!p) return fail("Choose a period that has data.");
    if (!g && !a) return fail("Choose a Search Console property, a GA4 property, or both.");
    if (a && !/^\\d+$/.test(a)) return fail("A GA4 property ID is a number like 336987406.");
    if (src === "paste" && !pasted().length) return fail("Paste at least one URL.");
    var c = $("compare").checked ? comparison(p) : null;
    var gaOpt = $("ga").options[$("ga").selectedIndex];
    var choice = { start: iso(p.start), end: iso(p.end), preset: preset, compare: !!c, compareMode: $("compareMode").value,
      compareStart: c ? iso(c.start) : "", compareEnd: c ? iso(c.end) : "", compareLabel: c ? c.label : "",
      gscSite: g, gaProp: a, gaName: $("ga").value === "__manual__" ? "" : (gaOpt ? gaOpt.getAttribute("data-name") || "" : ""),
      source: src, urls: $("urls").value, maxPages: $("maxPages").value, keywords: g ? $("keywords").value : 0,
      traffic: $("traffic").value, gaMatch: $("gaMatch").value, output: $("output").value };
    running = true; $("go").disabled = true; $("go").textContent = "Building…";
    show("prog", true); show("result", false); bar(0, false); $("steps").innerHTML = "";
    $("headline").textContent = "Building the report…"; $("chip").textContent = "Running"; $("chip").className = "chip";
    google.script.run.withSuccessHandler(done).withFailureHandler(function (e) {
      running = false; $("go").disabled = false; $("go").textContent = "Build report"; show("prog", false); fail("⚠ " + (e.message || String(e)));
    }).rptRun(choice);
    setTimeout(poll, 700);
  };
  function delta(cur, prev, pts, lowerBetter) {
    if (prev == null) return "";
    // pts: rate change in percentage points · lowerBetter (position): change in places · else: % change
    var d = pts ? (cur - prev) * 100 : lowerBetter ? cur - prev : prev ? (cur - prev) / prev * 100 : null;
    if (d === null) return "";
    if (Math.abs(d) < (lowerBetter ? 0.05 : 0.005)) return "<em>=</em>";
    var good = lowerBetter ? d < 0 : d > 0;
    var txt = (good ? "▲ " : "▼ ") + (d > 0 ? "+" : "") + (pts || lowerBetter ? d.toFixed(pts ? 2 : 1) : Math.round(d)) + (pts ? " pts" : lowerBetter ? "" : "%");
    return "<em class='" + (good ? "up" : "dn") + "'>" + txt + "</em>";
  }
  function done(r) {
    running = false; lastSheet = r.sheetName; $("go").disabled = false; $("go").textContent = "Build report";
    renderSteps({ phase: "done", steps: r.steps || [] });
    bar(100, true); $("headline").textContent = "Report ready ✅"; $("chip").textContent = "Complete ✓"; $("chip").className = "chip done";
    var t = [];
    if (r.gsc) {
      var gp = r.gscPrev;
      t.push(["Clicks", nf(r.gsc.clicks), delta(r.gsc.clicks, gp && gp.clicks), ""], ["Impressions", nf(r.gsc.impressions), delta(r.gsc.impressions, gp && gp.impressions), ""],
        ["CTR", (r.gsc.ctr * 100).toFixed(2) + "%", delta(r.gsc.ctr, gp && gp.ctr, true), ""], ["Avg position", r.gsc.position ? r.gsc.position.toFixed(1) : "-", gp && gp.position ? delta(r.gsc.position, gp.position, false, true) : "", ""]);
    }
    if (r.ga) {
      var ap = r.gaPrev;
      t.push(["Sessions", nf(r.ga.sessions), delta(r.ga.sessions, ap && ap.sessions), "ga"], ["Engagement rate", (r.ga.rate * 100).toFixed(1) + "%", delta(r.ga.rate, ap && ap.rate, true), "ga"]);
    }
    $("tiles").innerHTML = t.map(function (x) { return "<div class='tile " + x[3] + "'><b>" + x[1] + x[2] + "</b><span>" + x[0] + " · whole site</span></div>"; }).join("");
    $("resNote").textContent = nf(r.pages) + " page" + (r.pages === 1 ? "" : "s") + " → tab '" + r.sheetName + "'" + (r.notes.length ? " · " + r.notes.join(" ") : "");
    show("result", true);
    google.script.run.withSuccessHandler(function (i) { fillTabs(i.tabs); }).rptGetInfo();
  }
  $("openRes").onclick = function () { google.script.run.rptOpenTab(lastSheet); };

  // ---- Schedule ----
  var sched = null;
  /** The sidebar's current settings (cards 2–4 + Compare), without dates: each run picks its own period. */
  function settingsTemplate() {
    var g = $("gsc").value, a = gaValue(), gaOpt = $("ga").options[$("ga").selectedIndex];
    return { compare: $("compare").checked, compareMode: $("compareMode").value, gscSite: g, gaProp: a,
      gaName: $("ga").value === "__manual__" ? "" : (gaOpt ? gaOpt.getAttribute("data-name") || "" : ""),
      source: $("source").value, urls: $("urls").value, maxPages: $("maxPages").value, keywords: g ? $("keywords").value : 0,
      traffic: $("traffic").value, gaMatch: $("gaMatch").value };
  }
  function hourLabel(h) { var f = function (x) { return (x % 12 || 12) + (x < 12 ? " AM" : " PM"); }; return f(h) + " – " + f((h + 1) % 24); }
  function fillSchedulePickers() {
    $("sWeekday").innerHTML = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"].map(function (d, i) { return '<option value="' + i + '">' + d + (i === 2 ? " (recommended)" : "") + "</option>"; }).join("");
    $("sWeekday").value = "2";
    var o = ""; for (var d = 1; d <= 28; d++) o += '<option value="' + d + '">' + d + (d === 4 ? " (recommended)" : "") + "</option>";
    $("sMonthDay").innerHTML = o; $("sMonthDay").value = "4";
    o = ""; for (var h = 0; h < 24; h++) o += '<option value="' + h + '">' + hourLabel(h) + "</option>";
    $("sHour").innerHTML = o; $("sHour").value = "7";
  }
  function schedRefresh() {
    var m = $("sFreq").value === "monthly", k = Number($("sKeep").value);
    show("sWeekBox", !m); show("sMonthBox", m); show("sTo", $("sEmail").checked);
    $("keepHelp").textContent = k ? "Keeps the newest " + k + " scheduled report tabs (about " + (m ? k : Math.max(1, Math.round(k / 4.3))) + " month" + ((m ? k : Math.round(k / 4.3)) === 1 ? "" : "s") +
      " of history) and deletes older ones. Tabs you build yourself are never deleted." : "Nothing is deleted: every run adds a new tab.";
  }
  ["sFreq", "sKeep", "sEmail"].forEach(function (id) { $(id).onchange = schedRefresh; });
  function when(t) { return new Date(t).toLocaleString([], { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }); }
  function renderSchedule(i) {
    sched = i; show("sErr", false);
    $("sSave").disabled = false; $("sSave").textContent = i.exists ? "Update schedule" : "Save schedule";
    $("sNow").disabled = !i.exists; $("sNow").textContent = "Run now";
    if (!$("sTo").value && i.defaultEmail) $("sTo").value = i.defaultEmail;
    if (!i.exists) {
      show("sStat", false); show("sLinks", false);
      $("schedSub").textContent = "Off · build this report automatically"; $("schedSub").className = "";
      schedRefresh(); return;
    }
    var f = i.form;
    $("sFreq").value = f.freq; $("sWeekday").value = String(f.weekday); $("sMonthDay").value = String(f.monthDay); $("sHour").value = String(f.hour);
    $("sKeep").value = String(f.keep); $("sEmail").checked = !!f.email; if (f.to) $("sTo").value = f.to;
    schedRefresh();
    $("schedSub").textContent = i.active ? "✅ On · next: " + i.next : "⏸ Paused"; $("schedSub").className = i.active ? "on" : "";
    var l = i.last, html = (i.active ? "✅ <b>Active.</b> " : "⏸ <b>Paused.</b> ") + esc(i.summary) + (i.active ? "<br>⏭ Next run: <b>" + esc(i.next) + "</b>" : "");
    if (l) html += "<br>🕘 Last run " + when(l.at) + ": " + (l.ok ? "✅ " + esc(l.period) + ", " + nf(l.pages) + " pages → tab '" + esc(l.sheet) + "'" +
      (l.deleted ? " · " + l.deleted + " old report" + (l.deleted === 1 ? "" : "s") + " deleted" : "") : "<span class='bad'>❌ " + esc(l.error) + "</span>") +
      (l.emailed ? "<br>✉ Emailed to " + esc(l.emailed) : l.emailError ? "<br><span class='bad'>✉ Email failed: " + esc(l.emailError) + "</span>" : "");
    if (i.reports) html += "<br>📁 " + i.reports + " scheduled report" + (i.reports === 1 ? "" : "s") + " in this sheet";
    $("sStat").innerHTML = html; $("sStat").className = "sstat" + (i.active ? "" : " off"); show("sStat", true);
    show("sLinks", true); $("sPause").textContent = i.active ? "Pause" : "Resume"; $("sDel").textContent = "Delete schedule";
  }
  function schedFail(e) {
    $("sSave").disabled = false; $("sSave").textContent = sched && sched.exists ? "Update schedule" : "Save schedule";
    $("sNow").disabled = !(sched && sched.exists); $("sNow").textContent = "Run now";
    $("sErr").textContent = "⚠ " + (e.message || e); show("sErr", true);
  }
  $("sSave").onclick = function () {
    var t = settingsTemplate();
    show("sErr", false);
    if (!t.gscSite && !t.gaProp) return schedFail({ message: "Choose a Search Console property, a GA4 property, or both (card 2)." });
    if (t.gaProp && !/^\\d+$/.test(t.gaProp)) return schedFail({ message: "A GA4 property ID is a number like 336987406." });
    if ($("sEmail").checked && !/@/.test($("sTo").value)) return schedFail({ message: "Enter an email address, or untick the email option." });
    $("sSave").disabled = true; $("sSave").textContent = "Saving…";
    var form = { freq: $("sFreq").value, weekday: $("sWeekday").value, monthDay: $("sMonthDay").value, hour: $("sHour").value,
      keep: $("sKeep").value, email: $("sEmail").checked, to: $("sTo").value };
    google.script.run.withSuccessHandler(renderSchedule).withFailureHandler(schedFail).rptSaveSchedule(form, t);
  };
  $("sNow").onclick = function () {
    $("sNow").disabled = true; $("sNow").textContent = "Running…";
    google.script.run.withSuccessHandler(function (i) { renderSchedule(i); google.script.run.withSuccessHandler(function (x) { fillTabs(x.tabs); }).rptGetInfo(); })
      .withFailureHandler(schedFail).rptRunScheduleNow();
  };
  $("sPause").onclick = function () { google.script.run.withSuccessHandler(renderSchedule).withFailureHandler(schedFail).rptPauseSchedule(!!(sched && sched.active)); };
  $("sDel").onclick = function () {
    if ($("sDel").textContent === "Delete schedule") { $("sDel").textContent = "Click again to delete (reports already made stay)"; return; }
    google.script.run.withSuccessHandler(renderSchedule).withFailureHandler(schedFail).rptDeleteSchedule();
  };

  // ---- Load ----
  function fillTabs(tabs) {
    var keep = $("output").value;
    $("output").innerHTML = '<option value="__new__">Create New Sheet</option>' + tabs.map(function (t) { return '<option value="' + esc(t) + '">' + esc(t) + "</option>"; }).join("");
    if (tabs.indexOf(keep) !== -1) $("output").value = keep;
  }
  function sourceError(id, e) {
    if (!e.error) { show(id, false); return; }
    $(id).innerHTML = "<b>⚠ " + esc(e.error) + "</b>" + (e.enableUrl ? "<br><a class='enable' href='" + esc(e.enableUrl) + "' target='_blank' rel='noopener'>Enable it ↗</a>" : "");
    show(id, true);
  }
  function renderGuide(g) {
    var n = 0;
    $("guideSteps").innerHTML = (g || []).map(function (st) {
      return "<div class='stage'>" + esc(st.stage) + "</div>" + st.steps.map(function (s, k) {
        n++; return "<div class='gstep" + (k === st.steps.length - 1 ? " last" : "") + "'><div class='gnum'>" + n + "</div><div class='gtxt'><b>" + esc(s.title) + "</b><div>" + esc(s.text) + "</div>" +
          (s.link ? "<a href='" + esc(s.link) + "' target='_blank' rel='noopener'>" + esc(s.linkText || "Open") + " ↗</a>" : "") + "</div></div>";
      }).join("");
    }).join("");
  }
  fillPickers(); renderPeriod(); refresh(); fillSchedulePickers(); schedRefresh();
  google.script.run.withSuccessHandler(renderSchedule).rptScheduleInfo();
  google.script.run.withSuccessHandler(function (i) {
    info = i; var sv = i.saved || {};
    renderGuide(i.guide);
    $("gsc").innerHTML = '<option value="">— Don\\'t include Search Console —</option>' + i.gsc.sites.map(function (s) { return '<option value="' + esc(s) + '">' + esc(s) + "</option>"; }).join("");
    $("ga").innerHTML = '<option value="">— Don\\'t include GA4 —</option>' + i.ga.props.map(function (p) {
      return '<option value="' + esc(p.id) + '" data-name="' + esc(p.name) + '">' + esc(p.name) + " (" + esc(p.id) + ")" + (p.account ? " · " + esc(p.account) : "") + "</option>"; }).join("") +
      '<option value="__manual__">Enter a property ID…</option>';
    sourceError("gscErr", i.gsc); sourceError("gaErr", i.ga);
    if (i.ga.error) $("guide").open = true;
    if (sv.gscSite && i.gsc.sites.indexOf(sv.gscSite) !== -1) $("gsc").value = sv.gscSite;
    else if (i.gsc.sites.length === 1) $("gsc").value = i.gsc.sites[0];
    if (sv.gaProp) {
      if (i.ga.props.some(function (p) { return p.id === sv.gaProp; })) $("ga").value = sv.gaProp;
      else { $("ga").value = "__manual__"; $("gaId").value = sv.gaProp; }
      gaTouched = true;
    } else suggestGa();
    if (sv.keywords != null) $("keywords").value = String(sv.keywords);
    if (sv.traffic) $("traffic").value = sv.traffic;
    if (sv.gaMatch) $("gaMatch").value = sv.gaMatch;
    if (sv.maxPages != null) $("maxPages").value = String(sv.maxPages);
    if (sv.compare) $("compare").checked = true;
    if (sv.compareMode) $("compareMode").value = sv.compareMode;
    if (sv.preset && document.querySelector("#chips button[data-p='" + sv.preset + "']")) preset = sv.preset;
    fillTabs(i.tabs);
    if (i.selection && i.selection.count) $("selInfo").textContent = i.selection.count + " URL" + (i.selection.count === 1 ? "" : "s") + " in the selection (" + i.selection.a1 + ")";
    renderPeriod(); refresh();
  }).withFailureHandler(function (e) { $("gscErr").textContent = e.message || String(e); show("gscErr", true); $("guide").open = true; }).rptGetInfo();
</script></body></html>`;
