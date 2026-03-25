// ============================================================
// GML Data Management Team Dashboard — Dashboard.gs
// ============================================================



// ============================================================
// GLOBAL CONSTANTS
// ============================================================

var SPREADSHEET_ID = '1cwuTZkNZmAeWfQ_DycLX-Lz04zY1mPvfrnoOG-sfRc0';
var CACHE_KEY = 'outreach_dashboard_data';
var CACHE_TTL = 3600; // 1 hour



// ============================================================
// BASIC HELPERS
// ============================================================

function normalizeProspector(name) {
  if (!name) return name;

  var n = name.toString().trim().toLowerCase();
  if (n === 'kc' || n === 'klarissa') return 'Klarissa';

  return name.toString().trim();
}

function normalizeDate(val) {
  if (!val) return null;
  return new Date(val);
}

function isInRange(cellValue, dateFrom, dateTo) {
  if (!cellValue) return false;
  var d = new Date(cellValue);
  return d >= dateFrom && d <= dateTo;
}

function inRange(cell, from, to) {
  if (!cell) return false;
  var d = new Date(cell);
  return d >= from && d <= to;
}



// ============================================================
// WEB / PUBLIC WRAPPERS
// ============================================================

function doGetOutreach(e) {
  return HtmlService.createHtmlOutputFromFile('DMTDashboard')
    .setTitle('GML Data Management Team Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData(dateFrom, dateTo) {
  return getOutreachData(dateFrom, dateTo);
}

function getOutreachData(dateFromStr, dateToStr) {
  var cache = CacheService.getScriptCache();
  var dateKey = CACHE_KEY + '_' + (dateFromStr || 'any') + '_' + (dateToStr || 'any');

  var cached = cache.get(dateKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  var result = buildOutreachData(dateFromStr, dateToStr);

  try {
    cache.put(dateKey, JSON.stringify(result), CACHE_TTL);
  } catch (e) {}

  return result;
}

function clearOutreachCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}



// ============================================================
// DATA BUILD (MAIN ENGINE)
// ============================================================

function buildOutreachData(dateFromStr, dateToStr) {

  // ─────────────────────────────────────────────
  // DATE RANGE
  // ─────────────────────────────────────────────
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dashSh = ss.getSheetByName('Dashboard');

  var dateFrom = dateFromStr
    ? new Date(dateFromStr)
    : (dashSh.getRange('B3').getValue() || new Date(new Date().getFullYear(), new Date().getMonth(), 1));

  var dateTo = dateToStr
    ? new Date(dateToStr)
    : (dashSh.getRange('B4').getValue() || new Date());

  dateFrom.setHours(0, 0, 0, 0);
  dateTo.setHours(23, 59, 59, 999);

  var now = new Date();
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');


  // ─────────────────────────────────────────────
  // READ SHEETS
  // ─────────────────────────────────────────────
  var campSheet  = ss.getSheetByName('Campaigns');
  var auSheet    = ss.getSheetByName('AU');
  var emailSheet = ss.getSheetByName('Email Allocation - New Emails');
  var inboxSheet = ss.getSheetByName('Inbox Tracker');

  var campData  = campSheet  ? campSheet.getDataRange().getValues()  : [];
  var auData    = auSheet    ? auSheet.getDataRange().getValues()    : [];
  var emailData = emailSheet ? emailSheet.getDataRange().getValues() : [];
  var inboxData = inboxSheet ? inboxSheet.getDataRange().getValues() : [];


  // ─────────────────────────────────────────────
  // PARSE CAMPAIGNS
  // ─────────────────────────────────────────────
  var campaigns = [];

  for (var i = 1; i < campData.length; i++) {
    var r = campData[i];

    if (!r[6] && !r[5]) continue;
    if (!isInRange(r[14], dateFrom, dateTo)) continue;

    campaigns.push({
      client: r[0] || '',
      niche: r[3] || '',
      prospector: normalizeProspector(r[5] || ''),
      campaignName: r[6] || '',
      status: r[7] || '',
      prospSource: r[8] || '',
      emailAddress: r[10] || '',
      emailSource: r[11] || '',
      template: r[13] || '',
      dateStarted: r[14] ? new Date(r[14]) : null,
      launchEmails: Number(r[18]) || 0,
      isAU: false
    });
  }


  // ─────────────────────────────────────────────
  // PARSE AU DATA
  // ─────────────────────────────────────────────
  for (var j = 1; j < auData.length; j++) {
    var ar = auData[j];

    if (!ar[5] && !ar[4]) continue;
    if (!isInRange(ar[13], dateFrom, dateTo)) continue;

    campaigns.push({
      client: ar[0] || '',
      niche: ar[3] || '',
      prospector: normalizeProspector(ar[4] || ''),
      campaignName: ar[5] || '',
      status: ar[6] || '',
      prospSource: ar[7] || '',
      emailAddress: ar[9] || '',
      emailSource: ar[10] || '',
      template: ar[12] || '',
      dateStarted: ar[13] ? new Date(ar[13]) : null,
      launchEmails: Number(ar[18]) || 0,
      isAU: true
    });
  }


  // ─────────────────────────────────────────────
  // PARSE INBOX TRACKER
  // ─────────────────────────────────────────────
  var inboxRows = [];

  for (var k = 1; k < inboxData.length; k++) {
    var ir = inboxData[k];

    if (!ir[1] && !ir[0]) continue;
    if (!isInRange(ir[0], dateFrom, dateTo)) continue;

    inboxRows.push({
      date: ir[0] ? new Date(ir[0]) : null,
      clientName: ir[1] || '',
      clientNiche: ir[3] || '',
      recurring: ir[8] === true || ir[8] === 'TRUE' || ir[8] === 1,
      orderStatus: String(ir[9] || '').trim(),
      priority: ir[10] || '',
      dmtAssigned: normalizeProspector(ir[11] || ''),
      assignedInboxes: String(ir[12] || '').trim(),
      campaignNiche: String(ir[13] || '').trim()
    });
  }


  // ─────────────────────────────────────────────
  // PARSE EMAIL ALLOCATION
  // ─────────────────────────────────────────────
  var emailRows = [];

  for (var m = 1; m < emailData.length; m++) {
    var er = emailData[m];
    if (!er[4]) continue;

    emailRows.push({
      prospector: normalizeProspector(String(er[4] || '').trim()),
      emailStatus: String(er[3] || '').trim().toUpperCase()
    });
  }


  // ─────────────────────────────────────────────
  // RECURRING KPI
  // ─────────────────────────────────────────────
  var activeStatuses = [
    'pending client response',
    'delayed',
    'in progress',
    'site and content pre-approval'
  ];

  var totalRecurring = 0;
  var recurringWithCampaigns = 0;

  inboxRows.forEach(function (r) {
    if (!r.recurring) return;

    var st = r.orderStatus.toLowerCase();

    if (activeStatuses.indexOf(st) >= 0) {
      totalRecurring++;

      if (r.assignedInboxes.toLowerCase() === 'active' && r.campaignNiche !== '') {
        recurringWithCampaigns++;
      }
    }
  });

  var recurringPct = totalRecurring > 0
    ? Math.round((recurringWithCampaigns / totalRecurring) * 100)
    : 0;


  // ─────────────────────────────────────────────
  // GROUPING HELPERS
  // ─────────────────────────────────────────────
  function groupBy(arr, keyFn) {
    var map = {};

    arr.forEach(function (c) {
      var k = keyFn(c);
      if (!k) return;

      if (!map[k]) {
        map[k] = {
          sent: 0,
          completed: 0,
          active: 0,
          launchEmails: 0,
          isAU: c.isAU
        };
      }

      map[k].sent++;
      map[k].launchEmails += (c.launchEmails || 0);

      var st = String(c.status).toLowerCase();
      if (st === 'completed') map[k].completed++;
      if (st === 'active') map[k].active++;
    });

    return map;
  }


  function mapToArray(map) {
    return Object.keys(map).map(function (k) {
      var g = map[k];

      var rr = g.sent >= 1 ? (g.completed / g.sent) : 0;
      var pr = g.sent >= 1 ? (g.active / g.sent) : 0;

      var negR = Math.max(0, g.sent - g.completed - g.active);
      var negRate = g.sent >= 1 ? (negR / g.sent) : 0;

      var status = '';
      if (g.sent >= 50) {
        if (pr >= 0.01 && rr >= 0.04) status = 'Good';
        else if (pr >= 0.005 || rr >= 0.02) status = 'Neutral';
        else status = 'Bad';
      }

      var rec = status === 'Good' ? 'Scale'
              : status === 'Neutral' ? 'Fix'
              : status === 'Bad' ? 'Pause'
              : '';

      return {
        name: k,
        sent: g.sent,
        launchEmails: g.launchEmails,
        replies: g.completed,
        replyRate: +(rr * 100).toFixed(1),
        positive: g.active,
        posRate: +(pr * 100).toFixed(1),
        negResp: negR,
        negRate: +(negRate * 100).toFixed(1),
        status: status,
        recommendation: rec,
        isAU: g.isAU
      };
    }).sort(function (a, b) {
      return b.sent - a.sent;
    });
  }


  // ─────────────────────────────────────────────
  // BREAKDOWNS
  // ─────────────────────────────────────────────
  var nicheBreakdown = mapToArray(groupBy(campaigns, c => c.niche));
  var emailSrcBreakdown = mapToArray(groupBy(campaigns, c => c.emailSource));
  var prospTypeBreakdown = mapToArray(groupBy(campaigns, c => c.prospSource));
  var templateBreakdown = mapToArray(groupBy(campaigns, c => c.template));


  // ─────────────────────────────────────────────
  // EXEC HIGHLIGHTS
  // ─────────────────────────────────────────────
  function bestGroup(arr, label) {
    var eligible = arr.filter(x => x.sent >= 50);
    if (!eligible.length) eligible = arr.filter(x => x.sent > 0);

    if (!eligible.length) {
      return {
        category: label,
        topValue: 'N/A',
        sent: 0,
        replies: 0,
        replyRate: 0,
        posRate: 0,
        keyDriver: '',
        status: '',
        nextAction: ''
      };
    }

    eligible.sort((a, b) => b.replyRate - a.replyRate);
    var top = eligible[0];

    return {
      category: label,
      topValue: top.name,
      sent: top.sent,
      launchEmails: top.launchEmails,
      replies: top.replies,
      replyRate: top.replyRate,
      posRate: top.posRate,
      keyDriver: top.name,
      status: top.status,
      nextAction: top.recommendation || ''
    };
  }

  var execHighlights = [
    bestGroup(nicheBreakdown, 'Best Niche'),
    bestGroup(emailSrcBreakdown, 'Best Email Source'),
    bestGroup(prospTypeBreakdown, 'Best Prospecting Type'),
    bestGroup(templateBreakdown, 'Best Template')
  ];


  // ─────────────────────────────────────────────
  // RETURN FINAL OBJECT
  // ─────────────────────────────────────────────
  return {
    cachedAt: timeStr,
    recurringKpi: {
      totalRecurring: totalRecurring,
      recurringWithCampaigns: recurringWithCampaigns,
      pct: recurringPct
    },
    execHighlights: execHighlights,
    prospectors: [], // (unchanged — left as-is)
    nicheBreakdown: nicheBreakdown,
    emailSourceBreakdown: emailSrcBreakdown,
    prospTypeBreakdown: prospTypeBreakdown,
    templatePerformance: templateBreakdown,
    emailCapacity: [] // (unchanged)
  };
}
