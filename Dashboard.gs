// ============================================================
// GML Data Management Team Dashboard — Dashboard.gs
// ============================================================



// ============================================================
// GLOBAL CONSTANTS
// ============================================================

var SPREADSHEET_ID = '1cwuTZkNZmAeWfQ_DycLX-Lz04zY1mPvfrnoOG-sfRc0';
var CACHE_KEY = 'outreach_dashboard_data_v3';
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
// GML Data Management Team Dashboard — Dashboard.gs
// ============================================================

// ============================================================
// GLOBAL CONSTANTS
// ============================================================
var SPREADSHEET_ID = '1cwuTZkNZmAeWfQ_DycLX-Lz04zY1mPvfrnoOG-sfRc0';
var CACHE_PREFIX = 'outreach_dashboard_v2';
var CACHE_TTL = 3600; // 1 hour

// ============================================================
// BASIC HELPERS
// ============================================================
function normalizeProspector(name) {
  if (!name) return '';
  var n = name.toString().trim().toLowerCase();

  if (n === 'kc' || n === 'klarissa') return 'Klarissa';

  return name.toString().trim();
}

function normalizeDate(val) {
  if (!val) return null;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function isInRange(cellValue, dateFrom, dateTo) {
  if (!cellValue) return false;
  var d = new Date(cellValue);
  if (isNaN(d.getTime())) return false;
  return d >= dateFrom && d <= dateTo;
}

function safeStr(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function lc(v) {
  return safeStr(v).toLowerCase();
}

function pct(numerator, denominator) {
  if (!denominator) return 0;
  return +(numerator / denominator * 100).toFixed(1);
}

function getStatusAndRecommendation(sent, replyRatePct, posRatePct) {
  var status = '';
  var recommendation = '';

  if (sent >= 50) {
    if (posRatePct >= 1 && replyRatePct >= 4) status = 'Good';
    else if (posRatePct >= 0.5 || replyRatePct >= 2) status = 'Neutral';
    else status = 'Bad';
  }

  recommendation =
    status === 'Good' ? 'Scale' :
    status === 'Neutral' ? 'Fix' :
    status === 'Bad' ? 'Pause' : '';

  return {
    status: status,
    recommendation: recommendation
  };
}

function topValueFromRows(rows, fieldName) {
  var counts = {};

  rows.forEach(function (r) {
    var val = safeStr(r[fieldName]);
    if (!val) return;
    counts[val] = (counts[val] || 0) + 1;
  });

  var top = '';
  var max = 0;

  Object.keys(counts).forEach(function (k) {
    if (counts[k] > max) {
      max = counts[k];
      top = k;
    }
  });

  return top;
}

function makeMetricRow(name, rows) {
  var sent = rows.length;
  var launchEmails = 0;
  var replies = 0;
  var positive = 0;

  rows.forEach(function (r) {
    launchEmails += Number(r.launchEmails) || 0;

    var st = lc(r.status);
    if (st === 'completed') replies++;
    if (st === 'active') positive++;
  });

  var negResp = Math.max(0, sent - replies - positive);
  var replyRate = pct(replies, sent);
  var posRate = pct(positive, sent);
  var negRate = pct(negResp, sent);

  var sr = getStatusAndRecommendation(sent, replyRate, posRate);

  return {
    name: name,
    sent: sent,
    launchEmails: launchEmails,
    replies: replies,
    replyRate: replyRate,
    positive: positive,
    posRate: posRate,
    negResp: negResp,
    negRate: negRate,
    status: sr.status,
    recommendation: sr.recommendation
  };
}

function groupRows(campaigns, keyFn) {
  var map = {};

  campaigns.forEach(function (c) {
    var key = keyFn(c);
    key = safeStr(key);
    if (!key) return;

    if (!map[key]) map[key] = [];
    map[key].push(c);
  });

  return map;
}

function groupedMetricArray(campaigns, keyFn, extraBuilder) {
  var map = groupRows(campaigns, keyFn);

  var arr = Object.keys(map).map(function (k) {
    var base = makeMetricRow(k, map[k]);
    if (typeof extraBuilder === 'function') {
      var extra = extraBuilder(k, map[k]) || {};
      for (var prop in extra) base[prop] = extra[prop];
    }
    return base;
  });

  arr.sort(function (a, b) {
    return b.sent - a.sent;
  });

  return arr;
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
  var dateKey = CACHE_PREFIX + '_' + (dateFromStr || 'any') + '_' + (dateToStr || 'any');

  var cached = cache.get(dateKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  var result = buildOutreachData(dateFromStr, dateToStr);

  try {
    cache.put(dateKey, JSON.stringify(result), CACHE_TTL);
  } catch (e) {}

  return result;
}

function getProspectorBreakdown(dateFrom, dateTo) {
  var data = getOutreachData(dateFrom, dateTo);

  var breakdown = (data._rawCampaignsByProspector || []).map(function (p) {
    var rows = p.rows || [];

    var totalSent = rows.length;
    var totalLaunchEmails = 0;
    var totalReplies = 0;

    rows.forEach(function (r) {
      totalLaunchEmails += Number(r.launchEmails) || 0;
      if (lc(r.status) === 'completed') totalReplies++;
    });

    var emailMap = {};
    rows.forEach(function (r) {
      var email = safeStr(r.emailAddress);
      if (!email) return;

      if (!emailMap[email]) {
        emailMap[email] = {
          emailAddress: email,
          timesUsed: 0,
          nicheSet: {},
          replies: 0,
          sent: 0
        };
      }

      emailMap[email].timesUsed++;
      emailMap[email].sent++;
      if (safeStr(r.niche)) emailMap[email].nicheSet[r.niche] = true;
      if (lc(r.status) === 'completed') emailMap[email].replies++;
    });

    var emailRows = Object.keys(emailMap).map(function (email) {
      var e = emailMap[email];
      return {
        emailAddress: e.emailAddress,
        timesUsed: e.timesUsed,
        niches: Object.keys(e.nicheSet).sort().join(', '),
        replyRate: pct(e.replies, e.sent),
        overused: e.timesUsed >= 3
      };
    }).sort(function (a, b) {
      return b.timesUsed - a.timesUsed;
    });

    var nicheMap = {};
    rows.forEach(function (r) {
      var niche = safeStr(r.niche);
      if (!niche) return;

      if (!nicheMap[niche]) {
        nicheMap[niche] = [];
      }
      nicheMap[niche].push(r);
    });

    var nicheRows = Object.keys(nicheMap).map(function (niche) {
      var metric = makeMetricRow(niche, nicheMap[niche]);
      return {
        niche: niche,
        sent: metric.sent,
        launchEmails: metric.launchEmails,
        replies: metric.replies,
        replyRate: metric.replyRate,
        positive: metric.positive,
        negative: metric.negResp
      };
    }).sort(function (a, b) {
      return b.sent - a.sent;
    });

    return {
      name: p.name,
      totalSent: totalSent,
      totalLaunchEmails: totalLaunchEmails,
      totalReplies: totalReplies,
      overallReplyRate: pct(totalReplies, totalSent),
      uniqueEmails: emailRows.length,
      nicheCount: nicheRows.length,
      emailRows: emailRows,
      nicheRows: nicheRows
    };
  });

  return {
    breakdown: breakdown
  };
}

function clearOutreachCache(dateFromStr, dateToStr) {
  var cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY + '_' + (dateFromStr || 'any') + '_' + (dateToStr || 'any'));
}

// ============================================================
// MAIN ENGINE
// ============================================================
function buildOutreachData(dateFromStr, dateToStr) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dashSh = ss.getSheetByName('Dashboard');

  var dateFrom = dateFromStr
    ? new Date(dateFromStr)
    : (dashSh && dashSh.getRange('B3').getValue()
        ? new Date(dashSh.getRange('B3').getValue())
        : new Date(new Date().getFullYear(), new Date().getMonth(), 1));

  var dateTo = dateToStr
    ? new Date(dateToStr)
    : (dashSh && dashSh.getRange('B4').getValue()
        ? new Date(dashSh.getRange('B4').getValue())
        : new Date());

  dateFrom.setHours(0, 0, 0, 0);
  dateTo.setHours(23, 59, 59, 999);

  var now = new Date();
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');

  var campSheet = ss.getSheetByName('Campaigns');
  var auSheet = ss.getSheetByName('AU');
  var emailSheet = ss.getSheetByName('Email Allocation - New Emails');
  var inboxSheet = ss.getSheetByName('Inbox Tracker');

  var campData = campSheet ? campSheet.getDataRange().getValues() : [];
  var auData = auSheet ? auSheet.getDataRange().getValues() : [];
  var emailData = emailSheet ? emailSheet.getDataRange().getValues() : [];
  var inboxData = inboxSheet ? inboxSheet.getDataRange().getValues() : [];

  var campaigns = [];

  // Campaigns sheet
  for (var i = 1; i < campData.length; i++) {
    var r = campData[i];

    if (!r[6] && !r[5]) continue;
    if (!isInRange(r[14], dateFrom, dateTo)) continue;

    campaigns.push({
      client: safeStr(r[0]),
      niche: safeStr(r[3]),
      prospector: normalizeProspector(r[5]),
      campaignName: safeStr(r[6]),
      status: safeStr(r[7]),
      prospSource: safeStr(r[8]),
      emailAddress: safeStr(r[10]),
      emailSource: safeStr(r[11]),
      template: safeStr(r[13]),
      dateStarted: normalizeDate(r[14]),
      launchEmails: Number(r[18]) || 0,
      isAU: false
    });
  }

  // AU sheet
  for (var j = 1; j < auData.length; j++) {
    var ar = auData[j];

    if (!ar[5] && !ar[4]) continue;
    if (!isInRange(ar[13], dateFrom, dateTo)) continue;

    campaigns.push({
      client: safeStr(ar[0]),
      niche: safeStr(ar[3]),
      prospector: normalizeProspector(ar[4]),
      campaignName: safeStr(ar[5]),
      status: safeStr(ar[6]),
      prospSource: safeStr(ar[7]),
      emailAddress: safeStr(ar[9]),
      emailSource: safeStr(ar[10]),
      template: safeStr(ar[12]),
      dateStarted: normalizeDate(ar[13]),
      launchEmails: Number(ar[18]) || 0,
      isAU: true
    });
  }

  var inboxRows = [];
  for (var k = 1; k < inboxData.length; k++) {
    var ir = inboxData[k];

    if (!ir[1] && !ir[0]) continue;
    if (!isInRange(ir[0], dateFrom, dateTo)) continue;

    inboxRows.push({
      date: normalizeDate(ir[0]),
      clientName: safeStr(ir[1]),
      clientNiche: safeStr(ir[3]),
      recurring: ir[8] === true || ir[8] === 'TRUE' || ir[8] === 1,
      orderStatus: safeStr(ir[9]),
      priority: safeStr(ir[10]),
      dmtAssigned: normalizeProspector(ir[11]),
      assignedInboxes: safeStr(ir[12]),
      campaignNiche: safeStr(ir[13])
    });
  }

  var emailRows = [];
  for (var m = 1; m < emailData.length; m++) {
    var er = emailData[m];
    if (!er[4]) continue;

    emailRows.push({
      prospector: normalizeProspector(er[4]),
      emailStatus: safeStr(er[3]).toUpperCase()
    });
  }

  // ==========================================================
  // RECURRING KPI
  // ==========================================================
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

    var st = lc(r.orderStatus);
    if (activeStatuses.indexOf(st) >= 0) {
      totalRecurring++;

      if (lc(r.assignedInboxes) === 'active' && safeStr(r.campaignNiche) !== '') {
        recurringWithCampaigns++;
      }
    }
  });

  var recurringPct = totalRecurring > 0
    ? Math.round((recurringWithCampaigns / totalRecurring) * 100)
    : 0;

  // ==========================================================
  // BREAKDOWNS
  // ==========================================================
  var nicheBreakdown = groupedMetricArray(campaigns, function (c) {
    return c.niche;
  }, function (k, rows) {
    return {
      isAU: rows.some(function (r) { return r.isAU; })
    };
  });

  var emailSrcBreakdown = groupedMetricArray(campaigns, function (c) {
    return c.emailSource;
  });

  var prospTypeBreakdown = groupedMetricArray(campaigns, function (c) {
    return c.prospSource;
  });

  // ==========================================================
  // TEMPLATE PERFORMANCE (ENRICHED)
  // ==========================================================
  var templatePerformance = groupedMetricArray(campaigns, function (c) {
    return c.template;
  }, function (k, rows) {
    return {
      topNiche: topValueFromRows(rows, 'niche'),
      topEmailSrc: topValueFromRows(rows, 'emailSource'),
      topProsType: topValueFromRows(rows, 'prospSource')
    };
  });

  // ==========================================================
  // PROSPECTORS (ENRICHED)
  // ==========================================================
  var prospectors = groupedMetricArray(campaigns, function (c) {
    return c.prospector;
  }, function (k, rows) {
    return {
      bestNiche: topValueFromRows(rows, 'niche'),
      bestEmailSrc: topValueFromRows(rows, 'emailSource'),
      bestProsType: topValueFromRows(rows, 'prospSource'),
      bestTemplate: topValueFromRows(rows, 'template')
    };
  }).map(function (r) {
    return {
      name: r.name,
      bestNiche: r.bestNiche,
      bestEmailSrc: r.bestEmailSrc,
      bestProsType: r.bestProsType,
      bestTemplate: r.bestTemplate,
      sent: r.sent,
      launchEmails: r.launchEmails,
      replyRate: r.replyRate,
      posRate: r.posRate,
      status: r.status,
      nextAction: r.recommendation
    };
  });

  // ==========================================================
  // EXEC HIGHLIGHTS
  // ==========================================================
  function bestGroup(arr, label) {
    var eligible = arr.filter(function (x) { return x.sent >= 50; });
    if (!eligible.length) eligible = arr.filter(function (x) { return x.sent > 0; });

    if (!eligible.length) {
      return {
        category: label,
        topValue: 'N/A',
        sent: 0,
        launchEmails: 0,
        replies: 0,
        replyRate: 0,
        posRate: 0,
        keyDriver: '',
        status: '',
        nextAction: ''
      };
    }

    eligible.sort(function (a, b) {
      return b.replyRate - a.replyRate;
    });

    var top = eligible[0];

    return {
      category: label,
      topValue: top.name,
      sent: top.sent,
      launchEmails: top.launchEmails || 0,
      replies: top.replies || 0,
      replyRate: top.replyRate || 0,
      posRate: top.posRate || 0,
      keyDriver: top.name,
      status: top.status || '',
      nextAction: top.recommendation || top.nextAction || ''
    };
  }

  var execHighlights = [
    bestGroup(nicheBreakdown, 'Best Niche'),
    bestGroup(emailSrcBreakdown, 'Best Email Source'),
    bestGroup(prospTypeBreakdown, 'Best Prospecting Type'),
    bestGroup(templatePerformance, 'Best Template')
  ];

  // ==========================================================
  // EMAIL CAPACITY
  // ==========================================================
  var emailCapacityMap = {};

  emailRows.forEach(function (e) {
    var p = safeStr(e.prospector);
    if (!p) return;

    if (!emailCapacityMap[p]) {
      emailCapacityMap[p] = {
        prospector: p,
        total: 0,
        active: 0,
        resting: 0,
        other: 0
      };
    }

    emailCapacityMap[p].total++;

    var st = safeStr(e.emailStatus).toUpperCase();
    if (st === 'ACTIVE') emailCapacityMap[p].active++;
    else if (st === 'RESTING') emailCapacityMap[p].resting++;
    else emailCapacityMap[p].other++;
  });

  var emailCapacity = Object.keys(emailCapacityMap).map(function (p) {
    var g = emailCapacityMap[p];
    return {
      prospector: g.prospector,
      total: g.total,
      active: g.active,
      resting: g.resting,
      other: g.other,
      capacityPct: g.total ? Math.round((g.active / g.total) * 100) : 0
    };
  }).sort(function (a, b) {
    return a.prospector.localeCompare(b.prospector);
  });

  // Raw rows for breakdown page
  var rawCampaignsByProspector = Object.keys(groupRows(campaigns, function (c) {
    return c.prospector;
  })).map(function (name) {
    return {
      name: name,
      rows: groupRows(campaigns, function (c) { return c.prospector; })[name]
    };
  }).sort(function (a, b) {
    return b.rows.length - a.rows.length;
  });

  return {
    cachedAt: timeStr,
    recurringKpi: {
      totalRecurring: totalRecurring,
      recurringWithCampaigns: recurringWithCampaigns,
      pct: recurringPct
    },
    execHighlights: execHighlights,
    prospectors: prospectors,
    nicheBreakdown: nicheBreakdown,
    emailSourceBreakdown: emailSrcBreakdown,
    prospTypeBreakdown: prospTypeBreakdown,
    templatePerformance: templatePerformance,
    emailCapacity: emailCapacity,

    // internal for breakdown page
    _rawCampaignsByProspector: rawCampaignsByProspector
  };
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
   var nicheBreakdown = mapToArray(groupBy(campaigns, function(c) { return c.niche; }));
  var emailSrcBreakdown = mapToArray(groupBy(campaigns, function(c) { return c.emailSource; }));
  var prospTypeBreakdown = mapToArray(groupBy(campaigns, function(c) { return c.prospSource; }));
  var templateBreakdown = mapToArray(groupBy(campaigns, function(c) { return c.template; }));

  // Enrich templateBreakdown with top niche / email source / prosp source
  var tmplRowsMap = {};
  campaigns.forEach(function(c) {
    if (!c.template) return;
    if (!tmplRowsMap[c.template]) tmplRowsMap[c.template] = [];
    tmplRowsMap[c.template].push(c);
  });
  templateBreakdown.forEach(function(t) {
    var rows = tmplRowsMap[t.name] || [];
    t.topNiche   = topVal(rows, 'niche');
    t.topEmailSrc = topVal(rows, 'emailSource');
    t.topProsType = topVal(rows, 'prospSource');
  });



  // ─────────────────────────────────────────────
  // PROSPECTORS
  // ─────────────────────────────────────────────
  function topVal(rows, field) {
    var counts = {};
    rows.forEach(function(r) {
      var v = String(r[field] || '').trim();
      if (v) counts[v] = (counts[v] || 0) + 1;
    });
    var top = '', max = 0;
    Object.keys(counts).forEach(function(k) {
      if (counts[k] > max) { max = counts[k]; top = k; }
    });
    return top;
  }

  var prospMap = {};
  campaigns.forEach(function(c) {
    var k = c.prospector;
    if (!k) return;
    if (!prospMap[k]) prospMap[k] = { rows: [], sent: 0, completed: 0, active: 0, launchEmails: 0 };
    prospMap[k].rows.push(c);
    prospMap[k].sent++;
    prospMap[k].launchEmails += (c.launchEmails || 0);
    var st = String(c.status).toLowerCase();
    if (st === 'completed') prospMap[k].completed++;
    if (st === 'active') prospMap[k].active++;
  });

  var prospectors = Object.keys(prospMap).map(function(k) {
    var g = prospMap[k];
    var rr = g.sent >= 1 ? (g.completed / g.sent) : 0;
    var pr = g.sent >= 1 ? (g.active / g.sent) : 0;
    var status = '';
    if (g.sent >= 50) {
      if (pr >= 0.01 && rr >= 0.04) status = 'Good';
      else if (pr >= 0.005 || rr >= 0.02) status = 'Neutral';
      else status = 'Bad';
    }
    var rec = status === 'Good' ? 'Scale' : status === 'Neutral' ? 'Fix' : status === 'Bad' ? 'Pause' : '';
    return {
      name: k,
      bestNiche: topVal(g.rows, 'niche'),
      bestEmailSrc: topVal(g.rows, 'emailSource'),
      bestProsType: topVal(g.rows, 'prospSource'),
      bestTemplate: topVal(g.rows, 'template'),
      sent: g.sent,
      launchEmails: g.launchEmails,
      replyRate: +(rr * 100).toFixed(1),
      posRate: +(pr * 100).toFixed(1),
      status: status,
      nextAction: rec
    };
  }).sort(function(a, b) { return b.sent - a.sent; });


  // ─────────────────────────────────────────────
  // BEST KPI HIGHLIGHTS
  // ─────────────────────────────────────────────
  var bestNiche = (nicheBreakdown || [])
    .filter(function(n) { return n.sent >= 50; })
    .sort(function(a, b) { return b.replyRate - a.replyRate; })[0] || null;

  var bestProsType = (prospTypeBreakdown || [])
    .filter(function(n) { return n.sent >= 50; })
    .sort(function(a, b) { return b.replyRate - a.replyRate; })[0] || null;

  var topProspector = (prospectors || [])
    .filter(function(p) { return p.sent >= 1; })
    .sort(function(a, b) { return b.replyRate - a.replyRate; })[0] || null;


  // ─────────────────────────────────────────────
  // RAW CAMPAIGNS BY PROSPECTOR (for drill-down)
  // ─────────────────────────────────────────────
  var rawCampaignsByProspector = Object.keys(prospMap).map(function(name) {
    return { name: name, rows: prospMap[name].rows };
  }).sort(function(a, b) { return b.rows.length - a.rows.length; });


  // ─────────────────────────────────────────────
  // EXEC HIGHLIGHTS
  // ─────────────────────────────────────────────
  function bestGroup(arr, label) {
    var eligible = arr.filter(function(x) { return x.sent >= 50; });
    if (!eligible.length) eligible = arr.filter(function(x) { return x.sent > 0; });

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

    eligible.sort(function(a, b) { return b.replyRate - a.replyRate; });
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
    prospectors: prospectors,
    nicheBreakdown: nicheBreakdown,
    emailSourceBreakdown: emailSrcBreakdown,
    prospTypeBreakdown: prospTypeBreakdown,
    templatePerformance: templateBreakdown,
    emailCapacity: [],
    bestNiche: bestNiche ? { name: bestNiche.name, replyRate: bestNiche.replyRate } : null,
    bestProsType: bestProsType ? { name: bestProsType.name, replyRate: bestProsType.replyRate } : null,
    topProspector: topProspector ? { name: topProspector.name, replyRate: topProspector.replyRate } : null,
    _rawCampaignsByProspector: rawCampaignsByProspector
  };
}
