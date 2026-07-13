// ============================================================
// FIFA World Cup 2026 Prediction Game — Google Apps Script
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ============================================================

var ADMIN_PIN = '2026'; // Change this to your preferred PIN

// ── CORS & ROUTING ──────────────────────────────────────────

function doGet(e) {
  var action = e.parameter.action || '';
  var result;
  try {
    if (action === 'getPredictions') result = getPredictions();
    else if (action === 'getScores')  result = getScores();
    else if (action === 'ping')       result = { status: 'ok', time: new Date().toISOString() };
    else result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return respond(result);
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return respond({ error: 'Invalid JSON' }); }

  var result;
  try {
    if (data.action === 'savePrediction') {
      result = savePrediction(data);
    } else if (data.action === 'saveScore') {
      if (String(data.pin) !== String(ADMIN_PIN)) return respond({ error: 'Invalid PIN' });
      result = saveScore(data);
    } else if (data.action === 'fetchLiveScores') {
      if (String(data.pin) !== String(ADMIN_PIN)) return respond({ error: 'Invalid PIN' });
      result = fetchLiveScores();
    } else {
      result = { error: 'Unknown action: ' + data.action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }
  return respond(result);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ────────────────────────────────────────────

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function upsert(sheet, keyCol, keyVal, rowData) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var keyIdx = headers.indexOf(keyCol);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]) === String(keyVal)) {
      // Update in place
      var values = headers.map(function(h) {
        return rowData.hasOwnProperty(h) ? rowData[h] : data[i][headers.indexOf(h)];
      });
      sheet.getRange(i + 1, 1, 1, values.length).setValues([values]);
      return;
    }
  }
  // Insert new row
  sheet.appendRow(headers.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ''; }));
}

// ── PREDICTIONS ──────────────────────────────────────────────

function getPredictions() {
  var sheet = getSheet('Predictions', ['CompositeKey','MatchID','PlayerName','HomeGoals','AwayGoals','UpdatedAt']);
  return sheetToObjects(sheet);
}

function savePrediction(data) {
  var sheet = getSheet('Predictions', ['CompositeKey','MatchID','PlayerName','HomeGoals','AwayGoals','UpdatedAt']);
  var key = data.matchId + '_' + data.playerName;
  upsert(sheet, 'CompositeKey', key, {
    CompositeKey: key,
    MatchID:      data.matchId,
    PlayerName:   data.playerName,
    HomeGoals:    data.home,
    AwayGoals:    data.away,
    UpdatedAt:    new Date().toISOString()
  });
  return { success: true };
}

// ── SCORES ───────────────────────────────────────────────────

function getScores() {
  var sheet = getSheet('Scores', ['MatchID','HomeGoals','AwayGoals','Status','UpdatedAt']);
  return sheetToObjects(sheet);
}

function saveScore(data) {
  var sheet = getSheet('Scores', ['MatchID','HomeGoals','AwayGoals','Status','UpdatedAt']);
  upsert(sheet, 'MatchID', data.matchId, {
    MatchID:    data.matchId,
    HomeGoals:  data.home,
    AwayGoals:  data.away,
    Status:     data.status || 'FINISHED',
    UpdatedAt:  new Date().toISOString()
  });
  return { success: true };
}

// ── LIVE SCORE FETCH (football-data.org) ─────────────────────

var FOOTBALL_API_KEY = '948288f3b75442729d3885e14d1ec42e'; // Set your football-data.org API key here

// Team name mapping: football-data.org names → our names
var TEAM_MAP = {
  'Mexico':              'Mexico',
  'South Africa':        'South Africa',
  'Korea Republic':      'South Korea',
  'Czechia':             'Czechia',
  'Czech Republic':      'Czechia',
  'Canada':              'Canada',
  'Bosnia and Herzegovina': 'Bosnia & Herzegovina',
  'Qatar':               'Qatar',
  'Switzerland':         'Switzerland',
  'Brazil':              'Brazil',
  'Morocco':             'Morocco',
  'Haiti':               'Haiti',
  'Scotland':            'Scotland',
  'USA':                 'USA',
  'Paraguay':            'Paraguay',
  'Australia':           'Australia',
  'Turkey':              'Turkey',
  'Türkiye':             'Turkey',
  'Germany':             'Germany',
  'Curaçao':             'Curaçao',
  "Côte d'Ivoire":       'Ivory Coast',
  'Ecuador':             'Ecuador',
  'Netherlands':         'Netherlands',
  'Japan':               'Japan',
  'Sweden':              'Sweden',
  'Tunisia':             'Tunisia',
  'Spain':               'Spain',
  'Cabo Verde':          'Cape Verde',
  'Saudi Arabia':        'Saudi Arabia',
  'Uruguay':             'Uruguay',
  'Belgium':             'Belgium',
  'Egypt':               'Egypt',
  'Iran':                'Iran',
  'New Zealand':         'New Zealand',
  'France':              'France',
  'Senegal':             'Senegal',
  'Iraq':                'Iraq',
  'Norway':              'Norway',
  'Argentina':           'Argentina',
  'Algeria':             'Algeria',
  'Austria':             'Austria',
  'Jordan':              'Jordan',
  'Portugal':            'Portugal',
  'DR Congo':            'DR Congo',
  'Congo DR':            'DR Congo',
  'Democratic Republic of Congo': 'DR Congo',
  'Uzbekistan':          'Uzbekistan',
  'Colombia':            'Colombia',
  'England':             'England',
  'Croatia':             'Croatia',
  'Ghana':               'Ghana',
  'Panama':              'Panama'
};

// Our match schedule: matchId → {home, away} for lookup
var MATCH_LOOKUP = {
  1:  {home:'Mexico',      away:'South Africa'},
  2:  {home:'South Korea', away:'Czechia'},
  3:  {home:'Canada',      away:'Bosnia & Herzegovina'},
  4:  {home:'USA',         away:'Paraguay'},
  5:  {home:'Qatar',       away:'Switzerland'},
  6:  {home:'Brazil',      away:'Morocco'},
  7:  {home:'Haiti',       away:'Scotland'},
  8:  {home:'Australia',   away:'Turkey'},
  9:  {home:'Germany',     away:'Curaçao'},
  10: {home:'Netherlands', away:'Japan'},
  11: {home:'Ivory Coast', away:'Ecuador'},
  12: {home:'Sweden',      away:'Tunisia'},
  13: {home:'Spain',       away:'Cape Verde'},
  14: {home:'Belgium',     away:'Egypt'},
  15: {home:'Saudi Arabia',away:'Uruguay'},
  16: {home:'Iran',        away:'New Zealand'},
  17: {home:'France',      away:'Senegal'},
  18: {home:'Iraq',        away:'Norway'},
  19: {home:'Argentina',   away:'Algeria'},
  20: {home:'Austria',     away:'Jordan'},
  21: {home:'Portugal',    away:'DR Congo'},
  22: {home:'England',     away:'Croatia'},
  23: {home:'Ghana',       away:'Panama'},
  24: {home:'Uzbekistan',  away:'Colombia'},
  25: {home:'Czechia',     away:'South Africa'},
  26: {home:'Switzerland', away:'Bosnia & Herzegovina'},
  27: {home:'Canada',      away:'Qatar'},
  28: {home:'Mexico',      away:'South Korea'},
  29: {home:'USA',         away:'Australia'},
  30: {home:'Scotland',    away:'Morocco'},
  31: {home:'Brazil',      away:'Haiti'},
  32: {home:'Turkey',      away:'Paraguay'},
  33: {home:'Netherlands', away:'Sweden'},
  34: {home:'Germany',     away:'Ivory Coast'},
  35: {home:'Ecuador',     away:'Curaçao'},
  36: {home:'Tunisia',     away:'Japan'},
  37: {home:'Spain',       away:'Saudi Arabia'},
  38: {home:'Belgium',     away:'Iran'},
  39: {home:'Uruguay',     away:'Cape Verde'},
  40: {home:'New Zealand', away:'Egypt'},
  41: {home:'Argentina',   away:'Austria'},
  42: {home:'France',      away:'Iraq'},
  43: {home:'Norway',      away:'Senegal'},
  44: {home:'Jordan',      away:'Algeria'},
  45: {home:'Portugal',    away:'Uzbekistan'},
  46: {home:'England',     away:'Ghana'},
  47: {home:'Panama',      away:'Croatia'},
  48: {home:'Colombia',    away:'DR Congo'},
  49: {home:'Switzerland', away:'Canada'},
  50: {home:'Bosnia & Herzegovina', away:'Qatar'},
  51: {home:'Morocco',     away:'Haiti'},
  52: {home:'Scotland',    away:'Brazil'},
  53: {home:'South Africa',away:'South Korea'},
  54: {home:'Czechia',     away:'Mexico'},
  55: {home:'Curaçao',     away:'Ivory Coast'},
  56: {home:'Ecuador',     away:'Germany'},
  57: {home:'Tunisia',     away:'Netherlands'},
  58: {home:'Japan',       away:'Sweden'},
  59: {home:'Turkey',      away:'USA'},
  60: {home:'Paraguay',    away:'Australia'},
  61: {home:'Norway',      away:'France'},
  62: {home:'Senegal',     away:'Iraq'},
  63: {home:'Cape Verde',  away:'Saudi Arabia'},
  64: {home:'Uruguay',     away:'Spain'},
  65: {home:'New Zealand', away:'Belgium'},
  66: {home:'Egypt',       away:'Iran'},
  67: {home:'Panama',      away:'England'},
  68: {home:'Croatia',     away:'Ghana'},
  69: {home:'Colombia',    away:'Portugal'},
  70: {home:'DR Congo',    away:'Uzbekistan'},
  71: {home:'Algeria',     away:'Austria'},
  72: {home:'Jordan',      away:'Argentina'},
  // ── ROUND OF 32 ─────────────────────────────────────────────
  73: {home:'South Africa',        away:'Canada'},
  74: {home:'Brazil',              away:'Japan'},
  75: {home:'Germany',             away:'Paraguay'},
  76: {home:'Netherlands',         away:'Morocco'},
  77: {home:'Ivory Coast',         away:'Norway'},
  78: {home:'France',              away:'Sweden'},
  79: {home:'Mexico',              away:'Ecuador'},
  80: {home:'England',             away:'DR Congo'},
  81: {home:'Belgium',             away:'Senegal'},
  82: {home:'USA',                 away:'Bosnia & Herzegovina'},
  83: {home:'Spain',               away:'Austria'},
  84: {home:'Portugal',            away:'Croatia'},
  85: {home:'Switzerland',         away:'Algeria'},
  86: {home:'Australia',           away:'Egypt'},
  87: {home:'Argentina',           away:'Cape Verde'},
  88: {home:'Colombia',            away:'Ghana'},
  // ── ROUND OF 16 ─────────────────────────────────────────────
  89: {home:'Canada',      away:'Morocco'},
  90: {home:'Paraguay',    away:'France'},
  91: {home:'Brazil',      away:'Norway'},
  92: {home:'Mexico',      away:'England'},
  93: {home:'Portugal',    away:'Spain'},
  94: {home:'USA',         away:'Belgium'},
  95: {home:'Argentina',   away:'Egypt'},
  96: {home:'Switzerland', away:'Colombia'},
  // ── QUARTER-FINALS ──────────────────────────────────────────
  97:  {home:'France',     away:'Morocco'},
  98:  {home:'Spain',      away:'Belgium'},
  99:  {home:'Norway',     away:'England'},
  100: {home:'Argentina',  away:'Switzerland'},
  // ── SEMI-FINALS ─────────────────────────────────────────────
  101: {home:'France',  away:'Spain'},
  102: {home:'England', away:'Argentina'}
};

function fetchLiveScores() {
  if (!FOOTBALL_API_KEY) return { error: 'FOOTBALL_API_KEY not set in apps-script.gs' };
  try {
    var resp = UrlFetchApp.fetch(
      'https://api.football-data.org/v4/competitions/WC/matches?status=FINISHED',
      { headers: { 'X-Auth-Token': FOOTBALL_API_KEY } }
    );
    var body = resp.getContentText();
    var parsed = JSON.parse(body);
    var matches = parsed.matches || [];
    Logger.log('fetchLiveScores: HTTP ' + resp.getResponseCode() + ', matches=' + matches.length);
    if (parsed.error) Logger.log('API error: ' + parsed.error + ' | message: ' + parsed.message);
    var updated = 0;

    matches.forEach(function(m) {
      if (m.status !== 'FINISHED') return;
      var apiHome = TEAM_MAP[m.homeTeam.name] || m.homeTeam.name;
      var apiAway = TEAM_MAP[m.awayTeam.name] || m.awayTeam.name;
      Logger.log('Trying to match: ' + apiHome + ' vs ' + apiAway);

      // Find our matchId
      for (var id in MATCH_LOOKUP) {
        var ml = MATCH_LOOKUP[id];
        if (ml.home === apiHome && ml.away === apiAway) {
          saveScore({
            matchId: parseInt(id),
            home: m.score.fullTime.home,
            away: m.score.fullTime.away,
            status: 'FINISHED',
            pin: ADMIN_PIN
          });
          updated++;
          break;
        }
      }
    });
    Logger.log('fetchLiveScores complete: updated=' + updated);
    return { success: true, updated: updated };
  } catch(err) {
    Logger.log('fetchLiveScores error: ' + err.toString());
    return { error: err.toString() };
  }
}
