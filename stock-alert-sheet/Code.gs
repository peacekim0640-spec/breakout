/**
 * 주식 알림 구글 시트 (관심종목 + 시장폭)
 *
 * - 관심종목: GOOGLEFINANCE 시세(약 20분 지연)로 목표가·손절가·급등락을 확인해 텔레그램으로 알림
 * - 시장폭: 네이버 금융에서 코스피·코스닥 상승/보합/하락 종목 수를 읽어 매일 기록 (실험 기능)
 *
 * 설치 방법은 README.md 를 참고하세요.
 */

var SHEET_SETTINGS = '설정';
var SHEET_WATCH = '관심종목';
var SHEET_BREADTH = '시장폭';

var WATCH_HEADERS = ['시장', '종목코드', '종목명', '현재가', '등락률(%)', '목표가', '손절가', '급등락 기준(%)', '상태', '마지막 알림'];
var COL = { market: 0, code: 1, name: 2, price: 3, change: 4, target: 5, stop: 6, swing: 7, status: 8, lastAlert: 9 };

var SETTINGS_ROWS = [
  ['텔레그램 봇 토큰', '', 'BotFather 에게 받은 토큰 (이 시트를 다른 사람과 공유하지 마세요)'],
  ['텔레그램 채팅 ID', '', '내 채팅 ID (README 참고)'],
  ['알림 켜기', true, '체크 해제하면 알림을 보내지 않습니다'],
  ['상승비율 상단(%)', 70, '코스피 상승종목 비율이 이 값 이상이면 알림 (과열 신호)'],
  ['상승비율 하단(%)', 30, '코스피 상승종목 비율이 이 값 이하면 알림 (투매 신호)']
];

var BREADTH_MARKETS = [
  { name: '코스피', code: 'KOSPI', sosok: 0 },
  { name: '코스닥', code: 'KOSDAQ', sosok: 1 }
];

var TZ = 'Asia/Seoul';

// ─────────────────────────────── 메뉴 ───────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📈 주식 알림')
    .addItem('처음 설정하기', 'setup')
    .addSeparator()
    .addItem('지금 관심종목 확인', 'checkWatchlistNow')
    .addItem('지금 시장폭 기록', 'logBreadth')
    .addItem('텔레그램 테스트 메시지', 'sendTestMessage')
    .addSeparator()
    .addItem('자동 실행 켜기', 'installTriggers')
    .addItem('자동 실행 끄기', 'removeTriggers')
    .addToUi();
}

/** 시트 3개를 만들고 기본 양식을 채운다. 이미 있는 시트의 데이터는 건드리지 않는다. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(SHEET_SETTINGS);
    settings.getRange(1, 1, 1, 3).setValues([['항목', '값', '설명']]).setFontWeight('bold');
    settings.getRange(2, 1, SETTINGS_ROWS.length, 3).setValues(SETTINGS_ROWS);
    settings.getRange(4, 2).insertCheckboxes().setValue(true);
    settings.setColumnWidth(1, 160);
    settings.setColumnWidth(2, 320);
    settings.setColumnWidth(3, 420);
  }

  var watch = ss.getSheetByName(SHEET_WATCH);
  if (!watch) {
    watch = ss.insertSheet(SHEET_WATCH);
    watch.getRange(1, 1, 1, WATCH_HEADERS.length).setValues([WATCH_HEADERS]).setFontWeight('bold');
    watch.setFrozenRows(1);
    watch.getRange('B2:B200').setNumberFormat('@'); // 005930 의 앞자리 0 이 사라지지 않게
    watch.getRange(2, 1, 2, 2).setValues([['KRX', '005930'], ['KRX', '000660']]);
    watch.getRange(2, 6, 2, 3).setValues([['', '', 5], ['', '', 5]]);
    fillWatchFormulas_(watch, 200);
  }

  var breadth = ss.getSheetByName(SHEET_BREADTH);
  if (!breadth) {
    breadth = ss.insertSheet(SHEET_BREADTH);
    breadth.getRange(1, 1, 1, 7)
      .setValues([['날짜', '시장', '상승', '보합', '하락', '상승비율(%)', '비고']])
      .setFontWeight('bold');
    breadth.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    '설정 완료!\n\n' +
    '1) [설정] 시트에 텔레그램 봇 토큰과 채팅 ID를 입력하세요.\n' +
    '2) [관심종목] 시트에 종목을 추가하세요. (코스피 = KRX, 코스닥 = KOSDAQ)\n' +
    '3) 메뉴 → 텔레그램 테스트 메시지 로 확인한 뒤\n' +
    '4) 메뉴 → 자동 실행 켜기 를 누르세요.'
  );
}

/** 종목명·현재가·등락률 칸에 GOOGLEFINANCE 수식을 넣는다. */
function fillWatchFormulas_(sheet, lastRow) {
  var formulas = [];
  for (var r = 2; r <= lastRow; r++) {
    var ticker = 'A' + r + '&":"&B' + r;
    var guard = 'IF(B' + r + '="","",IFERROR(GOOGLEFINANCE(' + ticker + ',';
    formulas.push([
      '=' + guard + '"name"),"종목 확인불가"))',
      '=' + guard + '"price"),""))',
      '=' + guard + '"changepct"),""))'
    ]);
  }
  sheet.getRange(2, COL.name + 1, formulas.length, 3).setFormulas(formulas);
}

// ─────────────────────────────── 자동 실행 ───────────────────────────────

function installTriggers() {
  removeTriggers(true);
  ScriptApp.newTrigger('checkWatchlist').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('logBreadth').timeBased().everyDays(1).atHour(16).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('자동 실행을 켰습니다.\n\n· 관심종목: 장중(평일 09:00~15:30) 10분마다 확인\n· 시장폭: 평일 오후 4시쯤 하루 1번 기록');
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'checkWatchlist' || fn === 'logBreadth') ScriptApp.deleteTrigger(t);
  });
  if (silent !== true) SpreadsheetApp.getUi().alert('자동 실행을 껐습니다.');
}

/** 트리거용: 장중에만 동작한다. */
function checkWatchlist() {
  if (!isMarketOpen(new Date())) return;
  runWatchlistCheck_();
}

/** 메뉴용: 시간과 상관없이 바로 확인한다. */
function checkWatchlistNow() {
  var sent = runWatchlistCheck_();
  SpreadsheetApp.getActiveSpreadsheet().toast(sent + '건의 알림을 보냈습니다.', '관심종목 확인', 5);
}

// ─────────────────────────────── 관심종목 ───────────────────────────────

function runWatchlistCheck_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_WATCH);
  if (!sheet) throw new Error('[관심종목] 시트가 없습니다. 메뉴 → 처음 설정하기 를 먼저 실행하세요.');

  var settings = readSettings_();
  var props = PropertiesService.getDocumentProperties();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var now = Utilities.formatDate(new Date(), TZ, 'MM-dd HH:mm');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var rows = sheet.getRange(2, 1, lastRow - 1, WATCH_HEADERS.length).getValues();

  var sent = 0;
  rows.forEach(function (row, i) {
    if (!row[COL.code]) return;
    var stock = {
      name: row[COL.name] || row[COL.code],
      price: toNumber(row[COL.price]),
      change: toNumber(row[COL.change]),
      target: toNumber(row[COL.target]),
      stop: toNumber(row[COL.stop]),
      swing: toNumber(row[COL.swing])
    };

    var statusCell = sheet.getRange(i + 2, COL.status + 1);
    if (stock.price === null) {
      statusCell.setValue('시세 없음');
      return;
    }

    var hits = evaluateStock(stock);
    statusCell.setValue(hits.length ? hits.map(function (h) { return h.label; }).join(', ') : '정상');

    hits.forEach(function (hit) {
      // 같은 종목·같은 조건은 하루 한 번만 보낸다
      var key = 'alert|' + row[COL.market] + ':' + row[COL.code] + '|' + hit.type;
      if (props.getProperty(key) === today) return;
      if (settings.enabled && sendTelegram_(settings, hit.message)) {
        props.setProperty(key, today);
        sheet.getRange(i + 2, COL.lastAlert + 1).setValue(now + ' ' + hit.label);
        sent++;
      }
    });
  });
  return sent;
}

/**
 * 한 종목의 알림 조건을 검사한다. (순수 함수 — 테스트 가능)
 * @return {Array<{type: string, label: string, message: string}>}
 */
function evaluateStock(stock) {
  var hits = [];
  var head = '[' + stock.name + '] 현재가 ' + formatNumber(stock.price) + '원 (' + formatSigned(stock.change) + '%)';

  if (stock.target !== null && stock.target > 0 && stock.price >= stock.target) {
    hits.push({ type: 'target', label: '목표가 도달', message: '🎯 목표가 도달\n' + head + '\n목표가 ' + formatNumber(stock.target) + '원' });
  }
  if (stock.stop !== null && stock.stop > 0 && stock.price <= stock.stop) {
    hits.push({ type: 'stop', label: '손절가 이탈', message: '⚠️ 손절가 이탈\n' + head + '\n손절가 ' + formatNumber(stock.stop) + '원' });
  }
  if (stock.swing !== null && stock.swing > 0 && stock.change !== null) {
    if (stock.change >= stock.swing) {
      hits.push({ type: 'surge', label: '급등', message: '🚀 급등 ' + formatSigned(stock.change) + '%\n' + head });
    } else if (stock.change <= -stock.swing) {
      hits.push({ type: 'plunge', label: '급락', message: '📉 급락 ' + formatSigned(stock.change) + '%\n' + head });
    }
  }
  return hits;
}

// ─────────────────────────────── 시장폭 ───────────────────────────────

/** 코스피·코스닥 상승/보합/하락 종목 수를 [시장폭] 시트에 기록하고, 코스피 상승비율이 극단이면 알린다. */
function logBreadth() {
  var nowDate = new Date();
  var day = Number(Utilities.formatDate(nowDate, TZ, 'u')); // 1=월 … 7=일
  var isTrigger = !isInteractive_();
  if (isTrigger && day >= 6) return; // 주말에는 자동 기록하지 않음

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BREADTH);
  if (!sheet) throw new Error('[시장폭] 시트가 없습니다. 메뉴 → 처음 설정하기 를 먼저 실행하세요.');
  var settings = readSettings_();
  var today = Utilities.formatDate(nowDate, TZ, 'yyyy-MM-dd');

  var results = BREADTH_MARKETS.map(function (m) {
    try {
      var res = UrlFetchApp.fetch('https://finance.naver.com/sise/sise_index.naver?code=' + m.code, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      if (res.getResponseCode() !== 200) return { market: m, error: 'HTTP ' + res.getResponseCode() };
      var counts = parseBreadth(res.getContentText('EUC-KR'), m.sosok);
      return counts ? { market: m, counts: counts } : { market: m, error: '페이지에서 종목 수를 찾지 못함 (네이버 화면 구조 변경 가능성)' };
    } catch (e) {
      return { market: m, error: String(e.message || e) };
    }
  });

  results.forEach(function (r) {
    if (r.counts) {
      var c = r.counts;
      sheet.appendRow([today, r.market.name, c.up, c.flat, c.down, c.ratio, '']);
    } else {
      sheet.appendRow([today, r.market.name, '', '', '', '', '실패: ' + r.error]);
    }
  });

  var kospi = results[0].counts;
  if (kospi && settings.enabled) {
    var msg = breadthAlertMessage(kospi, settings.breadthHigh, settings.breadthLow);
    var key = 'breadth|' + today;
    var props = PropertiesService.getDocumentProperties();
    if (msg && props.getProperty(key) !== today && sendTelegram_(settings, msg)) props.setProperty(key, today);
  }

  if (!isTrigger) {
    var failed = results.filter(function (r) { return !r.counts; });
    ss.toast(failed.length ? '일부 실패: ' + failed.map(function (r) { return r.market.name + ' - ' + r.error; }).join(' / ') : '기록 완료', '시장폭', 8);
  }
}

/**
 * 네이버 금융 지수 페이지 HTML 에서 상승/보합/하락 종목 수를 읽는다. (순수 함수 — 테스트 가능)
 * 상한가는 상승에, 하한가는 하락에 포함한다.
 * @return {{up: number, flat: number, down: number, ratio: number} | null}
 */
function parseBreadth(html, sosok) {
  function countFor(page) {
    // 예: <a href="/sise/sise_rise.naver?sosok=0"> … <span class="num">412</span> … </a>
    var re = new RegExp('sise_' + page + '\\.naver(?:\\?sosok=' + sosok + ')?["\'][^>]*>([\\s\\S]*?)</a>');
    var m = html.match(re);
    if (!m) return null;
    var num = m[1].replace(/<[^>]*>/g, ' ').match(/(\d[\d,]*)/);
    return num ? Number(num[1].replace(/,/g, '')) : null;
  }

  var rise = countFor('rise');
  var steady = countFor('steady');
  var fall = countFor('fall');
  if (rise === null || steady === null || fall === null) return null;

  var upper = countFor('upper') || 0;
  var lower = countFor('lower') || 0;
  var up = rise + upper;
  var down = fall + lower;
  var total = up + steady + down;
  if (total === 0) return null;

  return { up: up, flat: steady, down: down, ratio: Math.round((up / total) * 1000) / 10 };
}

/** 코스피 상승비율이 설정한 상단/하단을 넘으면 알림 문구를 돌려준다. (순수 함수) */
function breadthAlertMessage(counts, high, low) {
  var body = '상승 ' + counts.up + ' / 보합 ' + counts.flat + ' / 하락 ' + counts.down + ' (상승비율 ' + counts.ratio + '%)';
  if (high !== null && counts.ratio >= high) return '🔥 코스피 시장폭 과열\n' + body + '\n기준 ' + high + '% 이상';
  if (low !== null && counts.ratio <= low) return '🧊 코스피 시장폭 급랭\n' + body + '\n기준 ' + low + '% 이하';
  return null;
}

// ─────────────────────────────── 텔레그램 ───────────────────────────────

function sendTestMessage() {
  var settings = readSettings_();
  var ok = sendTelegram_(settings, '✅ 주식 알림 시트 연결 성공!');
  SpreadsheetApp.getUi().alert(ok ? '텔레그램으로 테스트 메시지를 보냈습니다.' : '전송 실패: [설정] 시트의 토큰과 채팅 ID를 확인하세요.\n(봇에게 먼저 아무 메시지나 한 번 보내야 합니다)');
}

function sendTelegram_(settings, text) {
  if (!settings.token || !settings.chatId) return false;
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + settings.token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: settings.chatId, text: text }),
    muteHttpExceptions: true
  });
  return res.getResponseCode() === 200;
}

// ─────────────────────────────── 도우미 ───────────────────────────────

function readSettings_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) throw new Error('[설정] 시트가 없습니다. 메뉴 → 처음 설정하기 를 먼저 실행하세요.');
  var map = {};
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues().forEach(function (r) { map[r[0]] = r[1]; });
  return {
    token: String(map['텔레그램 봇 토큰'] || '').trim(),
    chatId: String(map['텔레그램 채팅 ID'] || '').trim(),
    enabled: map['알림 켜기'] === true || String(map['알림 켜기']).toUpperCase() === 'TRUE',
    breadthHigh: toNumber(map['상승비율 상단(%)']),
    breadthLow: toNumber(map['상승비율 하단(%)'])
  };
}

/** 평일 09:00~15:30 (한국 시간) 이면 true. 공휴일은 구분하지 않는다. */
function isMarketOpen(date) {
  var day = Number(Utilities.formatDate(date, TZ, 'u'));
  var hm = Number(Utilities.formatDate(date, TZ, 'HHmm'));
  return day <= 5 && hm >= 900 && hm <= 1530;
}

function isInteractive_() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

function toNumber(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function formatNumber(n) {
  return n === null ? '-' : Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatSigned(n) {
  if (n === null) return '-';
  var s = (Math.round(n * 100) / 100).toFixed(2);
  return n > 0 ? '+' + s : s;
}
