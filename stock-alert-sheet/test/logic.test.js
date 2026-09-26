// Code.gs 의 순수 함수를 Node 에서 검사한다: node stock-alert-sheet/test/logic.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
// vm 컨텍스트의 배열·객체는 다른 realm 이라 JSON 으로 비교한다
const same = (a, b) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b));

// Apps Script 의 Utilities.formatDate 중 이 파일이 쓰는 형식만 흉내 낸다 (Asia/Seoul 고정)
const Utilities = {
  formatDate(date, tz, fmt) {
    const kst = new Date(date.getTime() + 9 * 3600 * 1000);
    const pad = (n) => String(n).padStart(2, '0');
    if (fmt === 'u') return String(kst.getUTCDay() === 0 ? 7 : kst.getUTCDay());
    if (fmt === 'HHmm') return pad(kst.getUTCHours()) + pad(kst.getUTCMinutes());
    throw new Error('unsupported format ' + fmt);
  }
};

const ctx = vm.createContext({ Utilities });
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('✓ ' + name);
}

// ── evaluateStock ──
const base = { name: '삼성전자', price: 70000, change: 1.2, target: null, stop: null, swing: 5 };

test('조건 없으면 알림 없음', () => {
  same(ctx.evaluateStock(base).length, 0);
});

test('목표가 도달', () => {
  const hits = ctx.evaluateStock({ ...base, target: 70000 });
  same(hits.map((h) => h.type), ['target']);
  assert.ok(hits[0].message.includes('70,000원'));
});

test('손절가 이탈', () => {
  same(ctx.evaluateStock({ ...base, stop: 71000 }).map((h) => h.type), ['stop']);
});

test('급등·급락 기준', () => {
  same(ctx.evaluateStock({ ...base, change: 5 }).map((h) => h.type), ['surge']);
  same(ctx.evaluateStock({ ...base, change: -6.3 }).map((h) => h.type), ['plunge']);
  same(ctx.evaluateStock({ ...base, change: -4.9 }).length, 0);
});

test('등락률 문구에 부호 표시', () => {
  assert.ok(ctx.evaluateStock({ ...base, change: 7.456 })[0].message.includes('+7.46%'));
});

// ── parseBreadth ──
const html = `
<dl class="lst_kos_info">
  <dd class="dd">
    <a href="/sise/sise_upper.naver"><span class="ico_up">상한</span><span class="num">3</span></a>
    <a href="/sise/sise_rise.naver?sosok=0"><span class="ico_rise">상승</span> <span class="num">1,012</span></a>
    <a href="/sise/sise_steady.naver?sosok=0"><span>보합</span><span class="num">85</span></a>
    <a href="/sise/sise_fall.naver?sosok=0"><span>하락</span><span class="num">400</span></a>
    <a href="/sise/sise_lower.naver"><span>하한</span><span class="num">0</span></a>
  </dd>
</dl>`;

test('시장폭 파싱 (상한 포함, 쉼표 숫자)', () => {
  const c = ctx.parseBreadth(html, 0);
  same(c, { up: 1015, flat: 85, down: 400, ratio: 67.7 });
});

test('구조가 다르면 null', () => {
  assert.strictEqual(ctx.parseBreadth('<html>점검 중</html>', 0), null);
});

// ── breadthAlertMessage ──
test('시장폭 알림 기준', () => {
  assert.ok(ctx.breadthAlertMessage({ up: 700, flat: 0, down: 300, ratio: 70 }, 70, 30).includes('과열'));
  assert.ok(ctx.breadthAlertMessage({ up: 250, flat: 0, down: 750, ratio: 25 }, 70, 30).includes('급랭'));
  assert.strictEqual(ctx.breadthAlertMessage({ up: 500, flat: 0, down: 500, ratio: 50 }, 70, 30), null);
});

// ── parseInvestorFlow ──
test('투자자별 순매수 파싱', () => {
  const page = '<dl><dt>개인</dt><dd><span class="num">+1,234</span>억</dd>' +
    '<dt>외국인</dt><dd>-567억</dd><dt>기관계</dt><dd>- 890 억</dd></dl><a>개인정보처리방침</a>';
  same(ctx.parseInvestorFlow(page), { individual: 1234, foreign: -567, institution: -890 });
});

test('투자자별 순매수 없으면 null', () => {
  assert.strictEqual(ctx.parseInvestorFlow('<p>개인정보처리방침</p>'), null);
});

// ── parseDeposit ──
test('고객예탁금·신용잔고 파싱 (가장 최근 행)', () => {
  const page = `<table><tr><th>날짜</th><th>고객예탁금</th><th>증감</th><th>신용잔고</th></tr>
    <tr><td>26.09.24</td><td>612,345</td><td>-1,203</td><td>201,987</td><td>+55</td></tr>
    <tr><td>26.09.23</td><td>613,548</td><td>+900</td><td>201,932</td><td>-12</td></tr></table>`;
  same(ctx.parseDeposit(page), { date: '2026-09-24', customerDeposit: 612345, creditBalance: 201987 });
});

test('예탁금 표가 없으면 null', () => {
  assert.strictEqual(ctx.parseDeposit('<html>점검 중</html>'), null);
});

// ── isMarketOpen (UTC 로 넣고 한국 시간으로 판정) ──
test('장중 판정', () => {
  assert.strictEqual(ctx.isMarketOpen(new Date('2026-09-25T00:00:00Z')), true); // 금 09:00 KST
  assert.strictEqual(ctx.isMarketOpen(new Date('2026-09-25T06:30:00Z')), true); // 금 15:30
  assert.strictEqual(ctx.isMarketOpen(new Date('2026-09-25T06:31:00Z')), false); // 금 15:31
  assert.strictEqual(ctx.isMarketOpen(new Date('2026-09-24T23:59:00Z')), false); // 금 08:59
  assert.strictEqual(ctx.isMarketOpen(new Date('2026-09-26T02:00:00Z')), false); // 토 11:00
});

console.log(`\n${passed}개 테스트 통과`);
