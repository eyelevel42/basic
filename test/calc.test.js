// node --test test/*.test.js
const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const C = require(path.join(__dirname, '..', 'apps-script', 'Calc.gs'));

const prices = { 커피: { price: 3500, cost: 1100 }, 티: { price: 2500, cost: 800 } };
const cfg = { taxRate: 0.1, weekendServiceFee: 40000 };

test('표 예시: 커피 25 · 티 25 (평일)', () => {
  const s = C.calcSettlement({ 커피: 25, 티: 25 }, prices, false, cfg);
  const coffee = s.lines.find(l => l.menu === '커피');
  const tea = s.lines.find(l => l.menu === '티');
  assert.deepStrictEqual([coffee.pay, coffee.cost, coffee.tax, coffee.settle], [87500, 27500, 8750, 51250]);
  assert.deepStrictEqual([tea.pay, tea.cost, tea.tax, tea.settle], [62500, 20000, 6250, 36250]);
  assert.strictEqual(s.total.settle, 87500);
  assert.strictEqual(s.serviceFee, 0);
  assert.strictEqual(s.net, 87500);
});

test('주말이면 용역비 4만원 차감', () => {
  const s = C.calcSettlement({ 커피: 25, 티: 25 }, prices, true, cfg);
  assert.strictEqual(s.net, 47500);
});

test('주말 소량 주문은 음수(42maru 지급)', () => {
  const s = C.calcSettlement({ 커피: 10 }, prices, true, cfg); // 20,500 - 40,000
  assert.strictEqual(s.net, -19500);
});

test('단가표에 없는 메뉴는 오류', () => {
  assert.throws(() => C.calcSettlement({ 라떼: 1 }, prices, false, cfg));
});

test('주말 판정', () => {
  assert.strictEqual(C.isWeekend(new Date(2026, 9, 10)), true);  // 토
  assert.strictEqual(C.isWeekend(new Date(2026, 9, 12)), false); // 월
});

test('일정 파싱: ISO 형식', () => {
  const p = C.parseSchedule('2026-10-10 14:00-17:00 바이스퀘어', new Date(2026, 9, 1));
  assert.strictEqual(p.start.getTime(), new Date(2026, 9, 10, 14, 0).getTime());
  assert.strictEqual(p.end.getTime(), new Date(2026, 9, 10, 17, 0).getTime());
  assert.strictEqual(p.company, '바이스퀘어');
});

test('일정 파싱: 카톡 문장', () => {
  const p = C.parseSchedule('10월 10일 14시 ~ 17시, 대관', new Date(2026, 9, 1));
  assert.strictEqual(p.start.getTime(), new Date(2026, 9, 10, 14, 0).getTime());
  assert.strictEqual(p.end.getTime(), new Date(2026, 9, 10, 17, 0).getTime());
  assert.strictEqual(p.company, '');
});

test('일정 파싱: 연도 생략 + 지난 날짜면 다음 해', () => {
  const p = C.parseSchedule('1월 5일 10시 30분 ~ 12시', new Date(2026, 9, 12));
  assert.strictEqual(p.start.getTime(), new Date(2027, 0, 5, 10, 30).getTime());
});

test('일정 파싱: 잘못된 입력', () => {
  assert.strictEqual(C.parseSchedule('다음 주에 봬요'), null);
  assert.strictEqual(C.parseSchedule('2026-10-10 17:00-14:00'), null);
});

test('음료 회신 파싱: 인용문 무시', () => {
  const body = '커피 25잔\n티: 25\n\n2026년 10월 6일 (화) 오전 9:00, 42maru님이 작성:\n> 커피 : 잔\n> 티 : 99';
  assert.deepStrictEqual(C.parseDrinkOrder(body, ['커피', '티']), { 커피: 25, 티: 25 });
});

test('음료 회신 파싱: 아이스티와 티 구분', () => {
  assert.deepStrictEqual(C.parseDrinkOrder('아이스티 5, 티 3', ['티', '아이스티']), { 아이스티: 5, 티: 3 });
});

test('음료 회신 파싱: 수량 없음', () => {
  assert.deepStrictEqual(C.parseDrinkOrder('확인했습니다', ['커피', '티']), {});
});

test('금액 포맷', () => {
  assert.strictEqual(C.won(87500), '87,500');
  assert.strictEqual(C.won(-19500), '-19,500');
});
