/**
 * 순수 계산·파싱 함수 (Google API 미사용 → Node 테스트 가능)
 */

var WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

/** 주말(토·일) 여부 */
function isWeekend(date) {
  var d = date.getDay();
  return d === 0 || d === 6;
}

/** 2026.10.10(토) */
function formatDateKo(date) {
  return date.getFullYear() + '.' + pad2(date.getMonth() + 1) + '.' + pad2(date.getDate()) +
    '(' + WEEKDAYS[date.getDay()] + ')';
}

function formatTime(date) {
  return pad2(date.getHours()) + ':' + pad2(date.getMinutes());
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

/** 87500 → "87,500" */
function won(n) {
  var sign = n < 0 ? '-' : '';
  return sign + String(Math.abs(Math.round(n))).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 대관 일정 문자열 파싱.
 *  - "2026-10-10 14:00-17:00 바이스퀘어"
 *  - "10월 10일 14시 ~ 17시, 대관"   (연도 생략 시 now 기준으로 지나간 날짜면 다음 해)
 * @return {{start: Date, end: Date, company: string} | null}
 */
function parseSchedule(text, now) {
  now = now || new Date();
  var s = String(text).trim();
  var y, mo, d, sh, sm, eh, em, rest;

  var m = s.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})\s*(?:\([^)]*\))?\s+(\d{1,2}):(\d{2})\s*[-~]\s*(\d{1,2}):(\d{2})\s*(.*)$/);
  if (m) {
    y = +m[1]; mo = +m[2]; d = +m[3];
    sh = +m[4]; sm = +m[5]; eh = +m[6]; em = +m[7]; rest = m[8];
  } else {
    m = s.match(/(?:(\d{4})\s*년\s*)?(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*(?:\([^)]*\))?\s*(\d{1,2})\s*(?:시|:)\s*(\d{1,2})?\s*(?:분)?\s*[-~]\s*(\d{1,2})\s*(?:시|:)\s*(\d{1,2})?\s*(?:분)?\s*(.*)$/);
    if (!m) return null;
    mo = +m[2]; d = +m[3];
    sh = +m[4]; sm = +(m[5] || 0); eh = +m[6]; em = +(m[7] || 0); rest = m[8];
    if (m[1]) {
      y = +m[1];
    } else {
      y = now.getFullYear();
      var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      if (new Date(y, mo - 1, d) < today) y += 1;
    }
  }

  var start = new Date(y, mo - 1, d, sh, sm);
  var end = new Date(y, mo - 1, d, eh, em);
  if (start.getMonth() !== mo - 1 || start.getDate() !== d || end <= start) return null;

  var company = String(rest || '')
    .replace(/^[\s,·\-]+/, '')
    .replace(/대관/g, '')
    .replace(/[,·\-\s]+$/, '')
    .trim();
  return { start: start, end: end, company: company };
}

/**
 * 메일 회신 본문에서 음료 수량 추출 ("커피 25", "티: 25잔").
 * 인용된 원문(>, "wrote:", "작성:") 이후는 무시한다.
 * @param {string[]} menuNames 단가표 메뉴명
 * @return {Object<string, number>} 수량이 하나도 없으면 빈 객체
 */
function parseDrinkOrder(body, menuNames) {
  var lines = String(body).split(/\r?\n/);
  var own = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (/^\s*>/.test(line) || /wrote:\s*$/.test(line) || /작성:?\s*$/.test(line) ||
        /^-{2,}\s*(Original|원본)/i.test(line)) break;
    own.push(line);
  }
  var text = own.join('\n');
  // 긴 이름부터 매칭해 "아이스티"가 "티"로 잡히지 않게 한다
  var names = menuNames.slice().sort(function (a, b) { return b.length - a.length; });
  var result = {};
  names.forEach(function (name) {
    var re = new RegExp('(^|[^가-힣A-Za-z])' + escapeRegExp(name) + '\\s*[:：=]?\\s*(\\d+)\\s*잔?', 'g');
    var mm;
    while ((mm = re.exec(text)) !== null) {
      result[name] = (result[name] || 0) + Number(mm[2]);
    }
    text = text.replace(re, '$1');
  });
  return result;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 일별 음료 정산.
 * 메뉴별: 결제금액 = 판매가×수량, 원가 = 원가×수량, 세금 = 결제금액×10%,
 *         42maru 정산금액 = 결제금액 − 원가 − 세금
 * 주말이면 음료 용역비(부가세 포함)를 차감해 순정산액을 낸다.
 *
 * @param {Object<string, number>} order  {커피: 25, 티: 25}
 * @param {Object<string, {price:number, cost:number}>} priceTable
 * @param {boolean} weekend
 * @param {{taxRate:number, weekendServiceFee:number}} cfg
 */
function calcSettlement(order, priceTable, weekend, cfg) {
  var lines = [];
  var total = { qty: 0, pay: 0, cost: 0, tax: 0, settle: 0 };
  Object.keys(order).forEach(function (menu) {
    var qty = order[menu];
    var p = priceTable[menu];
    if (!p) throw new Error('단가표에 없는 메뉴: ' + menu);
    if (!qty) return;
    var pay = p.price * qty;
    var cost = p.cost * qty;
    var tax = Math.round(pay * cfg.taxRate);
    var settle = pay - cost - tax;
    lines.push({ menu: menu, qty: qty, price: p.price, unitCost: p.cost, pay: pay, cost: cost, tax: tax, settle: settle });
    total.qty += qty; total.pay += pay; total.cost += cost; total.tax += tax; total.settle += settle;
  });
  var serviceFee = weekend ? cfg.weekendServiceFee : 0;
  return { lines: lines, total: total, serviceFee: serviceFee, net: total.settle - serviceFee };
}

/** {커피: 25, 티: 25} → "커피 25잔, 티 25잔" */
function orderToText(order) {
  return Object.keys(order).filter(function (k) { return order[k]; })
    .map(function (k) { return k + ' ' + order[k] + '잔'; }).join(', ');
}

/** "커피 25잔, 티 25잔" → {커피: 25, 티: 25} */
function textToOrder(text) {
  var order = {};
  String(text || '').split(',').forEach(function (part) {
    var m = part.trim().match(/^(.+?)\s+(\d+)잔$/);
    if (m) order[m[1]] = Number(m[2]);
  });
  return order;
}

if (typeof module !== 'undefined') {
  module.exports = {
    isWeekend: isWeekend, formatDateKo: formatDateKo, formatTime: formatTime, won: won,
    parseSchedule: parseSchedule, parseDrinkOrder: parseDrinkOrder,
    calcSettlement: calcSettlement, orderToText: orderToText, textToOrder: textToOrder
  };
}
