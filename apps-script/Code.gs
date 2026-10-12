/**
 * 42maru 대관·카페42 정산 자동화 (Google Sheet에 붙이는 Apps Script)
 *
 * 설치: README의 "설치" 절차 참고. 최초 1회 setup() 실행.
 */

var CONFIG = {
  CUSTOMER_EMAIL: 'namgyu@bysq.co.kr',
  CUSTOMER_NAME: '바이스퀘어',
  CUSTOMER_CONTACT: '조남규 대표님',
  OWNER_EMAIL: '',              // 비우면 스크립트 실행 계정으로 알림
  RENTAL_FEE: 400000,           // 1회 대관료 (부가세 별도)
  TAX_RATE: 0.1,
  WEEKEND_SERVICE_FEE: 40000,   // 주말 음료 용역비 (부가세 포함, 42maru → 마켓샐러드)
  DRINK_REQUEST_DAYS_BEFORE: 4,
  REMIND_DAYS_BEFORE: 1,
  MAIL_SEARCH_WINDOW: 'newer_than:30d',
  DEFAULT_PRICES: [['커피', 3500, 1100], ['티', 2500, 800]]
};

var SHEET = {
  RENTAL: '대관원장',
  ORDER: '음료주문',
  SETTLE: '일별정산',
  PRICE: '단가표',
  MSG: '메시지함'
};

var HEADERS = {};
HEADERS[SHEET.RENTAL] = ['대관ID', '대관일', '요일', '주말', '시작', '종료', '고객사', '대관료(공급가)', '부가세',
  '대관료 합계', '상태', '계산서요청', '음료요청발송', '음료확정', '정산', '캘린더ID', '원본메일ID', '등록일'];
HEADERS[SHEET.ORDER] = ['대관ID', '대관일', '메뉴', '수량', '원본메일ID', '기록일'];
HEADERS[SHEET.SETTLE] = ['대관ID', '대관일', '요일', '고객사', '주문내역', '결제금액', '원가', '세금',
  '42maru 정산금액', '주말 용역비', '순정산액', '산정일'];
HEADERS[SHEET.PRICE] = ['메뉴', '판매가', '원가'];
HEADERS[SHEET.MSG] = ['생성일', '대관ID', '받는 곳', '채널', '문구'];

var COL = {}; // COL[sheet][header] = 1-based column
Object.keys(HEADERS).forEach(function (name) {
  COL[name] = {};
  HEADERS[name].forEach(function (h, i) { COL[name][h] = i + 1; });
});

// ───────────────────────── 설치 ─────────────────────────

function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#eef2f7');
    }
  });

  var price = ss.getSheetByName(SHEET.PRICE);
  if (price.getLastRow() === 1) {
    price.getRange(2, 1, CONFIG.DEFAULT_PRICES.length, 3).setValues(CONFIG.DEFAULT_PRICES);
  }

  var rental = ss.getSheetByName(SHEET.RENTAL);
  rental.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  rental.getRange('E:F').setNumberFormat('HH:mm');
  rental.getRange('H:J').setNumberFormat('#,##0');
  ss.getSheetByName(SHEET.ORDER).getRange('B:B').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEET.SETTLE).getRange('B:B').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEET.SETTLE).getRange('F:K').setNumberFormat('#,##0');
  ss.getSheetByName(SHEET.PRICE).getRange('B:C').setNumberFormat('#,##0');
  ss.getSheetByName(SHEET.MSG).getRange('E:E').setWrap(true);

  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processInbox').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(9).everyDays(1).create();

  SpreadsheetApp.getActive().toast('설치 완료: 10분마다 메일 확인, 매일 09시 음료 요청·알림');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('대관 자동화')
    .addItem('카톡으로 확정된 대관 등록', 'menuRegisterRental')
    .addItem('음료 수량 직접 입력', 'menuRecordDrinks')
    .addSeparator()
    .addItem('지금 메일 확인', 'processInbox')
    .addItem('오늘 일일 작업 실행', 'dailyJob')
    .addItem('전월 정산 리포트 보내기', 'sendLastMonthReport')
    .addToUi();
}

// ───────────────────────── 메뉴 ─────────────────────────

function menuRegisterRental() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('대관 등록', '예) 10월 10일 14시 ~ 17시  또는  2026-10-10 14:00-17:00', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var rental = registerRental(res.getResponseText(), '');
  ui.alert(rental ? '등록 완료: ' + rental.id + '\n메시지함 탭에서 반투소프트 단톡방용 문구를 확인하세요.'
    : '날짜·시간을 읽지 못했습니다. 예시 형식으로 다시 입력해 주세요.');
}

function menuRecordDrinks() {
  var ui = SpreadsheetApp.getUi();
  var idRes = ui.prompt('음료 수량 입력', '대관ID (예: R20261010-1)', ui.ButtonSet.OK_CANCEL);
  if (idRes.getSelectedButton() !== ui.Button.OK) return;
  var id = idRes.getResponseText().trim();
  if (!findRental(id)) { ui.alert('대관ID를 찾을 수 없습니다: ' + id); return; }
  var qRes = ui.prompt('음료 수량 입력', '예) 커피 25, 티 25', ui.ButtonSet.OK_CANCEL);
  if (qRes.getSelectedButton() !== ui.Button.OK) return;
  var order = parseDrinkOrder(qRes.getResponseText(), Object.keys(getPriceTable()));
  if (!Object.keys(order).length) { ui.alert('수량을 읽지 못했습니다.'); return; }
  var s = recordDrinkOrder(id, order, '');
  ui.alert('정산 완료: 순정산액 ' + won(s.net) + '원\n메시지함 탭에서 카페42용 문구를 확인하세요.');
}

// ───────────────────────── 메일 처리 ─────────────────────────

/** 10분 트리거: [대관확정] 메일과 [음료요청] 회신을 처리 */
function processInbox() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    processConfirmMails();
    processDrinkReplies();
  } finally {
    lock.releaseLock();
  }
}

function processConfirmMails() {
  GmailApp.search('subject:대관확정 ' + CONFIG.MAIL_SEARCH_WINDOW).forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      var m = msg.getSubject().match(/^\s*\[대관확정\]\s*(.+)$/); // 회신(Re:)은 제외
      if (!m || isSeen(msg.getId())) return;
      markSeen(msg.getId());
      registerRental(m[1], msg.getId());
    });
  });
}

function processDrinkReplies() {
  var menus = Object.keys(getPriceTable());
  GmailApp.search('subject:음료요청 from:' + CONFIG.CUSTOMER_EMAIL + ' ' + CONFIG.MAIL_SEARCH_WINDOW)
    .forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        if (isSeen(msg.getId())) return;
        if (msg.getFrom().toLowerCase().indexOf(CONFIG.CUSTOMER_EMAIL) < 0) return;
        var idMatch = msg.getSubject().match(/#(R\d{8}-\d+)/);
        if (!idMatch || !findRental(idMatch[1])) return;
        markSeen(msg.getId());
        var order = parseDrinkOrder(msg.getPlainBody(), menus);
        if (!Object.keys(order).length) {
          notifyOwner('[대관] 음료 회신을 읽지 못했습니다 #' + idMatch[1],
            '조남규 대표님 회신에서 수량을 찾지 못했습니다. 시트 메뉴 "음료 수량 직접 입력"으로 넣어 주세요.\n\n' +
            msg.getPlainBody().slice(0, 1000));
          return;
        }
        recordDrinkOrder(idMatch[1], order, msg.getId());
      });
    });
}

/** 처리한 메일 ID 기록 (같은 메일 재처리·반복 알림 방지) */
function isSeen(id) {
  return !!PropertiesService.getScriptProperties().getProperty('seen:' + id);
}

function markSeen(id) {
  PropertiesService.getScriptProperties().setProperty('seen:' + id, String(Date.now()));
}

// ───────────────────────── 1·2단계: 대관 등록 + 계산서 요청 ─────────────────────────

function registerRental(text, sourceId) {
  var p = parseSchedule(text, new Date());
  if (!p) {
    if (sourceId) notifyOwner('[대관] 확정 메일 형식 오류', '읽지 못한 제목: ' + text +
      '\n형식: [대관확정] 2026-10-10 14:00-17:00 바이스퀘어');
    return null;
  }
  var dup = findRentalByStart(p.start);
  if (dup) return dup;

  var company = p.company || CONFIG.CUSTOMER_NAME;
  var date = new Date(p.start.getFullYear(), p.start.getMonth(), p.start.getDate());
  var id = nextRentalId(date);
  var vat = Math.round(CONFIG.RENTAL_FEE * CONFIG.TAX_RATE);
  var weekend = isWeekend(date);

  var event = CalendarApp.getDefaultCalendar().createEvent('[대관] ' + company, p.start, p.end, {
    description: '대관ID ' + id + '\n대관료 ' + won(CONFIG.RENTAL_FEE) + '원 (부가세 별도)' +
      (weekend ? '\n주말: 음료 용역비 ' + won(CONFIG.WEEKEND_SERVICE_FEE) + '원 차감' : '')
  });

  appendRow(SHEET.RENTAL, {
    '대관ID': id, '대관일': date, '요일': WEEKDAYS[date.getDay()], '주말': weekend ? 'Y' : 'N',
    '시작': p.start, '종료': p.end, '고객사': company,
    '대관료(공급가)': CONFIG.RENTAL_FEE, '부가세': vat, '대관료 합계': CONFIG.RENTAL_FEE + vat,
    '상태': '계산서요청', '계산서요청': new Date(), '캘린더ID': event.getId(),
    '원본메일ID': sourceId, '등록일': new Date()
  });

  var when = formatDateKo(date) + ' ' + formatTime(p.start) + '~' + formatTime(p.end);
  var kakao = '[대관 공유]\n' +
    '일시: ' + when + '\n' +
    '대관: ' + company + '\n' +
    '세금계산서 발행 요청드립니다.\n' +
    '- 공간 대관료(공급가액): ' + won(CONFIG.RENTAL_FEE) + '원\n' +
    '- 부가세: ' + won(vat) + '원\n' +
    '- 합계: ' + won(CONFIG.RENTAL_FEE + vat) + '원';
  saveMessage(id, '반투소프트 단톡방 (이근영·김동환 대표)', kakao);
  notifyOwner('[대관 등록] ' + when + ' ' + company + ' #' + id,
    '대관을 등록하고 캘린더에 추가했습니다.\n\n아래 문구를 반투소프트 단톡방에 붙여넣어 주세요.\n\n' + kakao);

  var rental = findRental(id);
  if (daysUntil(date) <= CONFIG.DRINK_REQUEST_DAYS_BEFORE && daysUntil(date) >= 1) requestDrinks(rental);
  return rental;
}

// ───────────────────────── 3단계: 음료 요청 · 정산 ─────────────────────────

function requestDrinks(rental) {
  var menus = Object.keys(getPriceTable());
  var when = formatDateKo(rental.date) + ' ' + formatTime(rental.start) + '~' + formatTime(rental.end);
  var subject = '[음료요청] ' + when + ' ' + rental.company + ' 음료 수량 확인 #' + rental.id;
  var body = '안녕하세요, ' + CONFIG.CUSTOMER_CONTACT + '.\n42maru입니다.\n\n' +
    when + ' 대관 관련하여 음료 종류와 잔 수를 확인 부탁드립니다.\n' +
    '이 메일에 아래 형식으로 회신해 주시면 됩니다.\n\n' +
    menus.map(function (m) { return m + ' : 잔'; }).join('\n') + '\n\n' +
    '음료는 당일 카페42에서 결제됩니다.\n감사합니다.';
  GmailApp.sendEmail(CONFIG.CUSTOMER_EMAIL, subject, body);
  setRentalFields(rental.id, { '음료요청발송': new Date(), '상태': '음료요청' });
  saveMessage(rental.id, '조남규 대표 1:1 카톡 (메일 대신 보낼 때)',
    when + ' 대관 음료 종류와 잔 수 알려주세요.\n' + menus.map(function (m) { return m + ' : 잔'; }).join('\n'));
}

function recordDrinkOrder(rentalId, order, sourceId) {
  var rental = findRental(rentalId);
  var priceTable = getPriceTable();
  var s = calcSettlement(order, priceTable, rental.weekend,
    { taxRate: CONFIG.TAX_RATE, weekendServiceFee: CONFIG.WEEKEND_SERVICE_FEE });

  // 수정 회신이 오면 최신 내용으로 교체
  deleteRows(SHEET.ORDER, '대관ID', rentalId);
  s.lines.forEach(function (l) {
    appendRow(SHEET.ORDER, { '대관ID': rentalId, '대관일': rental.date, '메뉴': l.menu, '수량': l.qty,
      '원본메일ID': sourceId, '기록일': new Date() });
  });

  deleteRows(SHEET.SETTLE, '대관ID', rentalId);
  appendRow(SHEET.SETTLE, {
    '대관ID': rentalId, '대관일': rental.date, '요일': WEEKDAYS[rental.date.getDay()], '고객사': rental.company,
    '주문내역': orderToText(order), '결제금액': s.total.pay, '원가': s.total.cost, '세금': s.total.tax,
    '42maru 정산금액': s.total.settle, '주말 용역비': s.serviceFee, '순정산액': s.net, '산정일': new Date()
  });
  setRentalFields(rentalId, { '음료확정': new Date(), '정산': s.net, '상태': '정산완료' });

  var kakao = settlementText(rental, s);
  saveMessage(rentalId, '카페42 (마켓샐러드)', kakao);
  notifyOwner('[음료 확정·정산] ' + formatDateKo(rental.date) + ' ' + rental.company + ' 순정산 ' + won(s.net) + '원',
    '음료 주문이 확정되어 정산을 계산했습니다.\n아래 문구를 카페42에 전달해 주세요.\n\n' + kakao);
  return s;
}

function settlementText(rental, s) {
  var lines = ['[카페42 음료 주문·정산] ' + formatDateKo(rental.date) + ' ' +
    formatTime(rental.start) + '~' + formatTime(rental.end) + ' ' + rental.company,
    '주문: ' + s.lines.map(function (l) { return l.menu + ' ' + l.qty + '잔'; }).join(', '), ''];
  s.lines.forEach(function (l) {
    lines.push('- ' + l.menu + ' ' + l.qty + '잔: 결제 ' + won(l.pay) + ' / 원가 ' + won(l.cost) +
      ' / 세금 ' + won(l.tax) + ' → 정산 ' + won(l.settle) + '원');
  });
  lines.push('42maru 정산금액 합계: ' + won(s.total.settle) + '원');
  if (s.serviceFee) lines.push('주말 음료 용역비(42maru 지급, 부가세 포함): -' + won(s.serviceFee) + '원');
  lines.push(s.net >= 0 ? '순정산액(카페42 → 42maru): ' + won(s.net) + '원'
    : '순정산액(42maru → 카페42): ' + won(-s.net) + '원');
  return lines.join('\n');
}

// ───────────────────────── 일일 작업 · 월간 리포트 ─────────────────────────

/** 매일 09시: D-4 음료 요청, D-1 미회신 알림, 매월 1일 전월 리포트 */
function dailyJob() {
  getRentals().forEach(function (r) {
    var d = daysUntil(r.date);
    if (d >= 1 && d <= CONFIG.DRINK_REQUEST_DAYS_BEFORE && !r.drinkRequested) requestDrinks(r);
    if (d === CONFIG.REMIND_DAYS_BEFORE && !r.drinkConfirmed) {
      notifyOwner('[대관 D-1] 음료 수량 미회신 ' + formatDateKo(r.date) + ' ' + r.company + ' #' + r.id,
        '내일 대관인데 음료 수량 회신이 없습니다. 조남규 대표님께 확인 후 시트 메뉴 "음료 수량 직접 입력"으로 넣어 주세요.');
    }
  });
  if (new Date().getDate() === 1) sendLastMonthReport();
}

function sendLastMonthReport() {
  var now = new Date();
  var from = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var to = new Date(now.getFullYear(), now.getMonth(), 1);
  var rows = readObjects(SHEET.SETTLE).filter(function (r) { return r['대관일'] >= from && r['대관일'] < to; });
  var label = from.getFullYear() + '년 ' + (from.getMonth() + 1) + '월';
  if (!rows.length) { notifyOwner('[월간 정산] ' + label, label + ' 대관 정산 내역이 없습니다.'); return; }

  var sum = { pay: 0, settle: 0, fee: 0, net: 0 };
  var body = [label + ' 대관·카페42 정산 리포트', ''];
  rows.forEach(function (r) {
    sum.pay += r['결제금액']; sum.settle += r['42maru 정산금액']; sum.fee += r['주말 용역비']; sum.net += r['순정산액'];
    body.push(formatDateKo(r['대관일']) + ' ' + r['고객사'] + ' | ' + r['주문내역'] + ' | 정산 ' +
      won(r['42maru 정산금액']) + (r['주말 용역비'] ? ' - 용역비 ' + won(r['주말 용역비']) : '') +
      ' = ' + won(r['순정산액']) + '원');
  });
  body.push('', '대관 ' + rows.length + '건',
    '음료 결제금액 합계: ' + won(sum.pay) + '원',
    '42maru 정산금액 합계: ' + won(sum.settle) + '원',
    '주말 용역비 합계: ' + won(sum.fee) + '원',
    '순정산액 합계(카페42 → 42maru): ' + won(sum.net) + '원',
    '대관료 합계(공급가): ' + won(CONFIG.RENTAL_FEE * rows.length) + '원');
  notifyOwner('[월간 정산] ' + label + ' 순정산 ' + won(sum.net) + '원', body.join('\n'));
}

// ───────────────────────── 시트 헬퍼 ─────────────────────────

function sheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function appendRow(name, obj) {
  sheet(name).appendRow(HEADERS[name].map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

function readObjects(name) {
  var sh = sheet(name);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS[name].length).getValues().map(function (row, i) {
    var o = { _row: i + 2 };
    HEADERS[name].forEach(function (h, j) { o[h] = row[j]; });
    return o;
  });
}

function deleteRows(name, header, value) {
  var rows = readObjects(name).filter(function (r) { return r[header] === value; });
  rows.reverse().forEach(function (r) { sheet(name).deleteRow(r._row); });
}

function toRental(r) {
  return {
    id: r['대관ID'], date: r['대관일'], start: r['시작'], end: r['종료'], company: r['고객사'],
    weekend: r['주말'] === 'Y', drinkRequested: !!r['음료요청발송'], drinkConfirmed: !!r['음료확정'], _row: r._row
  };
}

function getRentals() {
  return readObjects(SHEET.RENTAL).filter(function (r) { return r['대관ID']; }).map(toRental);
}

function findRental(id) {
  return getRentals().filter(function (r) { return r.id === id; })[0] || null;
}

function findRentalByStart(start) {
  return getRentals().filter(function (r) {
    return r.date && r.date.getTime() === new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime() &&
      r.start && r.start.getHours() === start.getHours() && r.start.getMinutes() === start.getMinutes();
  })[0] || null;
}

function setRentalFields(id, fields) {
  var r = findRental(id);
  Object.keys(fields).forEach(function (h) {
    sheet(SHEET.RENTAL).getRange(r._row, COL[SHEET.RENTAL][h]).setValue(fields[h]);
  });
}

function nextRentalId(date) {
  var prefix = 'R' + date.getFullYear() + pad2(date.getMonth() + 1) + pad2(date.getDate()) + '-';
  var n = getRentals().filter(function (r) { return String(r.id).indexOf(prefix) === 0; }).length;
  return prefix + (n + 1);
}

function getPriceTable() {
  var table = {};
  readObjects(SHEET.PRICE).forEach(function (r) {
    if (r['메뉴']) table[String(r['메뉴']).trim()] = { price: Number(r['판매가']), cost: Number(r['원가']) };
  });
  return table;
}

function daysUntil(date) {
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((d - today) / 86400000);
}

function saveMessage(rentalId, to, text) {
  appendRow(SHEET.MSG, { '생성일': new Date(), '대관ID': rentalId, '받는 곳': to, '채널': '카톡', '문구': text });
}

function notifyOwner(subject, body) {
  var to = CONFIG.OWNER_EMAIL || Session.getEffectiveUser().getEmail();
  GmailApp.sendEmail(to, subject, body + '\n\n시트: ' + SpreadsheetApp.getActive().getUrl());
}
