// ============================================================
//  태오시스템 — Gmail PO 자동 수집 (Google Apps Script)
//  사용법:
//    1. Google 스프레드시트 생성 → 확장 프로그램 → Apps Script
//    2. 이 파일 전체 붙여넣기 → 저장
//    3. 배포 → 새 배포 → 유형: 웹 앱
//       - 실행 계정: 나(Me)
//       - 액세스 권한: 모든 사용자(anonymous 포함)
//    4. 발급된 URL을 po-viewer.html의 APPS_SCRIPT_URL에 입력
//    5. 아래 setupTrigger() 함수를 한 번 실행 (트리거 등록)
// ============================================================

// ── 설정 ──────────────────────────────────────────────────
// Gmail 검색 쿼리: 필요에 따라 조정
const PO_SEARCH_QUERY = '(subject:(PO OR "purchase order" OR "발주서" OR "발주" OR "order" OR "주문") OR (PO OR "purchase order" OR "발주")) newer_than:60d';
const SHEET_NAME = 'POs';
// ────────────────────────────────────────────────────────────

/**
 * GET 요청 처리 — JSON 또는 JSONP 반환
 */
function doGet(e) {
  // 최초 실행 시 트리거 자동 등록
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('trigger_set')) {
    setupTrigger();
    props.setProperty('trigger_set', 'true');
  }

  // ?refresh=1 파라미터로 수동 수집 가능
  if (e && e.parameter && e.parameter.refresh === '1') {
    fetchNewPOs();
  }

  const data = getAllPOs();
  const json = JSON.stringify({ success: true, data: data, updated: new Date().toISOString() });

  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 시트에서 전체 PO 목록 반환
 */
function getAllPOs() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
      return obj;
    })
    .reverse(); // 최신순
}

/**
 * Gmail에서 새 PO 이메일 수집 → 시트에 저장
 * 트리거로 자동 실행됨
 */
function fetchNewPOs() {
  const sheet = getOrCreateSheet();
  const existingIds = sheet.getDataRange().getValues().slice(1).map(r => r[11]);

  const threads = GmailApp.search(PO_SEARCH_QUERY);
  let added = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (existingIds.includes(msgId)) return;

      const subject = msg.getSubject();
      const from    = msg.getFrom();
      const date    = msg.getDate();
      const body    = msg.getPlainBody().substring(0, 3000);
      const parsed  = parsePO(subject, body, from);

      sheet.appendRow([
        Utilities.getUuid(),   // id
        parsed.poNumber,       // poNumber
        from,                  // sender
        parsed.company,        // company
        subject,               // subject
        date,                  // date
        parsed.amount,         // amount
        parsed.currency,       // currency
        parsed.items,          // items
        'new',                 // status
        body.substring(0, 800),// body
        msgId                  // emailId
      ]);
      added++;
    });
  });

  Logger.log('신규 PO 추가: ' + added + '건');
  return added;
}

/**
 * 이메일 제목/본문에서 PO 정보 추출
 */
function parsePO(subject, body, from) {
  const text = subject + '\n' + body;

  // PO 번호
  const poMatch = text.match(
    /(?:PO|P\.O\.|발주번호|주문번호|order\s*no|order\s*#)[.\s\-:#]*([A-Z0-9\-\/]{3,20})/i
  );
  const poNumber = poMatch ? poMatch[1].trim() : '';

  // 금액
  const amountMatch = text.match(
    /(?:USD|KRW|MNT|EUR|CNY)[\s]?([\d,]+(?:\.\d{1,2})?)|(?:[\d,]+(?:\.\d{1,2})?)[\s]?(?:USD|KRW|MNT|EUR|CNY)/i
  );
  const amount = amountMatch
    ? (amountMatch[1] || amountMatch[0]).replace(/[^0-9.]/g, '')
    : '';

  // 통화
  const currencyMatch = text.match(/\b(USD|KRW|MNT|EUR|CNY)\b/i);
  const currency = currencyMatch ? currencyMatch[1].toUpperCase() : 'USD';

  // 발신 회사 (이메일 표시명에서 추출)
  const companyMatch = from.match(/"?([^"<]+)"?\s*</) || from.match(/^([^@<]+)/);
  const company = companyMatch ? companyMatch[1].trim() : from;

  // 품목 (본문 앞부분에서 간략히)
  const lines = body.split('\n').filter(l => l.trim().length > 3).slice(0, 4);
  const items = lines.join(' | ').substring(0, 200);

  return { poNumber, amount, currency, company, items };
}

/**
 * 시트가 없으면 생성 + 헤더 설정
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['id','poNumber','sender','company','subject','date','amount','currency','items','status','body','emailId'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * 1시간마다 자동 수집 트리거 등록 (최초 1회만 실행)
 */
function setupTrigger() {
  // 기존 트리거 제거
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fetchNewPOs') ScriptApp.deleteTrigger(t);
  });
  // 새 트리거 등록
  ScriptApp.newTrigger('fetchNewPOs')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('트리거 등록 완료 — 1시간마다 fetchNewPOs() 실행');
}
