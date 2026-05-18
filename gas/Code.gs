/**
 * Code.gs — 민엽 & 현지 자산관리 대시보드 (Google Apps Script 버전)
 *
 * 배포 방법:
 *   확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱
 *   실행 계정: 나(내 계정)  /  액세스 권한: 모든 사용자
 *
 * 데이터 시트: 자산현황, 부동산계획, 결혼계획, 월별저축
 * 행 번호는 자산계획_v3.xlsx (patch_v3c 적용 후) 와 동일
 */

// ──────────────────────────────────────────────
// 진입점
// ──────────────────────────────────────────────
function doGet(e) {
  const d = loadData();
  buildDerivedData(d);

  const tmpl = HtmlService.createTemplateFromFile('dashboard');
  tmpl.d = d;
  return tmpl.evaluate()
    .setTitle('민엽 & 현지 · 자산관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ──────────────────────────────────────────────
// 엑셀과 동일한 로직으로 데이터 로드
// ──────────────────────────────────────────────
function loadData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d  = {};

  // 숫자 셀 읽기 (formula 셀도 계산값으로 반환)
  function g(sheet, r, c, def) {
    const v = sheet.getRange(r, c).getValue();
    return (typeof v === 'number' && isFinite(v)) ? v : (def || 0);
  }
  // 문자 셀 읽기
  function gs(sheet, r, c) {
    const v = sheet.getRange(r, c).getValue();
    return (v != null) ? String(v).trim() : '';
  }

  // ── 자산현황 ──────────────────────────────────
  const ws = ss.getSheetByName('자산현황');

  d.현금_m = 0; d.현금_h = 0;
  for (let r = 6; r <= 15; r++) { d.현금_m += g(ws,r,3); d.현금_h += g(ws,r,5); }
  d.현금_t = d.현금_m + d.현금_h;

  d.투자_m = 0; d.투자_h = 0;
  for (let r = 19; r <= 24; r++) { d.투자_m += g(ws,r,3); d.투자_h += g(ws,r,5); }
  d.투자_t = d.투자_m + d.투자_h;

  // 계약금 R28-31 (patch_v3c: 아파트+예식장+취득세+기타)
  d.계약금_m = 0; d.계약금_h = 0;
  const paid = [];
  for (let r = 28; r <= 31; r++) {
    const name = gs(ws, r, 2);
    const vm = g(ws,r,3), vh = g(ws,r,5);
    d.계약금_m += vm; d.계약금_h += vh;
    if (name && !name.includes('소')) paid.push([name, vm + vh]);
  }
  d.계약금_t = d.계약금_m + d.계약금_h;
  d.납부완료_items = paid;

  // 부채 R35-38 (insert_rows(29,3) 후 이동)
  d.부채_m = 0; d.부채_h = 0;
  for (let r = 35; r <= 38; r++) { d.부채_m += g(ws,r,3); d.부채_h += g(ws,r,5); }
  d.부채_t = d.부채_m + d.부채_h;

  d.순자산_m = d.현금_m + d.투자_m + d.계약금_m + d.부채_m;
  d.순자산_h = d.현금_h + d.투자_h + d.부채_h;
  d.순자산_t = d.순자산_m + d.순자산_h;
  d.가용_t   = d.현금_t + d.투자_t + d.부채_t;

  // ── 부동산계획 ────────────────────────────────
  const wsr = ss.getSheetByName('부동산계획');
  d.매매가   = g(wsr, 5, 4, 1210000000);
  d.잔금_자  = g(wsr, 7, 4, 220000000);
  d.대출     = g(wsr, 8, 4, 560000000);
  d.잔금_t   = d.잔금_자 + d.대출;
  d.취득세   = g(wsr, 9, 4, 50000000);
  d.보증금   = g(wsr, 10, 4, 300000000);
  d.인테리어 = g(wsr, 11, 4, 40000000);

  // ── 결혼계획 ──────────────────────────────────
  const wsw = ss.getSheetByName('결혼계획');
  d.결혼식   = g(wsw, 5, 4, 50000000);
  d.신혼여행 = g(wsw, 6, 4, 10000000);
  d.결혼총   = g(wsw, 7, 4) || (d.결혼식 + d.신혼여행);
  d.축의금   = g(wsw, 8, 4, 30000000);
  d.결혼순   = g(wsw, 9, 4) || (d.결혼총 - d.축의금);

  // ── 월별저축 ──────────────────────────────────
  const wsm = ss.getSheetByName('월별저축');
  const monthly = [];
  for (let r = 21; r <= 65; r++) {
    const bval = wsm.getRange(r, 2).getValue();
    if (!(bval instanceof Date) || isNaN(bval.getTime())) break;
    const ym     = Utilities.formatDate(bval, 'Asia/Seoul', 'yy.MM');
    const ymFull = Utilities.formatDate(bval, 'Asia/Seoul', 'yyyy-MM');
    const mIn  = g(wsm, r, 3), hIn  = g(wsm, r, 4);
    const mOut = g(wsm, r, 6), hOut = g(wsm, r, 7);
    const tIn = mIn + hIn, tOut = mOut + hOut;
    monthly.push({ ym, ymFull, 합계수입: tIn, 합계지출: tOut, 저축액: tIn - tOut });
  }
  d.월별     = monthly;
  d.누계저축 = monthly.filter(m => m.ymFull <= '2027-09').reduce((s, m) => s + m.저축액, 0);

  return d;
}

// ──────────────────────────────────────────────
// 파생 데이터 (차트 배열, 포맷 문자열 등)
// ──────────────────────────────────────────────
function buildDerivedData(d) {
  // D-Day & 프로그레스 바
  const now     = new Date();
  const wedding = new Date(2027, 6, 1);
  const movein  = new Date(2027, 8, 1);
  const start   = new Date(2024, 0, 1);
  d.dday_wed  = Math.ceil((wedding - now) / 86400000);
  d.dday_move = Math.ceil((movein  - now) / 86400000);
  d.d_bal     = d.dday_move;
  d.prog_wed  = Math.min(100, Math.floor((now - start) / (wedding - start) * 100));
  d.prog_bal  = Math.min(100, Math.floor((now - start) / (movein  - start) * 100));
  d.today     = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy. MM. dd');

  // 자금계획 수치
  const 보증금반환대출 = 100000000;
  d.보증금반환대출 = 보증금반환대출;
  d.총지출 = Math.abs(d.잔금_t) + Math.abs(d.취득세) + Math.abs(d.보증금)
           + Math.abs(d.인테리어) + Math.abs(d.결혼총);
  d.총가용 = d.가용_t + d.누계저축 + d.대출 + 보증금반환대출 + d.축의금;
  d.여유   = d.총가용 - d.총지출;

  // 차트 배열 (JSON 문자열로 미리 직렬화 → 템플릿에서 <?!= ?> 로 삽입)
  d.donut_labels = JSON.stringify(["현금·예금", "투자자산", "계약금"]);
  d.donut_values = JSON.stringify([
    Math.max(0, d.현금_t), Math.max(0, d.투자_t), Math.max(0, d.계약금_t)
  ]);

  d.bar_cats     = JSON.stringify(["현금·예금", "투자자산", "계약금", "부채"]);
  d.bar_minyeob  = JSON.stringify([d.현금_m, d.투자_m, d.계약금_m, d.부채_m]);
  d.bar_hyeonji  = JSON.stringify([d.현금_h, d.투자_h, 0,          d.부채_h]);

  const fundItems = [
    ["잔금 (자기자금+주담대)",  Math.abs(d.잔금_t),    "#EF4444"],
    ["취득세·부대비용",          Math.abs(d.취득세),    "#EF4444"],
    ["세입자 보증금 반환",       Math.abs(d.보증금),    "#EF4444"],
    ["인테리어·가전",            Math.abs(d.인테리어),  "#EF4444"],
    ["결혼식·신혼여행",          Math.abs(d.결혼총),    "#EF4444"],
    ["현재 가용자산",            d.가용_t,              "#2E75B6"],
    ["2027.09 저축 예상",        d.누계저축,            "#2E75B6"],
    ["주담대 대출",              d.대출,                "#0D9488"],
    ["보증금 반환 대출",         d.보증금반환대출,       "#0D9488"],
    ["예상 축의금",              d.축의금,              "#10B981"],
  ];
  d.fund_labels = JSON.stringify(fundItems.map(x => x[0]));
  d.fund_values = JSON.stringify(fundItems.map(x => x[1]));
  d.fund_colors = JSON.stringify(fundItems.map(x => x[2]));
  d.fund_count  = fundItems.length;

  const targetMonthly = d.월별.filter(m => m.ymFull <= '2027-09');
  d.months_ym = JSON.stringify(targetMonthly.map(m => m.ym));
  d.수입_arr  = JSON.stringify(targetMonthly.map(m => m.합계수입));
  d.지출_arr  = JSON.stringify(targetMonthly.map(m => m.합계지출));
  d.저축_arr  = JSON.stringify(targetMonthly.map(m => m.저축액));
  let acc = 0;
  d.누계_arr  = JSON.stringify(targetMonthly.map(m => { acc += m.저축액; return acc; }));

  // 납부완료 카드 HTML
  d.paid_html = d.납부완료_items.map(([name, val]) => {
    if (val > 0) return `<br>${name}: ${fmt(val)}`;
    return `<br><span style="color:#C0CAD8">${name}: -</span>`;
  }).join('');

  // 부동산·결혼 accordion HTML
  const reRows = [
    ["아파트 매매가",        fmt(d.매매가),                               ""],
    ["계약금 납부",          fmt(d.계약금_m),                             "2025.04 완료"],
    ["잔금 총액",            fmt(d.잔금_t),                               "2027.09"],
    ["&nbsp;&nbsp;└ 자기자금",    fmt(d.잔금_자),                         ""],
    ["&nbsp;&nbsp;└ 주담대 대출", fmt(d.대출),                            "예정"],
    ["취득세·부대비용",      fmt(d.취득세),                               "예상치"],
    ["세입자 보증금 반환",   fmt(d.보증금),                               "입주 시"],
    ["&nbsp;&nbsp;└ 보증금 반환 대출", fmt(d.보증금반환대출),              "예정"],
    ["인테리어·가전",        fmt(d.인테리어),                             "예상치"],
    ["대출 합계",            fmt(d.대출 + d.보증금반환대출),               "주담대+보증금"],
  ];
  d.re_rows_html = reRows.map(([n,v,note]) =>
    `<tr><td class="td-lbl">${n}</td><td class="td-val">${v}</td><td class="td-note">${note}</td></tr>`
  ).join('\n');

  const wedRows = [
    ["결혼식 합계",   fmt(d.결혼식),            ""],
    ["신혼여행 합계", fmt(d.신혼여행),           ""],
    ["결혼 총비용",   fmt(d.결혼총),             ""],
    ["예상 축의금",   "+" + fmt(d.축의금),       "들어올 돈"],
    ["결혼 순비용",   fmt(d.결혼순),             "총비용 - 축의금"],
  ];
  d.wed_rows_html = wedRows.map(([n,v,note]) =>
    `<tr><td class="td-lbl">${n}</td><td class="td-val">${v}</td><td class="td-note">${note}</td></tr>`
  ).join('\n');

  // 여유/부족 class
  d.여유_class = d.여유 >= 0 ? 'green' : 'red';
  d.여유_lbl   = d.여유 >= 0 ? '여유'  : '부족';
  d.여유_val   = fmt(Math.abs(d.여유));
}

// ──────────────────────────────────────────────
// 숫자 포맷 (억/만 단위, 음수=△)
// ──────────────────────────────────────────────
function fmt(n) {
  if (n == null) return '0';
  const neg = n < 0;
  n = Math.abs(n);
  let s;
  if (n >= 100000000) {
    const eok = Math.floor(n / 100000000);
    const man = Math.floor((n % 100000000) / 10000);
    s = man > 0 ? `${eok}억 ${man.toLocaleString()}만` : `${eok}억`;
  } else if (n >= 10000) {
    s = `${Math.floor(n / 10000).toLocaleString()}만`;
  } else {
    s = n.toLocaleString();
  }
  return (neg ? '△' : '') + s;
}
