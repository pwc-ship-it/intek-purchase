// Vercel Serverless Function
// 엔드포인트: POST /api/export
// Body: { items: [...], statusMap: {...}, searchMeta: { cond, summary } }
// Response: xlsx 파일 (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
//
// 의존성: exceljs (package.json에 추가 필요)
// npm install exceljs

import ExcelJS from 'exceljs';

// ── 색상 팔레트 ──────────────────────────────────────────
const C = {
  navy:      '1A3A5C',
  blue:      '1A56DB',
  blueLt:    'D6E4FF',
  bluePale:  'EBF5FF',
  teal:      '0694A2',
  green:     '057A55',
  greenLt:   'DEF7EC',
  amber:     'B45309',
  amberLt:   'FEF3C7',
  red:       'C81E1E',
  redLt:     'FDE8E8',
  gray900:   '111827',
  gray700:   '374151',
  gray500:   '6B7280',
  gray300:   'D1D5DB',
  gray200:   'E5E7EB',
  gray100:   'F3F4F6',
  gray50:    'F9FAFB',
  white:     'FFFFFF',
};

// ── 스타일 헬퍼 ──────────────────────────────────────────
const font = (bold = false, size = 10, color = C.gray900, italic = false) => ({
  name: 'Arial', bold, size, color: { argb: 'FF' + color }, italic,
});
const fill = (hex) => ({
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + hex },
});
const al = (horizontal = 'center', wrapText = false, indent = 0) => ({
  horizontal, vertical: 'middle', wrapText, indent,
});
const border = (color = C.gray200) => {
  const s = { style: 'thin', color: { argb: 'FF' + color } };
  return { top: s, bottom: s, left: s, right: s };
};
const btmBorder = () => ({
  bottom: { style: 'thin', color: { argb: 'FF' + C.gray200 } },
  left:   { style: 'thin', color: { argb: 'FF' + C.gray200 } },
  right:  { style: 'thin', color: { argb: 'FF' + C.gray200 } },
});
const hdrBorder = (c = C.blue) => ({
  top:    { style: 'thin',   color: { argb: 'FF' + c } },
  bottom: { style: 'medium', color: { argb: 'FF' + c } },
  left:   { style: 'thin',   color: { argb: 'FF' + c } },
  right:  { style: 'thin',   color: { argb: 'FF' + c } },
});

// 셀에 스타일 일괄 적용
function sc(cell, value, {fnt, fll, aln, bdr} = {}) {
  // exceljs에서 수식은 { formula: '...' } 객체로 설정해야 계산됨
  if (typeof value === 'string' && value.startsWith('=')) {
    cell.value = { formula: value.slice(1) };
  } else {
    cell.value = value;
  }
  if (fnt) cell.font      = fnt;
  if (fll) cell.fill      = fll;
  if (aln) cell.alignment = aln;
  if (bdr) cell.border    = bdr;
}

// 입고/배송 상태 판단
const today = () => new Date().toISOString().slice(0, 10);

function arrStatus(st) {
  if (st.arrival_done)       return 'done';
  if (!st.expected_arrival)  return 'none';
  if (st.expected_arrival < today()) return 'late';
  return 'pending';
}
function shipStatus(st) {
  if (st.ship_done)          return 'done';
  if (!st.ship_done_date)    return st.ship_start_date ? 'started' : 'none';
  if (st.ship_done_date < today()) return 'late';
  return 'pending';
}

// ─────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')
    return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { items = [], statusMap = {}, searchMeta = {} } = req.body || {};
    if (!items.length)
      return res.status(400).json({ error: '내보낼 데이터가 없습니다.' });

    const TODAY   = today();
    const docNos  = [...new Set(items.map(r => r.doc_no))];
    const wb      = new ExcelJS.Workbook();
    wb.creator    = '인텍플러스 구매관리';
    wb.created    = new Date();

    // ════════════════════════════════════════════════════════
    // 시트 1: 📦 데이터 원본
    // ════════════════════════════════════════════════════════
    const ws1 = wb.addWorksheet('📦 데이터', {
      views: [{ showGridLines: false, state: 'frozen', xSplit: 2, ySplit: 2 }],
      properties: { tabColor: { argb: 'FF' + C.navy } },
    });

    const DATA_COLS = [
      { header: '사이트',        key: 'site',    width: 20 },
      { header: '문서번호',      key: 'doc_no',  width: 24 },
      { header: '기안자',        key: 'user',    width: 10 },
      { header: '기안일',        key: 'date',    width: 12 },
      { header: '프로젝트명',    key: 'p_name',  width: 34 },
      { header: '코드',          key: 'p_code',  width: 20 },
      { header: '품명',          key: 'name',    width: 28 },
      { header: '규격',          key: 'spec',    width: 28 },
      { header: '수량',          key: 'qty',     width:  8 },
      { header: '입고예정일',    key: 'arr_d',   width: 13 },
      { header: '입고완료',      key: 'arr_c',   width: 10 },
      { header: '배송시작일',    key: 'ship_s',  width: 13 },
      { header: '배송완료예정일',key: 'ship_d',  width: 15 },
      { header: '배송완료',      key: 'ship_c',  width: 10 },
      { header: '메모',          key: 'memo',    width: 32 },
    ];
    ws1.columns = DATA_COLS;

    // 1행: 제목 배너
    ws1.getRow(1).height = 34;
    ws1.mergeCells(1, 1, 1, DATA_COLS.length);
    sc(ws1.getCell('A1'), `🛒  인텍플러스 구매관리 — 데이터 원본  |  ${TODAY}`, {
      fnt: font(true, 13, C.white), fll: fill(C.navy), aln: al('left'),
    });

    // 2행: 헤더
    ws1.getRow(2).height = 32;
    DATA_COLS.forEach((col, ci) => {
      const cell = ws1.getCell(2, ci + 1);
      sc(cell, col.header, {
        fnt: font(true, 10, C.white),
        fll: fill(C.navy),
        aln: al('center', true),
        bdr: hdrBorder(C.blue),
      });
    });

    // 데이터 행
    const SITE_ID = items[0]?.site || '';
    items.forEach((row, ri) => {
      const r   = ri + 3;
      const isE = ri % 2 === 0;
      const st  = statusMap[row.doc_no] || {};
      const asSt  = arrStatus(st);
      const shSt  = shipStatus(st);

      ws1.getRow(r).height = 22;

      // 입고완료 텍스트/스타일
      let arrText, arrFnt, arrFll;
      if (asSt === 'done')    { arrText = '✅ 완료';     arrFnt = font(true,10,C.green);  arrFll = fill(C.greenLt); }
      else if (asSt === 'late'){ arrText = '⚠️ 지연';    arrFnt = font(true,10,C.red);    arrFll = fill(C.redLt);   }
      else                    { arrText = '⏳ 미완료';   arrFnt = font(true,10,C.amber);  arrFll = fill(C.amberLt); }

      // 배송완료 텍스트/스타일
      let shipText, shipFnt, shipFll;
      if (shSt === 'done')     { shipText = '✅ 완료';    shipFnt = font(true,10,C.green);  shipFll = fill(C.greenLt); }
      else if (shSt === 'late'){ shipText = '❓ 확인필요'; shipFnt = font(true,10,C.red);    shipFll = fill(C.redLt);   }
      else if (shSt==='started'){ shipText = '🚚 배송중';  shipFnt = font(true,10,C.blue);   shipFll = fill(C.bluePale);}
      else                     { shipText = '—';           shipFnt = font(false,10,C.gray500); shipFll = fill(isE?C.gray50:C.white); }

      const bg  = fill(isE ? C.gray50 : C.white);
      const bgB = fill(C.bluePale);

      const rowData = [
        { v: SITE_ID,                fnt: font(false,10,C.gray500), fll: bg,          aln: al('left',false,1) },
        { v: row.doc_no||'',         fnt: font(true, 10,C.navy),    fll: isE?bgB:bg,  aln: al('left',false,1) },
        { v: row.requester||'',      fnt: font(false,10,C.gray900), fll: bg,          aln: al('center') },
        { v: row.doc_date||'',       fnt: font(false,10,C.gray900), fll: bg,          aln: al('center') },
        { v: row.p_name||'',         fnt: font(false,10,C.gray900), fll: bg,          aln: al('left',false,1) },
        { v: row.p_code||'',         fnt: font(false,10,C.gray500), fll: bg,          aln: al('center') },
        { v: row.name||'',           fnt: font(false,10,C.gray900), fll: bg,          aln: al('left',false,1) },
        { v: row.spec||'',           fnt: font(false,10,C.gray900), fll: bg,          aln: al('left',false,1) },
        { v: row.qty||0,             fnt: font(true, 10,C.blue),    fll: bgB,         aln: al('center') },
        { v: st.expected_arrival||'',fnt: font(false,10,C.gray900), fll: bg,          aln: al('center') },
        { v: arrText,                fnt: arrFnt,                   fll: arrFll,       aln: al('center') },
        { v: st.ship_start_date||'', fnt: font(false,10,C.gray900), fll: bg,          aln: al('center') },
        { v: st.ship_done_date||'',  fnt: font(false,10,C.gray900), fll: bg,          aln: al('center') },
        { v: shipText,               fnt: shipFnt,                  fll: shipFll,      aln: al('center') },
        { v: st.memo||'',            fnt: font(false,10,C.gray900), fll: bg,          aln: al('left',false,1) },
      ];

      rowData.forEach(({v, fnt: f, fll: fl, aln: a}, ci) => {
        const cell = ws1.getCell(r, ci + 1);
        sc(cell, v, { fnt: f, fll: fl, aln: a, bdr: btmBorder() });
      });
    });

    // AutoFilter
    ws1.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2 + items.length, column: DATA_COLS.length } };


    // ════════════════════════════════════════════════════════
    // 시트 2: 🔍 검색 (AGGREGATE/SMALL 다중조건 실시간 수식)
    // ════════════════════════════════════════════════════════
    const ws2 = wb.addWorksheet('🔍 검색', {
      views: [{ showGridLines: false, state: 'frozen', xSplit: 0, ySplit: 8 }],
      properties: { tabColor: { argb: 'FF' + C.teal } },
    });

    const SRCH_COLS = [
      { width: 20 }, { width: 24 }, { width: 10 }, { width: 12 }, { width: 34 },
      { width: 20 }, { width: 28 }, { width: 28 }, { width:  8 }, { width: 13 },
      { width: 10 }, { width: 13 }, { width: 15 }, { width: 10 }, { width: 32 },
    ];
    SRCH_COLS.forEach((col, ci) => { ws2.getColumn(ci + 1).width = col.width; });

    // 1행: 제목
    ws2.getRow(1).height = 34;
    ws2.mergeCells(1, 1, 1, 15);
    sc(ws2.getCell('A1'), `🔍  인텍플러스 구매관리 — 검색 & 필터  |  ${TODAY}`, {
      fnt: font(true, 13, C.white), fll: fill(C.navy), aln: al('left'),
    });

    // 2행: 구분선
    ws2.getRow(2).height = 10;
    ws2.mergeCells(2, 1, 2, 15);
    ws2.getCell('A2').fill = fill(C.gray100);

    // 3행: 섹션 제목
    ws2.getRow(3).height = 24;
    ws2.mergeCells(3, 1, 3, 8);
    sc(ws2.getCell('A3'), '🔎  검색 조건 입력  (값을 지우면 전체 표시)', {
      fnt: font(true, 11, C.navy), fll: fill(C.bluePale), aln: al('left', false, 1),
    });

    // 4행: 검색 레이블, 5행: 입력칸
    // 컬럼 배치: 1=사이트, 3=문서번호, 5=기안자, 7=프로젝트명, 9=코드, 11=품명, 13=규격, 15=입고완료
    ws2.getRow(4).height = 22;
    ws2.getRow(5).height = 30;

    const SEARCH_FIELDS = [
      [1,  '사이트',    '부분일치'],
      [3,  '문서번호',  '부분일치'],
      [5,  '기안자',    '부분일치'],
      [7,  '프로젝트명','부분일치'],
      [9,  '코드',      '부분일치'],
      [11, '품명',      '부분일치'],
      [13, '규격',      '부분일치'],
      [15, '입고완료',  '완료 / 미완료'],
    ];
    // 구분 컬럼(2,4,6,8,10,12,14) 좁게
    [2,4,6,8,10,12,14].forEach(ci => { ws2.getColumn(ci).width = 1.5; });

    SEARCH_FIELDS.forEach(([ci, label, note]) => {
      ws2.mergeCells(4, ci, 4, ci + 1);
      ws2.mergeCells(5, ci, 5, ci + 1);
      sc(ws2.getCell(4, ci), `${label}  (${note})`, {
        fnt: font(true, 9, C.navy), fll: fill(C.blueLt),
        aln: al('center', true), bdr: border(C.blue),
      });
      sc(ws2.getCell(5, ci), '', {
        fnt: font(false, 10, C.gray900), fll: fill(C.white),
        aln: al('left', false, 1),
        bdr: {
          top:    { style: 'medium', color: { argb: 'FF' + C.blue } },
          bottom: { style: 'medium', color: { argb: 'FF' + C.blue } },
          left:   { style: 'medium', color: { argb: 'FF' + C.blue } },
          right:  { style: 'medium', color: { argb: 'FF' + C.blue } },
        },
      });
    });

    // 6행: 안내 문구
    ws2.getRow(6).height = 18;
    ws2.mergeCells(6, 1, 6, 15);
    sc(ws2.getCell('A6'), '  ※ 각 조건을 입력하면 아래 결과가 자동 필터링됩니다. 비워두면 전체 표시.', {
      fnt: font(false, 9, C.gray500, true), fll: fill(C.gray50), aln: al('left'),
    });

    // 7행: 공백
    ws2.getRow(7).height = 8;

    // 8행: 결과 헤더
    ws2.getRow(8).height = 30;
    const SRCH_HDRS = ['No.','문서번호','기안자','기안일','프로젝트명','코드','품명','규격','수량',
                       '입고예정일','입고완료','배송시작일','배송완료예정일','배송완료','메모'];
    SRCH_HDRS.forEach((h, ci) => {
      sc(ws2.getCell(8, ci + 1), h, {
        fnt: font(true, 10, C.white),
        fll: fill(C.navy),
        aln: al('center', true),
        bdr: hdrBorder(C.teal),
      });
    });

    // 수식 파라미터
    const N  = items.length;
    const DA = `'📦 데이터'!$A$3:$A$${2 + N}`;
    const DB = `'📦 데이터'!$B$3:$B$${2 + N}`;
    const DC = `'📦 데이터'!$C$3:$C$${2 + N}`;
    const DD = `'📦 데이터'!$D$3:$D$${2 + N}`;
    const DE = `'📦 데이터'!$E$3:$E$${2 + N}`;
    const DF = `'📦 데이터'!$F$3:$F$${2 + N}`;
    const DG = `'📦 데이터'!$G$3:$G$${2 + N}`;
    const DH = `'📦 데이터'!$H$3:$H$${2 + N}`;
    const DI = `'📦 데이터'!$I$3:$I$${2 + N}`;
    const DJ = `'📦 데이터'!$J$3:$J$${2 + N}`;
    const DK = `'📦 데이터'!$K$3:$K$${2 + N}`;
    const DL = `'📦 데이터'!$L$3:$L$${2 + N}`;
    const DM = `'📦 데이터'!$M$3:$M$${2 + N}`;
    const DN = `'📦 데이터'!$N$3:$N$${2 + N}`;
    const DO = `'📦 데이터'!$O$3:$O$${2 + N}`;
    const ROW_REF = `'📦 데이터'!$B$2`;

    // 조건 배열 (8개 조건 — 5행 입력칸 참조)
    const COND = [
      `(($A$5="")+ISNUMBER(SEARCH($A$5,${DA})))`,
      `(($C$5="")+ISNUMBER(SEARCH($C$5,${DB})))`,
      `(($E$5="")+ISNUMBER(SEARCH($E$5,${DC})))`,
      `(($G$5="")+ISNUMBER(SEARCH($G$5,${DE})))`,
      `(($I$5="")+ISNUMBER(SEARCH($I$5,${DF})))`,
      `(($K$5="")+ISNUMBER(SEARCH($K$5,${DG})))`,
      `(($M$5="")+ISNUMBER(SEARCH($M$5,${DH})))`,
      `(($O$5="")+ISNUMBER(SEARCH($O$5,${DK})))`,
    ].join('*');

    const ROW_NUMS = `ROW(${DB})-ROW(${ROW_REF})`;

    // 결과 행 수식 삽입
    const MAX_ROWS = N + 20;
    for (let i = 1; i <= MAX_ROWS; i++) {
      const r   = 8 + i;
      const isE = i % 2 === 0;
      ws2.getRow(r).height = 22;

      const AGG = `AGGREGATE(15,6,${ROW_NUMS}/(${COND}>0),${i})`;
      const IE  = (d) => `=IFERROR(INDEX(${d},${AGG}),"")`;
      const bg  = fill(isE ? C.gray50 : C.bluePale);
      const bgB = fill(C.bluePale);

      const cells = [
        { v: `=IFERROR(${AGG},"")`,  fnt: font(false,9,C.gray500),  fll: bg,  aln: al('center') },
        { v: IE(DB),                  fnt: font(true, 10,C.navy),    fll: bgB, aln: al('left',false,1) },
        { v: IE(DC),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DD),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DE),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('left',false,1) },
        { v: IE(DF),                  fnt: font(false,10,C.gray500), fll: bg,  aln: al('center') },
        { v: IE(DG),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('left',false,1) },
        { v: IE(DH),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('left',false,1) },
        { v: IE(DI),                  fnt: font(true, 10,C.blue),    fll: bgB, aln: al('center') },
        { v: IE(DJ),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DK),                  fnt: font(true, 10,C.amber),   fll: bg,  aln: al('center') },
        { v: IE(DL),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DM),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DN),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('center') },
        { v: IE(DO),                  fnt: font(false,10,C.gray900), fll: bg,  aln: al('left',false,1) },
      ];

      cells.forEach(({v, fnt: f, fll: fl, aln: a}, ci) => {
        sc(ws2.getCell(r, ci + 1), v, { fnt: f, fll: fl, aln: a, bdr: btmBorder() });
      });
    }

    ws2.autoFilter = { from: { row: 8, column: 1 }, to: { row: 8 + MAX_ROWS, column: 15 } };


    // ════════════════════════════════════════════════════════
    // 시트 3: 📊 통계
    // ════════════════════════════════════════════════════════
    const ws3 = wb.addWorksheet('📊 통계', {
      views: [{ showGridLines: false }],
      properties: { tabColor: { argb: 'FF' + C.green } },
    });
    [22,13,13,13,13,22,13,13,13,13,14,14].forEach((w,ci) => { ws3.getColumn(ci+1).width = w; });

    // 1행: 제목
    ws3.getRow(1).height = 34;
    ws3.mergeCells(1, 1, 1, 12);
    sc(ws3.getCell('A1'), `📊  인텍플러스 구매관리 — 통계 요약  |  ${TODAY}`, {
      fnt: font(true, 13, C.white), fll: fill(C.navy), aln: al('left'),
    });
    ws3.getRow(2).height = 12;

    // KPI 계산
    const totalQty  = items.reduce((s, r) => s + (r.qty || 0), 0);
    const arrDone   = items.filter(r => (statusMap[r.doc_no]||{}).arrival_done).length;
    const arrPend   = items.length - arrDone;
    const shipDone  = items.filter(r => (statusMap[r.doc_no]||{}).ship_done).length;

    const KPIS = [
      { label: '전체 품목',    val: items.length, color: C.navy,    col: 1  },
      { label: '문서 수',      val: docNos.length,color: C.blue,    col: 3  },
      { label: '입고완료',     val: arrDone,       color: C.green,   col: 5  },
      { label: '입고 미완료',  val: arrPend,       color: C.amber,   col: 7  },
      { label: '배송완료',     val: shipDone,      color: C.teal,    col: 9  },
      { label: '총 수량',      val: totalQty,      color: C.gray700, col: 11 },
    ];
    ws3.getRow(3).height = 28;
    ws3.getRow(4).height = 44;
    KPIS.forEach(({ label, val, color, col }) => {
      ws3.mergeCells(3, col, 3, col + 1);
      ws3.mergeCells(4, col, 4, col + 1);
      sc(ws3.getCell(3, col), label, { fnt: font(true,9,C.white),  fll: fill(color), aln: al('center') });
      sc(ws3.getCell(4, col), val,   { fnt: font(true,22,C.white), fll: fill(color), aln: al('center') });
    });
    ws3.getRow(5).height = 14;

    // 기안자별 통계 계산
    const userMap = {};
    items.forEach(r => {
      const u  = r.requester || '(미상)';
      const st = statusMap[r.doc_no] || {};
      if (!userMap[u]) userMap[u] = { items: 0, arrDone: 0, shipDone: 0, qty: 0 };
      userMap[u].items++;
      if (st.arrival_done) userMap[u].arrDone++;
      if (st.ship_done)    userMap[u].shipDone++;
      userMap[u].qty += (r.qty || 0);
    });

    ws3.getRow(6).height = 26;
    ws3.mergeCells(6, 1, 6, 5);
    sc(ws3.getCell('A6'), '👤  기안자별 현황', { fnt: font(true,11,C.white), fll: fill(C.navy), aln: al('left',false,1) });
    ws3.getRow(7).height = 24;
    ['기안자','품목 수','입고완료','배송완료','총수량'].forEach((h,ci) => {
      sc(ws3.getCell(7, ci+1), h, { fnt: font(true,10,C.white), fll: fill(C.gray700), aln: al('center'), bdr: border(C.gray700) });
    });
    Object.entries(userMap).forEach(([u, v], ri) => {
      const r = 8 + ri;
      ws3.getRow(r).height = 22;
      const isE = ri % 2 === 0;
      const bg  = fill(isE ? C.gray100 : C.white);
      const all = v.items === v.arrDone;
      sc(ws3.getCell(r,1), u,          { fnt: font(true,10,C.navy),  fll: bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,2), v.items,    { fnt: font(false,10,C.gray900), fll: bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,3), v.arrDone,  { fnt: font(true,10,all?C.green:C.amber), fll: all?fill(C.greenLt):bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,4), v.shipDone, { fnt: font(false,10,C.gray900), fll: bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,5), v.qty,      { fnt: font(true,10,C.blue), fll: fill(C.bluePale), aln: al('center'), bdr: btmBorder() });
    });

    const projStartRow = 9 + Object.keys(userMap).length;
    ws3.getRow(projStartRow).height = 26;
    ws3.mergeCells(projStartRow, 1, projStartRow, 5);
    sc(ws3.getCell(projStartRow, 1), '📁  프로젝트별 현황', { fnt: font(true,11,C.white), fll: fill(C.navy), aln: al('left',false,1) });

    ws3.getRow(projStartRow + 1).height = 24;
    ['프로젝트명','품목 수','입고완료','입고율','총수량'].forEach((h,ci) => {
      sc(ws3.getCell(projStartRow+1, ci+1), h, { fnt: font(true,10,C.white), fll: fill(C.gray700), aln: al('center'), bdr: border(C.gray700) });
    });

    const projMap = {};
    items.forEach(r => {
      const p  = r.p_name || '(미상)';
      const st = statusMap[r.doc_no] || {};
      if (!projMap[p]) projMap[p] = { items: 0, arrDone: 0, qty: 0 };
      projMap[p].items++;
      if (st.arrival_done) projMap[p].arrDone++;
      projMap[p].qty += (r.qty || 0);
    });
    Object.entries(projMap).forEach(([p, v], ri) => {
      const r    = projStartRow + 2 + ri;
      const isE  = ri % 2 === 0;
      const bg   = fill(isE ? C.gray100 : C.white);
      const rate = v.items > 0 ? `${Math.round(v.arrDone / v.items * 100)}%` : '-';
      const all  = v.items === v.arrDone;
      ws3.getRow(r).height = 22;
      sc(ws3.getCell(r,1), p,         { fnt: font(true,10,C.teal),    fll: bg, aln: al('left',false,1), bdr: btmBorder() });
      sc(ws3.getCell(r,2), v.items,   { fnt: font(false,10,C.gray900), fll: bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,3), v.arrDone, { fnt: font(true,10,all?C.green:C.amber), fll: all?fill(C.greenLt):bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,4), rate,      { fnt: font(true,10,all?C.green:C.amber), fll: all?fill(C.greenLt):bg, aln: al('center'), bdr: btmBorder() });
      sc(ws3.getCell(r,5), v.qty,     { fnt: font(true,10,C.blue),    fll: fill(C.bluePale), aln: al('center'), bdr: btmBorder() });
    });


    // ════════════════════════════════════════════════════════
    // 시트 4: 📌 가이드
    // ════════════════════════════════════════════════════════
    const ws4 = wb.addWorksheet('📌 가이드', {
      views: [{ showGridLines: false }],
      properties: { tabColor: { argb: 'FF' + C.amber } },
    });
    ws4.getColumn(1).width = 4;
    ws4.getColumn(2).width = 22;
    ws4.getColumn(3).width = 56;
    ws4.getColumn(4).width = 16;

    ws4.getRow(1).height = 34;
    ws4.mergeCells(1, 1, 1, 4);
    sc(ws4.getCell('A1'), '📌  인텍플러스 구매관리 엑셀 — 사용 가이드', {
      fnt: font(true, 13, C.white), fll: fill(C.navy), aln: al('left'),
    });

    const GUIDE = [
      ['sec', '📦 시트 구성'],
      ['hdr', '시트명', '설명'],
      ['row', '',  '📦 데이터',   '원본 데이터. AutoFilter + 틀고정. 짝홀수 행 교번 배색'],
      ['row', '',  '🔍 검색',     '8개 조건 실시간 다중 필터링 (AGGREGATE 수식)'],
      ['row', '',  '📊 통계',     'KPI 5종 카드 + 기안자별 · 프로젝트별 집계'],
      ['row', '',  '📌 가이드',   '이 파일'],
      ['sp'],
      ['sec', '🔍 검색 시트 사용법'],
      ['row', '①', '단일 조건',  '품명 칸에 "ROLLER" 입력 → 해당 품목만 표시'],
      ['row', '②', '다중 조건',  '기안자 + 입고완료 동시 입력 → 모두 만족하는 항목 표시'],
      ['row', '③', '부분 일치',  '"IBAZ" → IBAZ-DSK, IBAZ-MSR 등 모두 검색'],
      ['row', '④', '초기화',     '5행 조건 셀을 Delete 키로 비우면 전체 표시'],
      ['sp'],
      ['sec', '⚠️ 주의사항'],
      ['warn','!', '수식 보호',   '9행 이하 결과 셀의 수식을 삭제하지 마세요'],
      ['warn','!', '스냅샷 파일', '내보내기 시점 데이터입니다. 실시간 연동 아님'],
      ['warn','!', '헤더 수정 금지', '2행 헤더를 수정하면 검색 수식이 깨집니다'],
    ];

    let gr = 2;
    GUIDE.forEach(g => {
      ws4.getRow(gr).height = g[0] === 'sec' ? 28 : g[0] === 'sp' ? 10 : 24;
      if (g[0] === 'sp') { gr++; return; }
      if (g[0] === 'sec') {
        ws4.mergeCells(gr, 1, gr, 4);
        sc(ws4.getCell(gr, 1), g[1], { fnt: font(true,11,C.white), fll: fill(C.navy), aln: al('left',false,1) });
      } else if (g[0] === 'hdr') {
        sc(ws4.getCell(gr, 2), g[1], { fnt: font(true,10,C.white), fll: fill(C.gray700), aln: al('center'), bdr: border(C.gray700) });
        ws4.mergeCells(gr, 3, gr, 4);
        sc(ws4.getCell(gr, 3), g[2], { fnt: font(true,10,C.white), fll: fill(C.gray700), aln: al('center'), bdr: border(C.gray700) });
      } else if (g[0] === 'row') {
        sc(ws4.getCell(gr, 1), g[1], { fnt: font(true,11,C.blue), aln: al('center') });
        sc(ws4.getCell(gr, 2), g[2], { fnt: font(true,10,C.gray900), fll: fill(C.blueLt), aln: al('left',false,1), bdr: border(C.gray300) });
        ws4.mergeCells(gr, 3, gr, 4);
        sc(ws4.getCell(gr, 3), g[3]||'', { fnt: font(false,10,C.gray700), fll: fill(C.gray50), aln: al('left',false,1), bdr: border(C.gray300) });
      } else if (g[0] === 'warn') {
        sc(ws4.getCell(gr, 1), g[1], { fnt: font(true,13,C.amber), aln: al('center') });
        sc(ws4.getCell(gr, 2), g[2], { fnt: font(true,10,C.amber), fll: fill(C.amberLt), aln: al('left',false,1), bdr: border(C.amberLt) });
        ws4.mergeCells(gr, 3, gr, 4);
        sc(ws4.getCell(gr, 3), g[3]||'', { fnt: font(true,10,C.amber), fll: fill(C.amberLt), aln: al('left',false,1), bdr: border(C.amberLt) });
      }
      gr++;
    });

    // 검색 시트를 기본 활성화
    wb.views = [{ activeTab: 1 }];

    // ── 응답 전송 ──────────────────────────────────────────
    const dateStr = TODAY;
    const filename = encodeURIComponent(`인텍플러스_구매관리_${dateStr}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);

    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('export handler error:', err);
    // 헤더가 이미 나간 경우 대비
    if (!res.headersSent) {
      res.status(500).json({ error: '엑셀 생성 실패', detail: err.message });
    }
  }
}
