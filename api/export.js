// Vercel Serverless Function
// 엔드포인트: POST /api/export
// Body: { items: [...], statusMap: {...}, searchMeta: { cond, summary } }
// Response: xlsx 파일
//
// 의존성: exceljs  →  package.json에 "exceljs": "^4.4.0" 필요

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
  gray200:   'E5E7EB',
  gray100:   'F3F4F6',
  gray50:    'F9FAFB',
  white:     'FFFFFF',
};

// ── 스타일 헬퍼 ──────────────────────────────────────────
const fnt = (bold=false, size=10, color=C.gray900, italic=false) =>
  ({ name:'Arial', bold, size, color:{ argb:'FF'+color }, italic });

const fll = (hex) =>
  ({ type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+hex } });

const aln = (horizontal='center', wrapText=false, indent=0) =>
  ({ horizontal, vertical:'middle', wrapText, indent });

const bdrAll = (color=C.gray200) => {
  const s = { style:'thin', color:{ argb:'FF'+color } };
  return { top:s, bottom:s, left:s, right:s };
};

const bdrBtm = () => ({
  bottom: { style:'thin',  color:{ argb:'FF'+C.gray200 } },
  left:   { style:'thin',  color:{ argb:'FF'+C.gray200 } },
  right:  { style:'thin',  color:{ argb:'FF'+C.gray200 } },
});

const bdrHdr = (c=C.blue) => ({
  top:    { style:'thin',   color:{ argb:'FF'+c } },
  bottom: { style:'medium', color:{ argb:'FF'+c } },
  left:   { style:'thin',   color:{ argb:'FF'+c } },
  right:  { style:'thin',   color:{ argb:'FF'+c } },
});

// 셀에 값 + 스타일 한 번에 적용
// ★ 수식(= 로 시작)은 자동으로 { formula } 형식으로 변환
function sc(cell, value, { f, fl, al, bd } = {}) {
  if (typeof value === 'string' && value.startsWith('=')) {
    cell.value = { formula: value.slice(1) };
  } else {
    cell.value = value;
  }
  if (f)  cell.font      = f;
  if (fl) cell.fill      = fl;
  if (al) cell.alignment = al;
  if (bd) cell.border    = bd;
}

const today = () => new Date().toISOString().slice(0, 10);

function arrSt(st) {
  if (st.arrival_done)                              return 'done';
  if (!st.expected_arrival)                         return 'none';
  if (st.expected_arrival < today())                return 'late';
  return 'pend';
}
function shipSt(st) {
  if (st.ship_done)                                 return 'done';
  if (!st.ship_done_date) return st.ship_start_date ? 'started' : 'none';
  if (st.ship_done_date < today())                  return 'late';
  return 'pend';
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
    const { items = [], statusMap = {} } = req.body || {};
    if (!items.length)
      return res.status(400).json({ error: '내보낼 데이터가 없습니다.' });

    const TODAY  = today();
    const docNos = [...new Set(items.map(r => r.doc_no))];
    const wb     = new ExcelJS.Workbook();
    wb.creator   = '인텍플러스 구매관리';
    wb.created   = new Date();

    // ══════════════════════════════════════════════════════
    // 시트명 주의: 이모지 없이 한글만 사용해야
    //   수식 내 시트 참조가 외부 링크로 오인되지 않음
    // ── 시트1: 데이터 ────────────────────────────────────
    // ══════════════════════════════════════════════════════
    const ws1 = wb.addWorksheet('데이터', {
      views: [{ showGridLines:false, state:'frozen', xSplit:2, ySplit:2 }],
      properties: { tabColor: { argb:'FF'+C.navy } },
    });

    const DCOLS = [
      { header:'사이트',        width:20 },
      { header:'문서번호',      width:24 },
      { header:'기안자',        width:10 },
      { header:'기안일',        width:12 },
      { header:'프로젝트명',    width:34 },
      { header:'코드',          width:20 },
      { header:'품명',          width:28 },
      { header:'규격',          width:28 },
      { header:'수량',          width: 8 },
      { header:'입고예정일',    width:13 },
      { header:'입고완료',      width:10 },
      { header:'배송시작일',    width:13 },
      { header:'배송완료예정일',width:15 },
      { header:'배송완료',      width:10 },
      { header:'메모',          width:32 },
    ];
    DCOLS.forEach((col, ci) => { ws1.getColumn(ci+1).width = col.width; });

    // 1행: 제목 배너
    ws1.getRow(1).height = 34;
    ws1.mergeCells(1, 1, 1, DCOLS.length);
    sc(ws1.getCell('A1'), `🛒  인텍플러스 구매관리 — 데이터 원본  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    // 2행: 헤더
    ws1.getRow(2).height = 32;
    DCOLS.forEach((col, ci) => {
      sc(ws1.getCell(2, ci+1), col.header, {
        f: fnt(true,10,C.white), fl: fll(C.navy),
        al: aln('center',true), bd: bdrHdr(C.blue),
      });
    });

    // 3행~: 데이터
    items.forEach((row, ri) => {
      const r   = ri + 3;
      const isE = ri % 2 === 0;
      const st  = statusMap[row.doc_no] || {};
      ws1.getRow(r).height = 22;

      const as = arrSt(st);
      const ss = shipSt(st);

      const arrText  = as==='done' ? '✅ 완료' : as==='late' ? '⚠️ 지연' : '⏳ 미완료';
      const arrF     = as==='done' ? fnt(true,10,C.green) : as==='late' ? fnt(true,10,C.red) : fnt(true,10,C.amber);
      const arrFl    = as==='done' ? fll(C.greenLt)       : as==='late' ? fll(C.redLt)       : fll(C.amberLt);

      const shipText = ss==='done' ? '✅ 완료' : ss==='late' ? '❓ 확인필요' : ss==='started' ? '🚚 배송중' : '—';
      const shipF    = ss==='done' ? fnt(true,10,C.green) : ss==='late' ? fnt(true,10,C.red) : ss==='started' ? fnt(true,10,C.blue) : fnt(false,10,C.gray500);
      const shipFl   = ss==='done' ? fll(C.greenLt)       : ss==='late' ? fll(C.redLt)       : ss==='started' ? fll(C.bluePale)    : fll(isE?C.gray50:C.white);

      const bg  = fll(isE ? C.gray50  : C.white);
      const bgB = fll(C.bluePale);

      [
        { v: row.site||'',             f: fnt(false,10,C.gray500), fl: bg,                  al: aln('left',false,1) },
        { v: row.doc_no||'',           f: fnt(true, 10,C.navy),    fl: isE ? bgB : bg,       al: aln('left',false,1) },
        { v: row.requester||'',        f: fnt(false,10,C.gray900), fl: bg,                  al: aln('center') },
        { v: row.doc_date||'',         f: fnt(false,10,C.gray900), fl: bg,                  al: aln('center') },
        { v: row.p_name||'',           f: fnt(false,10,C.gray900), fl: bg,                  al: aln('left',false,1) },
        { v: row.p_code||'',           f: fnt(false,10,C.gray500), fl: bg,                  al: aln('center') },
        { v: row.name||'',             f: fnt(false,10,C.gray900), fl: bg,                  al: aln('left',false,1) },
        { v: row.spec||'',             f: fnt(false,10,C.gray900), fl: bg,                  al: aln('left',false,1) },
        { v: row.qty||0,               f: fnt(true, 10,C.blue),    fl: bgB,                 al: aln('center') },
        { v: st.expected_arrival||'',  f: fnt(false,10,C.gray900), fl: bg,                  al: aln('center') },
        { v: arrText,                  f: arrF,                    fl: arrFl,               al: aln('center') },
        { v: st.ship_start_date||'',   f: fnt(false,10,C.gray900), fl: bg,                  al: aln('center') },
        { v: st.ship_done_date||'',    f: fnt(false,10,C.gray900), fl: bg,                  al: aln('center') },
        { v: shipText,                 f: shipF,                   fl: shipFl,              al: aln('center') },
        { v: st.memo||'',              f: fnt(false,10,C.gray900), fl: bg,                  al: aln('left',false,1) },
      ].forEach(({ v, f, fl, al: a }, ci) => {
        sc(ws1.getCell(r, ci+1), v, { f, fl, al: a, bd: bdrBtm() });
      });
    });

    ws1.autoFilter = {
      from: { row:2, column:1 },
      to:   { row:2+items.length, column:DCOLS.length },
    };

    // ══════════════════════════════════════════════════════
    // 시트2: 검색
    // ★ 수식에서 시트 참조: 데이터!$A$3 (이모지 없는 한글)
    // ══════════════════════════════════════════════════════
    const ws2 = wb.addWorksheet('검색', {
      views: [{ showGridLines:false, state:'frozen', ySplit:8 }],
      properties: { tabColor: { argb:'FF'+C.teal } },
    });

    const SCOLS = [20,24,10,12,34,20,28,28,8,13,10,13,15,10,32];
    SCOLS.forEach((w, ci) => { ws2.getColumn(ci+1).width = w; });

    // 1행: 제목
    ws2.getRow(1).height = 34;
    ws2.mergeCells(1,1,1,15);
    sc(ws2.getCell('A1'), `🔍  인텍플러스 구매관리 — 검색 & 필터  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    // 2행: 구분선
    ws2.getRow(2).height = 10;
    ws2.mergeCells(2,1,2,15);
    ws2.getCell('A2').fill = fll(C.gray100);

    // 3행: 섹션 제목
    ws2.getRow(3).height = 24;
    ws2.mergeCells(3,1,3,8);
    sc(ws2.getCell('A3'), '🔎  검색 조건 입력  (값을 지우면 전체 표시)', {
      f: fnt(true,11,C.navy), fl: fll(C.bluePale), al: aln('left',false,1),
    });

    // 4행: 검색 레이블 / 5행: 입력칸
    ws2.getRow(4).height = 22;
    ws2.getRow(5).height = 30;

    // 구분용 좁은 컬럼
    [2,4,6,8,10,12,14].forEach(ci => { ws2.getColumn(ci).width = 1.5; });

    const SFIELDS = [
      [1,  '사이트',    '부분일치'],
      [3,  '문서번호',  '부분일치'],
      [5,  '기안자',    '부분일치'],
      [7,  '프로젝트명','부분일치'],
      [9,  '코드',      '부분일치'],
      [11, '품명',      '부분일치'],
      [13, '규격',      '부분일치'],
      [15, '입고완료',  '완료 / 미완료'],
    ];
    SFIELDS.forEach(([ci, label, note]) => {
      ws2.mergeCells(4, ci, 4, ci+1);
      ws2.mergeCells(5, ci, 5, ci+1);
      sc(ws2.getCell(4, ci), `${label}  (${note})`, {
        f: fnt(true,9,C.navy), fl: fll(C.blueLt),
        al: aln('center',true), bd: bdrAll(C.blue),
      });
      sc(ws2.getCell(5, ci), '', {
        f: fnt(false,10,C.gray900), fl: fll(C.white), al: aln('left',false,1),
        bd: {
          top:    { style:'medium', color:{ argb:'FF'+C.blue } },
          bottom: { style:'medium', color:{ argb:'FF'+C.blue } },
          left:   { style:'medium', color:{ argb:'FF'+C.blue } },
          right:  { style:'medium', color:{ argb:'FF'+C.blue } },
        },
      });
    });

    // 6행: 안내
    ws2.getRow(6).height = 18;
    ws2.mergeCells(6,1,6,15);
    sc(ws2.getCell('A6'),
      '  ※ 각 조건을 입력하면 아래 결과가 자동 필터링됩니다. 비워두면 전체 표시.', {
      f: fnt(false,9,C.gray500,true), fl: fll(C.gray50), al: aln('left'),
    });

    // 7행: 공백
    ws2.getRow(7).height = 8;

    // 8행: 결과 헤더
    ws2.getRow(8).height = 30;
    ['No.','문서번호','기안자','기안일','프로젝트명','코드','품명','규격','수량',
     '입고예정일','입고완료','배송시작일','배송완료예정일','배송완료','메모']
      .forEach((h, ci) => {
        sc(ws2.getCell(8, ci+1), h, {
          f: fnt(true,10,C.white), fl: fll(C.navy),
          al: aln('center',true), bd: bdrHdr(C.teal),
        });
      });

    // ── 수식 (9행~) ──
    // 시트 참조: 이모지 없는 한글 시트명 "데이터"
    const N   = items.length;
    const ref = (col) => `데이터!$${col}$3:$${col}$${2+N}`;

    // 8개 조건 — 검색 입력칸(5행) 참조
    const COND = [
      `(($A$5="")+ISNUMBER(SEARCH($A$5,${ref('A')})))`,  // 사이트
      `(($C$5="")+ISNUMBER(SEARCH($C$5,${ref('B')})))`,  // 문서번호
      `(($E$5="")+ISNUMBER(SEARCH($E$5,${ref('C')})))`,  // 기안자
      `(($G$5="")+ISNUMBER(SEARCH($G$5,${ref('E')})))`,  // 프로젝트명
      `(($I$5="")+ISNUMBER(SEARCH($I$5,${ref('F')})))`,  // 코드
      `(($K$5="")+ISNUMBER(SEARCH($K$5,${ref('G')})))`,  // 품명
      `(($M$5="")+ISNUMBER(SEARCH($M$5,${ref('H')})))`,  // 규격
      `(($O$5="")+ISNUMBER(SEARCH($O$5,${ref('K')})))`,  // 입고완료
    ].join('*');

    const ROWNUM = `ROW(데이터!$B$3:$B$${2+N})-ROW(데이터!$B$2)`;

    const MAX = N + 20;
    for (let i = 1; i <= MAX; i++) {
      const r   = 8 + i;
      const isE = i % 2 === 0;
      ws2.getRow(r).height = 22;

      const AGG = `AGGREGATE(15,6,${ROWNUM}/(${COND}>0),${i})`;
      const IE  = (col) => `=IFERROR(INDEX(${ref(col)},${AGG}),"")`;

      const rowDef = [
        { v:`=IFERROR(${AGG},"")`,  f:fnt(false,9,C.gray500),  fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('B'), f:fnt(true,10,C.navy),    fl:fll(C.bluePale),              al:aln('left',false,1) },
        { v:IE('C'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('D'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('E'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('left',false,1) },
        { v:IE('F'), f:fnt(false,10,C.gray500),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('G'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('left',false,1) },
        { v:IE('H'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('left',false,1) },
        { v:IE('I'), f:fnt(true,10,C.blue),    fl:fll(C.bluePale),              al:aln('center') },
        { v:IE('J'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('K'), f:fnt(true,10,C.amber),   fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('L'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('M'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('N'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('center') },
        { v:IE('O'), f:fnt(false,10,C.gray900),fl:fll(isE?C.gray50:C.bluePale), al:aln('left',false,1) },
      ];

      rowDef.forEach(({ v, f, fl, al: a }, ci) => {
        sc(ws2.getCell(r, ci+1), v, { f, fl, al: a, bd: bdrBtm() });
      });
    }

    ws2.autoFilter = {
      from: { row:8, column:1 },
      to:   { row:8+MAX, column:15 },
    };

    // ══════════════════════════════════════════════════════
    // 시트3: 통계
    // ══════════════════════════════════════════════════════
    const ws3 = wb.addWorksheet('통계', {
      views: [{ showGridLines:false }],
      properties: { tabColor: { argb:'FF'+C.green } },
    });
    [22,13,13,13,13,22,13,13,13,13,14,14].forEach((w,ci) => { ws3.getColumn(ci+1).width = w; });

    ws3.getRow(1).height = 34;
    ws3.mergeCells(1,1,1,12);
    sc(ws3.getCell('A1'), `📊  인텍플러스 구매관리 — 통계 요약  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });
    ws3.getRow(2).height = 12;

    const totalQty = items.reduce((s,r) => s+(r.qty||0), 0);
    const nArrDone = items.filter(r => (statusMap[r.doc_no]||{}).arrival_done).length;
    const nArrPend = items.length - nArrDone;
    const nShipDone= items.filter(r => (statusMap[r.doc_no]||{}).ship_done).length;

    const KPIS = [
      { label:'전체 품목',   val:items.length, color:C.navy,    col:1  },
      { label:'문서 수',     val:docNos.length,color:C.blue,    col:3  },
      { label:'입고완료',    val:nArrDone,      color:C.green,   col:5  },
      { label:'입고 미완료', val:nArrPend,      color:C.amber,   col:7  },
      { label:'배송완료',    val:nShipDone,     color:C.teal,    col:9  },
      { label:'총 수량',     val:totalQty,      color:C.gray700, col:11 },
    ];
    ws3.getRow(3).height = 28;
    ws3.getRow(4).height = 44;
    KPIS.forEach(({ label, val, color, col }) => {
      ws3.mergeCells(3, col, 3, col+1);
      ws3.mergeCells(4, col, 4, col+1);
      sc(ws3.getCell(3, col), label, { f:fnt(true,9,C.white),  fl:fll(color), al:aln('center') });
      sc(ws3.getCell(4, col), val,   { f:fnt(true,22,C.white), fl:fll(color), al:aln('center') });
    });
    ws3.getRow(5).height = 14;

    // 기안자별
    const uMap = {};
    items.forEach(r => {
      const u = r.requester||'(미상)'; const st = statusMap[r.doc_no]||{};
      if (!uMap[u]) uMap[u] = { i:0, a:0, s:0, q:0 };
      uMap[u].i++; if(st.arrival_done)uMap[u].a++; if(st.ship_done)uMap[u].s++; uMap[u].q+=(r.qty||0);
    });
    ws3.getRow(6).height = 26;
    ws3.mergeCells(6,1,6,5);
    sc(ws3.getCell('A6'), '👤  기안자별 현황', { f:fnt(true,11,C.white), fl:fll(C.navy), al:aln('left',false,1) });
    ws3.getRow(7).height = 24;
    ['기안자','품목 수','입고완료','배송완료','총수량'].forEach((h,ci) => {
      sc(ws3.getCell(7,ci+1), h, { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
    });
    Object.entries(uMap).forEach(([u,v],ri) => {
      const r = 8+ri; ws3.getRow(r).height = 22;
      const isE = ri%2===0; const bg = fll(isE?C.gray100:C.white); const all = v.i===v.a;
      sc(ws3.getCell(r,1), u,   { f:fnt(true,10,C.navy),    fl:bg,                        al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,2), v.i, { f:fnt(false,10,C.gray900),fl:bg,                        al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,3), v.a, { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg, al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,4), v.s, { f:fnt(false,10,C.gray900),fl:bg,                        al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,5), v.q, { f:fnt(true,10,C.blue),    fl:fll(C.bluePale),            al:aln('center'), bd:bdrBtm() });
    });

    // 프로젝트별
    const pMap = {};
    items.forEach(r => {
      const p = r.p_name||'(미상)'; const st = statusMap[r.doc_no]||{};
      if (!pMap[p]) pMap[p] = { i:0, a:0, q:0 };
      pMap[p].i++; if(st.arrival_done)pMap[p].a++; pMap[p].q+=(r.qty||0);
    });
    const pStart = 9 + Object.keys(uMap).length;
    ws3.getRow(pStart).height = 26;
    ws3.mergeCells(pStart,1,pStart,5);
    sc(ws3.getCell(pStart,1), '📁  프로젝트별 현황', { f:fnt(true,11,C.white), fl:fll(C.navy), al:aln('left',false,1) });
    ws3.getRow(pStart+1).height = 24;
    ['프로젝트명','품목 수','입고완료','입고율','총수량'].forEach((h,ci) => {
      sc(ws3.getCell(pStart+1,ci+1), h, { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
    });
    Object.entries(pMap).forEach(([p,v],ri) => {
      const r = pStart+2+ri; ws3.getRow(r).height = 22;
      const isE = ri%2===0; const bg = fll(isE?C.gray100:C.white); const all = v.i===v.a;
      const rate = v.i>0 ? `${Math.round(v.a/v.i*100)}%` : '-';
      sc(ws3.getCell(r,1), p,    { f:fnt(true,10,C.teal),     fl:bg,                        al:aln('left',false,1), bd:bdrBtm() });
      sc(ws3.getCell(r,2), v.i,  { f:fnt(false,10,C.gray900), fl:bg,                        al:aln('center'),       bd:bdrBtm() });
      sc(ws3.getCell(r,3), v.a,  { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg, al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,4), rate, { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg, al:aln('center'), bd:bdrBtm() });
      sc(ws3.getCell(r,5), v.q,  { f:fnt(true,10,C.blue),     fl:fll(C.bluePale),            al:aln('center'),       bd:bdrBtm() });
    });

    // ══════════════════════════════════════════════════════
    // 시트4: 가이드
    // ══════════════════════════════════════════════════════
    const ws4 = wb.addWorksheet('가이드', {
      views: [{ showGridLines:false }],
      properties: { tabColor: { argb:'FF'+C.amber } },
    });
    [4,22,56,16].forEach((w,ci) => { ws4.getColumn(ci+1).width = w; });

    ws4.getRow(1).height = 34;
    ws4.mergeCells(1,1,1,4);
    sc(ws4.getCell('A1'), '📌  인텍플러스 구매관리 엑셀 — 사용 가이드', {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    const GUIDE = [
      ['sec','📦 시트 구성'],
      ['hdr','시트명','설명'],
      ['row','','데이터','원본 데이터. AutoFilter + 틀고정. 짝/홀수 행 교번 배색'],
      ['row','','검색','8개 조건 실시간 다중 필터링 (AGGREGATE 수식)'],
      ['row','','통계','KPI 5종 카드 + 기안자별 · 프로젝트별 집계'],
      ['row','','가이드','이 파일'],
      ['sp'],
      ['sec','🔍 검색 시트 사용법'],
      ['row','①','단일 조건','품명 칸에 "ROLLER" 입력 → 해당 품목만 표시'],
      ['row','②','다중 조건','기안자 + 입고완료 동시 입력 → 모두 만족하는 항목 표시'],
      ['row','③','부분 일치','"IBAZ" → IBAZ-DSK, IBAZ-MSR 등 모두 검색'],
      ['row','④','초기화','5행 조건 셀을 Delete 키로 비우면 전체 표시'],
      ['sp'],
      ['sec','⚠️ 주의사항'],
      ['warn','!','수식 보호','9행 이하 결과 셀의 수식을 삭제하지 마세요'],
      ['warn','!','스냅샷 파일','내보내기 시점 데이터입니다. 실시간 연동 아님'],
      ['warn','!','헤더 수정 금지','2행 헤더를 수정하면 검색 수식이 깨집니다'],
    ];
    let gr = 2;
    GUIDE.forEach(g => {
      ws4.getRow(gr).height = g[0]==='sec' ? 28 : g[0]==='sp' ? 10 : 24;
      if (g[0]==='sp') { gr++; return; }
      if (g[0]==='sec') {
        ws4.mergeCells(gr,1,gr,4);
        sc(ws4.getCell(gr,1), g[1], { f:fnt(true,11,C.white), fl:fll(C.navy), al:aln('left',false,1) });
      } else if (g[0]==='hdr') {
        sc(ws4.getCell(gr,2), g[1], { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
        ws4.mergeCells(gr,3,gr,4);
        sc(ws4.getCell(gr,3), g[2], { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
      } else if (g[0]==='row') {
        sc(ws4.getCell(gr,1), g[1], { f:fnt(true,11,C.blue), al:aln('center') });
        sc(ws4.getCell(gr,2), g[2], { f:fnt(true,10,C.gray900), fl:fll(C.blueLt), al:aln('left',false,1), bd:bdrAll(C.gray200) });
        ws4.mergeCells(gr,3,gr,4);
        sc(ws4.getCell(gr,3), g[3]||'', { f:fnt(false,10,C.gray700), fl:fll(C.gray50), al:aln('left',false,1), bd:bdrAll(C.gray200) });
      } else if (g[0]==='warn') {
        sc(ws4.getCell(gr,1), g[1], { f:fnt(true,13,C.amber), al:aln('center') });
        sc(ws4.getCell(gr,2), g[2], { f:fnt(true,10,C.amber), fl:fll(C.amberLt), al:aln('left',false,1), bd:bdrAll(C.amberLt) });
        ws4.mergeCells(gr,3,gr,4);
        sc(ws4.getCell(gr,3), g[3]||'', { f:fnt(true,10,C.amber), fl:fll(C.amberLt), al:aln('left',false,1), bd:bdrAll(C.amberLt) });
      }
      gr++;
    });

    // 검색 시트를 기본 활성화
    wb.views = [{ activeTab: 1 }];

    // ── 응답 전송 ──────────────────────────────────────────
    const filename = encodeURIComponent(`인텍플러스_구매관리_${TODAY}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('export handler error:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: '엑셀 생성 실패', detail: err.message });
    }
  }
}
