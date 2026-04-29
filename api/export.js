// Vercel Serverless Function
// 엔드포인트: POST /api/export
// Body: { items: [...], statusMap: {...}, searchMeta: { cond, summary } }
// Response: xlsx 파일
//
// 의존성: exceljs → package.json에 "exceljs": "^4.4.0" 필요

import ExcelJS from 'exceljs';

// ── 색상 팔레트 ──────────────────────────────────────────
const C = {
  navy:     '1A3A5C',
  blue:     '1A56DB',
  blueLt:   'D6E4FF',
  bluePale: 'EBF5FF',
  teal:     '0694A2',
  tealLt:   'CCFBF1',
  green:    '057A55',
  greenLt:  'DEF7EC',
  amber:    'B45309',
  amberLt:  'FEF3C7',
  red:      'C81E1E',
  redLt:    'FDE8E8',
  gray900:  '111827',
  gray700:  '374151',
  gray500:  '6B7280',
  gray300:  'D1D5DB',
  gray200:  'E5E7EB',
  gray100:  'F3F4F6',
  gray50:   'F9FAFB',
  white:    'FFFFFF',
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

const bdrBtm = (color=C.gray200) => ({
  bottom: { style:'thin', color:{ argb:'FF'+color } },
  left:   { style:'thin', color:{ argb:'FF'+color } },
  right:  { style:'thin', color:{ argb:'FF'+color } },
});

const bdrHdr = (c=C.blue) => ({
  top:    { style:'thin',   color:{ argb:'FF'+c } },
  bottom: { style:'medium', color:{ argb:'FF'+c } },
  left:   { style:'thin',   color:{ argb:'FF'+c } },
  right:  { style:'thin',   color:{ argb:'FF'+c } },
});

// 셀에 값+스타일 일괄 적용
function sc(cell, value, { f, fl, al, bd } = {}) {
  cell.value = value ?? '';
  if (f)  cell.font      = f;
  if (fl) cell.fill      = fl;
  if (al) cell.alignment = al;
  if (bd) cell.border    = bd;
}

// 날짜 기준
// KST(한국 시간) 기준 날짜 반환 — Vercel 서버는 UTC이므로 +9시간 보정
const today = () => {
  const d = new Date();
  d.setHours(d.getHours() + 9); // UTC → KST
  return d.toISOString().slice(0, 10);
};

// 입고/배송 상태
function arrInfo(st) {
  if (st.arrival_done)                             return { text:'✅ 완료',     f:fnt(true,10,C.green),  fl:fll(C.greenLt) };
  if (!st.expected_arrival)                        return { text:'미정',        f:fnt(false,10,C.gray500),fl:null };
  if (st.expected_arrival < today())               return { text:'⚠️ 지연',    f:fnt(true,10,C.red),    fl:fll(C.redLt)   };
  if (st.expected_arrival === today())             return { text:'📦 오늘',     f:fnt(true,10,C.amber),  fl:fll(C.amberLt) };
  return                                                  { text:'⏳ 미완료',   f:fnt(true,10,C.amber),  fl:fll(C.amberLt) };
}

function shipInfo(st) {
  if (st.ship_done)                                return { text:'✅ 완료',     f:fnt(true,10,C.green),  fl:fll(C.greenLt)  };
  if (st.ship_done_date && st.ship_done_date<today()) return { text:'❓ 확인필요', f:fnt(true,10,C.red),    fl:fll(C.redLt)    };
  if (st.ship_start_date)                          return { text:'🚚 배송중',   f:fnt(true,10,C.blue),   fl:fll(C.bluePale) };
  return                                                  { text:'—',           f:fnt(false,10,C.gray500),fl:null            };
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

    // ════════════════════════════════════════════════════
    // 공통: 데이터 행 스타일 적용 함수
    // ════════════════════════════════════════════════════
    function writeDataRow(ws, r, rowData, isEven) {
      ws.getRow(r).height = 22;
      const bg = fll(isEven ? C.gray50 : C.white);

      rowData.forEach(({ v, f, fl, al: a, bd }, ci) => {
        const cell = ws.getCell(r, ci + 1);
        sc(cell, v, {
          f:  f  ?? fnt(false, 10, C.gray900),
          fl: fl ?? bg,
          al: a  ?? aln('center'),
          bd: bd ?? bdrBtm(),
        });
      });
    }

    // ════════════════════════════════════════════════════
    // 시트 1: 데이터 원본
    // ════════════════════════════════════════════════════
    const ws1 = wb.addWorksheet('데이터', {
      views: [{ showGridLines:false, state:'frozen', xSplit:2, ySplit:2 }],
      properties: { tabColor:{ argb:'FF'+C.navy } },
    });

    const D_WIDTHS = [20,24,10,12,34,20,28,28,8,13,10,13,15,10,32];
    const D_HDRS   = ['사이트','문서번호','기안자','기안일','프로젝트명','코드','품명','규격','수량',
                      '입고예정일','입고완료','배송시작일','배송완료예정일','배송완료','메모'];
    D_WIDTHS.forEach((w, ci) => { ws1.getColumn(ci+1).width = w; });

    // 1행: 제목 배너
    ws1.getRow(1).height = 34;
    ws1.mergeCells(1, 1, 1, D_HDRS.length);
    sc(ws1.getCell('A1'), `🛒  인텍플러스 구매관리 — 데이터 원본  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    // 2행: 컬럼 헤더
    ws1.getRow(2).height = 32;
    D_HDRS.forEach((h, ci) => {
      sc(ws1.getCell(2, ci+1), h, {
        f: fnt(true,10,C.white), fl: fll(C.navy),
        al: aln('center',true), bd: bdrHdr(C.blue),
      });
    });

    // 3행~: 데이터 값 직접 삽입
    items.forEach((row, ri) => {
      const r   = ri + 3;
      const isE = ri % 2 === 0;
      const st  = statusMap[row.doc_no] || {};
      const bg  = fll(isE ? C.gray50 : C.white);
      const bgB = fll(C.bluePale);
      const arr  = arrInfo(st);
      const ship = shipInfo(st);

      writeDataRow(ws1, r, [
        { v: row.site||'',             f: fnt(false,10,C.gray500), fl: bg,         al: aln('left',false,1) },
        { v: row.doc_no||'',           f: fnt(true, 10,C.navy),    fl: isE?bgB:bg, al: aln('left',false,1) },
        { v: row.requester||'',        fl: bg },
        { v: row.doc_date||'',         fl: bg },
        { v: row.p_name||'',           fl: bg,    al: aln('left',false,1) },
        { v: row.p_code||'',           f: fnt(false,10,C.gray500), fl: bg },
        { v: row.name||'',             fl: bg,    al: aln('left',false,1) },
        { v: row.spec||'',             fl: bg,    al: aln('left',false,1) },
        { v: row.qty||0,               f: fnt(true,10,C.blue), fl: bgB },
        { v: st.expected_arrival||'',  fl: bg },
        { v: arr.text,                 f: arr.f,  fl: arr.fl || bg },
        { v: st.ship_start_date||'',   fl: bg },
        { v: st.ship_done_date||'',    fl: bg },
        { v: ship.text,                f: ship.f, fl: ship.fl || bg },
        { v: st.memo||'',              fl: bg,    al: aln('left',false,1) },
      ], isE);
    });

    ws1.autoFilter = {
      from: { row:2, column:1 },
      to:   { row:2+items.length, column:D_HDRS.length },
    };

    // ════════════════════════════════════════════════════
    // 시트 2: 검색 (값으로 직접 + AutoFilter)
    // ★ 수식 미사용 → @AGGREGATE 오류 원천 차단
    // ★ 엑셀 AutoFilter로 각 컬럼별 검색/필터 가능
    // ════════════════════════════════════════════════════
    const ws2 = wb.addWorksheet('검색', {
      views: [{ showGridLines:false, state:'frozen', ySplit:3 }],
      properties: { tabColor:{ argb:'FF'+C.teal } },
    });

    const S_WIDTHS = [5,24,10,12,34,20,28,28,8,13,10,13,15,10,32];
    const S_HDRS   = ['No.','문서번호','기안자','기안일','프로젝트명','코드','품명','규격','수량',
                      '입고예정일','입고완료','배송시작일','배송완료예정일','배송완료','메모'];
    S_WIDTHS.forEach((w, ci) => { ws2.getColumn(ci+1).width = w; });

    // 1행: 제목 배너
    ws2.getRow(1).height = 34;
    ws2.mergeCells(1, 1, 1, S_HDRS.length);
    sc(ws2.getCell('A1'), `🔍  인텍플러스 구매관리 — 검색 & 필터  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    // 2행: 안내 문구
    ws2.getRow(2).height = 20;
    ws2.mergeCells(2, 1, 2, S_HDRS.length);
    sc(ws2.getCell('A2'),
      '  ※ 각 컬럼 헤더의 ▼ 버튼을 클릭하면 필터/검색이 가능합니다. (엑셀 기본 AutoFilter)', {
      f: fnt(false,9,C.gray500,true), fl: fll(C.gray50), al: aln('left'),
    });

    // 3행: 컬럼 헤더
    ws2.getRow(3).height = 30;
    S_HDRS.forEach((h, ci) => {
      sc(ws2.getCell(3, ci+1), h, {
        f: fnt(true,10,C.white), fl: fll(C.navy),
        al: aln('center',true), bd: bdrHdr(C.teal),
      });
    });

    // 4행~: 실제 데이터 값 직접 삽입
    items.forEach((row, ri) => {
      const r   = ri + 4;
      const isE = ri % 2 === 0;
      const st  = statusMap[row.doc_no] || {};
      const bg  = fll(isE ? C.gray50 : C.white);
      const bgB = fll(C.bluePale);
      const arr  = arrInfo(st);
      const ship = shipInfo(st);

      writeDataRow(ws2, r, [
        { v: ri+1,                     f: fnt(false,9,C.gray500),  fl: bg },
        { v: row.doc_no||'',           f: fnt(true, 10,C.navy),    fl: isE?fll(C.bluePale):bg, al: aln('left',false,1) },
        { v: row.requester||'',        fl: bg },
        { v: row.doc_date||'',         fl: bg },
        { v: row.p_name||'',           fl: bg,    al: aln('left',false,1) },
        { v: row.p_code||'',           f: fnt(false,10,C.gray500), fl: bg },
        { v: row.name||'',             fl: bg,    al: aln('left',false,1) },
        { v: row.spec||'',             fl: bg,    al: aln('left',false,1) },
        { v: row.qty||0,               f: fnt(true,10,C.blue), fl: bgB },
        { v: st.expected_arrival||'',  fl: bg },
        { v: arr.text,                 f: arr.f,  fl: arr.fl || bg },
        { v: st.ship_start_date||'',   fl: bg },
        { v: st.ship_done_date||'',    fl: bg },
        { v: ship.text,                f: ship.f, fl: ship.fl || bg },
        { v: st.memo||'',              fl: bg,    al: aln('left',false,1) },
      ], isE);
    });

    // AutoFilter: 헤더 행 기준으로 설정 → 각 컬럼 ▼ 버튼으로 필터 가능
    ws2.autoFilter = {
      from: { row:3, column:1 },
      to:   { row:3+items.length, column:S_HDRS.length },
    };

    // 합계 행
    const totalRow = 4 + items.length;
    ws2.getRow(totalRow).height = 26;
    ws2.mergeCells(totalRow, 1, totalRow, 8);
    sc(ws2.getCell(totalRow, 1), `총 ${items.length}건 / 문서 ${docNos.length}개`, {
      f: fnt(true,10,C.white), fl: fll(C.gray700), al: aln('center'),
    });
    const totalQtySum = items.reduce((s,r) => s+(r.qty||0), 0);
    sc(ws2.getCell(totalRow, 9), totalQtySum, {
      f: fnt(true,12,C.white), fl: fll(C.navy), al: aln('center'),
    });
    const arrDoneCnt = items.filter(r=>(statusMap[r.doc_no]||{}).arrival_done).length;
    ws2.mergeCells(totalRow, 10, totalRow, 11);
    sc(ws2.getCell(totalRow, 10), `입고완료 ${arrDoneCnt}건`, {
      f: fnt(true,10,C.white), fl: fll(C.green), al: aln('center'),
    });
    const shipDoneCnt = items.filter(r=>(statusMap[r.doc_no]||{}).ship_done).length;
    ws2.mergeCells(totalRow, 12, totalRow, 15);
    sc(ws2.getCell(totalRow, 12), `배송완료 ${shipDoneCnt}건`, {
      f: fnt(true,10,C.white), fl: fll(C.teal), al: aln('center'),
    });

    // ════════════════════════════════════════════════════
    // 시트 3: 통계
    // ════════════════════════════════════════════════════
    const ws3 = wb.addWorksheet('통계', {
      views: [{ showGridLines:false }],
      properties: { tabColor:{ argb:'FF'+C.green } },
    });
    [22,13,13,13,13,22,13,13,13,13,14,14].forEach((w,ci) => { ws3.getColumn(ci+1).width = w; });

    ws3.getRow(1).height = 34;
    ws3.mergeCells(1, 1, 1, 12);
    sc(ws3.getCell('A1'), `📊  인텍플러스 구매관리 — 통계 요약  |  ${TODAY}`, {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });
    ws3.getRow(2).height = 12;

    // KPI 카드
    const totalQty  = items.reduce((s,r) => s+(r.qty||0), 0);
    const nArrDone  = items.filter(r => (statusMap[r.doc_no]||{}).arrival_done).length;
    const nArrPend  = items.length - nArrDone;
    const nShipDone = items.filter(r => (statusMap[r.doc_no]||{}).ship_done).length;

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

    // 기안자별 집계
    const uMap = {};
    items.forEach(r => {
      const u = r.requester||'(미상)';
      const st = statusMap[r.doc_no]||{};
      if (!uMap[u]) uMap[u] = { i:0, a:0, s:0, q:0 };
      uMap[u].i++; if(st.arrival_done)uMap[u].a++; if(st.ship_done)uMap[u].s++; uMap[u].q+=(r.qty||0);
    });

    ws3.getRow(6).height = 26;
    ws3.mergeCells(6, 1, 6, 5);
    sc(ws3.getCell('A6'), '👤  기안자별 현황', { f:fnt(true,11,C.white), fl:fll(C.navy), al:aln('left',false,1) });
    ws3.getRow(7).height = 24;
    ['기안자','품목 수','입고완료','배송완료','총수량'].forEach((h, ci) => {
      sc(ws3.getCell(7, ci+1), h, { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
    });
    Object.entries(uMap).forEach(([u, v], ri) => {
      const r = 8 + ri;
      ws3.getRow(r).height = 22;
      const isE = ri%2===0;
      const bg  = fll(isE ? C.gray100 : C.white);
      const all = v.i === v.a;
      sc(ws3.getCell(r,1), u,   { f:fnt(true,10,C.navy),              fl:bg,                        al:aln('center'),     bd:bdrBtm() });
      sc(ws3.getCell(r,2), v.i, { f:fnt(false,10,C.gray900),          fl:bg,                        al:aln('center'),     bd:bdrBtm() });
      sc(ws3.getCell(r,3), v.a, { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg,     al:aln('center'),     bd:bdrBtm() });
      sc(ws3.getCell(r,4), v.s, { f:fnt(false,10,C.gray900),          fl:bg,                        al:aln('center'),     bd:bdrBtm() });
      sc(ws3.getCell(r,5), v.q, { f:fnt(true,10,C.blue),              fl:fll(C.bluePale),            al:aln('center'),     bd:bdrBtm() });
    });

    // 프로젝트별 집계
    const pMap = {};
    items.forEach(r => {
      const p = r.p_name||'(미상)';
      const st = statusMap[r.doc_no]||{};
      if (!pMap[p]) pMap[p] = { i:0, a:0, q:0 };
      pMap[p].i++; if(st.arrival_done)pMap[p].a++; pMap[p].q+=(r.qty||0);
    });
    const pStart = 9 + Object.keys(uMap).length;
    ws3.getRow(pStart).height = 26;
    ws3.mergeCells(pStart, 1, pStart, 5);
    sc(ws3.getCell(pStart,1), '📁  프로젝트별 현황', { f:fnt(true,11,C.white), fl:fll(C.navy), al:aln('left',false,1) });
    ws3.getRow(pStart+1).height = 24;
    ['프로젝트명','품목 수','입고완료','입고율','총수량'].forEach((h, ci) => {
      sc(ws3.getCell(pStart+1, ci+1), h, { f:fnt(true,10,C.white), fl:fll(C.gray700), al:aln('center'), bd:bdrAll(C.gray700) });
    });
    Object.entries(pMap).forEach(([p, v], ri) => {
      const r   = pStart + 2 + ri;
      const isE = ri % 2 === 0;
      const bg  = fll(isE ? C.gray100 : C.white);
      const all = v.i === v.a;
      const rate = v.i > 0 ? `${Math.round(v.a/v.i*100)}%` : '-';
      ws3.getRow(r).height = 22;
      sc(ws3.getCell(r,1), p,    { f:fnt(true,10,C.teal),              fl:bg,                       al:aln('left',false,1), bd:bdrBtm() });
      sc(ws3.getCell(r,2), v.i,  { f:fnt(false,10,C.gray900),          fl:bg,                       al:aln('center'),       bd:bdrBtm() });
      sc(ws3.getCell(r,3), v.a,  { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg,    al:aln('center'),       bd:bdrBtm() });
      sc(ws3.getCell(r,4), rate, { f:fnt(true,10,all?C.green:C.amber), fl:all?fll(C.greenLt):bg,    al:aln('center'),       bd:bdrBtm() });
      sc(ws3.getCell(r,5), v.q,  { f:fnt(true,10,C.blue),              fl:fll(C.bluePale),           al:aln('center'),       bd:bdrBtm() });
    });

    // ════════════════════════════════════════════════════
    // 시트 4: 가이드
    // ════════════════════════════════════════════════════
    const ws4 = wb.addWorksheet('가이드', {
      views: [{ showGridLines:false }],
      properties: { tabColor:{ argb:'FF'+C.amber } },
    });
    [4,22,56,16].forEach((w,ci) => { ws4.getColumn(ci+1).width = w; });

    ws4.getRow(1).height = 34;
    ws4.mergeCells(1, 1, 1, 4);
    sc(ws4.getCell('A1'), '📌  인텍플러스 구매관리 엑셀 — 사용 가이드', {
      f: fnt(true,13,C.white), fl: fll(C.navy), al: aln('left'),
    });

    const GUIDE = [
      ['sec', '📦 시트 구성'],
      ['hdr', '시트명', '설명'],
      ['row', '',  '데이터',   '원본 데이터. AutoFilter + 틀고정. 짝/홀수 행 교번 배색'],
      ['row', '',  '검색',     '전체 데이터 + AutoFilter. 각 컬럼 ▼ 버튼으로 필터/검색 가능'],
      ['row', '',  '통계',     'KPI 5종 카드 + 기안자별 · 프로젝트별 집계'],
      ['row', '',  '가이드',   '이 파일'],
      ['sp'],
      ['sec', '🔍 검색 시트 사용법'],
      ['row', '①', 'AutoFilter 사용', '헤더 행의 ▼ 드롭다운 클릭 → "텍스트 필터" → "포함" 선택 후 검색어 입력'],
      ['row', '②', '문서번호 검색',   '문서번호 컬럼 ▼ → 텍스트 필터 → "인텍플러스-2026" 입력'],
      ['row', '③', '품명 검색',       '품명 컬럼 ▼ → 텍스트 필터 → "ROLLER" 입력 → 포함된 항목만 표시'],
      ['row', '④', '다중 조건',       '여러 컬럼에 동시에 필터 적용 가능 (AND 조건)'],
      ['row', '⑤', '필터 초기화',     '"데이터" 탭 상단 메뉴 → 데이터 → 필터 지우기 클릭'],
      ['sp'],
      ['sec', '⚠️ 주의사항'],
      ['warn','!', '스냅샷 파일', '이 파일은 내보내기 시점의 데이터입니다. 실시간 연동 아님'],
      ['warn','!', '헤더 수정 금지', '3행 헤더를 수정하면 AutoFilter가 깨집니다'],
      ['warn','!', '행 삽입 금지', '데이터 중간에 행을 삽입하지 마세요'],
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
        sc(ws4.getCell(gr,2), g[2], { f:fnt(true,10,C.gray900), fl:fll(C.blueLt), al:aln('left',false,1), bd:bdrAll(C.gray300) });
        ws4.mergeCells(gr,3,gr,4);
        sc(ws4.getCell(gr,3), g[3]||'', { f:fnt(false,10,C.gray700), fl:fll(C.gray50), al:aln('left',false,1), bd:bdrAll(C.gray300) });
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

    // ── 응답 전송 ──────────────────────────────────────
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
