// Vercel Serverless Function
// 엔드포인트: POST /api/extract
// Body: { text: string }
// Response: { doc_no, user, date, items: [...] }

const GEMINI_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';

// Gemini 호출 + Rate Limit 시 자동 재시도 (최대 3회)
async function callGemini(apiKey, prompt, retries = 3) {
  const url = `${GEMINI_URL}?key=${apiKey}`;
  const body = JSON.stringify({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
    },
  });

  for (let attempt = 1; attempt <= retries; attempt++) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body,
    });

    // 성공
    if (res.ok) return res;

    const errText = await res.text();

    // Rate Limit (429) 또는 서버 과부하(503) → 대기 후 재시도
    if ((res.status === 429 || res.status === 503) && attempt < retries) {
      // 재시도 대기: 1차=5초, 2차=15초
      const waitMs = attempt === 1 ? 5000 : 15000;
      console.warn(`Gemini rate limit (${res.status}), attempt ${attempt}/${retries}, waiting ${waitMs}ms...`);
      await new Promise(r => setTimeout(r, waitMs));
      continue;
    }

    // 그 외 에러 or 재시도 소진 → 에러 객체 반환
    const err = new Error(`Gemini API 오류 (${res.status})`);
    err.status  = res.status;
    err.detail  = errText.substring(0, 500);
    err.isQuota = res.status === 429;
    throw err;
  }
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')
    return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { text } = req.body || {};
    if (!text || typeof text !== 'string')
      return res.status(400).json({ error: 'text 파라미터가 필요합니다.' });

    const API_KEY = process.env.GEMINI_API_KEY;
    if (!API_KEY)
      return res.status(500).json({ error: '서버에 GEMINI_API_KEY가 설정되지 않았습니다.' });

    // 텍스트 길이 제한 (토큰 폭주 방지)
    const trimmed = text.replace(/\s+/g, ' ').trim().substring(0, 30000);

    const prompt = `아래 품의서 텍스트에서 다음 정보를 JSON으로만 추출하세요. 설명, 마크다운, 코드블록 없이 순수 JSON 객체만 출력하세요.

스키마:
{
  "doc_no": "문서번호 (예: 인텍플러스-2026-05315)",
  "user": "기안자 이름",
  "date": "기안일 (YYYY-MM-DD 형식, 없으면 빈 문자열)",
  "items": [
    {
      "p_name": "프로젝트명",
      "p_code": "프로젝트 코드",
      "name": "품명",
      "spec": "규격",
      "qty": 숫자
    }
  ]
}

규칙:
- 모든 품목을 빠짐없이 추출
- 값이 없으면 빈 문자열 "" 또는 0
- qty는 반드시 숫자(number) 타입
- date는 YYYY-MM-DD 형식으로 변환 (예: 2026년 4월 27일 → "2026-04-27")
- JSON 외 텍스트 절대 금지

품의서 데이터:
${trimmed}`;

    let geminiRes;
    try {
      geminiRes = await callGemini(API_KEY, prompt);
    } catch (err) {
      console.error('Gemini call failed:', err);
      if (err.isQuota) {
        return res.status(429).json({
          error: 'API 호출 한도를 초과했습니다. 1~2분 후 다시 시도해 주세요.',
          detail: err.detail,
        });
      }
      return res.status(502).json({ error: err.message, detail: err.detail });
    }

    const geminiData = await geminiRes.json();

    if (!geminiData.candidates || !geminiData.candidates[0]) {
      return res.status(502).json({
        error: 'Gemini가 응답을 생성하지 않았습니다. (safety block 또는 quota 문제 가능)',
        detail: geminiData,
      });
    }

    const aiText = geminiData.candidates[0].content?.parts?.[0]?.text || '';

    // JSON 파싱
    let parsed;
    try {
      parsed = JSON.parse(aiText);
    } catch {
      const match = aiText.match(/\{[\s\S]*\}/);
      if (!match)
        return res.status(502).json({
          error: 'Gemini 응답에서 JSON을 찾지 못했습니다.',
          raw: aiText.substring(0, 500),
        });
      parsed = JSON.parse(match[0]);
    }

    // 정규화
    const result = {
      doc_no: String(parsed.doc_no || '').trim(),
      user:   String(parsed.user   || '').trim(),
      date:   String(parsed.date   || '').trim(),
      items: Array.isArray(parsed.items)
        ? parsed.items.map(it => ({
            p_name: String(it.p_name || '').trim(),
            p_code: String(it.p_code || '').trim(),
            name:   String(it.name   || '').trim(),
            spec:   String(it.spec   || '').trim(),
            qty:    Number(it.qty)   || 0,
          }))
        : [],
    };

    return res.status(200).json(result);

  } catch (err) {
    console.error('extract handler error:', err);
    return res.status(500).json({ error: '서버 내부 오류', detail: err.message });
  }
}
