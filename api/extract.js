// Vercel Serverless Function
// 엔드포인트: POST /api/extract
// Body: { text: string }
// Response: { doc_no, user, items: [...] }

export default async function handler(req, res) {
  // CORS (같은 도메인에서만 쓸 거지만 혹시 몰라서)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { text } = req.body || {};
    if (!text || typeof text !== 'string') {
      return res.status(400).json({ error: 'text 파라미터가 필요합니다.' });
    }

    const API_KEY = process.env.GEMINI_API_KEY;
    if (!API_KEY) {
      return res.status(500).json({ error: '서버에 GEMINI_API_KEY가 설정되지 않았습니다.' });
    }

    // 텍스트 길이 제한 (토큰 폭주 방지)
    const trimmed = text.replace(/\s+/g, ' ').trim().substring(0, 30000);

    const prompt = `아래 품의서 텍스트에서 다음 정보를 JSON으로만 추출하세요. 설명, 마크다운, 코드블록 없이 순수 JSON 객체만 출력하세요.

스키마:
{
  "doc_no": "문서번호 (예: IP-2024-0001)",
  "user": "신청자 이름",
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
- JSON 외 텍스트 절대 금지

품의서 데이터:
${trimmed}`;

    // Gemini 2.5 Flash 호출 (structured output 지원)
    const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;

    const geminiRes = await fetch(geminiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.1,
          responseMimeType: 'application/json',
        },
      }),
    });

    if (!geminiRes.ok) {
      const errText = await geminiRes.text();
      console.error('Gemini API error:', geminiRes.status, errText);
      return res.status(502).json({
        error: `Gemini API 오류 (${geminiRes.status})`,
        detail: errText.substring(0, 500),
      });
    }

    const geminiData = await geminiRes.json();

    if (!geminiData.candidates || !geminiData.candidates[0]) {
      return res.status(502).json({
        error: 'Gemini가 응답을 생성하지 않았습니다. (safety block 또는 quota 문제 가능)',
        detail: geminiData,
      });
    }

    const aiText = geminiData.candidates[0].content?.parts?.[0]?.text || '';

    // JSON 파싱 (응답이 이미 JSON이어야 하지만 안전장치)
    let parsed;
    try {
      parsed = JSON.parse(aiText);
    } catch {
      const match = aiText.match(/\{[\s\S]*\}/);
      if (!match) {
        return res.status(502).json({
          error: 'Gemini 응답에서 JSON을 찾지 못했습니다.',
          raw: aiText.substring(0, 500),
        });
      }
      parsed = JSON.parse(match[0]);
    }

    // 정규화
    const result = {
      doc_no: String(parsed.doc_no || '').trim(),
      user: String(parsed.user || '').trim(),
      items: Array.isArray(parsed.items)
        ? parsed.items.map((it) => ({
            p_name: String(it.p_name || '').trim(),
            p_code: String(it.p_code || '').trim(),
            name: String(it.name || '').trim(),
            spec: String(it.spec || '').trim(),
            qty: Number(it.qty) || 0,
          }))
        : [],
    };

    return res.status(200).json(result);
  } catch (err) {
    console.error('extract handler error:', err);
    return res.status(500).json({
      error: '서버 내부 오류',
      detail: err.message,
    });
  }
}
