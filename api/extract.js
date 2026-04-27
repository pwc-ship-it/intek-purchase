// Vercel Serverless Function
// 엔드포인트: POST /api/extract
// Body: { text: string }
// Response: { doc_no, user, date, items: [...] }
//
// 필요 환경변수: GROQ_API_KEY
// Vercel > Settings > Environment Variables 에서 설정

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

    const API_KEY = process.env.GROQ_API_KEY;
    if (!API_KEY)
      return res.status(500).json({ error: '서버에 GROQ_API_KEY가 설정되지 않았습니다.' });

    // 텍스트 길이 제한
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

    // Groq API 호출
    const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEY}`,
      },
      body: JSON.stringify({
        model: 'llama-3.3-70b-versatile',  // 무료, 한국어 품질 우수
        temperature: 0.1,
        max_tokens: 4096,
        response_format: { type: 'json_object' }, // JSON만 반환
        messages: [{ role: 'user', content: prompt }],
      }),
    });

    if (!groqRes.ok) {
      const errText = await groqRes.text();
      console.error('Groq API error:', groqRes.status, errText);

      if (groqRes.status === 429) {
        return res.status(429).json({
          error: 'API 호출 한도를 초과했습니다. 잠시 후 다시 시도해 주세요.',
          detail: errText.substring(0, 300),
        });
      }
      return res.status(502).json({
        error: `Groq API 오류 (${groqRes.status})`,
        detail: errText.substring(0, 300),
      });
    }

    const groqData = await groqRes.json();
    const aiText = groqData.choices?.[0]?.message?.content || '';

    if (!aiText) {
      return res.status(502).json({ error: 'Groq가 응답을 생성하지 않았습니다.' });
    }

    // JSON 파싱
    let parsed;
    try {
      const clean = aiText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
      parsed = JSON.parse(clean);
    } catch {
      const match = aiText.match(/\{[\s\S]*\}/);
      if (!match) {
        return res.status(502).json({
          error: 'Groq 응답에서 JSON을 찾지 못했습니다.',
          raw: aiText.substring(0, 500),
        });
      }
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
