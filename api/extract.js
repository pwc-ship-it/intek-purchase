// Vercel Serverless Function
// 엔드포인트: POST /api/extract
// Body: { text: string }
// Response: { doc_no, user, items: [...] }
//
// 503/429 에러 시 자동 재시도 포함 (최대 3회, 지수 백오프)
// 주 모델 계속 실패 시 fallback 모델로 자동 전환

const MODEL = 'gemini-2.5-flash';
const FALLBACK_MODEL = 'gemini-2.0-flash';
const MAX_RETRIES = 3;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function callGemini(model, prompt, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  return fetch(url, {
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
}

async function callGeminiWithRetry(prompt, apiKey) {
  let lastError = null;

  // 1단계: 주 모델 재시도
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = await callGemini(MODEL, prompt, apiKey);

      if (res.ok) return { res, model: MODEL };

      if (res.status >= 500 || res.status === 429) {
        const text = await res.text();
        lastError = { status: res.status, text, model: MODEL };
        console.warn(`[${MODEL}] ${attempt}/${MAX_RETRIES} 실패 (${res.status})`);
        if (attempt < MAX_RETRIES) {
          await sleep(1000 * Math.pow(2, attempt - 1));
          continue;
        }
      } else {
        const text = await res.text();
        return { res: null, error: { status: res.status, text, model: MODEL } };
      }
    } catch (e) {
      console.warn(`[${MODEL}] 네트워크 에러: ${e.message}`);
      lastError = { status: 0, text: e.message, model: MODEL };
      if (attempt < MAX_RETRIES) await sleep(1000 * Math.pow(2, attempt - 1));
    }
  }

  // 2단계: fallback 모델 시도
  console.warn(`주 모델 실패. ${FALLBACK_MODEL}로 fallback...`);
  try {
    const res = await callGemini(FALLBACK_MODEL, prompt, apiKey);
    if (res.ok) return { res, model: FALLBACK_MODEL };
    const text = await res.text();
    return { res: null, error: { status: res.status, text, model: FALLBACK_MODEL } };
  } catch (e) {
    return { res: null, error: lastError || { status: 0, text: e.message, model: FALLBACK_MODEL } };
  }
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { text } = req.body || {};
    if (!text || typeof text !== 'string') {
      return res.status(400).json({ error: 'text 파라미터가 필요합니다.' });
    }

    const API_KEY = process.env.GEMINI_API_KEY;
    if (!API_KEY) {
      return res.status(500).json({ error: '서버에 GEMINI_API_KEY가 설정되지 않았습니다.' });
    }

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

    const { res: geminiRes, error: geminiErr, model: usedModel } = await callGeminiWithRetry(prompt, API_KEY);

    if (geminiErr || !geminiRes) {
      const detail = geminiErr || { status: 0, text: 'unknown' };
      console.error('Gemini 최종 실패:', detail);

      let userMsg = `Gemini API 오류 (${detail.status})`;
      if (detail.status === 503) {
        userMsg = 'Gemini 서버가 일시적으로 과부하입니다. 1~2분 후 다시 시도해 주세요.';
      } else if (detail.status === 429) {
        userMsg = 'API 호출 한도를 초과했습니다. 잠시 후 다시 시도해 주세요.';
      } else if (detail.status === 403) {
        userMsg = 'Gemini API 키 권한 오류입니다. Vercel 환경변수를 확인하세요.';
      }

      return res.status(502).json({
        error: userMsg,
        detail: String(detail.text).substring(0, 500),
        model: detail.model,
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
      _model: usedModel,
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
