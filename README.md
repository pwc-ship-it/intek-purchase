# 🛒 인텍플러스 구매관리 시스템 - 배포 가이드

## 📁 프로젝트 구조

```
intek-purchase/
├── api/
│   └── extract.js          ← Vercel 서버리스 함수 (Gemini 호출)
├── index.html              ← 프론트엔드 (사용자가 보는 화면)
├── package.json            ← Node.js 설정
├── vercel.json             ← Vercel 배포 설정
├── supabase_schema.sql     ← DB 테이블 생성 스크립트
└── README.md               ← 이 파일
```

> ⚠️ **중요**: 기존의 `.github/workflows/deploy.yml` 파일은 **삭제**하세요.
> GitHub Pages 배포와 Vercel 배포가 섞이면 꼬입니다.

---

## 🚀 배포 5단계

### 1단계: Supabase 프로젝트 준비

1. https://supabase.com 에서 프로젝트 생성 (기존에 있으면 그대로 사용)
2. 좌측 메뉴 **SQL Editor** 클릭
3. `supabase_schema.sql` 파일 내용을 복사해서 붙여넣고 **Run** 클릭
4. 좌측 **Settings → API** 에서 다음 2개 값 복사:
   - `Project URL` (예: `https://xxxxx.supabase.co`)
   - `anon public` key (매우 긴 문자열)

### 2단계: Gemini API 키 발급

1. https://aistudio.google.com/apikey 접속
2. **Create API key** 클릭 → 키 복사
3. ⚠️ 이 키는 **절대 Git에 올리지 마세요.** Vercel 환경변수로만 넣습니다.

### 3단계: `index.html` 수정 (Supabase 값 넣기)

`index.html` 파일 열어서 아래 2줄을 본인 값으로 교체:

```javascript
const SUPABASE_URL = 'https://YOUR-PROJECT.supabase.co';  // ← 1단계에서 복사한 값
const SUPABASE_ANON_KEY = 'YOUR-ANON-KEY';                // ← 1단계에서 복사한 값
```

> anon key는 공개돼도 안전합니다. RLS 정책이 권한을 제어합니다.

### 4단계: GitHub에 푸시

```bash
# 기존 deploy.yml 삭제
rm -rf .github

git add .
git commit -m "Migrate to Vercel with serverless function"
git push origin main
```

### 5단계: Vercel 배포 & 환경변수 설정

1. https://vercel.com → **Add New → Project**
2. GitHub 저장소 import
3. **Environment Variables** 섹션에서 추가:

   | Name | Value |
   |------|-------|
   | `GEMINI_API_KEY` | 2단계에서 발급받은 Gemini 키 |

4. **Deploy** 클릭
5. 배포 끝나면 `https://xxx.vercel.app` 주소로 접속

> 💡 환경변수를 나중에 추가했다면 **Deployments → ⋯ → Redeploy** 해야 반영됩니다.

---

## 🧪 동작 테스트

1. 배포된 사이트 접속
2. 품의서 HTML 파일 드래그앤드롭
3. Gemini가 분석 → 테이블에 품목 표시
4. 수정 필요하면 셀 편집 / 행 추가 / 삭제
5. **"데이터베이스에 저장"** 클릭
   - 같은 문서번호가 이미 저장돼 있으면 `⚠️ 문서번호가 중복되어 업로드 되지 않습니다.` 메시지가 뜨고 저장 안 됨
   - 수정하려면 Supabase Table Editor에서 해당 문서번호의 행을 직접 삭제 후 재업로드
6. **"저장된 구매 리스트"** 탭에서 확인

---

## 🛠 주요 변경 사항 (기존 코드 대비)

| 항목 | 기존 | 변경 |
|------|------|------|
| 배포 | GitHub Pages + Vercel 혼재 | Vercel 단일 |
| API 키 | 프론트엔드에 노출 🚨 | 서버리스 환경변수 ✅ |
| Gemini 모델 | `gemini-1.5-flash` (단종) | `gemini-2.5-flash` |
| JSON 추출 | 정규식으로 파싱 | `responseMimeType: application/json` 사용 |
| Supabase 저장 | 함수 미구현 | 중복 문서번호 거부 방식 완성 |
| 조회 기능 | 없음 | 검색 가능한 조회 탭 추가 |
| 행 편집 | 수정만 가능 | 추가/삭제 가능 |
| 에러 처리 | `alert()` | 토스트 알림 |

---

## 🆘 문제 해결

### "서버에 GEMINI_API_KEY가 설정되지 않았습니다"
→ Vercel 환경변수 설정 후 **Redeploy** 필수

### "Gemini API 오류 (400)"
→ API 키가 유효하지 않거나 quota 초과. aistudio에서 확인

### Supabase 저장 실패: "new row violates row-level security policy"
→ `supabase_schema.sql` 의 정책 부분이 실행 안 됐음. SQL 재실행

### 저장된 데이터가 안 보임
→ 브라우저 개발자 도구(F12) → Console 탭에서 에러 확인
→ Supabase Table Editor 에서 `purchase_items` 테이블 직접 확인
