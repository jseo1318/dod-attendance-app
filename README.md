# DOD 근태·연차·OT 통합관리 (Supabase + Netlify)

캡스 근태 엑셀을 업로드하면 지각/OT를 자동 계산하고, 연차·OT 잔여량을 관리하는 웹앱입니다.
데이터는 Supabase(PostgreSQL)에 저장되고, Netlify에 배포해서 팀원들과 링크로 공유합니다.

Node.js나 터미널 없이, 웹 화면 클릭만으로 끝까지 배포할 수 있도록 안내합니다.

---

## 1단계. Supabase 프로젝트 만들기

1. https://supabase.com 접속 → 회원가입/로그인
2. **New project** 클릭 → 프로젝트 이름(예: dod-attendance), 비밀번호 설정, 리전은 **Northeast Asia (Seoul)** 선택 → 생성 (1~2분 소요)
3. 왼쪽 메뉴 **SQL Editor** 클릭 → **New query**
4. 이 프로젝트 폴더 안의 `supabase_schema.sql` 파일을 열어서 내용 전체를 복사 → SQL Editor에 붙여넣기 → **Run** 클릭
   - 성공하면 employees / attendance / ledger / upload_log 테이블 4개가 생성됩니다
5. 왼쪽 메뉴 **Project Settings → API** 클릭 → 아래 두 값을 복사해둡니다 (3단계에서 사용)
   - **Project URL** (예: `https://abcdxyz.supabase.co`)
   - **anon public** 키 (긴 문자열)

⚠️ 이 앱은 로그인 기능이 없어서, 이 URL과 anon key를 아는 사람은 누구나 데이터를 읽고 쓸 수 있습니다.
지금까지 쓰시던 "링크 공유 = 같은 데이터 공유" 방식과 동일한 수준의 보안이라고 생각하시면 됩니다.

## 2단계. GitHub에 코드 올리기

1. https://github.com 에서 새 저장소(Repository) 생성 (예: `dod-attendance-app`), Private으로 설정 추천
2. 저장소 페이지의 **uploading an existing file** 링크 클릭 (또는 "Add file → Upload files")
3. 이 프로젝트 폴더 안의 모든 파일/폴더를 **node_modules, dist 폴더만 빼고** 통째로 드래그해서 업로드
   (`.env.example`, `netlify.toml`, `package.json`, `src` 폴더, `index.html`, `supabase_schema.sql` 등)
4. **Commit changes** 클릭

## 3단계. Netlify에 배포하기

1. https://app.netlify.com 접속 → 회원가입/로그인 (GitHub 계정으로 로그인하면 편해요)
2. **Add new site → Import an existing project** 클릭
3. GitHub 선택 → 방금 만든 저장소(`dod-attendance-app`) 선택
4. 빌드 설정은 자동으로 인식됩니다 (Build command: `npm run build`, Publish directory: `dist`) — 확인만 하고 넘어가기
5. **배포 전에** 반드시 환경변수를 등록해야 합니다:
   - **Add environment variables** 클릭 (또는 배포 후 Site settings → Environment variables)
   - `VITE_SUPABASE_URL` = 1단계에서 복사한 Project URL
   - `VITE_SUPABASE_ANON_KEY` = 1단계에서 복사한 anon public 키
6. **Deploy site** 클릭 → 1~2분 후 배포 완료
7. 발급된 주소(예: `https://dod-attendance.netlify.app`)로 접속해서 확인

## 4단계. 이후 수정할 때

- 코드를 수정하고 싶으면(예: 근무시간 규정 변경) `src/App.jsx` 파일을 고친 뒤 GitHub에 다시 업로드하면,
  Netlify가 자동으로 재배포합니다.
- 환경변수(Supabase URL/키)를 바꾸면 Netlify에서 **Trigger deploy → Deploy site**를 한 번 눌러줘야 반영됩니다.
- Supabase 대시보드의 **Table Editor**에서 직원/근태/연차 데이터를 직접 SQL 없이도 표 형태로 확인·수정할 수 있습니다.

## 로컬에서 미리 확인하고 싶다면 (선택, Node.js 필요)

```bash
npm install
cp .env.example .env   # .env 파일을 열어 실제 Supabase URL/키로 수정
npm run dev
```

---

### 파일 구성

- `src/App.jsx` — 전체 화면/로직 (근무시간 규정, 지각·OT 계산식 포함)
- `src/supabaseClient.js` — Supabase 연결 설정
- `supabase_schema.sql` — Supabase에 실행할 테이블 생성 스크립트
- `netlify.toml` — Netlify 빌드 설정
- `.env.example` — 필요한 환경변수 이름 예시
