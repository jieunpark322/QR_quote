# 🌐 무료 클라우드 배포 가이드 (Streamlit Community Cloud)

비개발자도 따라할 수 있게 단계별로 정리했습니다. 약 **30~40분** 소요됩니다.

## 결과물

배포를 완료하면 다음이 됩니다:

- 🌍 **`https://<원하는이름>.streamlit.app`** 형태의 영구 URL 생성
- ✅ 사용자가 박지은님 PC를 켜지 않아도 24시간 접속 가능
- ✅ HTTPS 자동 (보안 통신)
- ✅ **이메일 화이트리스트** 로 접근자 제한 가능
- ✅ 비용 0원

## ⚠️ 시작 전 확인사항

### 1. OneDrive 폴더에서 빼내기 (중요)

현재 프로젝트가 OneDrive 동기화 폴더에 있으면 git 동작과 충돌할 수 있습니다. **시작 전에 프로젝트 폴더를 OneDrive 밖으로 이동**해주세요.

예시:
- 현재: `C:\Users\박지은\OneDrive - (주)소프트먼트\...\Claude\contract-system\`
- 권장: `C:\Users\박지은\projects\contract-system\` 또는 `C:\Users\박지은\Documents\contract-system\`

폴더 통째로 잘라내기/붙여넣기 하면 됩니다.

### 2. 민감 정보 검토

`brands/softment/brand.json` 에 사업자번호와 입금계좌가 들어있습니다. **반드시 Private 리포지토리**로 배포해야 외부 노출 안 됩니다. 가이드대로 따라하면 Private 으로 설정됩니다.

---

## Step 1 — GitHub 계정 만들기 (5분)

이미 계정이 있으면 Step 2로.

1. https://github.com/signup 접속
2. 이메일·비밀번호·사용자명 입력 (예: `jieun-park`)
3. 이메일로 온 인증 코드 입력
4. "Continue for free" 선택

## Step 2 — GitHub Desktop 설치 (5분)

코딩 안 해도 GUI 로 GitHub 사용 가능한 도구.

1. https://desktop.github.com 에서 다운로드
2. 설치 후 실행
3. **"Sign in to GitHub.com"** → 방금 만든 계정으로 로그인
4. Configure Git: 이름·이메일 입력 (개인용도 OK)

## Step 3 — 프로젝트를 GitHub에 올리기 (10분)

1. GitHub Desktop에서 좌측 상단 **File → Add Local Repository**
2. **Choose...** 클릭 → 프로젝트 폴더 선택 (위에서 OneDrive 밖으로 옮긴 `contract-system` 폴더)
3. 알림이 나타날 수 있음:
   - "This directory does not appear to be a Git repository. Would you like to create a repository here instead?" → **"create a repository"** 클릭
4. 다음 정보 입력:
   - **Name**: `contract-system` (또는 원하는 이름)
   - **Local path**: 그대로 두기
   - **Description**: (선택) `B2B 견적서 자동 생성 시스템`
   - **Git ignore**: `None` (이미 `.gitignore` 가 있음)
   - **License**: `None`
5. **Create Repository** 클릭
6. 좌측에 변경사항 목록이 보임. 좌측 하단:
   - **Summary**: `Initial commit`
   - **Commit to main** 클릭
7. 상단 **Publish repository** 클릭
   - **Name**: 그대로
   - ✅ **"Keep this code private"** 반드시 체크 (사업자번호 보호)
   - **Publish Repository** 클릭

→ 이제 GitHub.com 의 본인 계정에 가서 보면 `contract-system` 리포지토리가 생성돼 있습니다.

## Step 4 — Streamlit Community Cloud 가입 (3분)

1. https://share.streamlit.io 접속
2. **"Continue with GitHub"** 클릭
3. GitHub 로그인 (자동 연동)
4. 권한 허용 (Streamlit이 본인 리포지토리에 접근 가능하도록)
5. 가입 완료

## Step 5 — 앱 배포 (5~10분)

1. Streamlit Cloud 대시보드에서 **"Create app"** 또는 **"New app"** 클릭
2. **"Deploy a public app from GitHub"** 선택 (Private 리포 OK)
3. 입력:
   - **Repository**: `<본인계정>/contract-system`
   - **Branch**: `main`
   - **Main file path**: `src/webapp.py`
   - **App URL**: 원하는 서브도메인 (예: `softment-quote`)
     → 최종 URL: `https://softment-quote.streamlit.app`
4. **Deploy!** 클릭
5. 로그가 표시됨. 다음 단계가 자동 진행됨:
   - Python 패키지 설치 (requirements.txt) — ~2분
   - **LibreOffice 설치** (packages.txt) — ~3~5분
   - 앱 시작
6. 완료되면 자동으로 견적서 작성 화면이 표시됨!

## Step 6 — 접근 제한 (중요, 3분)

기본값이 "Public" 이라 URL을 아는 누구나 들어올 수 있습니다. **꼭 제한 설정**하세요.

1. 배포된 앱 우측 상단의 **⋮ 메뉴 → Settings**
2. **"Sharing"** 탭
3. **"Only specific people"** 라디오 선택
4. **"Add viewer"** 클릭 → 접근 허용할 이메일 추가
   - 예: 본인 이메일, 동료 이메일
5. **Save**

이제 해당 이메일의 Google/GitHub 로그인 후에만 앱에 접속 가능합니다.

---

## ✅ 완료!

이제 누구든 `https://softment-quote.streamlit.app` (또는 본인이 정한 이름) 으로 접속하고, 박지은님이 추가한 이메일로 로그인하면 견적서를 만들 수 있습니다. PC 안 켜져 있어도 OK.

## 📝 추후 변경 시

### 카탈로그/브랜드 정보 영구 수정

⚠️ **중요**: 클라우드에서 웹 UI로 편집해도 **앱이 재시작되면 사라집니다** (Streamlit Cloud는 임시 파일시스템). 영구 변경은 다음 방법 중 하나로:

**방법 A. GitHub Desktop 사용 (권장)**
1. 로컬 프로젝트 폴더에서 `catalog/products.json` 또는 `brands/softment/brand.json` 편집
2. GitHub Desktop 열기 → 변경사항 확인
3. **Summary** 작성 → **Commit to main**
4. **Push origin** 클릭
5. 1~2분 후 클라우드 앱이 자동 재배포됨

**방법 B. GitHub 웹에서 직접 편집**
1. github.com 의 리포지토리에서 해당 JSON 파일 열기
2. 우측 상단 ✏️ 아이콘 (Edit this file)
3. 수정 후 **Commit changes** 클릭

### 코드 업데이트

같은 방식 — 로컬 편집 → Commit → Push → 자동 재배포.

## 🔒 보안 체크리스트

- [ ] 리포지토리가 **Private** 인가? (Step 3-7 확인)
- [ ] Streamlit 앱이 **"Only specific people"** 설정인가? (Step 6 확인)
- [ ] 접근 허용 이메일 목록이 실제 사용자만 포함하는가?
- [ ] `brand.json` 의 입금계좌가 노출돼도 괜찮은 사용자만 추가했는가?

## 🆘 트러블슈팅

| 증상 | 원인 / 해결 |
|---|---|
| "Module not found" 에러 | `requirements.txt` 에 패키지 누락 — 추가 후 push |
| PDF 변환 안 됨 | `packages.txt` 에 `libreoffice` 확인. 빌드 로그에서 설치 성공 여부 체크 |
| 한글 폰트 깨짐 | `packages.txt` 에 `fonts-noto-cjk` 있는지 확인 |
| 15분 후 sleep | 정상 동작. 다시 접속하면 30초 내 깨어남 |
| 메모리 부족 (1GB 제한) | 사용량 줄이거나 유료 플랜 |
| 변경 후 반영 안 됨 | 우측 상단 ⋮ → Reboot app 클릭 |

## 💸 무료 한도

- **앱 개수**: 무제한
- **공용 리소스**: CPU 1, RAM 1GB
- **트래픽**: 무제한 (적당히 쓰는 한)
- **Sleep**: 15분 미사용 시 (다음 접속 시 자동 wake)

소프트먼트 규모(소수 인원, 가끔 사용) 라면 충분합니다.

## 🚀 더 빠른 응답이 필요하면

Sleep 없이 즉시 응답 원하면 유료 옵션:
- Streamlit Cloud 유료 ($20/월~)
- Render.com Standard ($7/월)
- Railway ($5/월)
