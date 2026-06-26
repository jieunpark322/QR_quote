# 품목 관리 영구 저장 설정 (3분 소요 · 한 번만)

품목 관리에서 저장한 데이터가 시간이 지나도 사라지지 않도록 GitHub에 자동 백업되는 기능입니다.
**한 번 설정해 두면** 이후엔 누가 어디서 사용해도 영구 보존돼요.

## 1단계 — GitHub 토큰 발급 (2분)

1. 아래 링크를 그대로 클릭하세요 (이미 권한이 설정된 페이지로 바로 이동):
   👉 https://github.com/settings/tokens/new?scopes=repo&description=QR_quote+%EC%9E%90%EB%8F%99%EC%A0%80%EC%9E%A5

2. **Expiration** (만료기간): `No expiration` 선택 권장 (또는 1년)

3. 페이지 맨 아래 초록 버튼 **Generate token** 클릭

4. 화면에 `ghp_xxxxxxxxxxxxxxxxxx` 로 시작하는 긴 문자열이 표시됩니다.
   👉 **복사 아이콘 클릭** (한 번만 보여줘요!)

## 2단계 — Streamlit Cloud에 등록 (1분)

1. 아래 링크를 그대로 클릭:
   👉 https://share.streamlit.io

2. 본인의 앱 **qrquote** 클릭 → 우측 위 점 3개(⋮) → **Settings**

3. 왼쪽 메뉴 **Secrets** 클릭

4. 큰 텍스트 입력 박스에 **아래 4줄을 그대로 복사·붙여넣기**:

   ```toml
   GITHUB_TOKEN = "여기에_복사한_토큰_붙여넣기"
   GITHUB_OWNER = "jieunpark322"
   GITHUB_REPO = "QR_quote"
   GITHUB_BRANCH = "main"
   ```

5. 첫 줄의 `여기에_복사한_토큰_붙여넣기` 부분에 1단계에서 복사한 토큰(`ghp_...`)을 따옴표 안에 붙여넣기

6. 오른쪽 아래 **Save** 클릭 → 앱이 자동으로 다시 시작 (1~2분)

## 3단계 — 동작 확인 (30초)

1. 앱이 다시 켜지면 → 메뉴에서 **품목 관리** 진입
2. 아무 품목 한 줄 수정 → **💾 저장**
3. 표 아래에 **🌐 GitHub 영구 저장 완료** 메시지가 뜨면 성공 ✅

이제부터는 누가 언제 사용해도 변경 내용이 영구 보존됩니다.

---

## 문제 발생 시

- **"GITHUB_TOKEN 미설정"** 표시 → 2단계 secrets 4줄 중 첫 줄 토큰 따옴표 확인
- **"PUT 401/403"** → 토큰이 만료됐거나 권한 부족. 1단계부터 새로 발급
- **토큰을 잃어버렸을 때** → 1단계 다시 진행 (이전 토큰은 자동 무효화하려면 GitHub에서 Revoke)
