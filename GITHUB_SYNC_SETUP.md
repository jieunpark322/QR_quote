# GitHub 자동 commit 설정 (영구 저장)

품목 관리에서 저장한 데이터를 GitHub 저장소에 자동 commit하여 영구 보존하려면 아래 설정이 필요합니다.

## 1) GitHub Personal Access Token (PAT) 발급

1. GitHub 우상단 프로필 → **Settings**
2. 좌측 메뉴 맨 아래 **Developer settings**
3. **Personal access tokens → Fine-grained tokens** → **Generate new token**
4. 설정:
   - **Token name**: `QR_quote streamlit auto-commit`
   - **Expiration**: 1년 (또는 No expiration)
   - **Repository access**: Only select repositories → `jieunpark322/QR_quote` 선택
   - **Repository permissions** → **Contents** → **Read and write** ⚠ 필수
5. **Generate token** → 생성된 토큰 (`github_pat_...`) 복사 (한 번만 보임)

## 2) Streamlit Cloud Secrets 등록

1. https://share.streamlit.io 접속
2. 앱 (`qrquote`) → 우측 ⋮ → **Settings**
3. **Secrets** 탭 → 아래 내용 붙여넣기:

```toml
GITHUB_TOKEN = "github_pat_여기에_복사한_토큰_붙여넣기"
GITHUB_OWNER = "jieunpark322"
GITHUB_REPO = "QR_quote"
GITHUB_BRANCH = "main"
```

4. **Save** → 앱 자동 재시작

## 3) 동작 확인

품목 관리에서 아무 항목 수정 후 **💾 저장** 누름:
- 표 아래 메시지에 `🌐 GitHub 영구 저장 완료` 가 떠야 함
- GitHub 저장소의 `catalog/products.json` (또는 멤버십/야외형) 파일이 자동 업데이트되어 있음
- Streamlit Cloud 재배포되어도 데이터 유지됨

## 토큰이 노출되면

새 토큰 발급 후 Streamlit Cloud Secrets 갱신. 이전 토큰은 GitHub 설정에서 **Revoke**.
