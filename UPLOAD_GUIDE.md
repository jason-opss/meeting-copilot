# GitHub 업로드 가이드

`GITHUB UPLOAD` 폴더 안의 파일과 폴더를 GitHub 저장소 최상단에 올리면 됩니다.

중요: `GITHUB UPLOAD` 폴더 자체를 통째로 올리는 것이 아니라, 그 안의 내용물이 저장소 첫 화면에 보이게 올리세요.

## 올릴 파일

```text
index.html
app.js
style.css
README.md
UPLOAD_GUIDE.md
.nojekyll
STYLE_WANTED/DOCX_FINAL_WANTED.docx
```

## 올리지 않아도 되는 파일

```text
.env
node_modules/
SAMPLE/
SAMPLE_IM/
server-runtime*.log
tools/
```

## 새 GitHub 저장소에 올리는 방법

1. GitHub에 로그인합니다.
2. 오른쪽 위 `+` 버튼을 누르고 `New repository`를 선택합니다.
3. 저장소 이름을 입력합니다. 예: `meeting-copilot`
4. 공개 범위를 선택합니다.
   - 대회 제출용이면 보통 `Public`이 편합니다.
   - 링크를 아는 사람에게만 보여주고 싶으면 `Private`도 가능하지만, GitHub Pages 사용 조건은 계정/요금제에 따라 다를 수 있습니다.
5. 저장소를 만든 뒤 `Add file -> Upload files`를 누릅니다.
6. `GITHUB UPLOAD` 폴더 안의 파일과 폴더를 끌어다 놓습니다.
7. 커밋 메시지는 `Initial Meeting Copilot upload` 정도로 입력하고 `Commit changes`를 누릅니다.

## GitHub Pages 켜는 방법

1. 저장소 상단의 `Settings`를 누릅니다.
2. 왼쪽 메뉴에서 `Pages`를 누릅니다.
3. `Build and deployment`의 Source를 `Deploy from a branch`로 둡니다.
4. Branch는 `main`, Folder는 `/root`를 선택합니다.
5. `Save`를 누릅니다.
6. 잠시 기다리면 GitHub Pages 주소가 표시됩니다.

주소 예시:

```text
https://사용자명.github.io/저장소명/
```

## 배포 후 확인

1. GitHub Pages 주소로 접속합니다.
2. 왼쪽 아래 `설정`에서 Gemini API Key를 입력합니다.
3. `저장` 또는 `짧은 연결 테스트`를 눌러 모델 설정을 확인합니다.
4. 사전 준비 화면에서 IM 업로드 또는 세팅값 입력 후 `브리프 생성`을 실행합니다.
5. 최종 보고서 화면에서 DOCX 다운로드가 되는지 확인합니다.

## 보안 메모

- API Key는 저장소, localStorage, sessionStorage에 저장되지 않습니다.
- localStorage에는 미팅 데이터와 비민감 프로필만 저장됩니다.
- 이름, 부서, 선택 모델, 미팅 데이터 저장 설정은 편의를 위해 저장됩니다.
- 설정에서 미팅 데이터 저장을 끄거나, 전체 삭제하거나, 백업 후 저장소를 비울 수 있습니다.
- 실제 사내 기밀 IM을 사용할 때는 회사의 외부 AI/API 사용 정책을 확인하세요.
