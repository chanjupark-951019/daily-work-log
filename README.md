# daily-work-log

Claude Code 스킬. 자연어로 말한 오늘 업무를 `YYYY-MM-DD.yaml`로 기록하고, 금요일에 모아 A4 주간보고서 PDF/PPTX를 자동 생성한다.

- **일일 로그**: 카테고리 사전(`_categories.yaml`)에 맞춰 정제된 YAML 저장
- **주간 보고**: HTML 템플릿 → Playwright PDF + python-pptx PPTX (A4 세로/가로, B&W)

---

## 설치 (Claude Code 에이전트용 지시사항)

> 사용자가 이 repo URL과 함께 "이 스킬 설치해줘" 등을 요청하면 아래 절차를 그대로 수행. **각 단계 끝에 한 줄씩만 보고하고, 테스트·dry-run·코드 분석은 하지 말 것.** 경로에 한글/공백 있을 수 있으므로 전 단계에서 따옴표로 감쌀 것.

**1. 사전 점검**
- `git --version`, `python --version` 실행. `python`이 Windows Store alias(빈 출력 또는 'was not found')면 `python3 --version` 재시도.
- 둘 다 실패 시 "python.org에서 Python 3.x 받고 'Add to PATH' 체크 후 PC 재시작" 안내하고 중단. git 없어도 동일.

**2. 시스템 Chrome 경로 탐색**
다음 중 존재하는 첫 번째를 변수에 저장 (없으면 빈 값):
- Windows: `C:\Program Files\Google\Chrome\Application\chrome.exe`
- Windows: `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
- Windows: `$LOCALAPPDATA\Google\Chrome\Application\chrome.exe`
- Mac: `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`
- Linux: `/usr/bin/google-chrome`, `/usr/bin/chromium`

**3. 스킬 복사**
`git clone https://github.com/chanjupark-951019/daily-work-log`을 OS 표준 temp(`$env:TEMP` / `/tmp`)에 받고, 내부 `skills/daily-work-log/` 폴더만 `~/.claude/skills/`로 복사. 기존 같은 이름 폴더는 먼저 삭제. 임시 폴더 정리.

**4. 파이썬 패키지**
`python -m pip install pyyaml jinja2 playwright python-pptx`

**5. Chromium 설정**
- 2단계에서 Chrome을 찾았으면: `AGENT_BROWSER_EXECUTABLE_PATH` 환경변수를 그 경로로 영구 설정 (Windows: `setx`, Mac/Linux: `~/.zshrc` 또는 `~/.bashrc`에 `export`). **`playwright install` 실행 금지**.
- 못 찾았을 때만 `python -m playwright install chromium` 실행 (200MB, 5~15분 소요 안내).

**6. 사용자 데이터 루트 설정**
- 사용자에게 "일일 로그/주간보고 저장할 폴더 경로를 알려주세요 (예: `C:\Users\<name>\Desktop\주간업무`)" 1줄 질문.
- 입력받은 경로에 `일일업무로그/`, `보고서/` 하위 폴더 생성.
- `~/.claude/skills/daily-work-log/starter/_categories.yaml`을 사용자 루트로 복사 (이미 있으면 스킵).
- `DAILY_WORK_LOG_ROOT` 환경변수를 그 경로로 영구 설정.

**7. 완료 안내**
사용자에게 다음 3줄만 전달:
- "Claude Code를 한 번 종료했다가 다시 열어주세요. (Desktop: 시스템 트레이 우클릭 → Quit 후 재실행 / VSCode: Ctrl+Shift+P → 'Developer: Reload Window' / CLI: 새 터미널)"
- "재시작 후 '오늘 한 일 정리해줘 - <내역>'으로 일일 로그 기록."
- "월~금 로그 쌓이면 '이번주 주간보고 만들어줘'로 PDF+PPTX 자동 생성."

---

## 설치 후 사용

1. 채팅창: "오늘 한 일 정리해줘 - 광고분석 데이터 정리했고 블로그 초안 썼어" → 일일 로그 YAML 생성
2. 채팅창: "이번주 주간보고 만들어줘" → PDF/PPTX 4개 파일 자동 생성

## 폴더 구조

사용자 데이터 루트 (`DAILY_WORK_LOG_ROOT`):
```
주간업무/
├── _categories.yaml        ← 카테고리 사전
├── 일일업무로그/
│   └── YYYY-MM-DD.yaml     ← 매일 하나씩
└── 보고서/
    └── YYYY-MM/
        ├── 주간업무현황(...).pdf + .pptx
        └── 주간업무 실적 및 계획(...).pdf + .pptx
```

스킬 본체: `~/.claude/skills/daily-work-log/`

## 의존성
- Python 3.8+
- pip packages: `pyyaml`, `jinja2`, `playwright`, `python-pptx`
- Chromium (Playwright 기본) 또는 시스템 Chrome
