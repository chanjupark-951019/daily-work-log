# daily-work-log

Claude Code 스킬. 자연어로 말한 오늘 업무를 `YYYY-MM-DD.yaml`로 기록하고, 금요일에 모아 A4 주간보고서 PDF/PPTX를 자동 생성합니다.

- **일일 로그**: 카테고리 사전(`_categories.yaml`)에 맞춰 정제된 YAML 저장
- **주간 보고**: HTML 템플릿 → Playwright PDF + python-pptx PPTX (A4 세로/가로, B&W)

## 설치

Claude Code 채팅창에 아래 프롬프트를 통째로 복사해서 붙여넣으세요. (`INSTALL_PROMPT.md` 참고)

```
이 GitHub 스킬을 ~/.claude/skills/ 에 설치해줘: <GITHUB_URL>
```

자세한 프롬프트는 `INSTALL_PROMPT.md` 파일 내용을 복사해서 사용하세요.

## 설치 후 사용

1. 채팅창: "오늘 한 일 정리해줘 - 광고분석 데이터 정리했고 블로그 초안 썼어" → 일일 로그 YAML 생성
2. 채팅창: "이번주 주간보고 만들어줘" → PDF/PPTX 4개 파일 자동 생성

## 폴더 구조

설치 후 사용자 데이터 루트 (`DAILY_WORK_LOG_ROOT` 환경변수가 가리키는 경로):

```
주간업무/
├── _categories.yaml        ← 카테고리 사전 (starter에서 복사)
├── 일일업무로그/
│   └── YYYY-MM-DD.yaml     ← 매일 하나씩
└── 보고서/
    └── YYYY-MM/
        ├── 주간업무현황(...).pdf + .pptx
        └── 주간업무 실적 및 계획(...).pdf + .pptx
```

스킬 자체는 `~/.claude/skills/daily-work-log/`에 설치됩니다.

## 의존성

- Python 3.8+
- pip packages: `pyyaml`, `jinja2`, `playwright`, `python-pptx`
- Chromium (Playwright 기본) 또는 시스템 Chrome
