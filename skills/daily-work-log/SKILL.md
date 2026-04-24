---
name: daily-work-log
description: 사용자가 자연어로 말한 오늘(또는 특정 날짜)의 업무 내역을 YAML 일일 로그로 기록하고, 금요일에 이 로그들을 모아 주간업무보고 PDF/PPTX를 자동 생성한다. '오늘 한 일', '오늘 업무', '일일로그', '업무로그', '오늘 일한 거 정리해줘', '어제 작업 기록해줘', '이번주 주간보고 만들어줘', '주간업무현황 PDF' 등에 트리거. 카테고리는 사용자 `_categories.yaml`에 정의된 값만 사용하며, 사전에 없으면 먼저 사전을 업데이트한 뒤 로그를 쓴다.
---

# 일일 업무 로그 + 주간 보고 스킬

자연어 업무 내역 → 카테고리 매핑된 YAML 로그 저장 → 주간 PDF/PPTX 자동 생성.

## 경로 설정 (환경변수 기반)

이 스킬은 환경변수 `DAILY_WORK_LOG_ROOT`에 저장된 경로를 사용자 데이터 루트로 사용한다.

- **로그**: `$DAILY_WORK_LOG_ROOT/일일업무로그/YYYY-MM-DD.yaml`
- **카테고리 사전**: `$DAILY_WORK_LOG_ROOT/_categories.yaml`
- **주간보고 출력**: `$DAILY_WORK_LOG_ROOT/보고서/YYYY-MM/`
- **템플릿·스크립트**: 이 스킬 폴더 안의 `templates/`, `scripts/`

환경변수가 설정돼 있지 않으면 사용자에게 루트 경로를 물어보고 설정 방법을 안내한다:
- Windows: `setx DAILY_WORK_LOG_ROOT "C:\Users\...\주간업무"` 실행 후 터미널 재시작
- Mac/Linux: `~/.bashrc` 또는 `~/.zshrc`에 `export DAILY_WORK_LOG_ROOT="/path/to/주간업무"` 추가

## A. 일일 로그 기록

### 1. 날짜 결정
- 명시 없으면 **오늘 날짜** (환경의 currentDate)
- "어제", "지난 금요일" → 절대 날짜로 변환
- 파일명: `YYYY-MM-DD.yaml`

### 2. 카테고리 사전 로드
- **매번** `$DAILY_WORK_LOG_ROOT/_categories.yaml` 먼저 읽기
- `name` + `aliases` 전체를 매칭 후보로

### 3. 자연어 → 카테고리 매핑
- 한 발화가 여러 카테고리면 분리해 각각 entry로 기록
- 매칭 불가/애매 시: 사용자에게 새 카테고리 추가 여부 확인 → 승인 시 `_categories.yaml`에 **append만** (`name` + `overview` 한 줄 + `aliases: []`)

### 4. 기존 파일 병합 규칙
- 같은 날짜 파일 있으면: **read → merge → write**
  - 같은 category면 `done` 리스트 append
  - 새 category면 entries 끝에 append
  - 기존 내용 덮어쓰기·재정렬 금지

### 5. YAML 포맷

```yaml
date: 2026-04-27
weekday: 월
week_of_month: "2026년 4월 5주차"

entries:
  - category: "자동화 시스템 구축 3. 콘텐츠 제작"
    done:
      - "EdgeO 블로그 수정안 피드백 반영"
  - category: "광고성과분석"
    done:
      - "SA 비용 데이터 정리"
```

규칙:
- 공백 2칸 들여쓰기, 탭 금지, UTF-8
- `category` = `_categories.yaml`의 정식 `name` (alias 아님)
- `evidence`는 사용자 발화에 실제 경로/URL 있을 때만 — **지어내지 말 것**
- 콜론·따옴표는 큰따옴표로 감싸기

### 6. weekday / week_of_month
- `weekday`: 월/화/수/목/금/토/일
- `week_of_month`: 달력 기준 (일~토 한 주, 월 1일이 속한 주가 1주차)

### 7. 출력
```
기록 완료 → 일일업무로그/YYYY-MM-DD.yaml (요일, N주차)
- 카테고리명: N건
```

### 금지 사항
- 사전에 없는 카테고리 임의 사용 (항상 사전 먼저 append)
- 사용자가 말하지 않은 내용 추론·추가
- `evidence` 경로 조작
- 기존 파일 덮어쓰기·순서 변경

## B. 주간 보고 생성

월~금 5일치 로그가 쌓였을 때 실행. 사용자가 "이번주 주간보고", "주간업무현황 PDF", "금요일 보고" 등을 요청하면 수행.

### 실행 명령

스크립트는 이 스킬 폴더의 `scripts/generate_weekly.py`에 있다. 실행 전에 필요한 Python 패키지가 설치돼 있어야 한다 (jinja2, playwright, python-pptx, pyyaml).

```bash
python "$CLAUDE_SKILL_DIR/scripts/generate_weekly.py" \
  --monday 2026-04-27 \
  --week-label "5월 1주차" \
  [--plans plans_5월1주차.json]
```

`$CLAUDE_SKILL_DIR`은 이 스킬이 설치된 폴더 경로 — Claude Code는 통상 `~/.claude/skills/daily-work-log/`로 설치한다. 스킬 실행 시 이 경로를 파이썬에 전달한다.

### 선택: 차주 계획 override

`--plans`는 다음주 계획 내용을 별도 JSON으로 주입할 때 사용. 파일 생략 시 계획란은 빈 상태로 생성.

JSON 포맷 예시:
```json
{
  "status": {
    "광고성과분석": ["회귀분석 실행 (5/4~)", "예측모델 1차 산출"]
  },
  "gantt": {
    "광고성과분석": {
      "due": "5/8",
      "mon": ["회귀분석 실행"],
      "wed": ["예측모델 산출"]
    }
  }
}
```

### 출력물

`$DAILY_WORK_LOG_ROOT/보고서/YYYY-MM/` 하위에 4개 파일:
- `주간업무현황(박찬주_월 N주차).pdf` (A4 세로, B&W)
- `주간업무현황(박찬주_월 N주차).pptx` (편집 가능)
- `주간업무 실적 및 계획(박찬주_월 N주차_월 N+1주차).pdf` (A4 가로, B&W)
- `주간업무 실적 및 계획(박찬주_월 N주차_월 N+1주차).pptx` (편집 가능)

### 실행 순서 (사용자 요청 시)

1. `$DAILY_WORK_LOG_ROOT` 환경변수 확인 — 없으면 설정 안내 후 중단
2. 이번주 월요일 날짜와 주차 라벨 결정 (오늘 날짜 기준 지난 월요일, 또는 사용자 명시)
3. 해당 주 월~금 로그가 얼마나 쌓였는지 확인 → 비어있는 날 있으면 사용자에게 알림 (진행 여부 확인)
4. plans JSON이 필요한지 사용자에게 물음 (생략해도 OK)
5. `generate_weekly.py` 실행
6. 출력 파일 경로 보고

## 의존성

Python 3.8+ 및 패키지:
- pyyaml
- jinja2
- playwright (+ chromium 또는 시스템 Chrome)
- python-pptx
- pypdfium2 (선택, 프리뷰 렌더링용)

설치 시 환경변수 `AGENT_BROWSER_EXECUTABLE_PATH` 또는 playwright 자체 Chromium 중 하나가 있어야 PDF 렌더링 가능.
