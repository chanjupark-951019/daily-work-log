# 설치 프롬프트 (사내 공유용)

Claude Code 채팅창에 이 블록을 **통째로** 복사해서 붙여넣으세요.

---

이 GitHub 스킬을 ~/.claude/skills/ 에 설치해줘:
https://github.com/chanjupark-951019/daily-work-log

다음 순서로 진행해줘. 각 단계가 끝날 때마다 어느 단계까지 끝났는지 한 줄로 보고만 하고, 코드 분석이나 테스트, 검증 같은 부가 작업은 일절 하지 마.

1. 사전 점검: git --version 출력. python은 `python --version` 시도 후, 응답이 없거나 Windows Store alias이면(빈 출력이거나 'was not found' 메시지) 실패로 간주하고 `python3 --version`도 시도. 둘 다 실패하면 사용자에게 "python.org에서 Python 3.x 받고 'Add to PATH' 체크 후 PC 재시작" 안내 후 즉시 중단. git 없어도 같은 방식으로 중단 안내.

2. 시스템 Chrome 자동 감지. 다음 경로 중 존재하는 첫 번째 파일을 사용:
   - Windows: C:\Program Files\Google\Chrome\Application\chrome.exe
   - Windows: C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
   - Windows: $LOCALAPPDATA\Google\Chrome\Application\chrome.exe
   - Mac: /Applications/Google Chrome.app/Contents/MacOS/Google Chrome
   - Linux: /usr/bin/google-chrome 또는 /usr/bin/chromium
   발견하면 그 경로를 변수에 저장. 못 찾으면 변수는 빈 값.

3. git clone https://github.com/chanjupark-951019/daily-work-log 을 OS 표준 temp 폴더(Windows: $env:TEMP, Mac/Linux: /tmp 또는 mktemp)에 받은 뒤, 내부의 skills/daily-work-log 폴더만 ~/.claude/skills/ 안으로 복사. 기존에 같은 이름 폴더 있으면 먼저 삭제. 임시 폴더는 정리.

4. 파이썬 패키지 설치: `python -m pip install pyyaml jinja2 playwright python-pptx` 실행 (`pip` 직접 호출보다 `python -m pip`가 안전).

5. 2단계에서 Chrome을 찾았으면:
   - 환경변수 AGENT_BROWSER_EXECUTABLE_PATH를 그 경로로 영구 설정 (Windows: `setx AGENT_BROWSER_EXECUTABLE_PATH "<path>"`. Mac/Linux: ~/.bashrc 또는 ~/.zshrc에 `export AGENT_BROWSER_EXECUTABLE_PATH="<path>"` 추가).
   - `playwright install chromium` 절대 실행 금지. (200MB 다운로드는 회사 네트워크에서 오래 걸려서 건너뜀.)
   못 찾았으면 그제서야 `python -m playwright install chromium` 실행. 시간이 5~15분 걸린다고 사용자에게 미리 알려.

6. 사용자 데이터 루트 설정:
   - 사용자에게 "일일 로그와 주간보고서를 저장할 폴더 경로를 알려주세요 (예: C:\Users\<name>\Desktop\주간업무)" 한 줄 안내 후 입력받기.
   - 입력받은 경로에 `일일업무로그/`, `보고서/` 하위 폴더 생성.
   - `~/.claude/skills/daily-work-log/starter/_categories.yaml`을 사용자 루트에 복사 (이미 있으면 건드리지 말 것).
   - 환경변수 `DAILY_WORK_LOG_ROOT`를 그 경로로 영구 설정 (Windows: `setx`. Mac/Linux: `~/.bashrc` 또는 `~/.zshrc`에 `export` 추가).

7. 완료. 사용자에게 정확히 다음 세 줄을 안내:
   - "Claude Code를 한 번 종료했다가 다시 열어주세요. (Desktop 앱은 시스템 트레이에서 우클릭 → Quit 후 재실행 / VSCode 확장은 Ctrl+Shift+P → 'Developer: Reload Window' / CLI는 새 터미널 세션)"
   - "재시작 후 채팅창에 '오늘 한 일 정리해줘 - <자연어 업무 내역>' 형태로 일일 로그를 기록할 수 있습니다."
   - "월~금 로그 쌓이면 '이번주 주간보고 만들어줘'로 PDF+PPTX가 자동 생성됩니다."

주의 (반드시 지킬 것):
- /plugin 명령은 이 환경에서 안 됨. 사용 금지.
- 다른 에이전트 설정·MCP·hooks는 전혀 건드리지 마.
- 단계별 결과만 짧게 한 줄 보고. 코드 인스펙션·테스트 실행·dry-run 금지.
- bash/powershell 명령 실행 시 사용자 승인이 필요하면 그냥 진행하고, 거부되면 다음 단계로 넘기지 말고 사용자에게 "허용 클릭이 필요합니다"만 한 줄 안내.
- 임시 폴더는 OS 표준 temp 위치 사용.
- 사용자명/경로에 한글이나 공백이 있어도 따옴표로 감싸서 처리.
