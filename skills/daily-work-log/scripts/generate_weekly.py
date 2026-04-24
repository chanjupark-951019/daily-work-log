"""
주간업무 보고서 생성 파이프라인 (portable).
일일업무로그(월~금 YAML) + 카테고리 사전 → status/gantt HTML 렌더 → PDF + PPTX 동시 저장.

Usage:
  python generate_weekly.py --monday 2026-04-20 --week-label "4월 4주차"
  python generate_weekly.py --monday 2026-04-20 --week-label "4월 4주차" --plans plans_4주차.json

사용자 데이터 루트는 환경변수 DAILY_WORK_LOG_ROOT 사용.
  Windows: setx DAILY_WORK_LOG_ROOT "C:\\path\\to\\주간업무"
  Mac/Linux: export DAILY_WORK_LOG_ROOT="/path/to/주간업무"
"""
import argparse
import json
import os
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import yaml
from jinja2 import Environment, FileSystemLoader, select_autoescape
from playwright.sync_api import sync_playwright

SKILL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_DIR / 'scripts'))
from build_pptx import build_status_pptx, build_gantt_pptx

DATA_ROOT = os.environ.get('DAILY_WORK_LOG_ROOT')
if not DATA_ROOT:
    print('[ERROR] DAILY_WORK_LOG_ROOT 환경변수가 설정되지 않았습니다.', file=sys.stderr)
    print('  Windows: setx DAILY_WORK_LOG_ROOT "C:\\path\\to\\주간업무"  후 터미널 재시작', file=sys.stderr)
    print('  Mac/Linux: ~/.bashrc 또는 ~/.zshrc에 export DAILY_WORK_LOG_ROOT=... 추가', file=sys.stderr)
    sys.exit(2)

ROOT = Path(DATA_ROOT)
LOG_DIR = ROOT / '일일업무로그'
CATS_FILE = ROOT / '_categories.yaml'
TPL_DIR = SKILL_DIR / 'templates'

DAY_KEYS = ['mon', 'tue', 'wed', 'thu', 'fri']


def load_categories():
    data = yaml.safe_load(CATS_FILE.read_text(encoding='utf-8'))
    return {c['name']: c for c in data['categories']}


def load_daily_logs(monday: datetime):
    logs = {}
    for i, key in enumerate(DAY_KEYS):
        fp = LOG_DIR / f'{(monday + timedelta(days=i)):%Y-%m-%d}.yaml'
        logs[key] = yaml.safe_load(fp.read_text(encoding='utf-8')) if fp.exists() else None
    return logs


def collect_categories(logs):
    order, seen = [], set()
    for key in DAY_KEYS:
        log = logs.get(key)
        if not log:
            continue
        for entry in log.get('entries', []):
            cat = entry['category']
            if cat not in seen:
                seen.add(cat)
                order.append(cat)
    return order


def _short_date(d):
    if hasattr(d, 'month'):
        return f'{d.month}/{d.day}'
    mm, dd = str(d)[-5:].split('-')
    return f'{int(mm)}/{int(dd)}'


def build_status_data(logs, cats_dict, cat_order, plans_override):
    out = []
    for name in cat_order:
        meta = cats_dict.get(name, {})
        done = []
        for key in DAY_KEYS:
            log = logs.get(key)
            if not log:
                continue
            for entry in log.get('entries', []):
                if entry['category'] == name:
                    d = _short_date(log['date'])
                    for item in entry.get('done', []):
                        done.append(f'{item} ({d})')
        out.append({
            'name': name,
            'overview': meta.get('overview', ''),
            'done': done,
            'plan': plans_override.get(name, []),
        })
    return out


def build_gantt_data(logs, cat_order, monday, plans_override):
    def _dates(start):
        return {k: f'{(start+timedelta(days=i)).month}/{(start+timedelta(days=i)).day}'
                for i, k in enumerate(DAY_KEYS)}

    rows_actual = []
    for name in cat_order:
        row = {'task': name, 'due': '', **{k: [] for k in DAY_KEYS}}
        for key in DAY_KEYS:
            log = logs.get(key)
            if not log:
                continue
            for entry in log.get('entries', []):
                if entry['category'] == name:
                    row[key].extend(entry.get('done', []))
        rows_actual.append(row)

    rows_plan = []
    for name in cat_order:
        row = {'task': name, 'due': '', **{k: [] for k in DAY_KEYS}}
        p = plans_override.get(name, {})
        if isinstance(p, list):
            row['mon'] = p
        elif isinstance(p, dict):
            for k in DAY_KEYS:
                if p.get(k):
                    row[k] = p[k]
            row['due'] = p.get('due', '')
        rows_plan.append(row)

    return [
        {'title': '주간 업무 실적', 'dates': _dates(monday), 'rows': rows_actual},
        {'title': '주간 업무 계획', 'dates': _dates(monday + timedelta(days=7)), 'rows': rows_plan},
    ]


def render_to_pdf(html_text, out_pdf: Path, *, landscape=False):
    with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False, encoding='utf-8') as f:
        f.write(html_text)
        tmp = f.name
    try:
        with sync_playwright() as p:
            chrome_path = os.environ.get('AGENT_BROWSER_EXECUTABLE_PATH')
            if chrome_path and Path(chrome_path).exists():
                browser = p.chromium.launch(executable_path=chrome_path)
            else:
                browser = p.chromium.launch()
            page = browser.new_context().new_page()
            page.goto(Path(tmp).absolute().as_uri())
            page.emulate_media(media='print')
            page.pdf(path=str(out_pdf), format='A4', landscape=landscape,
                     print_background=True,
                     margin={'top': '0', 'bottom': '0', 'left': '0', 'right': '0'})
            browser.close()
    finally:
        Path(tmp).unlink(missing_ok=True)


def week_of_month(monday):
    first = monday.replace(day=1)
    first_sun = first - timedelta(days=(first.weekday() + 1) % 7)
    return ((monday - first_sun).days // 7) + 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--monday', required=True, help='해당 주 월요일, YYYY-MM-DD')
    ap.add_argument('--week-label', required=True, help='예: 4월 4주차')
    ap.add_argument('--plans', help='차주 계획 override JSON (선택)')
    ap.add_argument('--author', default='', help='작성자 이름(파일명에 포함, 선택)')
    args = ap.parse_args()

    monday = datetime.strptime(args.monday, '%Y-%m-%d')
    plans = json.loads(Path(args.plans).read_text(encoding='utf-8')) if args.plans else {}

    out_dir = ROOT / '보고서' / f'{monday.year}-{monday.month:02d}'
    out_dir.mkdir(parents=True, exist_ok=True)

    cats = load_categories()
    logs = load_daily_logs(monday)
    cat_order = collect_categories(logs)
    if not cat_order:
        print(f'[ERROR] 해당 주({args.monday}~)에 일일 로그가 하나도 없습니다.', file=sys.stderr)
        sys.exit(1)

    year_short = str(monday.year)[-2:]
    month = monday.month
    wk = week_of_month(monday)
    author_suffix = f'_{args.author}' if args.author else ''

    env = Environment(loader=FileSystemLoader(str(TPL_DIR)),
                      autoescape=select_autoescape(['html']))

    # Status
    status_cats = build_status_data(logs, cats, cat_order, plans.get('status', {}))
    status_html = env.get_template('status.html.j2').render(
        categories=status_cats, year_short=year_short, month=month, week_num=wk,
        week_label=args.week_label, author=args.author,
    )
    status_pdf = out_dir / f'주간업무현황({args.author+"_" if args.author else ""}{args.week_label}).pdf'
    status_pptx = out_dir / f'주간업무현황({args.author+"_" if args.author else ""}{args.week_label}).pptx'
    render_to_pdf(status_html, status_pdf, landscape=False)
    build_status_pptx(status_pptx, status_cats, year_short, month, wk, author=args.author)
    print(f'[OK] {status_pdf.name}')
    print(f'[OK] {status_pptx.name}')

    # Gantt
    sheets = build_gantt_data(logs, cat_order, monday, plans.get('gantt', {}))
    sheets[0]['title'] = f'주간 업무 실적 ({month}월 {wk}주차)'
    sheets[1]['title'] = f'주간 업무 계획 ({month}월 {wk+1}주차)'
    gantt_html = env.get_template('gantt.html.j2').render(
        sheets=sheets, page_title=f'주간업무 실적 및 계획 {args.week_label}', author=args.author,
    )
    gantt_pdf = out_dir / f'주간업무 실적 및 계획({args.author+"_" if args.author else ""}{args.week_label}_{month}월 {wk+1}주차).pdf'
    gantt_pptx = out_dir / f'주간업무 실적 및 계획({args.author+"_" if args.author else ""}{args.week_label}_{month}월 {wk+1}주차).pptx'
    render_to_pdf(gantt_html, gantt_pdf, landscape=True)
    build_gantt_pptx(gantt_pptx, sheets, author=args.author)
    print(f'[OK] {gantt_pdf.name}')
    print(f'[OK] {gantt_pptx.name}')


if __name__ == '__main__':
    main()
