#!/usr/bin/env python3
"""Basic repository quality checks for CI."""

from __future__ import annotations

import re
import sys
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parent.parent

REQUIRED_FILES = [
    'README.md',
    'CHANGELOG.md',
    'CASE_STUDY_RU.md',
    'CONTRIBUTING.md',
    'SECURITY.md',
    'docs/ARCHITECTURE_RU.md',
    'docs/QA_SMOKE_TESTS_RU.md',
    'docs/VBA_SOURCE_LAYOUT_RU.md',
    'src-vba/public/modules/modAIConfig.bas',
    'src-vba/public/modules/modAINetwork.bas',
    'src-vba/public/modules/modAICommands.bas',
    'src-vba/public/modules/modExcelHelper.bas',
    'src-vba/public/modules/modMain.bas',
    'src-vba/public/forms/frmChat.frm',
    'src-vba/public/forms/frmSettings.frm',
]

LOCAL_LINK_RE = re.compile(r'\[[^\]]+\]\(([^)]+)\)')


def fail(msg: str) -> None:
    print(f'[FAIL] {msg}')
    sys.exit(1)


def check_required_files() -> None:
    missing = [f for f in REQUIRED_FILES if not (ROOT / f).exists()]
    if missing:
        fail(f'Missing required files: {missing}')
    print('[OK] required files')


def check_local_markdown_links() -> None:
    md_files = list(ROOT.glob('*.md')) + list((ROOT / 'docs').glob('*.md'))
    for path in md_files:
        text = path.read_text(encoding='utf-8', errors='ignore')
        for target in LOCAL_LINK_RE.findall(text):
            if target.startswith('http://') or target.startswith('https://'):
                continue
            if target.startswith('#'):
                continue
            if target.startswith('mailto:'):
                continue
            target_path = (path.parent / unquote(target)).resolve()
            if not target_path.exists():
                fail(f'Broken local link in {path.relative_to(ROOT)}: {target}')
    print('[OK] local markdown links')


def main() -> None:
    check_required_files()
    check_local_markdown_links()
    print('[OK] validation complete')


if __name__ == '__main__':
    main()
