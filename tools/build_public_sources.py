#!/usr/bin/env python3
"""Build public VBA source layout from XLAM/XLSM file."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
EXTRACT_SCRIPT = ROOT / 'tools' / 'extract_and_annotate_vba.py'
SPLIT_SCRIPT = ROOT / 'tools' / 'split_aihelper_modules.py'
PUBLIC_DIR = ROOT / 'src-vba' / 'public'


def run(cmd: list[str]) -> None:
    subprocess.run(cmd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description='Build src-vba/public from add-in file')
    parser.add_argument('input_file', help='Path to .xlam/.xlsm file')
    args = parser.parse_args()

    input_file = Path(args.input_file)

    forms_dir = PUBLIC_DIR / 'forms'
    classes_dir = PUBLIC_DIR / 'classes'
    modules_dir = PUBLIC_DIR / 'modules'

    forms_dir.mkdir(parents=True, exist_ok=True)
    classes_dir.mkdir(parents=True, exist_ok=True)
    modules_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix='vba_build_') as tmp:
        tmp_dir = Path(tmp)

        run(['python', str(EXTRACT_SCRIPT), str(input_file), '--out', str(tmp_dir / 'src-vba')])

        annotated = tmp_dir / 'src-vba' / 'annotated'

        # Split monolithic AI helper into public modules
        run([
            'python',
            str(SPLIT_SCRIPT),
            '--source', str(annotated / 'modAIHelper.bas'),
            '--out', str(modules_dir),
        ])

        # Copy remaining source files to public layout
        shutil.copy2(annotated / 'modExcelHelper.bas', modules_dir / 'modExcelHelper.bas')
        shutil.copy2(annotated / 'modMain.bas', modules_dir / 'modMain.bas')

        shutil.copy2(annotated / 'frmChat.frm', forms_dir / 'frmChat.frm')
        shutil.copy2(annotated / 'frmSettings.frm', forms_dir / 'frmSettings.frm')

        shutil.copy2(annotated / 'ThisWorkbook.cls', classes_dir / 'ThisWorkbook.cls')
        shutil.copy2(annotated / 'Sheet1.cls', classes_dir / 'Sheet1.cls')

    print('Public source layout updated:')
    print(' -', modules_dir)
    print(' -', forms_dir)
    print(' -', classes_dir)


if __name__ == '__main__':
    main()
