#!/usr/bin/env python3
"""
Extract VBA code from an XLAM/XLSM file and generate two source trees:

1) raw/       - plain extracted modules
2) annotated/ - same modules with simple human-readable procedure comments

Usage:
  python tools/extract_and_annotate_vba.py AI_Assistant.xlam
"""

from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path


FILE_NAME_MAP = {
    "ЭтаКнига.cls": "ThisWorkbook.cls",
    "Лист1.cls": "Sheet1.cls",
}

MODULE_DESCRIPTIONS = {
    "frmChat.frm": "UI-форма основного чата: отправка запросов, запуск команд Excel, работа с вложением.",
    "frmSettings.frm": "UI-форма настроек: API-ключи, параметры LM Studio, выбор локальной модели.",
    "modAIHelper.bas": "Сетевой слой и логика AI: сбор prompt, HTTP-запросы, разбор ответа, выполнение AI-команд.",
    "modExcelHelper.bas": "Работа с Excel-контекстом: анализ книги, выборки данных и вспомогательные Excel-утилиты.",
    "modMain.bas": "Точки входа add-in: кнопка в меню Excel, запуск форм, быстрый сценарий анализа.",
    "ThisWorkbook.cls": "Класс книги (пустой в текущей версии).",
    "Sheet1.cls": "Класс листа (пустой в текущей версии).",
}


PROC_RE = re.compile(
    r"^(?P<indent>\s*)(?P<scope>Public|Private)\s+(?P<kind>Sub|Function)\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)


def run_olevba_code(input_file: Path) -> str:
    cmd = ["python", "-m", "oletools.olevba", "-c", str(input_file)]
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    return result.stdout


def normalize_macro_name(name: str) -> str:
    return FILE_NAME_MAP.get(name, name)


def apply_known_extraction_fixes(module_name: str, text: str) -> str:
    """Fixes for known extraction artifacts from binary VBA streams."""
    if module_name == "frmChat.frm":
        text = text.replace('lblAttachment.Caption = "?? " & GetFileName(attachedImagePath)', 'lblAttachment.Caption = "[IMG] " & GetFileName(attachedImagePath)')
        text = text.replace('parts = Split(fullPath, "")', 'parts = Split(fullPath, "\\")')
    return text


def parse_modules(olevba_output: str) -> dict[str, str]:
    modules: dict[str, list[str]] = {}
    current_name: str | None = None
    in_code = False

    for line in olevba_output.splitlines():
        if line.startswith("VBA MACRO "):
            macro_name = line[len("VBA MACRO ") :].strip()
            current_name = normalize_macro_name(macro_name)
            modules[current_name] = []
            in_code = False
            continue

        if current_name is None:
            continue

        if not in_code:
            if line.startswith("- - - - - - -"):
                in_code = True
            continue

        if line.startswith("-------------------------------------------------------------------------------"):
            current_name = None
            in_code = False
            continue

        modules[current_name].append(line)

    cleaned: dict[str, str] = {}
    for name, code_lines in modules.items():
        text = "\n".join(code_lines).strip("\n")
        if text == "(empty macro)" or not text.strip():
            text = "Option Explicit\n"
        text = apply_known_extraction_fixes(name, text)
        cleaned[name] = text + "\n"
    return cleaned


def split_args(signature_line: str) -> str:
    if "(" not in signature_line or ")" not in signature_line:
        return "нет"
    args = signature_line.split("(", 1)[1].rsplit(")", 1)[0].strip()
    if not args:
        return "нет"

    parts = []
    for raw in args.split(","):
        token = raw.strip()
        token = re.sub(r"\b(ByVal|ByRef|Optional|ParamArray)\b", "", token, flags=re.I).strip()
        token = token.split("As", 1)[0].strip()
        token = token.split("=", 1)[0].strip()
        if token:
            parts.append(token)
    return ", ".join(parts) if parts else "нет"


def describe_procedure(name: str, kind: str) -> str:
    low = name.lower()
    if low.startswith("get"):
        return "Читает данные из Excel/настроек и возвращает результат."
    if low.startswith("set"):
        return "Записывает значение в Excel или в настройки."
    if low.startswith("save"):
        return "Сохраняет данные в постоянное хранилище."
    if low.startswith("has") or low.startswith("is"):
        return "Проверяет условие и возвращает True/False."
    if low.startswith("send"):
        return "Отправляет запрос во внешний сервис и получает ответ."
    if low.startswith("build"):
        return "Собирает служебную строку или JSON для следующего шага."
    if low.startswith("parse") or low.startswith("extract"):
        return "Разбирает текст и извлекает структурированные данные."
    if low.startswith("execute"):
        return "Выполняет подготовленные команды в Excel."
    if low.startswith("create"):
        return "Создаёт новый объект или структуру в Excel."
    if low.startswith("delete") or low.startswith("remove"):
        return "Удаляет объект или очищает данные."
    if low.startswith("update"):
        return "Обновляет состояние интерфейса или данных."
    if low.endswith("_click") or low.endswith("_keydown") or low.endswith("_initialize"):
        return "Обработчик события формы."
    if kind.lower() == "function":
        return "Вспомогательная функция модуля."
    return "Вспомогательная процедура модуля."


def ensure_option_explicit_top(text: str) -> str:
    lines = text.splitlines()
    if any(line.strip().lower() == "option explicit" for line in lines[:40]):
        return text
    return "Option Explicit\n\n" + text


def annotate_module(module_name: str, raw_code: str) -> str:
    raw_code = ensure_option_explicit_top(raw_code).replace("\r\n", "\n")
    lines = raw_code.splitlines()

    func_count = sum(1 for line in lines if PROC_RE.match(line))
    description = MODULE_DESCRIPTIONS.get(module_name, "VBA module for Excel AI Assistant.")

    header = [
        "' =============================================================================",
        f"' {module_name}",
        f"' Что внутри: {description}",
        f"' Количество процедур: {func_count}",
        "' Стиль комментариев: простой язык для быстрой поддержки.",
        "' =============================================================================",
        "",
    ]

    out: list[str] = header
    for line in lines:
        m = PROC_RE.match(line)
        if m:
            kind = m.group("kind")
            name = m.group("name")
            args = split_args(line)
            desc = describe_procedure(name, kind)

            out.append("' ---")
            out.append(f"' Что делает: {desc}")
            out.append(f"' Вход: {args}")
            out.append(f"' Выход: {'значение функции' if kind.lower() == 'function' else 'нет (процедура)'}")
            out.append("' ---")

            out.append(line.rstrip())
            continue

        out.append(line.rstrip())

    return "\n".join(out).rstrip() + "\n"


def write_files(base_dir: Path, modules: dict[str, str], annotate: bool) -> None:
    base_dir.mkdir(parents=True, exist_ok=True)
    for module_name, code in modules.items():
        path = base_dir / module_name
        text = annotate_module(module_name, code) if annotate else code
        path.write_text(text, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract and annotate VBA code from XLAM/XLSM")
    parser.add_argument("input_file", help="Path to .xlam/.xlsm file")
    parser.add_argument("--out", default="src-vba", help="Output directory (default: src-vba)")
    args = parser.parse_args()

    input_file = Path(args.input_file)
    out_dir = Path(args.out)
    raw_dir = out_dir / "raw"
    annotated_dir = out_dir / "annotated"

    text = run_olevba_code(input_file)
    modules = parse_modules(text)

    write_files(raw_dir, modules, annotate=False)
    write_files(annotated_dir, modules, annotate=True)

    print(f"Extracted {len(modules)} modules")
    print(f"Raw:       {raw_dir}")
    print(f"Annotated: {annotated_dir}")


if __name__ == "__main__":
    main()
