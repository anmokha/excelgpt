#!/usr/bin/env python3
"""Split monolithic modAIHelper.bas into smaller modules for public source layout."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


PROC_RE = re.compile(r'^(Public|Private)\s+(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b', re.I)
END_RE = re.compile(r'^End\s+(Sub|Function)\b', re.I)

CONFIG_PROCS = {
    'GetApiKey',
    'SaveApiKey',
    'HasApiKey',
    'GetLMStudioSetting',
    'SaveLMStudioSetting',
    'IsLocalModelEnabled',
    'HasLocalModel',
}

NETWORK_PROCS = {
    'SendToAI',
    'ImageToBase64',
    'EncodeBase64',
    'GetMimeType',
    'BuildRequestJSONWithImage',
    'BuildSystemPrompt',
    'BuildRequestJSON',
    'EscapeJSON',
    'SendHTTPRequest',
    'ParseResponse',
    'IsLMStudioAvailable',
    'GetLMStudioModels',
    'SendToLocalAI',
    'SendLocalHTTPRequest',
}

COMMAND_PROCS = {
    'ExtractCommands',
    'ExecuteCommands',
    'ParseColor',
    'LocalizeFormula',
    'GetColumnNumber',
    'ReplaceFunc',
    'ReplaceConstant',
    'GetChartType',
    'FindPivotTable',
    'GetChartIndex',
    'ExecuteSingleCommand',
}


def find_proc_blocks(lines: list[str]) -> dict[str, list[str]]:
    blocks: dict[str, list[str]] = {}
    i = 0
    last_end = -1
    n = len(lines)
    while i < n:
        m = PROC_RE.match(lines[i].strip())
        if not m:
            i += 1
            continue

        name = m.group(3)

        # include comment block immediately above procedure
        start = i
        j = i - 1
        while j > last_end and (lines[j].strip() == '' or lines[j].lstrip().startswith("'")):
            j -= 1
        start = j + 1

        k = i + 1
        while k < n and not END_RE.match(lines[k].strip()):
            k += 1
        if k >= n:
            raise RuntimeError(f'Cannot find End Sub/Function for {name}')

        end = k
        block = lines[start:end + 1]
        blocks[name] = block

        last_end = end
        i = end + 1

    return blocks


def render_module(module_name: str, module_desc: str, const_lines: list[str], proc_names: list[str], blocks: dict[str, list[str]]) -> str:
    out: list[str] = []
    out.append("' =============================================================================")
    out.append(f"' {module_name}")
    out.append(f"' Что внутри: {module_desc}")
    out.append("' Автоматически собран из модульной структуры проекта.")
    out.append("' =============================================================================")
    out.append("")
    out.append('Option Explicit')
    out.append("")

    for line in const_lines:
        out.append(line)
    if const_lines:
        out.append("")

    for name in proc_names:
        block = blocks.get(name)
        if not block:
            raise RuntimeError(f'Missing procedure block: {name}')
        out.extend(block)
        out.append("")

    return "\n".join(out).rstrip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Split modAIHelper into config/network/commands modules")
    parser.add_argument("--source", default="src-vba/annotated/modAIHelper.bas", help="Path to source modAIHelper.bas")
    parser.add_argument("--out", default="src-vba/public/modules", help="Output directory for split modules")
    args = parser.parse_args()

    src = Path(args.source)
    out_dir = Path(args.out)

    text = src.read_text(encoding='utf-8')
    lines = text.splitlines()
    blocks = find_proc_blocks(lines)

    out_dir.mkdir(parents=True, exist_ok=True)

    config_consts = [
        "' Настройки LM Studio по умолчанию",
        'Private Const LMSTUDIO_DEFAULT_IP As String = "127.0.0.1"',
        'Private Const LMSTUDIO_DEFAULT_PORT As String = "1234"',
        "",
        "' Хранение ключей (в реестре)",
        'Private Const REG_PATH As String = "HKEY_CURRENT_USER\\Software\\ExcelAIAssistant\\"',
    ]

    network_consts = [
        "' Константы API",
        'Private Const DEEPSEEK_URL As String = "https://api.deepseek.com/chat/completions"',
        'Private Const OPENROUTER_URL As String = "https://openrouter.ai/api/v1/chat/completions"',
        'Private Const DEEPSEEK_MODEL As String = "deepseek-chat"',
        'Private Const CLAUDE_MODEL As String = "anthropic/claude-4.5-sonnet-20250929"',
        'Private Const GPT_MODEL As String = "openai/gpt-5.2"',
        'Private Const GEMINI_MODEL As String = "google/gemini-3-pro-preview"',
        'Private Const GEMINI_FLASH_MODEL As String = "google/gemini-3-flash-preview-20251217"',
    ]

    config_order = [name for name in blocks if name in CONFIG_PROCS]
    network_order = [name for name in blocks if name in NETWORK_PROCS]
    command_order = [name for name in blocks if name in COMMAND_PROCS]

    # Safety check: ensure every known proc was mapped
    all_target = CONFIG_PROCS | NETWORK_PROCS | COMMAND_PROCS
    found = set(config_order) | set(network_order) | set(command_order)
    missing_known = sorted(all_target - found)
    if missing_known:
        raise RuntimeError(f'Mapped procedures missing in source: {missing_known}')

    config_text = render_module(
        'modAIConfig.bas',
        'Доступ к API-ключам и настройкам LM Studio через реестр Windows.',
        config_consts,
        config_order,
        blocks,
    )

    network_text = render_module(
        'modAINetwork.bas',
        'Запросы к облачным моделям и LM Studio, сбор JSON, HTTP и парсинг ответов.',
        network_consts,
        network_order,
        blocks,
    )

    commands_text = render_module(
        'modAICommands.bas',
        'Парсинг и выполнение AI-команд в Excel (формат commands).',
        [],
        command_order,
        blocks,
    )

    (out_dir / 'modAIConfig.bas').write_text(config_text, encoding='utf-8')
    (out_dir / 'modAINetwork.bas').write_text(network_text, encoding='utf-8')
    (out_dir / 'modAICommands.bas').write_text(commands_text, encoding='utf-8')

    print('Generated:')
    print(' -', out_dir / 'modAIConfig.bas')
    print(' -', out_dir / 'modAINetwork.bas')
    print(' -', out_dir / 'modAICommands.bas')


if __name__ == '__main__':
    main()
