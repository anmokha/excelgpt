# Архитектура

## Поток выполнения

1. Пользователь вводит задачу в `frmChat`.
2. `modExcelHelper` собирает контекст книги и выделения.
3. `modAINetwork` формирует prompt и отправляет запрос в провайдера.
4. Модель возвращает текстовый ответ.
5. `modAICommands` извлекает блок `commands`.
6. `modAICommands` исполняет команды через объектную модель Excel.
7. `frmChat` показывает результат пользователю.

## Слои

- UI слой:
  - `src-vba/public/forms/frmChat.frm`
  - `src-vba/public/forms/frmSettings.frm`
- Конфигурация:
  - `src-vba/public/modules/modAIConfig.bas`
- Сетевой слой:
  - `src-vba/public/modules/modAINetwork.bas`
- Командный движок:
  - `src-vba/public/modules/modAICommands.bas`
- Excel helper слой:
  - `src-vba/public/modules/modExcelHelper.bas`
- Точки входа add-in:
  - `src-vba/public/modules/modMain.bas`

## Поставщики моделей

- DeepSeek API
- OpenRouter API (Claude/GPT/Gemini)
- LM Studio (`/v1/chat/completions`, `/v1/models`)

## Границы ответственности

- `modAIConfig`:
  - ключи API и настройки LM Studio
- `modAINetwork`:
  - JSON, HTTP, retry/ошибки, парсинг ответа
- `modAICommands`:
  - протокол `commands`, исполнение Excel-действий
- `modExcelHelper`:
  - контекст книги, выборка данных

## Риски

- AI может сгенерировать агрессивные команды (`DELETE_*`, `CLEAR_*`).
- Регистровое хранение ключей требует аккуратной операционной практики.
- VBA сложнее автоматизировать тестами, поэтому важны smoke-сценарии.
