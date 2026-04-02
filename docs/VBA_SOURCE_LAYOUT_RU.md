# VBA Source Layout

Этот документ описывает рабочую структуру исходников в `src-vba/public`.

## Папки

- `src-vba/public/modules` — стандартные VBA-модули (`.bas`)
- `src-vba/public/forms` — UserForm-код (`.frm`)
- `src-vba/public/classes` — классы книги/листов (`.cls`)

## Модули

### modAIConfig.bas
Отвечает за:
- чтение/сохранение API-ключей
- чтение/сохранение настроек LM Studio
- включение/отключение локального режима

### modAINetwork.bas
Отвечает за:
- сбор системного prompt
- сбор JSON-тела запросов
- HTTP-запросы в DeepSeek/OpenRouter/LM Studio
- парсинг ответов моделей
- отправку изображений (Vision)

### modAICommands.bas
Отвечает за:
- извлечение блока `commands`
- выполнение Excel-команд
- вспомогательные функции (цвета, формулы, chart/pivot helpers)

### modExcelHelper.bas
Отвечает за:
- сбор контекста книги/листа
- сбор данных выделения
- утилиты для работы с диапазонами

### modMain.bas
Отвечает за:
- создание кнопки AI Assistant в меню Excel
- запуск форм `frmChat` и `frmSettings`
- quick-analyze сценарий

## Рекомендуемый порядок импорта в VBA Editor

1. `classes/*`
2. `modules/modAIConfig.bas`
3. `modules/modAINetwork.bas`
4. `modules/modAICommands.bas`
5. `modules/modExcelHelper.bas`
6. `modules/modMain.bas`
7. `forms/*`

## Принципы поддержки

- Логику новых интеграций API добавлять в `modAINetwork.bas`
- Логику выполнения Excel-команд добавлять в `modAICommands.bas`
- Настройки и ключи хранить только через `modAIConfig.bas`
- Не смешивать UI-код с сетевой логикой
