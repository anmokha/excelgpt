# Changelog

Все заметные изменения проекта публикуются в этом файле.

Формат основан на Keep a Changelog, версия — SemVer.

## [0.2.0] - 2026-04-02

### Added
- Публичная модульная структура исходников `src-vba/public`.
- Разделение `modAIHelper` на `modAIConfig`, `modAINetwork`, `modAICommands`.
- Скрипт `tools/build_public_sources.py` для обновления публичного layout из `.xlam`.
- CI workflow `.github/workflows/ci.yml`.
- Репозиторная валидация `tools/validate_repo.py`.
- Документы `ARCHITECTURE_RU`, `QA_SMOKE_TESTS_RU`, `CASE_STUDY_RU`.

### Changed
- README обновлён и структурирован для публичной документации.
- Документация по исходникам и поддержке структурирована.

## [0.1.0] - 2026-02-05

### Added
- Базовая надстройка `AI_Assistant.xlam`.
- Поддержка облачных и локальных моделей.
- Командный протокол для автоматизации действий в Excel.
