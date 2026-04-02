# Contributing

## Scope

Contributions are welcome for:
- reliability and safety improvements
- Excel command coverage
- LM Studio / provider integration quality
- documentation and examples

## Development Flow

1. Keep source edits in `src-vba/public/*`.
2. Keep procedures small and single-purpose where possible.
3. Add simple comments for non-obvious logic blocks.
4. Keep user-facing texts in Russian unless a change requires localization.

## Pull Request Checklist

- [ ] Code is placed in correct module (`modAIConfig`, `modAINetwork`, `modAICommands`, etc.)
- [ ] No secrets/keys in source
- [ ] Manual smoke test performed in Excel
- [ ] README/docs updated if behavior changed

## Notes

This repository stores VBA source files directly for code review and collaboration.
