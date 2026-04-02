# Security Policy

## Supported Versions

The latest code on `main` is supported.

## Reporting a Vulnerability

Please report vulnerabilities privately to the maintainer before opening a public issue.

Include:
- affected module/file
- reproduction steps
- impact assessment
- suggested fix (if available)

## Security Notes

- Do not commit API keys, registry dumps, or real user data.
- Prefer local LM Studio mode for sensitive datasets.
- Review command execution paths carefully (`modAICommands.bas`) for destructive operations.
