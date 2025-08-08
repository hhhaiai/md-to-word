# Repository Guidelines

## Project Structure & Modules
- `src/core/`: Orchestrates the pipeline (`MarkdownPreprocessor` → `PandocProcessor` → `WordPostprocessor`).
- `src/formatters/`: Specialized formatters (page, paragraph, title, table, list, image, base).
- `src/utils/`: Utilities and guards (config/path validators, constants, exceptions, XPath cache).
- `src/config/`: Tunable settings used across processors.
- `md_to_word.py`: CLI entry point. `docs/` architecture and security notes. `examples/` sample inputs.

## Build, Test, and Dev Commands
- Create venv (optional): `python3 -m venv venv && source venv/bin/activate`.
- Install deps: `pip3 install -r requirements.txt`.
- Config check only: `python3 md_to_word.py --check-config`.
- Convert example: `python3 md_to_word.py examples/example.md -o examples/example.docx`.
- Verbose debug: `python3 md_to_word.py input.md -o out.docx -v --force`.

## Coding Style & Naming
- Follow PEP 8 (4-space indents, 120-col soft wrap). Use type hints where practical.
- Names: modules/files `snake_case`, classes `PascalCase`, functions/vars `snake_case`, constants `UPPER_SNAKE` (see `src/utils/constants.py`).
- Keep processors small and composable; add logic via new formatter classes rather than bloating existing ones.
- Exceptions: raise project-specific ones from `src/utils/exceptions.py`.

## Testing Guidelines
- No formal test suite yet. For manual checks: run on `examples/example.md` and confirm `.docx` is produced and opens correctly.
- When adding tests, prefer `pytest` in `tests/` with `test_*.py`. Use small Markdown fixtures; assert CLI exit codes and file creation; avoid binary `.docx` diffs.

## Commit & Pull Requests
- Use Conventional Commits: `feat:`, `fix:`, `docs:`, `refactor:`, `chore:` (see `git log`). Imperative, present tense, concise scope.
- PRs should include: purpose/motivation, key changes, run commands used for validation, screenshots of resulting Word output when UI-visible changes apply, and links to related issues.
- Update `docs/` if behavior/config changes; note security implications.

## Security & Configuration Tips
- Pandoc is required; ensure it’s installed and callable. Never build shell strings—use `subprocess.run([...])` as in `PandocProcessor`.
- Validate all paths with existing helpers; forbid traversal (`..`) and unsafe targets.
- Respect env vars `OBSIDIAN_VAULT_NAME`, `OBSIDIAN_ATTACHMENTS_FOLDER`, `OBSIDIAN_VAULT_PATH` (see `docs/configuration.md`); don’t hardcode local paths.
