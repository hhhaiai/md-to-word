# Refactor Plan and Tasks

## Scope
Minimal, production-safe refactors to improve correctness and consistency without changing external behavior. Work proceeds in P0 → P1 order. P2 covers documentation sync and decisions captured post-review.

## P0: Immediate Fixes

1) Fix case-insensitive regex usage in image caption detection
- File: `src/formatters/image_formatter.py`
- Issue: `pattern.search(text, re.IGNORECASE)` misuses the second argument (pos). CI-insensitive never applied.
- Action: Move `re.IGNORECASE` to compile-time in `src/utils/constants.py` for `IMAGE_FILENAME_PATTERNS`, and call `pattern.search(text)`.

2) Unify `DocumentConfig` import path
- Files: `src/core/word_postprocessor.py` (and keep others as-is)
- Issue: Mixed imports (`from ..config.config import DocumentConfig` vs `from ..config import DocumentConfig`).
- Action: Use `from ..config import DocumentConfig` everywhere in core to avoid ambiguity.

3) Remove duplicated pandoc args
- Files: `src/core/pandoc_processor.py`, `src/config/config.py`
- Issue: `--preserve-tabs` and `--wrap=none` configured in both places.
- Action: Keep args only in `config.py` and remove duplicates from `_get_pandoc_args()`.

## P1: Consistency Improvements

4) Deduplicate math formula detection
- Files: `src/core/word_postprocessor.py`, `src/formatters/base_formatter.py`
- Issue: `_has_math_formula` duplicated.
- Action: Remove the duplicate in `WordPostprocessor` and reuse `BaseFormatter` implementation.

5) Narrow image re-insertion scope to Obsidian syntax only
- File: `src/core/word_postprocessor.py`
- Issue: `process_and_insert_images` parses both Markdown `![...](...)` and Obsidian `![[...]]`, overlapping with pandoc’s native image handling.
- Action: Limit re-insertion/rewriting to Obsidian `![[...]]` only and rely on pandoc for standard Markdown images.

6) Add non-interactive CLI flag for overwrite
- File: `md_to_word.py`
- Issue: Overwrite confirmation uses `input()` which breaks automation.
- Action: Add `--force` to skip prompt and overwrite if output exists.

## P2: Docs and Behavior Decisions (Done)
- Align security docs to subprocess list-args; add note about `--force`
- Architecture doc: adjust pandoc args example and exception hierarchy; clarify image handling (Obsidian reinsert, Markdown via Pandoc)
- README: update safety statement and CLI usage, include `--force`
- Decision: Do NOT remove filename-like noise from normal paragraphs. Reverted previous attempt; `_remove_image_captions_from_all_paragraphs` kept as a no-op for compatibility and can be deleted in a future breaking change.

## Out of Scope (future)
- Optionally move Obsidian syntax normalization to preprocessor or Lua filter.
- Reconcile docs and exception taxonomy (`MarkdownParsingError`) if needed.
 - Remove `_remove_image_captions_from_all_paragraphs` entirely in next major if no external users.

## Risk Assessment
- Changes are local and backward-compatible. Primary behavioral change is restricting post-processor’s image text parsing to Obsidian syntax, which matches design intent and avoids double-processing.
 - Added pandoc flags for inline math parsing and resource-path to ensure image and math reliability.

## Tasks Checklist
- [x] P0-1: constants – add `re.IGNORECASE` to `IMAGE_FILENAME_PATTERNS`
- [x] P0-2: image_formatter – use `pattern.search(text)` without flags
- [x] P0-3: word_postprocessor – unify `DocumentConfig` import
- [x] P0-4: pandoc_processor – remove duplicate args in `_get_pandoc_args`
- [x] P1-1: word_postprocessor – remove duplicate logic by delegating math detection to shared impl
- [x] P1-2: word_postprocessor – restrict `process_and_insert_images` to Obsidian images
- [x] P1-3: md_to_word – add `--force` to skip overwrite prompt
- [x] P2-1: security/README/architecture docs synced
- [x] P2-2: decision recorded – no noise removal in normal paragraphs; revert related code path
- [x] P1-extra: preprocessor – ensure ordered list items become separate paragraphs (insert blank lines)
- [x] P1-extra: processor – pass `--resource-path` so markdown images resolve relative to input
- [x] P1-extra: processor – enable inline math with `markdown+tex_math_dollars+tex_math_single_backslash`

## Redundancy / Dead Code Review
- `src/utils/constants.py`: removed unused `ORDERED_LIST_PATTERN_PREPROCESSOR` (merged intent into comments). No remaining references.
- `src/formatters/image_formatter.py`:
  - `_remove_image_captions_from_all_paragraphs` is now a no-op per product decision and no longer referenced in flow. Safe to delete in next major. Kept to avoid accidental API break.
  - Call site removed from `format_images`.
- No other dead pathways detected. Formatters are all referenced by `WordPostprocessor`. XPath helpers are used across image/table formatters. Config validator entry-point remains callable via CLI.


