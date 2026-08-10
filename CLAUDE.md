# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Response language

대답을 할 때는 해요체를 사용해주세요.

## Repository overview

This is a collection of small, independent Excel-processing tasks, each in its own top-level Korean-named folder. Each task folder is self-contained: its own scripts, input `.xlsx` files, and output folder. There is no shared build system, package manifest, or top-level entry point — treat each folder as its own project and check for a `CLAUDE.md` inside it before working there.

- `기초 실적 합산/` — department expense aggregation tool, single input file per month, output written under its own `출력/` folder. Has its own `CLAUDE.md` — read it before editing anything inside this folder.
- `복잡한 실적 합산/` — multi-team performance aggregation tool. Teams submit Excel files in 4 different layouts each month; the tool normalizes and merges them. Has its own `CLAUDE.md` — read it before editing anything inside this folder.
- `크롤링 실습/` — pulls public data out of an ArcGIS Experience dashboard (Philippine MGB mining tenements) via the REST API and writes CSVs. Unlike the other folders, its inputs are remote services rather than local `.xlsx` files. Has its own `CLAUDE.md` — read it before editing anything inside this folder.

Scripts are run directly with `py`/`python` (no venv, no requirements file, no test suite in this repo). Dependencies used: `pandas` + an Excel engine (`기초 실적 합산`), `openpyxl` directly (`복잡한 실적 합산`), and stdlib only (`크롤링 실습`).

**Network access from this machine**: Python's `urllib`/`ssl` fails HTTPS certificate verification here (corporate SSL inspection: `Missing Authority Key Identifier`), and plain `curl` fails with `CRYPT_E_NO_REVOCATION_CHECK`. Working pattern: shell out to `curl -sL --ssl-no-revoke` (see `크롤링 실습/신청광구_통합.py`). Don't waste time re-diagnosing this.

## Cross-cutting rules

- Never modify or overwrite source input `.xlsx` files — treat them as read-only.
- Every processing run emits a validation/warning report alongside its output, rather than silently dropping unrecognized data.
- Don't hardcode row/column positions, team lists, sheet counts, or file names where the task folder's own `CLAUDE.md` says not to — locate data by structural markers (header labels) instead, since these inputs are designed to change shape month to month.

See each task folder's `CLAUDE.md` for its specific architecture, commands, and known open issues.

## Skills

- `.claude/skills/excel-merge/` — how to aggregate Excel files that arrive in different
  layouts per submitter (marker-based header detection, account-name aliasing, the traps
  that actually bit us, validation and output conventions). Invoke with `/excel-merge`, or
  just read it before starting any new merge task in this repo.
