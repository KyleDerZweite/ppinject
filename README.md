# ppinject

[![CI](https://github.com/KyleDerZweite/ppinject/actions/workflows/ci.yml/badge.svg)](https://github.com/KyleDerZweite/ppinject/actions/workflows/ci.yml)

A surgical XML injector scaffold for `.pptx` files.

`ppinject` is intended to become an API-first library for targeted PresentationML mutations inside the PPTX archive while preserving layout, formatting, relationships, and unsupported XML whenever feasible.

> Status: **Scaffold only - implementation pending**

## Why this project exists

PowerPoint files are ZIP containers built on Open Packaging Conventions and PresentationML. Many higher-level libraries deserialize and rebuild large parts of the document model. This project is being structured around targeted XML-first edits instead, so future injection workflows can minimize collateral changes.

## Planned direction

- Target specific slides, shapes, and text-bearing XML nodes directly
- Preserve neighboring XML, unknown tags, and relationship structure
- Keep the library API as the primary interface
- Implement and validate the format-specific behavior on Windows

## Current repository scope

This repository currently contains:

- package scaffolding for `ppinject`
- collaboration and contribution docs
- initial architecture and roadmap notes
- CI and local quality tooling

It does not yet contain a functional PPTX injector.

## Development setup (uv)

### Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)

### Quick start

```bash
uv sync --dev
uv run pre-commit install
```

### Quality checks

```bash
uv run pre-commit run --all-files
uv run ruff check .
uv run mypy .
uv run pytest -q
uv build
uv run twine check dist/*
```

## License

`ppinject` is open source under the GNU General Public License v3.0 or later (GPL-3.0-or-later).
See [LICENSE](LICENSE) for the full license text.
