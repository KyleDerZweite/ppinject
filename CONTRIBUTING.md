# Contributing

Thanks for contributing to `ppinject`.

The published beta provides `render_template_slide` and `RenderReport` for paragraph-scoped text replacement and existing picture-shape replacement. Preserve this narrow API and the mutation boundaries in `docs/ARCHITECTURE.md`.

## License and contributions

This project is licensed under GPLv3 (or later).
By submitting contributions, you agree your changes are provided under the same license.

## Setup

Install the uv version required by `[tool.uv].required-version` in `pyproject.toml`.
The development interpreter is selected by `.python-version`; consumers require Python 3.11 or newer.

```bash
uv sync --locked --dev
uv run --locked pre-commit install
```

## Before opening a PR

```bash
uv run --locked pre-commit run --all-files
uv run --locked ruff check .
uv run --locked mypy .
uv run --locked pytest -q
uv build
uv run --locked twine check dist/*
```

## Contribution rules

- Keep changes focused and small.
- Update relevant docs under `docs/` when behavior or plan changes.
- Add tests with behavior changes.
- Avoid unrelated cleanup in feature PRs.
- Prefer extending reusable high-level helpers over adding app-specific logic.
- Keep CLI additions minimal and generic. The library API is the primary interface.
- Use anonymized and synthetic data only in docs, tests, and issue reports.

## Beta guidance

- Keep compatibility claims limited to tested scenarios.
- For XML mutations, test unchanged neighboring structure, unknown extensions, relationships and unrelated archive parts.
- Keep the Windows PowerPoint opening check separate from automated ZIP/XML checks.
- Prefer PresentationML and OPC terminology in docs and code.
- Preserve XML and relationship safety requirements as first-class design constraints.

## Branch naming

Suggested prefixes:

- `feat/`
- `fix/`
- `docs/`
- `chore/`
