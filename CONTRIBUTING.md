# Contributing

Thanks for contributing to `ppinject`.

Current direction: API-first Python library for targeted `.pptx` XML mutation, with implementation to follow after scaffold setup.

## License and contributions

This project is licensed under GPLv3 (or later).
By submitting contributions, you agree your changes are provided under the same license.

## Setup

```bash
uv sync --dev
```

## Before opening a PR

```bash
uv run ruff check .
uv run mypy .
uv run pytest -q
uv build
uv run twine check dist/*
```

## Contribution rules

- Keep changes focused and small.
- Update relevant docs under `docs/` when behavior or plan changes.
- Add tests with behavior changes.
- Avoid unrelated cleanup in feature PRs.
- Prefer extending reusable high-level helpers over adding app-specific logic.
- Keep CLI additions minimal and generic. The library API is the primary interface.
- Use anonymized and synthetic data only in docs, tests, and issue reports.

## Early-stage guidance

- Do not claim a stable public injection API until it exists.
- Prefer PresentationML and OPC terminology in docs and code.
- Preserve XML and relationship safety requirements as first-class design constraints.

## Branch naming

Suggested prefixes:

- `feat/`
- `fix/`
- `docs/`
- `chore/`
