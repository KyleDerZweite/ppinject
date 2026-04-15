# Publishing to PyPI

This document describes the release flow for `ppinject`.

## One-time setup

1. In PyPI, configure a Trusted Publisher or Pending Publisher for this GitHub repository:
   - Owner: `KyleDerZweite`
   - Repository: `ppinject`
   - Workflow file: `.github/workflows/publish-pypi.yml`
   - Environment: `pypi`
2. In GitHub repository settings, create the `pypi` environment.
3. Restrict the `pypi` environment to:
   - branch: `main`
   - tag: `v*`
4. Do not add a PyPI API token secret when using Trusted Publisher.

## Local preflight checks

```bash
uv sync --dev
uv run pre-commit run --all-files
uv run ruff check .
uv run mypy .
uv run pytest -q
uv build
uv run twine check dist/*
```

## Manual release gate

Before the first public release, open at least one generated output in PowerPoint on Windows and confirm there is no repair prompt for the currently supported template-rendering flow.

## Release process

1. Bump `version` in `pyproject.toml`.
2. Update release notes and docs so the supported scope stays narrow and accurate.
3. Commit and push to `main`.
4. Confirm GitHub `CI` passes on `main`.
5. Create a GitHub Release with a matching tag such as `v0.1.0b2`.
6. Mark beta releases as pre-releases in GitHub Releases.
7. The `publish-pypi.yml` workflow runs automatically and publishes to PyPI.

## Verify installation

```bash
python -m pip install --upgrade ppinject
python -c "import ppinject; print(ppinject.__version__)"
```

## Release notes guidance

Release notes should state:

1. What concrete template-rendering scenario is supported.
2. What is intentionally out of scope.
3. Any XML compatibility fixes or known PowerPoint limitations.

## Security notes

- Do not store PyPI API tokens in repository secrets when using Trusted Publisher.
- Keep templates, images, tests, and examples synthetic and anonymized.
- Never publish customer presentations or logs with sensitive identifiers.
