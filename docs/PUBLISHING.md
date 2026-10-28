# Publishing to PyPI

This document describes the release flow for `ppinject`.

## Existing release state

`v0.1.0b2` was published on 2026-04-15. As checked on 2026-10-28, the GitHub
Release has `prerelease: false` even though the package version is a beta.
This metadata discrepancy remains for the release owner to correct. No repeat
publication is needed. Windows PowerPoint validation evidence is not recorded
in this repository and remains an open release check.

## Publisher configuration

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

The source distribution explicitly includes the package, tests, examples, docs and
package metadata. Keep this selection in `pyproject.toml` so local environments
and caches cannot enter release archives.

## Local preflight checks

```bash
uv sync --locked --dev
uv run --locked pre-commit run --all-files
uv run --locked ruff check .
uv run --locked mypy .
uv run --locked pytest -q
uv build
uv run --locked twine check dist/*
```

## Manual release gate

Before the next release, open a generated output from the bundled synthetic template in PowerPoint on Windows. Confirm there is no repair prompt and check replaced text, images and unchanged layout. Record the source commit, template, Windows and PowerPoint versions, steps, result and validator in the release notes. Automated tests do not close this check.

## Release process

1. Bump `version` in `pyproject.toml` and the fallback in `src/ppinject/__init__.py`, update the version test, then run `uv lock`.
2. Update release notes and docs so the supported scope stays narrow and accurate.
3. Commit and push to `main`.
4. Confirm GitHub `CI` passes on `main`.
5. Create a GitHub Release with a matching tag such as `v0.1.0b2`.
6. Mark beta releases as pre-releases in GitHub Releases.
7. The `publish-pypi.yml` workflow runs automatically and publishes to PyPI.

## Verify installation

Use a fresh consumer environment and replace the version below with the release being verified.

```bash
python -m pip install --upgrade ppinject==0.1.0b2
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
