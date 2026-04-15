# Publishing

Publishing automation is still manual, but the repository should now be kept in a state that is ready for a first honest package release.

## Preflight checklist

1. Confirm the supported behavior statement in `README.md` is still accurate.
2. Run `uv run ruff check .`.
3. Run `uv run mypy .`.
4. Run `uv run pytest -q`.
5. Run `uv build`.
6. Run `uv run twine check dist/*`.
7. Open at least one generated output in PowerPoint on Windows and confirm there is no repair prompt.

## Release notes guidance

Release notes should state:

1. What concrete template-rendering scenario is supported.
2. What is intentionally out of scope.
3. Any XML compatibility fixes or known PowerPoint limitations.

## Before enabling automated publication

1. Add a release workflow and trusted publisher configuration together.
2. Decide whether versioning should stay beta or move to a broader stability signal.
3. Keep examples and tests synthetic so the public repository does not depend on customer data.
