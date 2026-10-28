# AGENTS.md

This repository is prepared for human + AI collaboration.

## Current phase

- Published beta with the narrow `render_template_slide` API and `RenderReport`
- Keep paragraph-scoped text replacement and existing picture-shape targeting as the supported boundary
- Windows PowerPoint validation remains open in `docs/ROADMAP.md`; ZIP/XML tests do not establish it
- Contributors should update `docs/ROADMAP.md` when the intended scope changes

## Working agreements for coding agents

1. Keep edits surgical and scope-limited to the requested task.
2. Prefer preserving existing structure and naming.
3. Do not introduce broad refactors during feature work.
4. When touching behavior, update docs in `docs/` in the same change.
5. When changing PPTX XML writes, test unchanged neighboring XML and unrelated archive parts.
6. Do not add dependencies without documenting rationale in the change.
7. Avoid em dashes (`--`) and avoid emoji usage in repository text/docs.
8. Keep the library API as the primary interface; avoid task-specific wrappers.

## Implementation guardrails

- Prioritize XML node-level mutation over full presentation reconstruction.
- Preserve unsupported or unknown tags and attributes exactly where feasible.
- Keep generated XML ordering deterministic and PowerPoint-compatible.
- Treat relationships, namespace declarations, and slide structure as safety-critical.
- Prefer explicit targeting of shapes, placeholders, or tagged regions over heuristic rewrites.

## Data safety and sharing

- Do not commit customer presentations, business identifiers, or API tokens into this repository.
- Use synthetic fixtures in tests and docs.
- Keep integration examples anonymized.
- If you publish logs, redact presentation paths and business identifiers.

## Release and packaging expectations

- Use the uv version required in `pyproject.toml` and the development Python version in `.python-version`.
- Use `uv sync --locked --dev` and `uv run --locked` for checks; update `uv.lock` deliberately with dependency changes.
- Keep the package install/build flow in `pyproject.toml`.
- Prefer small, incremental PRs aligned with roadmap phases.
- Follow `docs/PUBLISHING.md` for the existing Trusted Publisher workflow and release validation gate.

## Scope

- API-first PPTX archive traversal and targeted XML mutation
- Preservation-first handling of slide XML and relationships
- Windows-oriented implementation and validation for format-sensitive behavior
- Minimal generic tooling only after the core library API exists
