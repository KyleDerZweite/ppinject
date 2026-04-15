# ppinject

[![CI](https://github.com/KyleDerZweite/ppinject/actions/workflows/ci.yml/badge.svg)](https://github.com/KyleDerZweite/ppinject/actions/workflows/ci.yml)

A surgical XML injector for `.pptx` files.

`ppinject` is an API-first library for targeted PresentationML mutations inside the PPTX archive while preserving layout, relationships, and neighboring XML whenever feasible.

> Status: **Beta - narrow API-first MVP**

## Release note

`0.1.0b2` is the first public PyPI beta release.
It supports deterministic rendering of an existing PPTX template slide by replacing text placeholders and existing picture shapes.

The repository ships a synthetic example template at `examples/Template_Slide.pptx`.

## Who this is for

`ppinject` is useful when you already have a PowerPoint template with approved layout and formatting,
and your application only needs to fill known placeholders and swap existing image slots without rebuilding the presentation.

Typical workflow:

1. Prepare deterministic text values in your application code.
2. Render charts or other image assets separately.
3. Call `render_template_slide` with placeholder values and image replacements.
4. Open the output in PowerPoint with layout and neighboring XML preserved where feasible.

## Current supported scenario

The current release supports one production scenario:

- render one existing PPTX template slide to a new output file
- replace text placeholders that resolve within a single paragraph, even if PowerPoint split them across multiple `<a:t>` nodes
- replace existing picture shapes by shape name while keeping the original media part paths and extensions
- return a structured render report with replaced and missing keys

Public entry point:

```python
from ppinject import render_template_slide
```

## Quick example

```python
from ppinject import render_template_slide

report = render_template_slide(
	"examples/Template_Slide.pptx",
	"rendered_slide_part_a.pptx",
	text_replacements={
		"data1": "Demo Value 1",
		"data2": "Demo Value 2",
		"data3": "Demo Value 3",
		"data4": "Demo Value 4",
		"year-1": "2025",
		"year": "2026",
		"k-3": "Metric A",
		"k-2": "Metric B",
		"k-1": "Metric C",
		"b-v-3": "100",
		"b-v-2": "200",
		"b-v-1": "300",
		"c-v-3": "110",
		"c-v-2": "220",
		"c-v-1": "330",
		"v-p-3": "10",
		"v-p-2": "12",
		"v-p-1": "15",
		"headline": "Synthetic example output",
	},
	image_replacements={
		"Grafik 46": "example_chart.png",
		"Grafik 54": "example_chart.png",
	},
)

print(report)
```

## Current constraints

- the first supported scenario is deterministic rendering of an existing PPTX template slide
- placeholder replacement currently assumes the placeholder text belongs to one paragraph, even if PowerPoint split it across multiple `<a:t>` nodes
- image replacement currently targets existing picture shapes and keeps the original media part names and extensions
- the current public API is intentionally small: `render_template_slide` plus `RenderReport`
- broader slide authoring and general-purpose PresentationML mutation are still out of scope

It does not yet contain a broad general-purpose PPTX object model.

## Documentation

- See `docs/USAGE.md` for API details and a variance-report-style example.
- See `docs/USAGE.md` for API details and the bundled example template workflow.
- See `docs/ARCHITECTURE.md` for XML mutation boundaries and tradeoffs.
- See `docs/ROADMAP.md` for planned expansion beyond the initial template-rendering scenario.
- See `examples/render_template_slide.py` for a small example wired to the bundled template.

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

## Publish-readiness checklist

1. Keep the supported behavior statement narrow and honest.
2. Validate produced files in PowerPoint on Windows, not only via ZIP/XML inspection.
3. Use synthetic templates and assets in tests and examples.
4. Run lint, type check, tests, build, and `twine check` before every release.
5. Add release notes for every published version that describe supported and unsupported cases clearly.

## License

`ppinject` is open source under the GNU General Public License v3.0 or later (GPL-3.0-or-later).
See [LICENSE](LICENSE) for the full license text.
