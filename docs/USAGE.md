# Usage

Primary interface is the Python API: `render_template_slide`.
For architecture and design constraints, refer to `ARCHITECTURE.md`.

The repository ships a synthetic example template at `examples/Template_Slide.pptx`.

## Install / prepare

For local development:

```bash
uv sync --dev
```

For package consumers after publication:

```bash
pip install ppinject
```

## Python API

```python
from ppinject import render_template_slide

report = render_template_slide(
	"examples/Template_Slide.pptx",
	"output.pptx",
	text_replacements={
		"data1": "Demo Value 1",
		"data2": "Demo Value 2",
		"data3": "Demo Value 3",
		"data4": "Demo Value 4",
		"year-1": "2025",
		"year": "2026",
		"headline": "Synthetic example output",
	},
	image_replacements={
		"Grafik 46": "example_chart.png",
		"Grafik 54": "example_chart.png",
	},
)

print(report)
```

Notes:

- `render_template_slide` is the recommended entry point for the current release.
- `text_replacements` accepts keys with or without surrounding braces.
- `image_replacements` maps existing picture shape names to replacement image files.
- `slide_number` defaults to `1` and can be overridden when the template contains multiple slides.
- In the bundled example template, the two picture shapes are currently named `Grafik 46` and `Grafik 54`.

## Report object

`render_template_slide` returns a `RenderReport`:

```python
RenderReport(
	output_file=Path("output.pptx"),
	text_replacement_count=41,
	replaced_text_keys=("data1", "headline", "year"),
	missing_text_keys=(),
	replaced_image_names=("Grafik 46", "Grafik 54"),
	missing_image_names=(),
)
```

This is intended to make application-level validation simple:

- fail fast if `missing_text_keys` is not empty
- fail fast if `missing_image_names` is not empty
- log `text_replacement_count` for template regression checks

## Bundled example template

```python
from pathlib import Path
from ppinject import render_template_slide

examples_dir = Path("examples")

render_template_slide(
	examples_dir / "Template_Slide.pptx",
	examples_dir / "rendered_example_output.pptx",
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
		"Grafik 46": examples_dir / "example_chart.png",
		"Grafik 54": examples_dir / "example_chart.png",
	},
)
```

Note:

- The bundled example template is synthetic and safe to publish.
- The two picture shapes in that template currently resolve to the same placeholder artwork, so the example uses the same replacement image for both.
- For demonstrations with two independent images, use your own template with separate picture parts.

## Production checklist

1. Keep the source template immutable and always render to a new output path.
2. Validate that `missing_text_keys` and `missing_image_names` are empty.
3. Keep placeholder design stable inside the PowerPoint template.
4. Validate the output in PowerPoint on Windows.
5. Use synthetic sample templates for bug reports and pull requests.

## Current limitations

1. Placeholder replacement is paragraph-based, not arbitrary multi-paragraph flow-based.
2. Image replacement assumes the target shape already exists in the template.
3. Media replacement keeps the original target file extension, so replacement files must match it.
4. The current public API is focused on deterministic template rendering, not presentation authoring.
