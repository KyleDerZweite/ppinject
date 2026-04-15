"""High-level API for targeted single-slide PPTX template rendering."""

from __future__ import annotations

from pathlib import Path
from typing import Mapping

from ppinject.injector import RenderReport, render_slide_package


def render_template_slide(
    template_path: str | Path,
    output_path: str | Path,
    *,
    text_replacements: Mapping[str, object],
    image_replacements: Mapping[str, str | Path],
    slide_number: int = 1,
) -> RenderReport:
    """Render one PPTX template slide by replacing text placeholders and pictures.

    The current MVP is intentionally narrow:
    - placeholders are expected to live fully inside one paragraph
    - image replacements target existing picture shapes by name
    - slide targeting defaults to the first slide
    """
    return render_slide_package(
        template_path=Path(template_path),
        output_path=Path(output_path),
        text_replacements=text_replacements,
        image_replacements={name: Path(path) for name, path in image_replacements.items()},
        slide_number=slide_number,
    )
