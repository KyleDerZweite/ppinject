from __future__ import annotations

from pathlib import Path

from ppinject import render_template_slide


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    template_path = base_dir / "Template_Slide.pptx"
    output_path = base_dir / "rendered_output.pptx"
    example_chart = base_dir / "example_chart.png"

    report = render_template_slide(
        template_path,
        output_path,
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
            "Grafik 46": example_chart,
            "Grafik 54": example_chart,
        },
    )

    print(report)


if __name__ == "__main__":
    main()