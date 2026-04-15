from __future__ import annotations

from base64 import b64decode
from pathlib import Path
import zipfile

from ppinject import __version__, render_template_slide

PNG_BYTES = b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9p6m4m8AAAAASUVORK5CYII="
)


def _build_fixture(template_path: Path) -> None:
    slide_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr><p:cNvPr id=\"1\" name=\"Title\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p><a:r><a:t>{report_title} {site_code} {location}</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
            <p:graphicFrame>
                <p:nvGraphicFramePr><p:cNvPr id=\"2\" name=\"Table 1\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
                <a:graphic>
                    <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">
                        <a:tbl>
                            <a:tr h="0"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>{address_line}</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                            <a:tr h="0"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>{reference_value} kWh</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                        </a:tbl>
                    </a:graphicData>
                </a:graphic>
            </p:graphicFrame>
            <p:pic>
                <p:nvPicPr><p:cNvPr id="3" name="Image Slot 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                <p:blipFill><a:blip r:embed=\"rId2\"/></p:blipFill>
            </p:pic>
            <p:pic>
                <p:nvPicPr><p:cNvPr id="4" name="Image Slot 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                <p:blipFill><a:blip r:embed=\"rId3\"/></p:blipFill>
            </p:pic>
        </p:spTree>
    </p:cSld>
</p:sld>
"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image13.png\"/>
    <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image14.png\"/>
</Relationships>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
    <Default Extension=\"xml\" ContentType=\"application/xml\"/>
    <Default Extension=\"png\" ContentType=\"image/png\"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>
"""

    with zipfile.ZipFile(template_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("ppt/slides/slide1.xml", slide_xml)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", rels_xml)
        archive.writestr("ppt/media/image13.png", PNG_BYTES)
        archive.writestr("ppt/media/image14.png", PNG_BYTES)


def test_version_is_defined() -> None:
    assert __version__ == "0.1.0b1"


def test_render_template_slide_replaces_text_and_images(tmp_path: Path) -> None:
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "rendered.pptx"
    image_a = tmp_path / "chart_a.png"
    image_b = tmp_path / "chart_b.png"

    _build_fixture(template_path)
    image_a.write_bytes(PNG_BYTES + b"A")
    image_b.write_bytes(PNG_BYTES + b"B")

    report = render_template_slide(
        template_path,
        output_path,
        text_replacements={
            "report_title": "Template Report",
            "site_code": "SITE-001",
            "location": "Demo City",
            "address_line": "100 Example Street",
            "reference_value": "8210",
        },
        image_replacements={
            "Image Slot 1": image_a,
            "Image Slot 2": image_b,
        },
    )

    assert report.output_file == output_path
    assert report.missing_text_keys == ()
    assert report.missing_image_names == ()
    assert report.replaced_image_names == ("Image Slot 1", "Image Slot 2")

    with zipfile.ZipFile(output_path, "r") as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
        assert "Template Report SITE-001 Demo City" in slide_xml
        assert "100 Example Street" in slide_xml
        assert "8210 kWh" in slide_xml
        assert archive.read("ppt/media/image13.png") == image_a.read_bytes()
        assert archive.read("ppt/media/image14.png") == image_b.read_bytes()


def test_render_template_slide_reports_missing_targets(tmp_path: Path) -> None:
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "rendered_missing.pptx"
    image_a = tmp_path / "chart_a.png"

    _build_fixture(template_path)
    image_a.write_bytes(PNG_BYTES)

    report = render_template_slide(
        template_path,
        output_path,
        text_replacements={
            "missing-token": "value",
            "address_line": "Example Address",
        },
        image_replacements={
            "Missing Picture": image_a,
        },
    )

    assert report.replaced_text_keys == ("address_line",)
    assert report.missing_text_keys == ("missing-token",)
    assert report.replaced_image_names == ()
    assert report.missing_image_names == ("Missing Picture",)
