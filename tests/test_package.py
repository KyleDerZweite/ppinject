from __future__ import annotations

import zipfile
from base64 import b64decode
from pathlib import Path

import pytest
from lxml import etree as ET

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
    assert __version__ == "0.1.0b2"


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


@pytest.mark.parametrize("output_kind", ["same", "symlink", "hardlink"])
def test_render_rejects_output_alias_of_template(tmp_path: Path, output_kind: str) -> None:
    template_path = tmp_path / "template.pptx"
    _build_fixture(template_path)
    original = template_path.read_bytes()
    output_path = template_path
    if output_kind == "symlink":
        output_path = tmp_path / "alias.pptx"
        try:
            output_path.symlink_to(template_path)
        except OSError as exc:
            if getattr(exc, "winerror", None) == 1314:
                pytest.skip("Windows user lacks symlink creation privilege")
            raise
    elif output_kind == "hardlink":
        output_path = tmp_path / "alias.pptx"
        output_path.hardlink_to(template_path)

    with pytest.raises(ValueError, match="different file"):
        render_template_slide(
            template_path, output_path,
            text_replacements={"report_title": "Updated"}, image_replacements={},
        )
    assert template_path.read_bytes() == original


def _rewrite_fixture(template_path: Path, changes: dict[str, bytes | None]) -> None:
    with zipfile.ZipFile(template_path) as archive:
        parts = {item.filename: archive.read(item) for item in archive.infolist()}
    for name, payload in changes.items():
        if payload is None:
            del parts[name]
        else:
            parts[name] = payload
    with zipfile.ZipFile(template_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


def test_render_preserves_neighboring_xml_and_unrelated_parts(tmp_path: Path) -> None:
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "nested" / "output.pptx"
    image_path = tmp_path / "replacement.png"
    image_path.write_bytes(PNG_BYTES + b"replacement")
    _build_fixture(template_path)
    with zipfile.ZipFile(template_path) as archive:
        slide = archive.read("ppt/slides/slide1.xml")
    slide = slide.replace(
        b"</p:spTree>",
        b'''<p:sp xmlns:custom="urn:ppinject:test" custom:flag="keep">
            <p:spPr><a:xfrm rot="123"><a:off x="12" y="34"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/>
            <a:t>Unchanged &amp; formatted</a:t></a:r></a:p></p:txBody>
            <custom:extension custom:value="unknown"><!--keep comment--></custom:extension>
        </p:sp></p:spTree>''',
    )
    _rewrite_fixture(template_path, {
        "ppt/slides/slide1.xml": slide,
        "ppt/slides/slide2.xml": slide,
        "customXml/item1.xml": b'<unknown flag="preserve">\n untouched \n</unknown>',
    })
    original_template = template_path.read_bytes()
    report = render_template_slide(
        template_path, output_path,
        text_replacements={"report_title": "A & <B>"},
        image_replacements={"Image Slot 1": image_path},
    )
    assert report.text_replacement_count == 1
    assert report.replaced_text_keys == ("report_title",)
    assert report.replaced_image_names == ("Image Slot 1",)
    assert template_path.read_bytes() == original_template
    with zipfile.ZipFile(template_path) as source, zipfile.ZipFile(output_path) as output:
        assert output.testzip() is None
        assert source.namelist() == output.namelist()
        changed = {
            name for name in source.namelist() if source.read(name) != output.read(name)
        }
        assert changed == {"ppt/slides/slide1.xml", "ppt/media/image13.png"}
        for item in source.infolist():
            written = output.getinfo(item.filename)
            assert (written.date_time, written.compress_type, written.external_attr) == (
                item.date_time, item.compress_type, item.external_attr,
            )
        before = ET.fromstring(source.read("ppt/slides/slide1.xml"))
        after = ET.fromstring(output.read("ppt/slides/slide1.xml"))
        # Remove only the targeted text difference, then compare the entire XML tree.
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        before_text = before.findall(".//a:t", ns)
        after_text = after.findall(".//a:t", ns)
        assert after_text[0].text == "A & <B> {site_code} {location}"
        after_text[0].text = before_text[0].text
        assert ET.tostring(after, method="c14n") == ET.tostring(before, method="c14n")
        assert output.read("ppt/media/image13.png") == image_path.read_bytes()


def test_render_split_runs_on_selected_slide(tmp_path: Path) -> None:
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "output.pptx"
    _build_fixture(template_path)
    with zipfile.ZipFile(template_path) as archive:
        slide = archive.read("ppt/slides/slide1.xml")
        rels = archive.read("ppt/slides/_rels/slide1.xml.rels")
    slide = slide.replace(
        b"<a:t>{report_title} {site_code} {location}</a:t>",
        b"<a:rPr b='1'/><a:t>{report_</a:t></a:r><a:r><a:rPr i='1'/>"
        b"<a:t>title} / {report_title} / {empty}</a:t>",
    )
    _rewrite_fixture(template_path, {
        "ppt/slides/slide2.xml": slide,
        "ppt/slides/_rels/slide2.xml.rels": rels,
    })
    report = render_template_slide(
        template_path, output_path, slide_number=2,
        text_replacements={"{report_title}": "Demo", "empty": None, "missing": "unused"},
        image_replacements={},
    )
    assert report.text_replacement_count == 1  # Counts changed paragraphs, not occurrences.
    assert report.replaced_text_keys == ("empty", "report_title")
    assert report.missing_text_keys == ("missing",)
    with zipfile.ZipFile(template_path) as source, zipfile.ZipFile(output_path) as output:
        assert source.read("ppt/slides/slide1.xml") == output.read("ppt/slides/slide1.xml")
        root = ET.fromstring(output.read("ppt/slides/slide2.xml"))
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        paragraph = root.findall(".//a:p", ns)[0]
        assert "".join(node.text or "" for node in paragraph.findall(".//a:t", ns)) == "Demo / Demo / "
        assert len(paragraph.findall("a:r/a:rPr", ns)) == 2


@pytest.mark.parametrize("existing_output", [False, True])
@pytest.mark.parametrize("failure", [
    "invalid_slide", "missing_slide", "missing_rels", "malformed_xml", "bad_zip",
    "missing_image", "wrong_extension",
])
def test_render_input_failures_leave_output_untouched(
    tmp_path: Path, failure: str, existing_output: bool,
) -> None:
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "output" / "rendered.pptx"
    image_path = tmp_path / "replacement.png"
    _build_fixture(template_path)
    image_path.write_bytes(PNG_BYTES)
    slide_number = 1
    error: type[Exception]
    match: str
    if failure == "invalid_slide":
        slide_number, error, match = 0, ValueError, "positive"
    elif failure == "missing_slide":
        slide_number, error, match = 99, KeyError, "slide99.xml"
    elif failure == "missing_rels":
        _rewrite_fixture(template_path, {"ppt/slides/_rels/slide1.xml.rels": None})
        error, match = KeyError, "slide1.xml.rels"
    elif failure == "malformed_xml":
        _rewrite_fixture(template_path, {"ppt/slides/slide1.xml": b"<broken>"})
        error, match = ET.XMLSyntaxError, "Premature end"
    elif failure == "bad_zip":
        template_path.write_bytes(b"not a zip archive")
        error, match = zipfile.BadZipFile, "not a zip file"
    elif failure == "missing_image":
        image_path = tmp_path / "absent.png"
        error, match = FileNotFoundError, "does not exist"
    else:
        image_path = tmp_path / "replacement.jpg"
        image_path.write_bytes(b"wrong format")
        error, match = ValueError, "must use '.png'"
    original_template = template_path.read_bytes()
    if existing_output:
        output_path.parent.mkdir()
        output_path.write_bytes(b"existing output")
    with pytest.raises(error, match=match):
        render_template_slide(
            template_path, output_path, slide_number=slide_number,
            text_replacements={"report_title": "Updated"},
            image_replacements={"Image Slot 1": image_path},
        )
    assert template_path.read_bytes() == original_template
    if existing_output:
        assert output_path.read_bytes() == b"existing output"
    else:
        assert not output_path.parent.exists()


def test_bundled_template_archive_round_trip(tmp_path: Path) -> None:
    examples = Path(__file__).resolve().parents[1] / "examples"
    template_path = examples / "Template_Slide.pptx"
    image_path = examples / "example_chart.png"
    output_path = tmp_path / "bundled-output.pptx"
    original = template_path.read_bytes()
    report = render_template_slide(
        template_path, output_path,
        text_replacements={"data1": "A & B", "headline": "Synthetic regression output"},
        image_replacements={"Grafik 46": image_path, "Grafik 54": image_path},
    )
    assert report.replaced_text_keys == ("data1", "headline")
    assert report.missing_text_keys == report.missing_image_names == ()
    assert report.replaced_image_names == ("Grafik 46", "Grafik 54")
    assert template_path.read_bytes() == original
    with zipfile.ZipFile(template_path) as source, zipfile.ZipFile(output_path) as output:
        assert output.testzip() is None
        assert source.namelist() == output.namelist()
        changed = {name for name in source.namelist() if source.read(name) != output.read(name)}
        assert "ppt/slides/slide1.xml" in changed
        assert all(name == "ppt/slides/slide1.xml" or name.startswith("ppt/media/")
                   for name in changed)
        for name in output.namelist():
            if name.endswith((".xml", ".rels")):
                ET.fromstring(output.read(name))
        slide = ET.fromstring(output.read("ppt/slides/slide1.xml"))
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        text = "".join(node.text or "" for node in slide.findall(".//a:t", ns))
        assert "A & B" in text
        assert "Synthetic regression output" in text
