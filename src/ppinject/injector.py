from __future__ import annotations

from dataclasses import dataclass
import posixpath
from pathlib import Path, PurePosixPath
from typing import Mapping
import zipfile

from lxml import etree as ET

SLIDE_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass(frozen=True)
class RenderReport:
    output_file: Path
    text_replacement_count: int
    replaced_text_keys: tuple[str, ...]
    missing_text_keys: tuple[str, ...]
    replaced_image_names: tuple[str, ...]
    missing_image_names: tuple[str, ...]


def _parse_xml(xml_bytes: bytes) -> ET._Element:
    parser = ET.XMLParser(remove_blank_text=False, recover=False)
    return ET.fromstring(xml_bytes, parser=parser)


def _normalize_placeholder_map(text_replacements: Mapping[str, object]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for key, value in text_replacements.items():
        token = str(key)
        if not token.startswith("{"):
            token = "{" + token + "}"
        normalized[token] = "" if value is None else str(value)
    return normalized


def _replace_text_nodes(slide_root: ET._Element, placeholder_map: Mapping[str, str]) -> tuple[int, set[str]]:
    replacement_count = 0
    replaced_keys: set[str] = set()

    for paragraph in slide_root.findall(".//a:p", SLIDE_NS):
        text_nodes = paragraph.findall(".//a:t", SLIDE_NS)
        if not text_nodes:
            continue

        current_text = "".join(text_node.text or "" for text_node in text_nodes)
        updated_text = current_text
        for placeholder, replacement in placeholder_map.items():
            if placeholder not in updated_text:
                continue
            updated_text = updated_text.replace(placeholder, replacement)
            replaced_keys.add(placeholder[1:-1])

        if updated_text != current_text:
            text_nodes[0].text = updated_text
            for text_node in text_nodes[1:]:
                text_node.text = ""
            replacement_count += 1

    return replacement_count, replaced_keys


def _resolve_media_targets(slide_root: ET._Element, rels_root: ET._Element) -> dict[str, str]:
    relationships = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("pr:Relationship", SLIDE_NS)
        if "Id" in rel.attrib and "Target" in rel.attrib
    }
    media_targets: dict[str, str] = {}

    for picture in slide_root.findall(".//p:pic", SLIDE_NS):
        c_nv_pr = picture.find("p:nvPicPr/p:cNvPr", SLIDE_NS)
        blip = picture.find(".//a:blip", SLIDE_NS)
        if c_nv_pr is None or blip is None:
            continue

        name = c_nv_pr.attrib.get("name", "")
        rel_id = blip.attrib.get(f"{{{REL_NS}}}embed")
        if not name or not rel_id:
            continue

        target = relationships.get(rel_id)
        if not target:
            continue

        resolved = posixpath.normpath(str(PurePosixPath("ppt/slides") / PurePosixPath(str(target))))
        media_targets[name] = resolved

    return media_targets


def _replace_media_parts(
    media_targets: Mapping[str, str],
    image_replacements: Mapping[str, Path],
) -> tuple[dict[str, bytes], set[str]]:
    updated_parts: dict[str, bytes] = {}
    replaced_images: set[str] = set()

    for picture_name, image_path in image_replacements.items():
        media_part = media_targets.get(picture_name)
        if media_part is None:
            continue
        if not image_path.exists():
            raise FileNotFoundError(f"Image replacement path does not exist: {image_path}")

        target_suffix = PurePosixPath(media_part).suffix.lower()
        source_suffix = image_path.suffix.lower()
        if target_suffix and source_suffix and target_suffix != source_suffix:
            raise ValueError(
                f"Image replacement for '{picture_name}' must use '{target_suffix}' but got '{source_suffix}'."
            )

        updated_parts[media_part] = image_path.read_bytes()
        replaced_images.add(picture_name)

    return updated_parts, replaced_images


def render_slide_package(
    *,
    template_path: Path,
    output_path: Path,
    text_replacements: Mapping[str, object],
    image_replacements: Mapping[str, Path],
    slide_number: int,
) -> RenderReport:
    if slide_number <= 0:
        raise ValueError("slide_number must be positive")

    slide_part = f"ppt/slides/slide{slide_number}.xml"
    rels_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    placeholder_map = _normalize_placeholder_map(text_replacements)

    with zipfile.ZipFile(template_path, "r") as source_zip:
        slide_root = _parse_xml(source_zip.read(slide_part))
        rels_root = _parse_xml(source_zip.read(rels_part))

        text_replacement_count, replaced_text_keys = _replace_text_nodes(slide_root, placeholder_map)
        media_targets = _resolve_media_targets(slide_root, rels_root)
        updated_media_parts, replaced_images = _replace_media_parts(
            media_targets=media_targets,
            image_replacements=image_replacements,
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        serialized_slide = ET.tostring(
            slide_root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as target_zip:
            for item in source_zip.infolist():
                payload = source_zip.read(item.filename)
                if item.filename == slide_part:
                    payload = serialized_slide
                elif item.filename in updated_media_parts:
                    payload = updated_media_parts[item.filename]
                target_zip.writestr(item, payload)

    requested_text_keys = {key.strip("{}") for key in placeholder_map}
    requested_image_names = set(image_replacements)
    return RenderReport(
        output_file=output_path,
        text_replacement_count=text_replacement_count,
        replaced_text_keys=tuple(sorted(replaced_text_keys)),
        missing_text_keys=tuple(sorted(requested_text_keys - replaced_text_keys)),
        replaced_image_names=tuple(sorted(replaced_images)),
        missing_image_names=tuple(sorted(requested_image_names - replaced_images)),
    )
