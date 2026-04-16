from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom

from .models import StructureOperationType, StructurePlan

SMART_QUOTE_REPLACEMENTS = {
    "\u201c": "&#x201C;",
    "\u201d": "&#x201D;",
    "\u2018": "&#x2018;",
    "\u2019": "&#x2019;",
}


def _pretty_print_xml(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except Exception:
        return


def _escape_smart_quotes(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_REPLACEMENTS.items():
            content = content.replace(char, entity)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        return


def unpack_pptx(input_pptx: Path, output_dir: Path) -> None:
    if not input_pptx.exists():
        raise FileNotFoundError(f"PPTX introuvable: {input_pptx}")

    output_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(input_pptx, "r") as archive:
        archive.extractall(output_dir)

    for path in list(output_dir.rglob("*.xml")) + list(output_dir.rglob("*.rels")):
        _pretty_print_xml(path)
        _escape_smart_quotes(path)


def _load_dom(path: Path) -> defusedxml.minidom.Document:
    return defusedxml.minidom.parse(str(path))


def _save_dom(path: Path, dom: defusedxml.minidom.Document) -> None:
    path.write_bytes(dom.toxml(encoding="utf-8"))


def _presentation_paths(unpacked_dir: Path) -> tuple[Path, Path]:
    pres = unpacked_dir / "ppt" / "presentation.xml"
    rels = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"
    if not pres.exists() or not rels.exists():
        raise FileNotFoundError("presentation.xml ou presentation.xml.rels manquant")
    return pres, rels


def _slide_id_nodes(unpacked_dir: Path):
    pres_path, _ = _presentation_paths(unpacked_dir)
    dom = _load_dom(pres_path)
    sld_id_lst_nodes = dom.getElementsByTagName("p:sldIdLst")
    if not sld_id_lst_nodes:
        raise RuntimeError("<p:sldIdLst> introuvable")
    sld_id_lst = sld_id_lst_nodes[0]
    nodes = [
        n
        for n in sld_id_lst.childNodes
        if n.nodeType == n.ELEMENT_NODE and n.tagName == "p:sldId"
    ]
    return dom, sld_id_lst, nodes


def _rid_to_slide_target(unpacked_dir: Path) -> dict[str, str]:
    _, rels_path = _presentation_paths(unpacked_dir)
    rels_dom = _load_dom(rels_path)
    mapping: dict[str, str] = {}
    for rel in rels_dom.getElementsByTagName("Relationship"):
        rel_type = rel.getAttribute("Type")
        target = rel.getAttribute("Target")
        if rel_type.endswith("/slide") and target.startswith("slides/"):
            mapping[rel.getAttribute("Id")] = target
    return mapping


def list_slide_sequence(unpacked_dir: Path) -> list[str]:
    _, _, slide_nodes = _slide_id_nodes(unpacked_dir)
    mapping = _rid_to_slide_target(unpacked_dir)
    out: list[str] = []
    for node in slide_nodes:
        rid = node.getAttribute("r:id")
        target = mapping.get(rid)
        if target:
            out.append(Path(target).name)
    return out


def _next_slide_number(slides_dir: Path) -> int:
    numbers = []
    for path in slides_dir.glob("slide*.xml"):
        match = re.match(r"slide(\d+)\.xml", path.name)
        if match:
            numbers.append(int(match.group(1)))
    return max(numbers) + 1 if numbers else 1


def _next_rid(rels_path: Path) -> str:
    content = rels_path.read_text(encoding="utf-8")
    ids = [int(match) for match in re.findall(r'Id="rId(\d+)"', content)]
    return f"rId{max(ids) + 1 if ids else 1}"


def _next_slide_id(unpacked_dir: Path) -> int:
    dom, _, slide_nodes = _slide_id_nodes(unpacked_dir)
    ids = []
    for node in slide_nodes:
        value = node.getAttribute("id")
        if value.isdigit():
            ids.append(int(value))
    del dom
    return max(ids) + 1 if ids else 256


def _add_slide_override_content_type(unpacked_dir: Path, slide_name: str) -> None:
    content_types_path = unpacked_dir / "[Content_Types].xml"
    dom = _load_dom(content_types_path)
    root = dom.documentElement
    part_name = f"/ppt/slides/{slide_name}"

    for override in root.getElementsByTagName("Override"):
        if override.getAttribute("PartName") == part_name:
            _save_dom(content_types_path, dom)
            return

    override = dom.createElement("Override")
    override.setAttribute("PartName", part_name)
    override.setAttribute(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    )
    root.appendChild(override)
    _save_dom(content_types_path, dom)


def _add_presentation_relationship(unpacked_dir: Path, slide_name: str) -> str:
    _, rels_path = _presentation_paths(unpacked_dir)
    dom = _load_dom(rels_path)
    root = dom.documentElement
    rid = _next_rid(rels_path)

    rel = dom.createElement("Relationship")
    rel.setAttribute("Id", rid)
    rel.setAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    )
    rel.setAttribute("Target", f"slides/{slide_name}")
    root.appendChild(rel)
    _save_dom(rels_path, dom)
    return rid


def _insert_slide_id(
    unpacked_dir: Path, rid: str, target_index: int | None = None
) -> None:
    dom, sld_id_lst, slide_nodes = _slide_id_nodes(unpacked_dir)
    node = dom.createElement("p:sldId")
    node.setAttribute("id", str(_next_slide_id(unpacked_dir)))
    node.setAttribute("r:id", rid)

    if target_index is None:
        sld_id_lst.appendChild(node)
    else:
        insert_at = max(1, target_index)
        if insert_at > len(slide_nodes):
            sld_id_lst.appendChild(node)
        else:
            sld_id_lst.insertBefore(node, slide_nodes[insert_at - 1])

    pres_path, _ = _presentation_paths(unpacked_dir)
    _save_dom(pres_path, dom)


def delete_slide(unpacked_dir: Path, slide_index: int) -> str:
    dom, sld_id_lst, slide_nodes = _slide_id_nodes(unpacked_dir)
    if slide_index < 1 or slide_index > len(slide_nodes):
        raise IndexError(f"slide_index invalide: {slide_index}")
    node = slide_nodes[slide_index - 1]
    rid = node.getAttribute("r:id")
    sld_id_lst.removeChild(node)

    pres_path, _ = _presentation_paths(unpacked_dir)
    _save_dom(pres_path, dom)
    return rid


def reorder_slides(unpacked_dir: Path, new_order: list[int]) -> None:
    dom, sld_id_lst, slide_nodes = _slide_id_nodes(unpacked_dir)
    count = len(slide_nodes)
    expected = list(range(1, count + 1))
    if sorted(new_order) != expected:
        raise ValueError(f"new_order doit etre une permutation de {expected}")

    for node in list(slide_nodes):
        sld_id_lst.removeChild(node)

    for idx in new_order:
        sld_id_lst.appendChild(slide_nodes[idx - 1])

    pres_path, _ = _presentation_paths(unpacked_dir)
    _save_dom(pres_path, dom)


def duplicate_slide(
    unpacked_dir: Path, slide_index: int, target_index: int | None = None
) -> str:
    sequence = list_slide_sequence(unpacked_dir)
    if slide_index < 1 or slide_index > len(sequence):
        raise IndexError(f"slide_index invalide: {slide_index}")

    source_name = sequence[slide_index - 1]
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    rels_dir.mkdir(parents=True, exist_ok=True)

    next_number = _next_slide_number(slides_dir)
    dest_name = f"slide{next_number}.xml"

    shutil.copy2(slides_dir / source_name, slides_dir / dest_name)

    source_rels = rels_dir / f"{source_name}.rels"
    dest_rels = rels_dir / f"{dest_name}.rels"
    if source_rels.exists():
        shutil.copy2(source_rels, dest_rels)
        rels_dom = _load_dom(dest_rels)
        changed = False
        for rel in list(rels_dom.getElementsByTagName("Relationship")):
            rel_type = rel.getAttribute("Type")
            if rel_type.endswith("/notesSlide"):
                rel.parentNode.removeChild(rel)
                changed = True
        if changed:
            _save_dom(dest_rels, rels_dom)

    _add_slide_override_content_type(unpacked_dir, dest_name)
    rid = _add_presentation_relationship(unpacked_dir, dest_name)

    insertion_position = target_index
    if insertion_position is None:
        insertion_position = slide_index + 1
    _insert_slide_id(unpacked_dir, rid, target_index=insertion_position)
    return dest_name


def add_slide_from_layout(
    unpacked_dir: Path, layout_index: int, target_index: int | None = None
) -> str:
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    rels_dir.mkdir(parents=True, exist_ok=True)

    layout_file = f"slideLayout{layout_index}.xml"
    layout_path = unpacked_dir / "ppt" / "slideLayouts" / layout_file
    if not layout_path.exists():
        raise FileNotFoundError(f"layout introuvable: {layout_file}")

    next_number = _next_slide_number(slides_dir)
    slide_name = f"slide{next_number}.xml"
    slide_xml = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>
"""

    (slides_dir / slide_name).write_text(slide_xml, encoding="utf-8")

    rels_xml = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/{layout_file}\"/>
</Relationships>
"""
    (rels_dir / f"{slide_name}.rels").write_text(rels_xml, encoding="utf-8")

    _add_slide_override_content_type(unpacked_dir, slide_name)
    rid = _add_presentation_relationship(unpacked_dir, slide_name)
    _insert_slide_id(unpacked_dir, rid, target_index=target_index)
    return slide_name


def _all_referenced_paths(unpacked_dir: Path) -> set[Path]:
    referenced: set[Path] = set()
    root = unpacked_dir.resolve()
    for rels_file in unpacked_dir.rglob("*.rels"):
        dom = _load_dom(rels_file)
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if (
                not target
                or target.startswith("http://")
                or target.startswith("https://")
            ):
                continue
            if target.startswith("/"):
                candidate = root / target.lstrip("/")
            else:
                candidate = (rels_file.parent.parent / target).resolve()
            try:
                referenced.add(candidate.relative_to(root))
            except ValueError:
                continue
    return referenced


def _cleanup_orphaned_slides(unpacked_dir: Path) -> list[Path]:
    removed: list[Path] = []
    sequence = set(list_slide_sequence(unpacked_dir))
    slides_dir = unpacked_dir / "ppt" / "slides"
    slides_rels = slides_dir / "_rels"

    if not slides_dir.exists():
        return removed

    for slide_file in slides_dir.glob("slide*.xml"):
        if slide_file.name in sequence:
            continue
        slide_file.unlink(missing_ok=True)
        removed.append(slide_file.relative_to(unpacked_dir))
        rels_file = slides_rels / f"{slide_file.name}.rels"
        if rels_file.exists():
            rels_file.unlink()
            removed.append(rels_file.relative_to(unpacked_dir))

    _, pres_rels_path = _presentation_paths(unpacked_dir)
    rels_dom = _load_dom(pres_rels_path)
    rels_root = rels_dom.documentElement
    changed = False
    for rel in list(rels_root.getElementsByTagName("Relationship")):
        rel_type = rel.getAttribute("Type")
        if not rel_type.endswith("/slide"):
            continue
        target_name = Path(rel.getAttribute("Target")).name
        if target_name in sequence:
            continue
        rels_root.removeChild(rel)
        changed = True

    if changed:
        _save_dom(pres_rels_path, rels_dom)

    return removed


def _cleanup_unreferenced_resources(
    unpacked_dir: Path, referenced: set[Path]
) -> list[Path]:
    removed: list[Path] = []
    resource_dirs = [
        "media",
        "embeddings",
        "charts",
        "diagrams",
        "drawings",
        "ink",
        "tags",
        "notesSlides",
    ]
    for dir_name in resource_dirs:
        resource_dir = unpacked_dir / "ppt" / dir_name
        if not resource_dir.exists():
            continue
        for file_path in resource_dir.glob("*"):
            if not file_path.is_file():
                continue
            rel = file_path.relative_to(unpacked_dir)
            if rel in referenced:
                continue
            file_path.unlink(missing_ok=True)
            removed.append(rel)

    theme_dir = unpacked_dir / "ppt" / "theme"
    if theme_dir.exists():
        for theme_file in theme_dir.glob("theme*.xml"):
            rel = theme_file.relative_to(unpacked_dir)
            if rel in referenced:
                continue
            theme_file.unlink(missing_ok=True)
            removed.append(rel)
            rels_file = theme_dir / "_rels" / f"{theme_file.name}.rels"
            if rels_file.exists():
                rels_file.unlink()
                removed.append(rels_file.relative_to(unpacked_dir))

    return removed


def _update_content_types(unpacked_dir: Path, removed: list[Path]) -> None:
    content_types_path = unpacked_dir / "[Content_Types].xml"
    if not content_types_path.exists() or not removed:
        return
    removed_part_names = {"/" + str(path).replace("\\", "/") for path in removed}

    dom = _load_dom(content_types_path)
    root = dom.documentElement
    changed = False
    for override in list(root.getElementsByTagName("Override")):
        part_name = override.getAttribute("PartName")
        if part_name in removed_part_names:
            root.removeChild(override)
            changed = True
    if changed:
        _save_dom(content_types_path, dom)


def clean_unreferenced_files(unpacked_dir: Path) -> list[str]:
    removed: list[Path] = []

    removed.extend(_cleanup_orphaned_slides(unpacked_dir))
    referenced = _all_referenced_paths(unpacked_dir)
    removed.extend(_cleanup_unreferenced_resources(unpacked_dir, referenced))
    _update_content_types(unpacked_dir, removed)

    return [str(path).replace("\\", "/") for path in removed]


def apply_structure_plan(unpacked_dir: Path, plan: StructurePlan) -> list[str]:
    logs: list[str] = []
    for step, operation in enumerate(plan.operations, start=1):
        if operation.op == StructureOperationType.delete_slide:
            delete_slide(unpacked_dir, operation.slide_index or 1)
            logs.append(f"[{step}] delete_slide {operation.slide_index}")

        elif operation.op == StructureOperationType.duplicate_slide:
            created = duplicate_slide(
                unpacked_dir,
                slide_index=operation.slide_index or 1,
                target_index=operation.target_index,
            )
            logs.append(
                f"[{step}] duplicate_slide {operation.slide_index} -> {created} @ {operation.target_index or 'auto'}"
            )

        elif operation.op == StructureOperationType.add_layout_slide:
            created = add_slide_from_layout(
                unpacked_dir,
                layout_index=operation.layout_index or 1,
                target_index=operation.target_index,
            )
            logs.append(
                f"[{step}] add_layout_slide layout={operation.layout_index} -> {created} @ {operation.target_index or 'end'}"
            )

        elif operation.op == StructureOperationType.reorder_slides:
            reorder_slides(unpacked_dir, operation.new_order)
            logs.append(f"[{step}] reorder_slides -> {operation.new_order}")

    removed = clean_unreferenced_files(unpacked_dir)
    if removed:
        logs.append(f"clean: removed {len(removed)} orphan files")
    return logs


def _condense_xml(xml_file: Path) -> None:
    dom = _load_dom(xml_file)
    for element in dom.getElementsByTagName("*"):
        if element.tagName.endswith(":t"):
            continue
        for child in list(element.childNodes):
            if (
                child.nodeType == child.TEXT_NODE
                and child.nodeValue
                and child.nodeValue.strip() == ""
            ):
                element.removeChild(child)
    xml_file.write_bytes(dom.toxml(encoding="utf-8"))


def pack_pptx(input_directory: Path, output_file: Path) -> None:
    if not input_directory.is_dir():
        raise NotADirectoryError(f"Dossier introuvable: {input_directory}")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir) / "content"
        shutil.copytree(input_directory, temp_root)

        for pattern in ("*.xml", "*.rels"):
            for xml_file in temp_root.rglob(pattern):
                _condense_xml(xml_file)

        with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as archive:
            for file_path in temp_root.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_root))
