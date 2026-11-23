"""Bounded OOXML content audit. No rendering or Word layout interpretation."""
from __future__ import annotations

import hashlib
import json
import posixpath
from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


class AuditInputError(Exception):
    pass


def q(namespace: str, name: str) -> str:
    return "{" + namespace + "}" + name


def digest(value: object) -> str:
    data = json.dumps(value, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


@dataclass
class Snapshot:
    records: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    derived_fields: list = field(default_factory=list)
    parts: list[str] = field(default_factory=list)
    counts: dict = field(default_factory=lambda: {
        "paragraphs": 0, "tables": 0, "images": 0, "hyperlinks": 0,
        "equations": 0, "fields": 0,
    })


def canonical_math(node: ET.Element):
    """Compare OMML structure while ignoring Word font-run properties."""
    if node.tag == q(W, "rPr"):
        return None
    if node.tag == q(M, "r") and all(child.tag in {q(M, "t"), q(W, "rPr")} for child in node):
        return ["math_text", "".join(child.text or "" for child in node if child.tag == q(M, "t"))]
    children = []
    for child in node:
        item = canonical_math(child)
        if item is None:
            continue
        if item[0] == "math_text" and children and children[-1][0] == "math_text":
            children[-1][1] += item[1]
        else:
            children.append(item)
    text = node.text or ""
    if node.tag != q(M, "t") and not text.strip():
        text = ""
    return [node.tag, sorted(node.attrib.items()), text, children]


class PackageAudit:
    def __init__(self, path: Path, scope: str):
        self.path, self.scope = path, scope
        self.snapshot = Snapshot()
        self.xml_cache = {}
        self.rels_cache = {}
        self.note_maps = {}
        self.note_used = {"footnote": set(), "endnote": set(), "comment": set()}
        self.note_active = set()
        self.header_used = set()
        self.warning_keys = set()
        try:
            self.archive = ZipFile(path)
            names = self.archive.namelist()
            if len(names) != len(set(names)):
                raise AuditInputError("DOCX ZIP contains duplicate member names")
            self.names = set(names)
            if "[Content_Types].xml" not in self.names:
                raise AuditInputError("ZIP is not a DOCX package: missing [Content_Types].xml")
            self.xml("[Content_Types].xml")
            main = [v for v in self.relationships("").values() if v["type"].endswith("/officeDocument")]
            if len(main) != 1 or main[0]["external"]:
                raise AuditInputError("DOCX must have one internal officeDocument relationship")
            self.main_part = main[0]["target"]
            self.main = self.xml(self.main_part)
            if self.main.tag != q(W, "document"):
                raise AuditInputError("Unsupported DOCX main namespace or document root")
            self.load_notes()
            self.styles = self.load_styles()
        except (OSError, BadZipFile, RuntimeError, ET.ParseError, KeyError) as exc:
            raise AuditInputError(str(exc)) from exc

    def close(self):
        self.archive.close()

    def warn(self, code: str, part: str, detail: str):
        key = (code, part, detail)
        if key not in self.warning_keys:
            self.warning_keys.add(key)
            self.snapshot.warnings.append({"code": code, "part": part, "detail": detail})

    def xml(self, name: str) -> ET.Element:
        if name not in self.xml_cache:
            try:
                raw = self.archive.read(name)
                if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
                    raise AuditInputError("DTD/entity declarations are unsupported: " + name)
                self.xml_cache[name] = ET.fromstring(raw)
            except (KeyError, ET.ParseError, BadZipFile, RuntimeError) as exc:
                raise AuditInputError("Cannot parse DOCX part " + name + ": " + str(exc)) from exc
        return self.xml_cache[name]

    def relationships(self, part: str) -> dict:
        if part in self.rels_cache:
            return self.rels_cache[part]
        relname = (posixpath.dirname(part) + "/_rels/" + posixpath.basename(part) + ".rels").lstrip("/") if part else "_rels/.rels"
        result = {}
        if relname in self.names:
            for element in self.xml(relname):
                rid = element.get("Id")
                if not rid or rid in result:
                    raise AuditInputError("Missing or duplicate relationship Id in " + relname)
                target = element.get("Target", "")
                external = element.get("TargetMode") == "External"
                if not external:
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith("/") else target.lstrip("/")
                    if target.startswith("../") or target == "..":
                        raise AuditInputError("Relationship escapes package: " + relname)
                result[rid] = {"type": element.get("Type", ""), "target": target, "external": external}
        self.rels_cache[part] = result
        return result

    def relationship(self, part: str, rid: str | None):
        relation = self.relationships(part).get(rid)
        if relation is None:
            self.warn("missing_relationship", part, "Unresolved relationship: " + str(rid))
        elif not relation["external"] and relation["target"] not in self.names:
            self.warn("missing_target_part", part, "Missing target: " + relation["target"])
            return None
        return relation

    def load_notes(self):
        rels = self.relationships(self.main_part)
        for kind, suffix in (("footnote", "footnotes"), ("endnote", "endnotes"), ("comment", "comments")):
            if kind == "comment" and self.scope != "all":
                continue
            candidates = []
            for relation in rels.values():
                if not relation["type"].endswith("/" + suffix):
                    continue
                if relation["external"] or relation["target"] not in self.names:
                    self.warn("missing_note_part", self.main_part, suffix + " relationship has no internal target")
                else:
                    candidates.append(relation["target"])
            conventional = posixpath.dirname(self.main_part) + "/" + suffix + ".xml"
            if conventional in self.names and conventional not in candidates:
                self.warn("unrelated_note_part", conventional, "Part exists without the expected main-document relationship")
                candidates.append(conventional)
            mapping = {}
            for part in candidates:
                self.include(part)
                for item in self.xml(part):
                    if item.tag != q(W, kind):
                        continue
                    if kind != "comment" and item.get(q(W, "type"), "normal") != "normal":
                        continue
                    item_id = item.get(q(W, "id"))
                    if item_id in mapping:
                        self.warn("duplicate_note_id", part, kind + " ID " + str(item_id))
                    mapping[item_id] = (part, item)
            self.note_maps[kind] = mapping

    def load_styles(self):
        result = {}
        for rel in self.relationships(self.main_part).values():
            if rel["type"].endswith("/styles") and not rel["external"] and rel["target"] in self.names:
                for style in self.xml(rel["target"]).findall(q(W, "style")):
                    result[style.get(q(W, "styleId"))] = style
        return result

    def paragraph_numbering(self, paragraph, part):
        numbered = paragraph.find("./" + q(W, "pPr") + "/" + q(W, "numPr")) is not None
        style_ref = paragraph.find("./" + q(W, "pPr") + "/" + q(W, "pStyle"))
        style_id = style_ref.get(q(W, "val")) if style_ref is not None else None
        if style_id is None:
            style_id = next((key for key, value in self.styles.items() if value.get(q(W, "default")) in {"1", "true"} and value.get(q(W, "type")) == "paragraph"), None)
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            style = self.styles.get(style_id)
            if style is None:
                self.warn("missing_style_definition", part, "Cannot resolve paragraph style " + style_id)
                break
            numbered |= style.find(".//" + q(W, "numPr")) is not None
            base = style.find(q(W, "basedOn"))
            style_id = base.get(q(W, "val")) if base is not None else None
        if numbered:
            self.warn("automatic_numbering", part, "Visible numbering/bullets require Word-aware review; this tool compares stored authored text")

    def include(self, part: str):
        if part not in self.snapshot.parts:
            self.snapshot.parts.append(part)

    def fragment(self, part: str, root: ET.Element):
        self.include(part)
        return StoryWalker(self, part).extract(root)

    def note(self, kind: str, item_id: str | None, origin: str):
        self.note_used[kind].add(item_id)
        item = self.note_maps.get(kind, {}).get(item_id)
        if item is None:
            self.warn("missing_note", origin, "Missing " + kind + " ID " + str(item_id))
            return {"missing": kind}
        guard = (kind, item_id)
        if guard in self.note_active:
            self.warn("recursive_note", origin, "Recursive note reference")
            return {"recursive": kind}
        self.note_active.add(guard)
        try:
            part, node = item
            payload = {"records": self.fragment(part, node)}
            if kind == "comment":
                payload["author"] = node.get(q(W, "author"), "")
                payload["initials"] = node.get(q(W, "initials"), "")
            return payload
        finally:
            self.note_active.remove(guard)

    def image(self, node: ET.Element, part: str):
        result = []
        for image in node.iter():
            if image.tag == q(A, "blip"):
                rid = image.get(q(R, "embed")) or image.get(q(R, "link"))
            elif image.tag == q(V, "imagedata"):
                rid = image.get(q(R, "id"))
            else:
                continue
            self.snapshot.counts["images"] += 1
            relation = self.relationship(part, rid)
            if relation is None:
                result.append({"missing_image": True})
            elif relation["external"]:
                self.warn("external_image", part, "External image bytes are not available for verification")
                result.append({"external_target": relation["target"]})
            else:
                if not relation["type"].endswith("/image"):
                    self.warn("unexpected_image_relationship", part, relation["type"])
                raw = self.archive.read(relation["target"])
                result.append({"sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)})
        return result

    def run(self) -> Snapshot:
        body = self.main.find(q(W, "body"))
        if body is None:
            raise AuditInputError("DOCX has no document body")
        self.snapshot.records = [["main", self.fragment(self.main_part, body)]]
        if self.scope == "all":
            furniture = []
            for section in body.iter(q(W, "sectPr")):
                for reference in section:
                    if reference.tag not in {q(W, "headerReference"), q(W, "footerReference")}:
                        continue
                    relation = self.relationship(self.main_part, reference.get(q(R, "id")))
                    kind = "header" if reference.tag == q(W, "headerReference") else "footer"
                    if relation is None or relation["external"]:
                        self.warn("unavailable_page_furniture", self.main_part, kind + " is missing or external")
                        content = [["missing", kind]]
                    else:
                        self.header_used.add(relation["target"])
                        content = self.fragment(relation["target"], self.xml(relation["target"]))
                    furniture.append([kind, reference.get(q(W, "type"), "default"), content])
            for relation in self.relationships(self.main_part).values():
                if relation["type"].split("/")[-1] in {"header", "footer"} and not relation["external"] and relation["target"] not in self.header_used:
                    self.warn("unreferenced_page_furniture", self.main_part, relation["target"])
                    if relation["target"] in self.names:
                        furniture.append(["unreferenced_" + relation["type"].split("/")[-1], self.fragment(relation["target"], self.xml(relation["target"]))])
            self.snapshot.records.append(["page_furniture", furniture])
        for kind, mapping in self.note_maps.items():
            orphans = []
            for item_id in mapping:
                if item_id not in self.note_used[kind]:
                    self.warn("unreferenced_note", mapping[item_id][0], kind + " has no checked anchor")
                    orphans.append(self.note(kind, item_id, self.main_part))
            if orphans:
                self.snapshot.records.append(["unreferenced_" + kind, sorted(orphans, key=digest)])
        if not self.snapshot.counts["paragraphs"] and not any(self.snapshot.counts[key] for key in ("images", "equations", "tables", "fields")):
            self.warn("empty_checked_content", self.main_part, "No authored paragraph text or supported objects were found")
        for relation in self.relationships(self.main_part).values():
            ending = relation["type"].split("/")[-1]
            if ending in {"aFChunk", "subDocument", "commentsExtended", "commentsExtensible", "commentsIds", "people", "glossaryDocument"}:
                if self.scope == "all" or not ending.startswith("comments") and ending != "people":
                    self.warn("unsupported_content_part", self.main_part, ending + ": " + relation["target"])
        return self.snapshot


class StoryWalker:
    PROPERTY_TAGS = {q(W, name) for name in ("pPr", "rPr", "tblPr", "tblGrid", "trPr", "tcPr", "sectPr")}
    REVISION_NAMES = {"ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"}
    WRAPPERS = {"body", "document", "p", "r", "sdt", "sdtContent", "customXml", "smartTag", "dir", "bdo", "footnote", "endnote", "comment", "hdr", "ftr"}
    IGNORED = {"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd", "lastRenderedPageBreak", "footnoteRef", "endnoteRef", "annotationRef"}

    def __init__(self, audit: PackageAudit, part: str):
        self.audit, self.part = audit, part
        self.blocks, self.inline, self.stack = [], [], []

    def emit(self, kind, value):
        target = self.stack[-1]["cache"] if self.stack else self.inline
        if kind == "text" and target and target[-1][0] == "text":
            target[-1][1] += value
        elif kind != "text" or value:
            target.append([kind, value])

    def block(self, record):
        if self.stack:
            self.stack[-1]["cache"].append(record)
        else:
            self.blocks.append(record)

    def start_field(self, instruction="", simple=False):
        descriptor = {"kind": "simple" if simple else "complex", "instruction": instruction, "nested": []}
        if self.stack:
            self.stack[-1]["descriptor"]["nested"].append(descriptor)
        self.emit("field", descriptor)
        self.stack.append({"descriptor": descriptor, "cache": [], "separated": simple})
        self.audit.snapshot.counts["fields"] += 1
        self.audit.warn("field_requires_review", self.part, "Field instructions are compared; cached results are derived data requiring manual/Word review")

    def finish_field(self):
        if not self.stack:
            self.audit.warn("malformed_field", self.part, "Field end without field start")
            return
        current = self.stack.pop()
        self.audit.snapshot.derived_fields.append({
            "part": self.part, "instruction": current["descriptor"]["instruction"],
            "cache_sha256": digest(current["cache"]), "cache_preview": repr(current["cache"])[:600],
        })

    def extract(self, root):
        self.walk(root)
        while self.stack:
            self.audit.warn("malformed_field", self.part, "Unclosed complex field; some cached text cannot be classified")
            self.finish_field()
        return self.blocks

    def walk(self, node):
        tag = node.tag
        local = tag.rsplit("}", 1)[-1]
        if tag in self.PROPERTY_TAGS:
            if any(element.tag.rsplit("}", 1)[-1].endswith("Change") for element in node.iter()):
                self.audit.warn("tracked_revision", self.part, "Tracked property revisions require review")
            return
        if tag == q(W, "p"):
            self.audit.paragraph_numbering(node, self.part)
            previous, self.inline = self.inline, []
            for child in node:
                self.walk(child)
            if self.inline:
                self.audit.snapshot.counts["paragraphs"] += 1
                # Authored text/field descriptors preceding a multi-paragraph
                # field must not be swallowed into that field's cached result.
                self.blocks.append(["paragraph", self.inline])
            self.inline = previous
            return
        if tag in {q(W, "tbl"), q(W, "tr"), q(W, "tc")}:
            if tag == q(W, "tbl"):
                self.audit.snapshot.counts["tables"] += 1
            name = {q(W, "tbl"): "table", q(W, "tr"): "row", q(W, "tc"): "cell"}[tag]
            self.block([name + "_start"])
            for child in node:
                self.walk(child)
            self.block([name + "_end"])
            return
        if tag in {q(W, "t"), q(W, "delText")}:
            self.emit("text", node.text or "")
            if tag == q(W, "delText"):
                self.audit.warn("tracked_revision", self.part, "Deleted text is present; no revision acceptance is inferred")
            return
        if tag == q(W, "tab"):
            self.emit("text", "\t")
            return
        if tag in {q(W, "br"), q(W, "cr")}:
            if node.get(q(W, "type"), "textWrapping") == "textWrapping":
                self.emit("text", "\n")
            return
        if tag in {q(W, "noBreakHyphen"), q(W, "softHyphen")}:
            self.emit("text", "\u2011" if local == "noBreakHyphen" else "\u00ad")
            return
        if tag == q(W, "hyperlink"):
            relation = self.audit.relationship(self.part, node.get(q(R, "id"))) if node.get(q(R, "id")) else None
            target = {"anchor": node.get(q(W, "anchor"), ""), "docLocation": node.get(q(W, "docLocation"), "")}
            if target["anchor"]:
                self.audit.warn("internal_anchor_review", self.part, "Bookmark target name is compared; bookmark existence, range and resulting navigation require manual review")
            if relation:
                target.update({"target": relation["target"], "external": relation["external"]})
                if not relation["type"].endswith("/hyperlink"):
                    self.audit.warn("unexpected_hyperlink_relationship", self.part, relation["type"])
            elif node.get(q(R, "id")):
                target["missing_relationship"] = True
            elif not target["anchor"]:
                self.audit.warn("missing_hyperlink_target", self.part, "Hyperlink has neither relationship nor anchor")
            self.audit.snapshot.counts["hyperlinks"] += 1
            self.emit("hyperlink_start", target)
            for child in node:
                self.walk(child)
            self.emit("hyperlink_end", None)
            return
        if tag in {q(W, "footnoteReference"), q(W, "endnoteReference"), q(W, "commentReference"), q(W, "commentRangeStart"), q(W, "commentRangeEnd")}:
            kind = "footnote" if local.startswith("footnote") else "endnote" if local.startswith("endnote") else "comment"
            if kind != "comment" or self.audit.scope == "all":
                self.emit(local, self.audit.note(kind, node.get(q(W, "id")), self.part))
            return
        if tag == q(W, "fldSimple"):
            self.start_field(node.get(q(W, "instr"), ""), simple=True)
            for child in node:
                self.walk(child)
            self.finish_field()
            return
        if tag == q(W, "fldChar"):
            kind = node.get(q(W, "fldCharType"))
            if kind == "begin":
                self.start_field()
            elif kind == "separate" and self.stack:
                self.stack[-1]["separated"] = True
            elif kind == "end":
                self.finish_field()
            else:
                self.audit.warn("malformed_field", self.part, "Unexpected field marker " + str(kind))
            return
        if tag in {q(W, "instrText"), q(W, "delInstrText")}:
            if self.stack:
                self.stack[-1]["descriptor"]["instruction"] += node.text or ""
            else:
                self.emit("orphan_field_instruction", node.text or "")
                self.audit.warn("malformed_field", self.part, "Instruction text outside a field")
            return
        if tag in {q(M, "oMath"), q(M, "oMathPara")}:
            self.audit.snapshot.counts["equations"] += 1
            self.emit("omml", canonical_math(node))
            return
        if tag in {q(W, "drawing"), q(W, "pict"), q(W, "object")}:
            images = self.audit.image(node, self.part)
            for image in images:
                self.emit("image", image)
            for element in node.iter():
                relationship_attributes = {key for key in element.attrib if key.startswith("{" + R + "}")}
                checked_image_attributes = {q(R, "embed"), q(R, "link")} if element.tag == q(A, "blip") else {q(R, "id")} if element.tag == q(V, "imagedata") else set()
                if relationship_attributes - checked_image_attributes or element.tag == q(A, "blip") and element.get(q(R, "embed")) and element.get(q(R, "link")):
                    self.audit.warn("unverified_drawing_relationship", self.part, "Drawing has an additional action, alternate image or object relationship; its behavior/target needs manual review")
                if element.tag.endswith("}docPr"):
                    alt = {key: element.get(key) for key in ("descr", "title") if element.get(key)}
                    if alt:
                        self.emit("image_alternative_text", alt)
            graphic_data = list(node.iter(q(A, "graphicData")))
            picture_only = tag == q(W, "drawing") and images and graphic_data and all(element.get("uri", "").endswith("/picture") for element in graphic_data) and not any(element.tag in {q(W, "txbxContent"), q(W, "object")} for element in node.iter())
            if not picture_only:
                self.audit.warn("complex_drawing_object", self.part, "Drawing/VML/OLE/chart/textbox requires visual and object review")
                for text_box in node.iter(q(W, "txbxContent")):
                    self.emit("textbox_text", self.audit.fragment(self.part, text_box))
            return
        if tag == q(W, "sym"):
            self.emit("symbol", sorted(node.attrib.items()))
            self.audit.warn("font_symbol", self.part, "Font symbol code compared; glyph rendering is not verified")
            return
        if tag == q(MC, "AlternateContent"):
            self.audit.warn("alternate_content", self.part, "Choice/Fallback content requires application-aware review")
            return
        if tag.startswith("{" + W + "}"):
            if local in self.REVISION_NAMES or local.endswith("Change"):
                self.audit.warn("tracked_revision", self.part, "Revision content retained for audit; no accepted/rejected view is inferred")
            elif local in self.IGNORED:
                return
            elif local == "sdtPr":
                if node.find(".//" + q(W, "dataBinding")) is not None:
                    self.audit.warn("bound_content_control", self.part, "Content control has an external/custom XML data binding")
                return
            elif local not in self.WRAPPERS:
                self.audit.warn("unsupported_word_element", self.part, local + " requires review; nested stored text is inspected when present")
        else:
            self.audit.warn("unsupported_namespace_element", self.part, tag + " requires review")
        for child in node:
            self.walk(child)


def read_docx(path: Path, scope: str) -> Snapshot:
    audit = PackageAudit(path, scope)
    try:
        return audit.run()
    except (OSError, BadZipFile, RuntimeError, ET.ParseError, KeyError) as exc:
        raise AuditInputError(str(exc)) from exc
    finally:
        audit.close()
