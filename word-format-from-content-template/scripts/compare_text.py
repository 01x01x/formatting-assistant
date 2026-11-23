#!/usr/bin/env python3
"""Audit selected stored content in UTF-8 text/Markdown or DOCX pairs.

0 PASS: the explicitly checked representation matches, without detected gaps.
1 FAIL: authored text, checked structure, or supported objects differ.
2 ERROR: input/encoding/package/report I/O error.
3 REVIEW: coverage needs manual review, or only whitespace-ignoring text matches.
This is not a rendering, layout, or full-DOCX-equivalence validator.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
from itertools import islice
import json
import re
import sys
from pathlib import Path

from docx_text_audit import AuditInputError, Snapshot, digest, read_docx


SCOPE_DESCRIPTIONS = {
    "all": "Main document in stored order, table row/cell boundaries, referenced footnotes/endnotes, headers/footers, comments and comment anchors; supported hyperlinks, image bytes and OMML. Layout/rendering, generated field results, numbering labels and complex objects are not proven equivalent.",
    "content": "Main document in stored order, table row/cell boundaries and referenced footnotes/endnotes; supported hyperlinks, image bytes and OMML in those stories. Headers/footers and comments are intentionally excluded. Layout/rendering, generated field results, numbering labels and complex objects are not proven equivalent.",
}


def normalize(text: str, *, loose: bool = False) -> str:
    # Only transport-level normalization is safe by default.
    if text.startswith("\ufeff"):
        text = text[1:]
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\s+", "", text) if loose else text


def read_text(path: Path) -> str:
    try:
        return normalize(path.read_bytes().decode("utf-8"))
    except (OSError, UnicodeDecodeError) as exc:
        raise AuditInputError("Cannot read UTF-8 text " + str(path) + ": " + str(exc)) from exc


def input_kind(path: Path) -> str:
    extension = path.suffix.lower()
    if extension in {".txt", ".md"}:
        return "text"
    if extension == ".docx":
        return "docx"
    raise AuditInputError("Unsupported input extension " + (extension or "(none)") + "; use UTF-8 .txt/.md or .docx")


def loose_records(value):
    """Ignore whitespace only in authored strings, never in URLs/object IDs."""
    if isinstance(value, list):
        if len(value) == 2 and value[0] == "text" and isinstance(value[1], str):
            return ["text", re.sub(r"\s+", "", value[1])]
        return [loose_records(item) for item in value]
    if isinstance(value, dict):
        return {key: loose_records(item) for key, item in value.items()}
    return value


def representation_lines(value, kind):
    if kind == "text":
        return value.splitlines(keepends=True)
    output = []
    def visit(item, path=""):
        if isinstance(item, list) and item and isinstance(item[0], str):
            if item[0] in {"main", "page_furniture"}:
                visit(item[1], item[0])
            else:
                output.append(path + " " + json.dumps(item, ensure_ascii=False, sort_keys=True) + "\n")
        elif isinstance(item, list):
            for child in item:
                visit(child, path)
        else:
            output.append(path + " " + json.dumps(item, ensure_ascii=False, sort_keys=True) + "\n")
    visit(value)
    return output


def summary(snapshot: Snapshot, representation, kind):
    if kind == "text":
        return {"normalized_chars": len(representation), "normalized_sha256": hashlib.sha256(representation.encode("utf-8")).hexdigest()}
    return {"checked_representation_sha256": digest(representation), "included_parts": snapshot.parts, "counts": snapshot.counts, "warnings": snapshot.warnings, "derived_fields": snapshot.derived_fields}


def write_report(path: Path | None, report):
    if path:
        try:
            path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            raise AuditInputError("Cannot write report " + str(path) + ": " + str(exc)) from exc


def same_input_file(left: Path, right: Path) -> bool:
    """Protect inputs even when a report has another name via a hard link."""
    try:
        if left.resolve() == right.resolve():
            return True
        try:
            return left.samefile(right)
        except FileNotFoundError:
            return False
    except (OSError, RuntimeError) as exc:
        raise AuditInputError("Cannot verify report/input file identity: " + str(exc)) from exc


def configure_console():
    # Preserve the terminal's encoding; unsupported glyphs remain visible as escapes
    # rather than crashing halfway through a diff. JSON reports stay lossless UTF-8.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="backslashreplace")
            except (OSError, ValueError):
                pass


def main(argv=None) -> int:
    configure_console()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("source", type=Path)
    parser.add_argument("final", type=Path)
    parser.add_argument("--loose", action="store_true", help="Whitespace diagnostic only; always REVIEW/3 unless a non-whitespace mismatch requires FAIL/1")
    parser.add_argument("--show-diff", action="store_true", help="Print at most 240 unified-diff lines")
    parser.add_argument("--report", type=Path, help="Write the structured JSON audit report")
    parser.add_argument("--scope", choices=("all", "content"), default="all")
    args = parser.parse_args(argv)
    report = {"schema_version": 1, "source": str(args.source), "final": str(args.final), "scope": args.scope, "loose_requested": args.loose}
    report_path_safe = False
    try:
        if args.report:
            if any(same_input_file(args.report, path) for path in (args.source, args.final)):
                raise AuditInputError("Report path must not overwrite either input (including linked aliases)")
            report_path_safe = True
        source_kind, final_kind = input_kind(args.source), input_kind(args.final)
        if source_kind != final_kind:
            raise AuditInputError("Input types differ: " + source_kind + " vs " + final_kind)
        report["input_kind"] = source_kind
        source_snapshot, final_snapshot = Snapshot(), Snapshot()
        if source_kind == "text":
            source, final = read_text(args.source), read_text(args.final)
            if not source.strip() or not final.strip():
                raise AuditInputError("Empty text input cannot establish successful extraction/content identity")
            report["scope_description"] = "All supplied UTF-8 text; leading BOM and CR/LF encoding only are normalized. DOCX stories are not inferred from extracted text; --scope is a label only for plain text."
            loose_equal = re.sub(r"\s+", "", source) == re.sub(r"\s+", "", final)
        else:
            source_snapshot, final_snapshot = read_docx(args.source, args.scope), read_docx(args.final, args.scope)
            source, final = source_snapshot.records, final_snapshot.records
            report["scope_description"] = SCOPE_DESCRIPTIONS[args.scope]
            loose_equal = loose_records(source) == loose_records(final)
        equal = source == final
        warnings = [{"input": "source", **warning} for warning in source_snapshot.warnings] + [{"input": "final", **warning} for warning in final_snapshot.warnings]
        fields_changed = source_snapshot.derived_fields != final_snapshot.derived_fields
        if fields_changed:
            warnings.append({"input": "pair", "code": "derived_field_results_changed", "part": "", "detail": "Field cache audit records differ; cached results are derived data, not automatically authored-text changes"})
        if not equal and not (args.loose and loose_equal):
            status, exit_code = "FAIL", 1
        elif not equal or warnings or args.loose:
            status, exit_code = "REVIEW", 3
        else:
            status, exit_code = "PASS", 0
        report.update({"status": status, "exit_code": exit_code, "strict_equal": equal, "whitespace_ignoring_equal": loose_equal if args.loose else None, "source_summary": summary(source_snapshot, source, source_kind), "final_summary": summary(final_snapshot, final, source_kind), "warnings": warnings, "derived_field_records_changed": fields_changed})
        diff = list(islice(difflib.unified_diff(representation_lines(source, source_kind), representation_lines(final, source_kind), fromfile=str(args.source), tofile=str(args.final)), 240)) if not equal else []
        report["diff_preview"] = [line.rstrip("\n") for line in diff]
        report["diff_preview_limit_lines"] = 240
        write_report(args.report, report)
        print(f"{status}: checked {source_kind} representation; scope={args.scope}; strict_equal={equal}.")
        print(report["scope_description"])
        if args.loose and not equal and loose_equal:
            print("REVIEW: only whitespace-ignoring text matches; strict differences remain.")
        for warning in warnings[:20]:
            print("REVIEW: " + warning["code"] + " [" + warning["input"] + "] " + warning["detail"])
        if len(warnings) > 20:
            print(f"REVIEW: {len(warnings) - 20} more warnings are recorded in the JSON report.")
        if args.show_diff:
            for line in diff:
                print(line, end="" if line.endswith("\n") else "\n")
        return exit_code
    except AuditInputError as exc:
        report.update({"status": "ERROR", "exit_code": 2, "error": str(exc)})
        if args.report and report_path_safe:
            try:
                write_report(args.report, report)
            except AuditInputError:
                pass
        print("ERROR: " + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
