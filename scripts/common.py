#!/usr/bin/env python3
"""Shared configuration, template, filename, and privacy helpers for study-fast."""

from __future__ import annotations

import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
STYLE_PATH = ROOT / "assets" / "formal-style.json"

BLOCK_PRIVACY_PATTERNS = {
    "phone": re.compile(r"(?<!\d)(?:1[3-9]\d{9}|\+\d[\d ()-]{8,}\d|\(\d{2,4}\)[\d -]{6,}\d)(?!\d)"),
    "private_key": re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
}
WARN_PRIVACY_PATTERNS = {
    "email": re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I),
    "absolute_windows_path": re.compile(r"(?<![A-Za-z0-9_])[A-Za-z]:[\\/]"),
    "unix_user_path": re.compile(r"/(?:home|Users)/[^/\s]+"),
    "unc_path": re.compile(r"\\\\[^\\\s]+\\[^\\\s]+"),
    "private_url": re.compile(
        r"https?://(?:localhost|127\.0\.0\.1|10\.\d+\.\d+\.\d+|192\.168\.\d+\.\d+|[^/\s]+\.(?:local|internal))(?:[/:]|$)",
        re.I,
    ),
}
SECRET_ASSIGNMENT = re.compile(
    r"\b(?:password|passwd|token|api[_-]?key|secret)\s*[:=]\s*([^\s,;]+)", re.I
)
SAFE_PLACEHOLDERS = re.compile(
    r"^(?:<[^>]+>|\$\{[^}]+\}|YOUR[_-]?[A-Z0-9_-]*|EXAMPLE|PLACEHOLDER|REDACTED|XXX+)$",
    re.I,
)
RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL", *("COM%d" % value for value in range(1, 10)),
    *("LPT%d" % value for value in range(1, 10)),
}


def load_style_catalog():
    payload = json.loads(STYLE_PATH.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or not isinstance(payload.get("templates"), dict):
        raise ValueError("assets/formal-style.json must contain a templates object.")
    default = payload.get("default")
    if default not in payload["templates"]:
        raise ValueError("assets/formal-style.json default template is missing.")
    return payload


def get_template(name=None):
    catalog = load_style_catalog()
    template_name = name or catalog["default"]
    try:
        template = catalog["templates"][template_name]
    except KeyError as exc:
        raise ValueError("Unknown template: %s" % template_name) from exc
    return template_name, template


def collect_strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from collect_strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from collect_strings(child)


def privacy_findings(text, prohibited_terms=()):
    """Return blocking and warning findings, allowing obvious safe placeholders."""
    blocked = []
    warnings = []
    for label, pattern in BLOCK_PRIVACY_PATTERNS.items():
        if pattern.search(text):
            blocked.append(label)
    for label, pattern in WARN_PRIVACY_PATTERNS.items():
        if pattern.search(text):
            warnings.append(label)
    for match in SECRET_ASSIGNMENT.finditer(text):
        value = match.group(1).strip('"\'')
        if not SAFE_PLACEHOLDERS.fullmatch(value):
            blocked.append("secret_assignment")
    lowered = text.casefold()
    for term in prohibited_terms:
        if term.casefold() in lowered:
            blocked.append("prohibited_term:" + term)
    return {"block": sorted(set(blocked)), "warn": sorted(set(warnings))}


def privacy_errors_and_warnings(text, prohibited_terms=(), mode="balanced"):
    findings = privacy_findings(text, prohibited_terms)
    blocked = list(findings["block"])
    warnings = list(findings["warn"])
    if mode == "strict":
        blocked.extend(warnings)
        warnings = []
    return sorted(set(blocked)), sorted(set(warnings))


def filename_error(filename):
    if not isinstance(filename, str) or not filename.strip():
        return "must be a non-empty string."
    if Path(filename).name != filename or "/" in filename or "\\" in filename:
        return "must be a basename without directories."
    if Path(filename).suffix.casefold() != ".docx":
        return "must use the .docx extension."
    if any(ord(character) < 32 for character in filename):
        return "must not contain control characters."
    if any(character in '<>:"|?*' for character in filename):
        return "contains a character that is unsafe on Windows."
    if Path(filename).stem.rstrip(" .").upper() in RESERVED_WINDOWS_NAMES:
        return "uses a reserved Windows device name."
    if filename[-1] in {" ", "."}:
        return "must not end with a space or period."
    return None


def resolve_layout(document):
    assessment = document.get("document_type") in {"question_bank", "mock_exam"}
    defaults = {"header": False, "cover": not assessment, "toc": not assessment,
                "table_style": "three_line", "question_color": "1F4E79" if assessment else "000000",
                "answer_mode": "separate" if document.get("document_type") == "mock_exam" else "adjacent"}
    layout = document.get("layout", {})
    if not isinstance(layout, dict) or set(layout) - set(defaults):
        raise ValueError("layout contains unsupported fields or is not an object.")
    result = dict(defaults, **layout)
    if any(not isinstance(result[k], bool) for k in ("header", "cover", "toc")):
        raise ValueError("header, cover and toc must be booleans.")
    if result["table_style"] not in {"three_line", "grid"} or result["question_color"] not in {"000000", "1F4E79"}:
        raise ValueError("Unsupported table_style or question_color.")
    if result["answer_mode"] not in {"adjacent", "separate"}:
        raise ValueError("answer_mode must be adjacent or separate; every question requires an answer.")
    return result
