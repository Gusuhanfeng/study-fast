#!/usr/bin/env python3
"""Check study-fast for structure, portability, privacy, syntax, and template consistency."""

from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path


IGNORED = {".git", ".venv", "venv", "__pycache__", ".pytest_cache", "node_modules"}
TEXT_SUFFIXES = {".py", ".md", ".json", ".txt", ".yml", ".yaml"}
REQUIRED = {
    ".github/workflows/test.yml", ".gitignore", "CONTRIBUTING.md", "LICENSE",
    "README.md", "README.en.md", "SECURITY.md", "SKILL.md",
    "agents/openai.yaml", "requirements.txt",
    "assets/formal-style.json", "assets/pack-config.example.json",
    "assets/pack-config.schema.json", "examples/selected-handbook.json",
    "references/scenario-playbooks.md", "references/output-contract.md",
    "references/privacy-and-evidence.md", "references/formal-document-template.md",
    "references/learning-design.md",
    "scripts/build_pack.py", "scripts/audit_docx.py", "scripts/common.py", "scripts/self_check.py",
    "tests/test_tooling.py", "evals/evals.json",
}
PRIVACY_PATTERNS = {
    "windows_user_path": re.compile(r"\b[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/]", re.I),
    "unix_user_path": re.compile(r"/(?:home|Users)/[^/\s]+"),
    "email": re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I),
    "phone": re.compile(r"(?<!\d)(?:1[3-9]\d{9}|\+\d[\d ()-]{8,}\d|\(\d{2,4}\)[\d -]{6,}\d)(?!\d)"),
    "private_key": re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
}


def files_under(root):
    for path in root.rglob("*"):
        if path.is_file() and not any(part in IGNORED for part in path.relative_to(root).parts):
            yield path


def main():
    root = Path(__file__).resolve().parents[1]
    errors = []
    texts = {}
    for path in files_under(root):
        if path.suffix.lower() in TEXT_SUFFIXES:
            try:
                texts[path] = path.read_text(encoding="utf-8")
            except UnicodeDecodeError as exc:
                errors.append("Not UTF-8: %s (%s)" % (path.relative_to(root), exc))
    for path, content in texts.items():
        if path.suffix.lower() == ".py":
            try:
                ast.parse(content)
            except SyntaxError as exc:
                errors.append("Python syntax: %s (%s)" % (path.relative_to(root), exc))
        elif path.suffix.lower() == ".json":
            try:
                json.loads(content)
            except json.JSONDecodeError as exc:
                errors.append("JSON syntax: %s (%s)" % (path.relative_to(root), exc))
    skill = texts.get(root / "SKILL.md", "")
    frontmatter = re.match(r"^---\n(.*?)\n---\n", skill, re.S)
    if not frontmatter:
        errors.append("SKILL.md has no valid frontmatter.")
    else:
        lines = [line for line in frontmatter.group(1).splitlines() if line.strip()]
        keys = [line.split(":", 1)[0].strip() for line in lines if not line[0].isspace() and ":" in line]
        if keys[:2] != ["name", "description"] or any(key not in {"name", "description", "metadata"} for key in keys):
            errors.append("SKILL.md frontmatter must start with name and description and may include metadata.")
        if not re.search(r"^name:\s*study-fast\s*$", frontmatter.group(1), re.M):
            errors.append("SKILL.md name does not match the directory.")
    if len(skill.splitlines()) > 500:
        errors.append("SKILL.md exceeds 500 lines.")
    if "TODO" in skill or "[TODO" in skill:
        errors.append("SKILL.md contains an unfinished placeholder.")
    for relative in sorted(REQUIRED):
        if not (root / relative).is_file():
            errors.append("Missing required file: " + relative)
    link_pattern = re.compile(r"\[[^\]]*\]\(([^)]+)\)")
    for path, content in texts.items():
        if path.suffix.lower() != ".md":
            continue
        for link in link_pattern.findall(content):
            if "://" in link or link.startswith("#"):
                continue
            target = link.split("#", 1)[0]
            if target and not (path.parent / target).resolve().exists():
                errors.append("Broken link: %s -> %s" % (path.relative_to(root), link))
    for path, content in texts.items():
        relative = path.relative_to(root)
        if relative.parts and relative.parts[0] in {"scripts", "tests"}:
            continue
        for label, pattern in PRIVACY_PATTERNS.items():
            if pattern.search(content):
                errors.append("Privacy candidate: %s (%s)" % (relative, label))
    binaries = [str(path.relative_to(root)) for path in files_under(root)
                if path.suffix.lower() in {".docx", ".doc", ".pdf", ".pptx", ".xlsx"}]
    if binaries:
        errors.append("Generated or private binary files are bundled: " + ", ".join(binaries))
    style_path = root / "assets" / "formal-style.json"
    if style_path.is_file():
        style = json.loads(style_path.read_text(encoding="utf-8"))
        templates = style.get("templates", {})
        if style.get("default") not in templates:
            errors.append("Formal style default does not name a template.")
        required_templates = {"formal_zh_cn", "formal_cross_platform"}
        if not required_templates.issubset(templates):
            errors.append("Formal style catalog is missing a required template.")
        for name, template in templates.items():
            fonts = template.get("fonts", {})
            if not fonts.get("cjk") or not fonts.get("latin"):
                errors.append("Template %s has incomplete font settings." % name)
            if set(template.get("colors", {}).values()) != {"000000", "FFFFFF"}:
                errors.append("Template %s is not black and white." % name)
    eval_path = root / "evals" / "evals.json"
    if eval_path.is_file():
        payload = json.loads(eval_path.read_text(encoding="utf-8"))
        if payload.get("skill_name") != "study-fast":
            errors.append("Eval skill_name does not match.")
        if len(payload.get("evals", [])) < 6:
            errors.append("Evals do not cover all six primary scenarios.")
        for item in payload.get("evals", []):
            if not item.get("assertions"):
                errors.append("Eval %s has no assertions." % item.get("id"))
    result = {
        "root": ".",
        "files_checked": len(texts),
        "python_files": sum(path.suffix.lower() == ".py" for path in texts),
        "privacy_or_portability_findings": sum("Privacy candidate" in item for item in errors),
        "errors": sorted(set(errors)),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
