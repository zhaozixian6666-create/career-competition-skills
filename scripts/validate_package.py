"""Check the distributable skill package without third-party dependencies."""

from pathlib import Path
import re
import sys


ROOT = Path(__file__).resolve().parent.parent
SKILLS = ROOT / "skills"
errors = []


def check_skill(folder: Path) -> None:
    entry = folder / "SKILL.md"
    if not entry.is_file():
        errors.append(f"Missing {entry.relative_to(ROOT)}")
        return

    content = entry.read_text(encoding="utf-8")
    match = re.match(r"\A---\r?\n(.*?)\r?\n---\r?\n", content, re.S)
    if not match:
        errors.append(f"Invalid frontmatter in {entry.relative_to(ROOT)}")
        return

    fields = {}
    for line in match.group(1).splitlines():
        field = re.fullmatch(r"([a-z-]+):\s*(.+)", line)
        if not field:
            errors.append(f"Unsupported frontmatter line in {entry.relative_to(ROOT)}: {line}")
            continue
        fields[field.group(1)] = field.group(2).strip()

    name = fields.get("name", "")
    if name != folder.name or not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name):
        errors.append(f"Skill name does not match its folder: {folder.name}")
    if not fields.get("description"):
        errors.append(f"Missing description in {entry.relative_to(ROOT)}")
    if set(fields) - {"name", "description"}:
        errors.append(f"Unexpected frontmatter fields in {entry.relative_to(ROOT)}")
    if "[TODO:" in content:
        errors.append(f"Unfinished placeholder in {entry.relative_to(ROOT)}")


def check_markdown(path: Path) -> None:
    content = path.read_text(encoding="utf-8")
    label = path.relative_to(ROOT)
    if re.search(r"\b[A-Za-z]:\\", content):
        errors.append(f"Machine-specific path in {label}")

    for link in re.findall(r"\[[^\]]+\]\(([^)]+)\)", content):
        target = link.split("#", 1)[0]
        if not target or re.match(r"^[a-z]+://", target):
            continue
        if not (path.parent / target).is_file():
            errors.append(f"Broken link in {label}: {link}")


if not SKILLS.is_dir():
    errors.append("Missing skills directory")
else:
    folders = sorted(path for path in SKILLS.iterdir() if path.is_dir())
    for folder in folders:
        check_skill(folder)
    for path in ROOT.rglob("*.md"):
        if "dist" not in path.relative_to(ROOT).parts:
            check_markdown(path)

if errors:
    print("Package validation failed:")
    for error in errors:
        print(f"- {error}")
    sys.exit(1)

print(f"Validated {len(folders)} skills and all local Markdown links.")
