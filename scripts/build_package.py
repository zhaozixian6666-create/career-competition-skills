"""Build a portable release archive from the public project files."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
import argparse


ROOT = Path(__file__).resolve().parent.parent
DOCUMENTS = [
    "README.md",
    "LICENSE",
    "CONTRIBUTING.md",
    "第四届大赛规则基线.md",
    "历届资料索引与变化.md",
]
DIRECTORIES = ["skills", "media", "examples"]
TOOLS = ["scripts/build_package.py", "scripts/validate_package.py"]


def build(version: str) -> Path:
    files = [ROOT / name for name in [*DOCUMENTS, *TOOLS]]
    files.extend(
        path
        for dirname in DIRECTORIES
        for path in (ROOT / dirname).rglob("*")
        if path.is_file() and path.suffix != ".pyc" and "__pycache__" not in path.parts
    )
    missing = [path for path in files if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Missing release files: {missing}")

    target = ROOT / "dist" / f"career-competition-skills-v{version}.zip"
    target.parent.mkdir(exist_ok=True)
    with ZipFile(target, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(files, key=lambda item: item.relative_to(ROOT).as_posix()):
            relative = path.relative_to(ROOT).as_posix()
            info = ZipInfo(relative, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            archive.writestr(info, path.read_bytes(), compress_type=ZIP_DEFLATED, compresslevel=9)

    with ZipFile(target) as archive:
        broken = archive.testzip()
        if broken:
            raise RuntimeError(f"Archive CRC failed: {broken}")
        assert set(archive.namelist()) == {path.relative_to(ROOT).as_posix() for path in files}
    return target


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default="0.3.1")
    args = parser.parse_args()
    print(build(args.version))
