"""Replace all fonts in a PPTX with Noto Sans JP and remove embedded fonts."""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

DEFAULT_FONT = "Noto Sans JP"
TYPEFACE_RE = re.compile(r'typeface="[^"]*"')
EMBEDDED_FONT_LST_RE = re.compile(
    r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", re.DOTALL
)
FONT_REL_RE = re.compile(
    r'<Relationship[^>]+Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/font"[^>]*/>\s*'
)


def replace_fonts_in_xml(name: str, data: str, target_font: str) -> tuple[str, int]:
    if name == "ppt/presentation.xml":
        return _replace_presentation_xml(data, target_font)

    if name == "ppt/_rels/presentation.xml.rels":
        return FONT_REL_RE.subn("", data)

    if not name.endswith(".xml"):
        return data, 0

    return TYPEFACE_RE.subn(f'typeface="{target_font}"', data)


def _replace_presentation_xml(data: str, target_font: str) -> tuple[str, int]:
    count = 0
    data, n = EMBEDDED_FONT_LST_RE.subn("", data)
    count += n
    data = data.replace('embedTrueTypeFonts="1"', 'embedTrueTypeFonts="0"')
    data, n = TYPEFACE_RE.subn(f'typeface="{target_font}"', data)
    count += n
    return data, count


def make_backup(src: Path) -> Path:
    backup = src.with_name(f"{src.stem}_backup{src.suffix}")
    if backup.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = src.with_name(f"{src.stem}_backup_{ts}{src.suffix}")
    shutil.copy2(src, backup)
    return backup


def replace_fonts(src: Path, dst: Path | None = None, target_font: str = DEFAULT_FONT) -> dict[str, int | str]:
    dst = dst or src
    in_place = dst.resolve() == src.resolve()
    backup: Path | None = None

    if in_place:
        backup = make_backup(src)

    tmp = dst.with_suffix(".tmp.pptx")
    if tmp.exists():
        tmp.unlink()

    total = 0
    files_changed = 0
    fonts_removed = 0

    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename.startswith("ppt/fonts/"):
                    fonts_removed += 1
                    continue

                data = zin.read(item.filename)
                if item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    new_text, n = replace_fonts_in_xml(item.filename, text, target_font)
                    if new_text != text:
                        files_changed += 1
                        total += n
                        data = new_text.encode("utf-8")
                zout.writestr(item, data)

    if in_place and dst.exists():
        dst.unlink()
    shutil.move(str(tmp), str(dst))

    return {
        "output": str(dst),
        "backup": str(backup) if backup else "",
        "target_font": target_font,
        "replacements": total,
        "files_changed": files_changed,
        "embedded_fonts_removed": fonts_removed,
    }


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(
            "Usage:\n"
            f"  replace_fonts.py <file.pptx> [font_name]\n"
            f"\nDefault font: {DEFAULT_FONT}"
        )
        return 0

    src = Path(sys.argv[1]).resolve()
    target_font = sys.argv[2] if len(sys.argv) >= 3 else DEFAULT_FONT

    if not src.is_file():
        print(f"File not found: {src}", file=sys.stderr)
        return 1

    stats = replace_fonts(src, target_font=target_font)
    if stats["backup"]:
        print(f"Backup: {stats['backup']}")
    print(f"Output: {stats['output']}")
    print(f"Target font: {stats['target_font']}")
    print(f"Typeface replacements: {stats['replacements']}")
    print(f"Embedded font files removed: {stats['embedded_fonts_removed']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
