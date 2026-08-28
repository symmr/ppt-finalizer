---
name: ppt-finalizer
description: >-
  PPT Finalizer: unify fonts in PowerPoint (.pptx), analyze media usage,
  remove orphan media (with size preview), and optionally delete unused
  layouts/masters. Browser tool (default) or Python CLI for font-only edits.
  Use for フォント変更, Noto Sans JP, PPT Finalizer, 孤立メディア削除,
  未使用レイアウト削除, or font replace without translating.
---

# PPT Finalizer

PowerPoint (.pptx) の仕上げツール。フォント統一（既定: **Noto Sans JP**）、メディア容量分析、孤立メディア削除、未使用レイアウト／マスター整理。翻訳は行いません。

## Prerequisites

- Source `.pptx` must not be open in PowerPoint

## Workflow A — Browser (no Python)

```
- [ ] Step 1: Open the web tool, set options, finalize the .pptx
- [ ] Step 2: Report output filename, size change, and cleanup stats
```

Open in a browser (file stays local; nothing is uploaded):

```
C:\Users\syamamur\.cursor\skills\ppt-replace-fonts\web\replace-fonts.html
```

1. Set **タイトル** / **本文** fonts (default both: **Noto Sans JP**; presets, PPTX-detected fonts, or custom)
   - Title placeholders (`title`, `ctrTitle`, `subTitle`) and `titleStyle` → title font
   - Body placeholders, `bodyStyle`, tables/charts/theme leftovers → body font
   - Free text boxes and shapes without title placeholder → body font
2. Open or drop a `.pptx`
   - **上書き保存**: **ファイルを開く（上書き可）** (Chrome/Edge), or drag-drop when the browser grants a file handle
   - **別名ダウンロード**: any browser; drop or legacy file picker
3. Review **整理オプション** (preview shows planned deletion size/count)
   - **孤立メディアを削除** (default ON)
   - **未使用レイアウト／マスターを削除** (advanced, OFF by default)
   - **スピーカーノートを削除** (OFF by default)
   - **ファイルプロパティを削除** (OFF by default): author, company, custom props, thumbnail
4. Expand **フォント分析** / **メディア容量分析** to inspect the deck (collapsed by default)
5. Click **仕上げて上書き（バックアップ付き）** or **仕上げてダウンロード**

Processing order: font replace → optional structure cleanup → optional notes/properties → optional orphan media removal.

Overwrite flow (File System Access API):

- Creates `<name>_backup.pptx` beside the original (timestamp if name exists)
- Writes finalized file back to the original path

Download output: `<original>_finalized.pptx`

## GitHub Pages

Static site entry point:

```
docs/index.html
```

After editing `web/replace-fonts.html`, sync before deploy:

```powershell
Copy-Item -Force "C:\Users\syamamur\.cursor\skills\ppt-replace-fonts\web\replace-fonts.html" "C:\Users\syamamur\.cursor\skills\ppt-replace-fonts\docs\index.html"
```

Deploy (after `gh auth login`):

```powershell
cd "C:\Users\syamamur\.cursor\skills\ppt-replace-fonts"
.\deploy.ps1
```

Public URL: `https://symmr.github.io/ppt-finalizer/`
Repo: `https://github.com/symmr/ppt-finalizer`

Notes:

- Static hosting only; JSZip from jsDelivr CDN (network on first visit)
- Overwrite save: Chrome/Edge over HTTPS (GitHub Pages OK; `file://` does not support overwrite)

## Workflow B — Python CLI (font replace only)

Requires Python 3 (stdlib only).

```powershell
python "C:\Users\syamamur\.cursor\skills\ppt-replace-fonts\scripts\replace_fonts.py" "<absolute-path-to.pptx>"
```

Custom font:

```powershell
python "C:\Users\syamamur\.cursor\skills\ppt-replace-fonts\scripts\replace_fonts.py" "<absolute-path-to.pptx>" "Meiryo"
```

CLI does **not** remove orphan media or unused layouts — use the browser tool for cleanup.

## Step 2 — Report

Tell the user:

- Output path / download filename
- Backup path (if overwrite)
- Target font, typeface replacement count, embedded fonts removed
- Orphan media removed (count and MB) if option was enabled
- Layouts/masters removed if advanced option was enabled
- File size before → after and reduction %

## User input examples

| User says | Action |
| --- | --- |
| `PPT Finalizer` / フォント統一 + cleanup | Browser tool with options |
| `孤立メディア削除` + file | Browser; enable orphan media (default) |
| `未使用レイアウト削除` | Browser; enable advanced structure option |
| `翻訳不要。フォントだけ` | Browser or CLI (no translation) |
| Font to Meiryo | Set font field, then finalize |
| Browser-only / ブラウザで | Open `web/replace-fonts.html` |

## Do not

- Translate slide text (use `ppt-translate` skill)
- Use ExJector for font-only requests
