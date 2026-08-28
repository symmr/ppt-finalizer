# Changelog

[PPT Finalizer](https://symmr.github.io/ppt-finalizer/) の変更履歴。  
画面下部のバージョン表示は `docs/version.json` を参照。

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠。

## [1.3.0] - 2026-08-28

### Changed

- 処理ロジックを `docs/app.js` に分離（`docs/index.html` は UI のみ）
- リポジトリを GitHub Pages 向けに整理（`web/`・`deploy.ps1`・`SKILL.md` を公開リポジトリから除外）

### Removed

- Python CLI（`scripts/replace_fonts.py`）— Web 版が上位互換のため

## [1.2.0] - 2026-08-28

### Added

- 画面下部にバージョン表示（`docs/version.json` から取得）

### Fixed

- タイトル／本文フォント指定時にすべて本文フォントになる不具合
  - `titleStyle` 置換後に本文フォントの一括置換で上書きされていた処理順序を修正
  - `<p:spPr>` 等を `<p:sp>` と誤検出していた shape パースを修正

### Changed

- README からローカル起動・開発手順・ Python CLI の記載を削除

## [1.1.0] - 2026-08-28

### Added

- タイトル用・本文用フォントの個別指定（プレースホルダー種別と `titleStyle` / `bodyStyle` に応じて適用）
- メディア容量分析に圧縮後サイズ列を追加（展開後サイズと併記）
- GitHub Pages 公開（`/docs`）

## [1.0.0] - 2026-08-28

### Added

- ブラウザ内 PPTX 仕上げツール（ファイルはサーバーへ送信しない）
- フォント統一（既定: Noto Sans JP）と埋込フォント削除
- フォント分析（使用フォント・埋込ファイルサイズ・使用箇所数）
- メディア容量分析（使用スライドマップ、画像ホバープレビュー）
- 孤立メディア削除
- 未使用レイアウト／マスター削除（オプション）
- スピーカーノート削除（オプション）
- ファイルプロパティ削除（オプション）
- 上書き保存（File System Access API、バックアップ付き）と別名ダウンロード

[1.3.0]: https://github.com/symmr/ppt-finalizer/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/symmr/ppt-finalizer/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/symmr/ppt-finalizer/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/symmr/ppt-finalizer/releases/tag/v1.0.0
