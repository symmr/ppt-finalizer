# Changelog

[PPT Finalizer](https://symmr.github.io/ppt-finalizer/) の変更履歴。  
リポジトリ: [github.com/symmr/ppt-finalizer](https://github.com/symmr/ppt-finalizer/)  
画面下部のバージョン表示は `docs/version.json` を参照。

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠。

## [1.1.0] - 2026-08-31

### Changed

- ファイル／フォント／メディア分析と処理結果を右側パネルに表示（2 カラムレイアウト）

## [1.0] - 2026-08-28

初回公開リリース。

### Added

- ブラウザ内 PPTX 仕上げ（ファイルはサーバーへ送信しない）
- タイトル／本文フォントの個別統一（既定: Noto Sans JP）と埋込フォント削除
- ファイル分析（ファイルサイズ・スライド数・削減見込み）
- フォント分析（使用フォント・埋込サイズ・使用箇所）
- メディア容量分析（サイズ順一覧・画像ホバープレビュー）
- 整理オプション: 孤立メディア削除、未使用レイアウト／マスター削除、スピーカーノート削除、ファイルプロパティ削除
- 上書き保存（File System Access API、バックアップ付き）と別名ダウンロード
- フォント選択・整理オプションの設定記憶（localStorage）

[1.1.0]: https://github.com/symmr/ppt-finalizer/compare/v1.0...v1.1.0
[1.0]: https://github.com/symmr/ppt-finalizer/releases/tag/v1.0
