# PPT Finalizer

PowerPoint (`.pptx`) の仕上げをブラウザ内で行うツール。ファイルはサーバーへ送信されません。

## 機能

- フォント統一（既定: Noto Sans JP）
- フォント分析（使用フォント・埋込ファイルサイズ）
- メディア容量分析（ホバーで画像プレビュー）
- 孤立メディア削除
- 未使用レイアウト／マスター削除（オプション）
- スピーカーノート削除（オプション）
- ファイルプロパティ削除（オプション）

## 使い方（Web）

**GitHub Pages:** https://symmr.github.io/ppt-finalizer/

1. `.pptx` をドロップまたは開く
2. 整理オプションを確認
3. **仕上げてダウンロード** または **仕上げて上書き**（Chrome / Edge）

上書き保存は File System Access API 対応ブラウザ（Chrome / Edge）で HTTPS 上のみ利用できます。

## ローカル

`web/replace-fonts.html` をブラウザで開いても同じです（`file://` では上書き保存不可）。

## Python CLI（フォント置換のみ）

```powershell
python scripts/replace_fonts.py "path\to\deck.pptx"
```

孤立メディア削除などは Web 版を使用してください。

## 開発

`web/replace-fonts.html` を編集したら GitHub Pages 用に同期:

```powershell
Copy-Item -Force web/replace-fonts.html docs/index.html
```

## ライセンス

MIT
