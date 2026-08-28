# PPT Finalizer

PowerPoint (`.pptx`) の仕上げをブラウザ内で行うツール。ファイルはサーバーへ送信されません。

**https://symmr.github.io/ppt-finalizer/**

## 機能

- タイトル／本文フォントの統一（既定: Noto Sans JP）
- フォント分析（使用フォント・埋込ファイルサイズ）
- メディア容量分析（ホバーで画像プレビュー）
- 孤立メディア削除
- 未使用レイアウト／マスター削除（オプション）
- スピーカーノート削除（オプション）
- ファイルプロパティ削除（オプション）

## 使い方

1. 上記 URL を Chrome または Edge で開く
2. `.pptx` をドロップまたは開く
3. タイトル／本文フォントと整理オプションを確認
4. **仕上げてダウンロード** または **仕上げて上書き**

上書き保存は File System Access API 対応ブラウザ（Chrome / Edge）で HTTPS 上のみ利用できます。

## Python CLI（フォント置換のみ）

```powershell
python scripts/replace_fonts.py "path\to\deck.pptx"
```

孤立メディア削除などは Web 版を使用してください。

## ライセンス

MIT
