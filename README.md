# PPT ExJector

PowerPoint (`.pptx`) のテキストをブラウザ内で抽出し、翻訳文を書き戻すツール。ファイルはサーバへ送信されません。

- **Web:** https://symmr.github.io/ppt-exjector/
- **GitHub:** https://github.com/symmr/ppt-exjector/

フォント統一・画像圧縮・孤立メディア削除は [PPT Finalizer](https://symmr.github.io/ppt-finalizer/) で行います。

## 使い方

1. `.pptx` をドロップする
2. 抽出テキストと翻訳プロンプトをコピー、または `_to_translate.txt` をダウンロードする
3. 翻訳済みテキスト（`uid_0001` 行を残したもの）をドロップするか、画面に貼り付ける
4. `_translated.pptx` をダウンロードする

翻訳そのものはこのページでは行いません。コピーしたプロンプトを LLM に渡し、返ってきたテキストをドロップしてください。

## サンプル

ブラウザ確認用の小さなデッキ:

- [docs/sample/smoke-test.pptx](docs/sample/smoke-test.pptx)
- Pages: https://symmr.github.io/ppt-exjector/sample/smoke-test.pptx

2 枚。見出し、色の違う run（`[0]...[/0]`）、表、スピーカーノート、製品名混在。再生成:

```sh
python test/fixtures/build_smoke_pptx.py
```

## 開発

```sh
npm install
npm test
```

リポジトリ構成: `docs/index.html`（UI）+ `docs/app.js` + `docs/pptx-text.js`（抽出・書き戻し）+ `docs/version.json`。

## ライセンス

MIT
