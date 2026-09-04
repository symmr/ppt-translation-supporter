# PPT Translation Supporter

PowerPoint (`.pptx`) のテキストをブラウザ内で抽出し、翻訳文を書き戻すツール。ファイルはサーバへ送信されません。

- **Web:** https://symmr.github.io/ppt-translation-supporter/
- **GitHub:** https://github.com/symmr/ppt-translation-supporter/

デッキ全体のフォント統一・画像圧縮・孤立メディア削除は [PPT Finalizer](https://symmr.github.io/ppt-finalizer/) で行います。

## 使い方

1. `.pptx` をドロップする
2. 抽出テキストと翻訳プロンプトをコピー、または `_to_translate.txt` をダウンロードする
3. 翻訳済みテキスト（`uid_0001` 行を残したもの）をドロップするか、画面に貼り付ける
4. `_translated.pptx` をダウンロードする

翻訳そのものはこのページでは行いません。コピーしたプロンプトを LLM に渡し、返ってきたテキストをドロップまたは貼り付けてください。

プロンプトは画面上で編集でき、内容はブラウザに保存されます。訳先の言語や文体を変えたい場合は書き換えてください。「プロンプトを既定に戻す」で既定値に戻せます。

書き戻すときにタイトルと本文のフォントを指定できます。候補は [PPT Finalizer](https://symmr.github.io/ppt-finalizer/) と同じおすすめフォントに加え、その PPTX が実際に使っているフォント（使用箇所数つき）と自由入力から選べます。「指定しない」なら元のフォントを変更しません。指定した場合は訳文を書き戻した箇所にだけ適用され、`latin` / `ea` / `cs` を設定します。「タイトル」はタイトルプレースホルダー（`title` / `ctrTitle` / `subTitle`）の図形が対象で、テキストボックスだけで作られたスライドはすべて本文として扱われます。

タグ（`[0]...[/0]`）は run が2つ以上ある段落にだけ付きます。run が1つの段落の訳文にタグらしき文字列があっても、それは本文そのものとして扱い、加工せずに書き戻します（uid は警告に表示されます）。run が2つ以上ある段落で、タグが原文の run と1対1で対応しない場合は、先頭 run にまとめて書き戻します。文字は入りますが run ごとの書式（色分けなど）は失われるため、該当件数と uid が警告に表示されます。

## サンプル

手動確認用の小さなデッキは `test/fixtures/smoke-test.pptx`。2 枚。見出し、色の違う run（`[0]...[/0]`）、表、スピーカーノート、製品名混在。再生成:

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
