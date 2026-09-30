# PPT Translation Supporter

PowerPoint (`.pptx`) と Word (`.docx`) のテキストをブラウザ内で抽出し、翻訳文を書き戻すツール。ファイルはサーバへ送信されません。DOCX は本文、表、テキストボックス（`word/document.xml`）が対象です。ヘッダー、フッター、脚注は対象外です。

- **Web:** https://symmr.github.io/ppt-translation-supporter/
- **GitHub:** https://github.com/symmr/ppt-translation-supporter/

デッキ全体のフォント統一・画像圧縮・孤立メディア削除は [PPT Finalizer](https://symmr.github.io/ppt-finalizer/) で行います。

## 使い方

1. `.pptx` または `.docx` をドロップする（別のファイルを置くと、それまでの抽出テキスト・貼り付けた訳文・書き戻し結果はクリアされます）
2. 図の中の文字を載せる場合は、抽出欄の上にある「画像からテキスト抽出」を押す。本文に貼られた png / jpeg / webp だけが一覧になる。画像を選ぶとモーダルで読み取り結果が出る。不要な行は「消す」か「一括削除」で除く。原文は欄で直せる。「完了」でその画像のテキストボックスを載せ（同じ画像を再度開いて完了すると差し替え）、抽出テキストがやり直される。DOCX ではテキストボックスを図と同じグループに入れるので、図と一緒に動く
3. 抽出テキストと翻訳プロンプトをコピー、または `_to_translate.txt` をダウンロードする
4. 翻訳済みテキスト（`uid_0001` 行を残したもの）をドロップするか、画面に貼り付ける
5. `_translated.pptx` または `_translated.docx` をダウンロードする

翻訳そのものはこのページでは行いません。コピーしたプロンプトを LLM に渡し、返ってきたテキストをドロップまたは貼り付けてください。

プロンプトは画面上で編集でき、内容はブラウザに保存されます。訳先の言語や文体を変えたい場合は書き換えてください。「プロンプトを既定に戻す」で既定値に戻せます。

書き戻すときにタイトルと本文のフォントを指定できます。候補は [PPT Finalizer](https://symmr.github.io/ppt-finalizer/) と同じおすすめフォントに加え、そのファイルが実際に使っているフォント（使用箇所数つき）と自由入力から選べます。既定は Noto Sans JP で、元のフォントをそのまま使う場合は「指定しない」を選びます。指定した場合は訳文を書き戻した箇所にだけ適用されます。PPTX では `latin` / `ea` / `cs`、DOCX では `ascii` / `hAnsi` / `eastAsia` / `cs` を設定します。「タイトル」は、PPTX ではタイトルプレースホルダー（`title` / `ctrTitle` / `subTitle`）の図形、DOCX では見出しスタイル（`Title` / `Subtitle` / `Heading`）の段落が対象です。それ以外は本文フォントです。

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

リポジトリ構成: `docs/index.html`（UI）+ `docs/app.js` + `docs/pptx-text.js`（PPTX の抽出・書き戻し）+ `docs/docx-text.js`（DOCX の抽出・書き戻し）+ `docs/version.json`。

## ライセンス

MIT
