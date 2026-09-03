# Changelog

[PPT Translation Supporter](https://symmr.github.io/ppt-translate-supporter/) の変更履歴。  
リポジトリ: [github.com/symmr/ppt-translate-supporter](https://github.com/symmr/ppt-translate-supporter/)

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠。

## [1.2.0] - 2026-09-03

### Changed

- 名称を PPT Translation Supporter に修正（英語として自然な表記に）
- トップの説明文とステップ2に「翻訳自体はこのツールでは行わない」旨を明記し、LLM に貼り付ける手順を明確化
- uid 不一致時、件数表示を実際に一致した件数（`一致 X 件 / 抽出 Y 件`）に変更し、見かけ上の件数一致による誤解を防止
- uid 不一致時は成功メッセージと警告メッセージを分けて出さず、1つの警告メッセージに完了状況をまとめて表示
- クリップボードへのコピー失敗時にエラーメッセージを表示するよう修正（従来は無反応だった）

## [1.1.0] - 2026-09-03

### Changed

- 名称を PPT Translate Supporter に変更
- サンプル PPTX を Pages から外し、`test/fixtures` のみに置く

## [1.0.2] - 2026-09-03

### Added

- ブラウザ確認用のサンプル PPTX（`docs/sample/smoke-test.pptx`）

## [1.0.1] - 2026-09-03

### Added

- 翻訳テキストの貼り付け（ドロップ以外）

### Changed

- 翻訳プロンプト第5項から製品名の例示を削除

## [1.0.0] - 2026-09-03

### Added

- PPTX からのテキスト抽出（図形・グループ・表・ノート）
- run 単位の `[0]...[/0]` タグ
- 翻訳プロンプトの表示とコピー
- 翻訳済み txt の書き戻しと `_translated.pptx` ダウンロード
- ノート枠が空の PPTX でも抽出が落ちないガード
