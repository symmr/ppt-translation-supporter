# Changelog

[PPT ExJector](https://symmr.github.io/ppt-exjector/) の変更履歴。  
リポジトリ: [github.com/symmr/ppt-exjector](https://github.com/symmr/ppt-exjector/)

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠。

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
