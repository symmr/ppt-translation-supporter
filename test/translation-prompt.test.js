"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");

const {
  TRANSLATION_PROMPT,
  TRANSLATION_PROMPTS,
  TRANSLATION_DIRECTIONS,
  DEFAULT_DIRECTION,
  translationPrompt,
  detectDirection,
  formatExtractFile,
} = require("../docs/pptx-text.js");

test("default direction is English to Japanese and keeps the original prompt", () => {
  assert.equal(DEFAULT_DIRECTION, "en-ja");
  assert.equal(translationPrompt("en-ja"), TRANSLATION_PROMPT);
  assert.match(TRANSLATION_PROMPT, /日本語に翻訳してください/);
});

test("every listed direction has its own prompt", () => {
  assert.deepEqual(TRANSLATION_DIRECTIONS.map((d) => d.id), ["en-ja", "ja-en"]);
  for (const { id } of TRANSLATION_DIRECTIONS) {
    assert.ok(TRANSLATION_PROMPTS[id], `prompt for ${id}`);
  }
  assert.notEqual(TRANSLATION_PROMPTS["en-ja"], TRANSLATION_PROMPTS["ja-en"]);
});

test("Japanese to English prompt targets English and keeps the uid and tag rules", () => {
  const prompt = translationPrompt("ja-en");
  assert.match(prompt, /英語に翻訳してください/);
  assert.doesNotMatch(prompt, /日本語に翻訳してください/);
  assert.match(prompt, /uid_0001/);
  assert.match(prompt, /\[0\]\.\.\.\[\/0\]/);
});

test("an unknown direction falls back to the default prompt", () => {
  assert.equal(translationPrompt("fr-ja"), TRANSLATION_PROMPT);
  assert.equal(translationPrompt(undefined), TRANSLATION_PROMPT);
});

test("detectDirection picks English to Japanese for an English deck", () => {
  const text = formatExtractFile(["uid_0001", "[0]Quarterly results[/0]", "uid_0002", "Revenue grew 12% year over year"]);
  const found = detectDirection(text);
  assert.equal(found.direction, "en-ja");
  assert.equal(found.ja, 0);
});

test("detectDirection picks Japanese to English for a Japanese deck", () => {
  const text = formatExtractFile(["uid_0001", "四半期の業績", "uid_0002", "売上は前年比 12% 増加した"]);
  assert.equal(detectDirection(text).direction, "ja-en");
});

test("detectDirection counts only characters, not uid lines, tags, digits or symbols", () => {
  // uid_ lines and [n] tags hold Latin letters that are not content
  const text = formatExtractFile(["uid_0001", "[0][/0]", "uid_0002", "2026/10/07 ---"]);
  const found = detectDirection(text);
  assert.equal(found.jaShare, null);
  assert.equal(found.direction, "en-ja");
});

test("detectDirection keeps a Japanese deck full of product names as Japanese", () => {
  // 30 Japanese characters against 40 Latin letters is 43% Japanese
  const text = formatExtractFile(["uid_0001", "Splunk Observability Cloud と OpenTelemetry Collector を使ってサービスの状態を可視化する方法を説明する"]);
  const found = detectDirection(text);
  assert.ok(found.jaShare >= 0.3, `share ${found.jaShare}`);
  assert.equal(found.direction, "ja-en");
});

test("detectDirection keeps an English deck with a few Japanese names as English", () => {
  const text = formatExtractFile(["uid_0001", "Presented by 山田 太郎 at the annual customer summit for the platform team and partners"]);
  assert.equal(detectDirection(text).direction, "en-ja");
});

test("detectDirection treats the 30% line as Japanese", () => {
  // 3 Japanese characters and 7 Latin letters
  assert.equal(detectDirection("あいうabcdefg").direction, "ja-en");
  assert.equal(detectDirection("あいabcdefgh").direction, "en-ja");
});