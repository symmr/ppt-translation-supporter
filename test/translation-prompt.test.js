"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");

const {
  TRANSLATION_PROMPT,
  TRANSLATION_PROMPTS,
  TRANSLATION_DIRECTIONS,
  DEFAULT_DIRECTION,
  translationPrompt,
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
