"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const {
  PROMPT_STYLES,
  PROMPT_AMOUNT_UNITS,
  FX_SOURCES,
  validFxRate,
  parseFxResponse,
  buildPromptRules,
  buildFullPrompt,
  basePromptOf,
  GLOSSARY_TEMPLATE,
  parseGlossaryCsv,
  normalizeJa,
  matchGlossary,
} = require("../docs/prompt-options.js");
const { TRANSLATION_PROMPT, extractTextsFromZip } = require("../docs/pptx-text.js");

test("the defaults keep the previous behavior: 常体 and amounts as written", () => {
  const rules = buildPromptRules({ direction: "en-ja", style: "plain", amountUnit: "kmb" });
  assert.match(rules, /常体\(だ・である調\)/);
  assert.match(rules, /原文の表記のまま残す/);
  assert.doesNotMatch(rules, /為替換算/);
});

test("the base English prompt no longer fixes a writing style", () => {
  assert.doesNotMatch(TRANSLATION_PROMPT, /常体|敬体/);
});

test("敬体 keeps titles in 常体 and lists the title uids", () => {
  const rules = buildPromptRules({
    direction: "en-ja",
    style: "polite",
    titleIds: ["uid_0001", "uid_0007"],
  });
  assert.match(rules, /敬体\(です・ます調\)/);
  assert.match(rules, /体言止め/);
  assert.match(rules, /uid_0001, uid_0007/);
});

test("敬体 without any title uid does not print an empty list", () => {
  const rules = buildPromptRules({ direction: "en-ja", style: "polite", titleIds: [] });
  assert.doesNotMatch(rules, /次の uid はタイトル/);
});

test("万・億 option converts K/M/B for English to Japanese", () => {
  const rules = buildPromptRules({ direction: "en-ja", amountUnit: "man" });
  assert.match(rules, /万・億/);
  assert.match(rules, /\$250K → 25万ドル/);
});

test("English to Japanese always turns an approximate tilde into 約, with or without exchange", () => {
  for (const amountUnit of ["kmb", "man"]) {
    for (const fx of [undefined, { enabled: true, rate: 150 }]) {
      const rules = buildPromptRules({ direction: "en-ja", amountUnit, fx });
      assert.match(rules, /「~」を訳文に残さない/, `${amountUnit} ${fx ? "fx" : "no fx"}`);
      assert.match(rules, /いつも「約」にする/);
    }
  }
  assert.match(buildPromptRules({ direction: "en-ja", amountUnit: "kmb" }), /~\$100K → 約\$100K/);
  assert.match(buildPromptRules({ direction: "en-ja", amountUnit: "man" }), /~\$100K → 約10万ドル/);
  assert.doesNotMatch(buildPromptRules({ direction: "ja-en" }), /いつも「約」にする/);
});

test("Japanese to English always uses $ and K/M/B and has no style rule", () => {
  for (const amountUnit of PROMPT_AMOUNT_UNITS.map((o) => o.id)) {
    const rules = buildPromptRules({ direction: "ja-en", style: "polite", amountUnit });
    assert.match(rules, /250万ドル → \$2\.5M/);
    assert.doesNotMatch(rules, /敬体|常体/);
  }
});

test("exchange rate: off by default, and an unusable rate adds nothing", () => {
  assert.doesNotMatch(buildPromptRules({ direction: "en-ja", fx: { enabled: false, rate: "150" } }), /為替換算/);
  assert.doesNotMatch(buildPromptRules({ direction: "en-ja", fx: { enabled: true, rate: "" } }), /為替換算/);
  assert.doesNotMatch(buildPromptRules({ direction: "en-ja", fx: { enabled: true, rate: "abc" } }), /為替換算/);
});

test("exchange rate English to Japanese multiplies, rounds and adds 約", () => {
  const rules = buildPromptRules({
    direction: "en-ja",
    fx: { enabled: true, rate: "151.2", date: "2026-10-07", source: "Frankfurter" },
  });
  assert.match(rules, /米ドル.*円に換算/);
  assert.match(rules, /1 USD = 151\.2 JPY\(2026-10-07、Frankfurter\)/);
  assert.match(rules, /掛けて円/);
  assert.match(rules, /「約」/);
  assert.match(rules, /最大でも 3 桁/);
  assert.match(rules, /注記は書かない/);
});

test("exchange rate Japanese to English divides and adds ~", () => {
  const rules = buildPromptRules({ direction: "ja-en", fx: { enabled: true, rate: 150 } });
  assert.match(rules, /米ドルに換算/);
  assert.match(rules, /150 で割って/);
  assert.match(rules, /半角のチルダ「~」/);
  assert.match(rules, /範囲.*意味では使わない/);
});

test("rules are numbered in order and declared to win over the base prompt", () => {
  const rules = buildPromptRules({
    direction: "en-ja",
    fx: { enabled: true, rate: 150 },
  });
  assert.match(rules, /^# 表記ルール（上のルールと食い違う場合は、こちらを優先する）/);
  assert.match(rules, /\n1\. 文体/);
  assert.match(rules, /\n2\. 金額/);
  assert.match(rules, /\n3\. 為替換算/);
});

test("buildFullPrompt joins the base prompt and the rules with one blank line", () => {
  assert.equal(buildFullPrompt("base\n", "rules\n"), "base\n\nrules\n");
  assert.equal(buildFullPrompt("base", ""), "base\n");
});

test("basePromptOf takes back the base prompt from the combined text", () => {
  const rules = buildPromptRules({ direction: "en-ja", style: "plain", amountUnit: "kmb" });
  assert.equal(basePromptOf(buildFullPrompt("edited base\n", rules)), "edited base\n");
  assert.equal(basePromptOf("no rules here\n\n"), "no rules here\n");
});

test("validFxRate accepts positive numbers only", () => {
  assert.equal(validFxRate("150.5"), 150.5);
  for (const bad of ["", "0", "-3", "abc", null, undefined, "1e9"]) {
    assert.equal(validFxRate(bad), null, String(bad));
  }
});

test("parseFxResponse reads both sources and rejects bad replies", () => {
  assert.deepEqual(
    parseFxResponse("frankfurter", { base: "USD", date: "2026-10-06", rates: { JPY: 149.8 } }),
    { rate: 149.8, date: "2026-10-06" }
  );
  assert.deepEqual(
    parseFxResponse("er-api", { result: "success", time_last_update_unix: 1791244801, rates: { JPY: 150.2 } }),
    { rate: 150.2, date: "2026-10-06" }
  );
  assert.equal(parseFxResponse("frankfurter", { rates: {} }), null);
  assert.equal(parseFxResponse("er-api", { result: "error", rates: { JPY: 150 } }), null);
  assert.equal(parseFxResponse("other", {}), null);
  assert.equal(parseFxResponse("frankfurter", null), null);
});

test("each exchange-rate source has a name, a site and an https URL", () => {
  for (const source of FX_SOURCES) {
    assert.ok(source.name && source.site && source.id);
    assert.match(source.url, /^https:\/\//);
  }
  assert.ok(PROMPT_STYLES.length === 2);
});

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

test("PPTX extraction marks title placeholders and leaves body text unmarked", async () => {
  const shape = (id, ph, text) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>
    <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;
  const zip = new JSZip();
  zip.file("ppt/presentation.xml", `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`);
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      ${shape(2, '<p:ph type="title"/>', "Quarterly review")}
      ${shape(3, '<p:ph type="body" idx="1"/>', "Revenue grew steadily")}
    </p:spTree></p:cSld></p:sld>`
  );
  const extracted = await extractTextsFromZip(zip);
  assert.equal(extracted.metadata.length, 2);
  assert.equal(extracted.metadata[0].title, true);
  assert.notEqual(extracted.metadata[1].title, true);
});

test("parseGlossaryCsv reads two columns, with or without the header row", () => {
  const csv = "\uFEFF英語,日本語\r\n\"Notable Event\",注目すべきイベント\n\"a, b\",\"x\"\"y\"\nSplunk,\n# memo,x\n\n";
  assert.deepEqual(parseGlossaryCsv(csv), [
    { en: "Notable Event", ja: "注目すべきイベント" },
    { en: "a, b", ja: 'x"y' },
    { en: "Splunk", ja: "" },
  ]);
  assert.deepEqual(parseGlossaryCsv("alert,アラート"), [{ en: "alert", ja: "アラート" }]);
  assert.equal(parseGlossaryCsv(GLOSSARY_TEMPLATE).length, 3);
});

test("English terms match whole words in any case, with a plural", () => {
  const entries = parseGlossaryCsv("Notable Event,注目すべきイベント\nalert,アラート\nES,拡張セキュリティ\nSOAR,");
  const text = "uid_0001\n[0]Notable  events[/0] and ALERTS in ESG";
  assert.deepEqual(matchGlossary(entries, text, "en-ja").map((e) => e.en), ["Notable Event", "alert"]);
  assert.deepEqual(matchGlossary(entries, "ES and SOAR.", "en-ja").map((e) => e.en), ["ES", "SOAR"]);
});

test("Japanese terms match across width, spacing and a closing long-vowel mark", () => {
  assert.equal(normalizeJa("サーバー"), normalizeJa("ｻｰﾊﾞ"));
  assert.equal(normalizeJa("データ・センター"), normalizeJa("データセンタ"));
  assert.notEqual(normalizeJa("コード"), normalizeJa("コド"));
  const entries = parseGlossaryCsv("server,サーバー\nNotable Event,注目すべきイベント\nalert,アラート");
  const text = "ｻｰﾊﾞの 注目すべき イベント";
  assert.deepEqual(matchGlossary(entries, text, "ja-en").map((e) => e.ja), ["サーバー", "注目すべきイベント"]);
});

test("only matched terms become a glossary rule, in the direction's order", () => {
  const entries = parseGlossaryCsv("alert,アラート\nalert,警告\nSplunk,\nunused,未使用");
  const toJa = buildPromptRules({ direction: "en-ja", glossary: matchGlossary(entries, "Splunk alert", "en-ja") });
  assert.match(toJa, /用語集/);
  assert.match(toJa, /alert → アラート \/ 警告\(文脈に合うものを使う\)/);
  assert.match(toJa, /Splunk\(訳さず、原文のまま残す\)/);
  assert.doesNotMatch(toJa, /unused/);
  const toEn = buildPromptRules({ direction: "ja-en", glossary: matchGlossary(entries, "アラートの件", "ja-en") });
  assert.match(toEn, /アラート → alert/);
  assert.doesNotMatch(buildPromptRules({ direction: "en-ja", glossary: [] }), /用語集/);
});
