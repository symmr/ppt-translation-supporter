"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const {
  extractTextsFromZip,
  injectTextsToZip,
  parseUidDelimitedText,
  formatExtractFile,
} = require("../docs/pptx-text.js");

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// A paragraph with a line break and a trailing a:endParaRPr, the shape Google
// Slides exports for wrapped body text.
function slideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:r><a:rPr lang="ja-JP"/><a:t>現実に</a:t></a:r>
            <a:br><a:rPr lang="ja-JP"/></a:br>
            <a:r><a:rPr lang="ja-JP"/><a:t>近い環境で</a:t></a:r>
            <a:endParaRPr lang="ja-JP"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

async function buildZip() {
  const zip = new JSZip();
  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId2" Type="${REL_NS}/slide" Target="slides/slide1.xml"/></Relationships>`
  );
  zip.file("ppt/slides/slide1.xml", slideXml());
  return zip;
}

test("a paragraph with a line break keeps endParaRPr as its last child after write-back", async () => {
  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  const translated = formatExtractFile(["uid_0001", "Experience threat hunting in a realistic environment."]);

  await injectTextsToZip(zip, parseUidDelimitedText(translated), extracted.metadata);

  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  const para = xml.match(/<a:p>[\s\S]*?<\/a:p>/)[0];
  const endAt = para.indexOf("<a:endParaRPr");
  assert.ok(endAt > -1, "endParaRPr is kept");
  assert.ok(para.indexOf("Experience threat hunting") < endAt, "translated run comes before endParaRPr");
  assert.doesNotMatch(para.slice(endAt), /<a:(r|br|t)\b/, "nothing but the closing tag follows endParaRPr");
});

test("a translation with its own line break is placed before endParaRPr too", async () => {
  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  const translated = formatExtractFile(["uid_0001", "first line\nsecond line"]);

  await injectTextsToZip(zip, parseUidDelimitedText(translated), extracted.metadata);

  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  const para = xml.match(/<a:p>[\s\S]*?<\/a:p>/)[0];
  const endAt = para.indexOf("<a:endParaRPr");
  assert.ok(para.indexOf("second line") < endAt);
  assert.ok(para.indexOf("<a:br") < endAt);
  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /uid_0001\nfirst line\nsecond line/);
});

// A wrapped paragraph where one phrase is bold and coloured, as in the BOTS
// deck: "豪華景品" must keep its own formatting after translation.
function emphasisSlideXml() {
  const normal = `<a:rPr lang="en-US" sz="1600" b="0"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="Noto Sans JP"/></a:rPr>`;
  const strong = `<a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Noto Sans JP"/></a:rPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:r>${normal}<a:t>顕著な成績をあげられたチームには</a:t></a:r>
            <a:br>${normal}</a:br>
            <a:r>${normal}<a:t>Splunkより</a:t></a:r>
            <a:r>${strong}<a:t>豪華景品</a:t></a:r>
            <a:r>${normal}<a:t>を用意しております</a:t></a:r>
            <a:r>${normal}<a:t>。</a:t></a:r>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

async function buildEmphasisZip() {
  const zip = await buildZip();
  zip.file("ppt/slides/slide1.xml", emphasisSlideXml());
  return zip;
}

function firstParagraph(xml) {
  return xml.match(/<a:p>[\s\S]*?<\/a:p>/)[0];
}

test("a wrapped paragraph with a bold phrase is extracted with one tag per formatting group", async () => {
  const extracted = await extractTextsFromZip(await buildEmphasisZip());
  assert.match(
    extracted.text,
    /uid_0001\n\[0\]顕著な成績をあげられたチームには\nSplunkより\[\/0\]\[1\]豪華景品\[\/1\]\[2\]を用意しております。\[\/2\]/
  );
});

test("a wrapped paragraph whose runs share one format stays untagged", async () => {
  const extracted = await extractTextsFromZip(await buildZip());
  assert.match(extracted.text, /uid_0001\n現実に\n近い環境で/);
  assert.doesNotMatch(extracted.text, /\[0\]/);
});

test("the bold phrase keeps its formatting in the translation", async () => {
  const zip = await buildEmphasisZip();
  const extracted = await extractTextsFromZip(zip);
  const translated = formatExtractFile([
    "uid_0001",
    "[0]Teams with outstanding results\nwill receive Splunk's [/0][1]premium prizes[/1][2].[/2]",
  ]);

  const result = await injectTextsToZip(zip, parseUidDelimitedText(translated), extracted.metadata);
  assert.deepEqual(result.flattened, []);

  const para = firstParagraph(await zip.file("ppt/slides/slide1.xml").async("string"));
  const bold = para.match(/<a:r><a:rPr[^>]*b="1"[\s\S]*?<\/a:r>/)[0];
  assert.match(bold, /accent1/);
  assert.match(bold, />premium prizes</);
  const normal = para.match(/<a:r><a:rPr[^>]*b="0"[\s\S]*?<\/a:r>/)[0];
  assert.match(normal, /dk1/);
  assert.match(para, /<a:br>/);
  assert.ok(para.lastIndexOf("</a:r>") < para.indexOf("<a:endParaRPr"), "runs stay in front of endParaRPr");

  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /\[1\]premium prizes\[\/1\]/);
});

test("a translation that dropped the tags is flattened and reported", async () => {
  const zip = await buildEmphasisZip();
  const extracted = await extractTextsFromZip(zip);
  const translated = formatExtractFile(["uid_0001", "Teams with outstanding results\nwill receive premium prizes."]);

  const result = await injectTextsToZip(zip, parseUidDelimitedText(translated), extracted.metadata);
  assert.deepEqual(result.flattened, ["uid_0001"]);

  const para = firstParagraph(await zip.file("ppt/slides/slide1.xml").async("string"));
  assert.doesNotMatch(para, /accent1/);
  assert.match(para, /premium prizes/);
  assert.ok(para.lastIndexOf("</a:r>") < para.indexOf("<a:endParaRPr"));
});

test("text outside every tag is not guessed at", async () => {
  const zip = await buildEmphasisZip();
  const extracted = await extractTextsFromZip(zip);
  const translated = formatExtractFile(["uid_0001", "Intro [0]Teams[/0][1]premium prizes[/1][2].[/2]"]);

  const result = await injectTextsToZip(zip, parseUidDelimitedText(translated), extracted.metadata);
  assert.deepEqual(result.flattened, ["uid_0001"]);
});
