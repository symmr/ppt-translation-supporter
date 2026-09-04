"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const {
  extractTextsFromZip,
  injectTextsToZip,
  parseUidDelimitedText,
  validateTranslations,
  formatExtractFile,
} = require("../docs/pptx-text.js");

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = REL_NS;

function rel(id, type, target) {
  return `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `${rels.join("")}</Relationships>`;
}

function slideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>Hello World</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Runs"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:r><a:rPr lang="en-US"/><a:t>Red</a:t></a:r>
            <a:r><a:rPr lang="en-US"/><a:t>Blue</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="4" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblGrid><a:gridCol w="1000"/><a:gridCol w="1000"/></a:tblGrid>
              <a:tr h="200000">
                <a:tc>
                  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function notesWithBody() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes Placeholder"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Speaker note</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;
}

function notesWithoutBody() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:notes>`;
}

async function buildZip(options = {}) {
  const zip = new JSZip();
  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="${P}" xmlns:r="${R}">` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `</p:presentation>`
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relsXml([rel("rId2", "slide", "slides/slide1.xml")])
  );
  zip.file("ppt/slides/slide1.xml", slideXml());
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relsXml([
      rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
      rel("rId2", "notesSlide", "../notesSlides/notesSlide1.xml"),
    ])
  );
  zip.file(
    "ppt/notesSlides/notesSlide1.xml",
    options.emptyNotes ? notesWithoutBody() : notesWithBody()
  );
  return zip;
}

test("extracts shapes, run tags, table cells, and notes", async () => {
  const extracted = await extractTextsFromZip(await buildZip());
  assert.equal(extracted.uidCount, 5);
  assert.match(extracted.text, /uid_0001\nHello World/);
  assert.match(extracted.text, /uid_0002\n\[0\]Red\[\/0\]\[1\]Blue\[\/1\]/);
  assert.match(extracted.text, /uid_0003\nA1/);
  assert.match(extracted.text, /uid_0004\nA2/);
  assert.match(extracted.text, /uid_0005\nSpeaker note/);
});

test("skips notes slides that have no body text frame", async () => {
  const extracted = await extractTextsFromZip(await buildZip({ emptyNotes: true }));
  assert.equal(extracted.uidCount, 4);
  assert.doesNotMatch(extracted.text, /Speaker note/);
});

test("injects translations back onto the same runs", async () => {
  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  const ja = [
    "uid_0001",
    "こんにちは世界",
    "uid_0002",
    "[0]赤[/0][1]青[/1]",
    "uid_0003",
    "エー1",
    "uid_0004",
    "エー2",
    "uid_0005",
    "スピーカーノート",
  ].join("\n");
  const translations = parseUidDelimitedText(ja);
  const report = validateTranslations(extracted.metadata, translations);
  assert.equal(report.ok, true);

  const result = await injectTextsToZip(zip, translations, extracted.metadata);
  assert.equal(result.injected, 5);
  assert.equal(result.missing, 0);
  assert.deepEqual(result.flattened, []);

  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /uid_0001\nこんにちは世界/);
  assert.match(again.text, /uid_0002\n\[0\]赤\[\/0\]\[1\]青\[\/1\]/);
  assert.match(again.text, /uid_0005\nスピーカーノート/);
});

test("reports paragraphs whose formatting was lost to a tag mismatch", async () => {
  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  // uid_0002 is a two-run paragraph; the translation drops its [n] tags, which
  // is what an LLM commonly does. The text still lands, but both runs collapse
  // into the first one, so the per-run formatting is gone.
  const ja = [
    "uid_0001",
    "こんにちは世界",
    "uid_0002",
    "赤青",
    "uid_0003",
    "エー1",
    "uid_0004",
    "エー2",
    "uid_0005",
    "スピーカーノート",
  ].join("\n");
  const translations = parseUidDelimitedText(ja);
  const report = validateTranslations(extracted.metadata, translations);
  assert.equal(report.ok, true, "uid counts still line up, so uid validation stays silent");

  const result = await injectTextsToZip(zip, translations, extracted.metadata);
  assert.equal(result.injected, 5);
  assert.deepEqual(result.flattened, ["uid_0002"]);

  // the whole translation is crammed into run 0 and run 1 is left empty, so
  // whatever formatting run 1 carried now applies to nothing
  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /uid_0002\n\[0\]赤青\[\/0\]\[1\]\[\/1\]/);
});

test("applies the requested fonts to title and body runs only when asked", async () => {
  const ja = [
    "uid_0001", "こんにちは世界",
    "uid_0002", "[0]赤[/0][1]青[/1]",
    "uid_0003", "エー1",
    "uid_0004", "エー2",
    "uid_0005", "スピーカーノート",
  ].join("\n");

  // no options -> the file keeps whatever fonts it had
  const untouched = await buildZip();
  const meta0 = await extractTextsFromZip(untouched);
  await injectTextsToZip(untouched, parseUidDelimitedText(ja), meta0.metadata);
  const before = await untouched.file("ppt/slides/slide1.xml").async("string");
  assert.doesNotMatch(before, /<a:latin/);

  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  await injectTextsToZip(zip, parseUidDelimitedText(ja), extracted.metadata, {
    titleFont: "Meiryo UI",
    bodyFont: "Yu Gothic",
  });

  const slide = await zip.file("ppt/slides/slide1.xml").async("string");

  // uid_0001 sits in the title placeholder; the run carrying it takes titleFont
  const titleRun = slide.slice(0, slide.indexOf("こんにちは世界"));
  assert.match(titleRun, /<a:latin typeface="Meiryo UI"/);
  assert.match(titleRun, /<a:ea typeface="Meiryo UI"/);
  assert.match(titleRun, /<a:cs typeface="Meiryo UI"/);

  // uid_0002 sits in a plain shape, so it takes bodyFont instead
  const bodyRun = slide.slice(slide.indexOf("こんにちは世界"), slide.indexOf("青") + 10);
  assert.match(bodyRun, /<a:latin typeface="Yu Gothic"/);
  assert.doesNotMatch(bodyRun, /Meiryo UI/);
  // and the title font never leaks into the body
  assert.equal((slide.match(/Meiryo UI/g) || []).length, 3, "title font set on exactly one run");

  // the round trip still reads back cleanly after the font attributes are added
  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /uid_0001\nこんにちは世界/);
  assert.match(again.text, /uid_0002\n\[0\]赤\[\/0\]\[1\]青\[\/1\]/);
});

test("parseUidDelimitedText ignores BOM and keeps blank body lines", () => {
  const parsed = parseUidDelimitedText(formatExtractFile(["uid_0001", "line1", "", "line3"]));
  assert.equal(parsed.uid_0001, "line1\n\nline3");
});
