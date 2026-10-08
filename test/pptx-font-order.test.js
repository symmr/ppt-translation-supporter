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

// Runs that already carry some of the font elements, in the combinations that
// real decks have: only cs, ea + cs, and a fill before the fonts.
const RUN_PROPS = [
  `<a:rPr lang="en-US"><a:cs typeface="CiscoSans"/></a:rPr>`,
  `<a:rPr lang="en-US"><a:ea typeface="ＭＳ Ｐゴシック"/><a:cs typeface="CiscoSans"/></a:rPr>`,
  `<a:rPr lang="en-US"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:cs typeface="CiscoSans"/><a:hlinkClick r:id="rId9"/></a:rPr>`,
  `<a:rPr lang="en-US"/>`,
];

async function buildZip() {
  const shapes = RUN_PROPS.map(
    (rPr, i) => `<p:sp>
      <p:nvSpPr><p:cNvPr id="${i + 2}" name="S${i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r>${rPr}<a:t>text ${i}</a:t></a:r></a:p>
      </p:txBody></p:sp>`
  ).join("");
  const zip = new JSZip();
  zip.file("ppt/presentation.xml", `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`);
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}
    </p:spTree></p:cSld></p:sld>`
  );
  return zip;
}

test("font elements end up in schema order (latin, ea, cs) whatever the run had before", async () => {
  const zip = await buildZip();
  const extracted = await extractTextsFromZip(zip);
  const lines = [];
  extracted.metadata.forEach((item, i) => lines.push(item.id, `訳 ${i}`));
  await injectTextsToZip(zip, parseUidDelimitedText(formatExtractFile(lines)), extracted.metadata, {
    bodyFont: "Noto Sans JP",
  });

  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  const runs = [...xml.matchAll(/<a:rPr\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:rPr>)/g)];
  assert.equal(runs.length, RUN_PROPS.length);
  for (const [whole, inner] of runs) {
    const order = [...(inner || "").matchAll(/<a:(\w+)\b/g)].map((m) => m[1]);
    const fonts = order.filter((n) => ["latin", "ea", "cs"].includes(n));
    assert.deepEqual(fonts, ["latin", "ea", "cs"], whole);
    for (const name of fonts) {
      assert.match(inner, new RegExp(`<a:${name}[^>]*typeface="Noto Sans JP"`));
    }
    // fills stay in front of the fonts, a hyperlink stays behind them
    if (order.includes("solidFill")) assert.ok(order.indexOf("solidFill") < order.indexOf("latin"));
    if (order.includes("hlinkClick")) assert.ok(order.indexOf("hlinkClick") > order.indexOf("cs"));
  }
});
