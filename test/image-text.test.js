"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const {
  extractTextsFromZip,
  injectTextsToZip,
  parseUidDelimitedText,
  listSlideRasterPictures,
  addPictureTextBoxes,
  replacePictureTextBoxes,
} = require("../docs/pptx-text.js");
const { clusterWords, profileFromSlider, segmentToEmu, inheritEdit, resolveOcrScale, suppressOverlaps, textFromRegionWords } = require("../docs/image-text.js");

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${REL_NS}">${rels.join("")}</Relationships>`;
}

function rel(id, type, target) {
  return `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
}

function slideWithPictures() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="5" name="drawing"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="6" name="meta"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId4"/></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm>
      </p:spPr>
    </p:pic>
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="7" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="1000" y="2000"/>
          <a:ext cx="2000" cy="2000"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="1000" cy="1000"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="8" name="grouped"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId5"/></p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="100" y="100"/><a:ext cx="100" cy="100"/></a:xfrm>
        </p:spPr>
      </p:pic>
    </p:grpSp>
  </p:spTree></p:cSld>
</p:sld>`;
}

async function buildZip() {
  const zip = new JSZip();
  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`
  );
  zip.file("ppt/_rels/presentation.xml.rels", relsXml([
    rel("rId2", "slide", "slides/slide1.xml"),
  ]));
  zip.file("ppt/slides/slide1.xml", slideWithPictures());
  zip.file("ppt/slides/_rels/slide1.xml.rels", relsXml([
    rel("rId3", "image", "../media/image1.png"),
    rel("rId4", "image", "../media/image2.emf"),
    rel("rId5", "image", "../media/image3.jpg"),
  ]));
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideWithPictures());
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", relsXml([
    rel("rId3", "image", "../media/layout.png"),
  ]));
  zip.file("ppt/media/image1.png", Buffer.from("png"));
  zip.file("ppt/media/image3.jpg", Buffer.from("jpg"));
  return zip;
}

test("lists slide rasters and skips emf plus layout art", async () => {
  const pictures = await listSlideRasterPictures(await buildZip());
  assert.deepEqual(pictures.map((pic) => pic.media), [
    "ppt/media/image1.png",
    "ppt/media/image3.jpg",
  ]);
  assert.equal(pictures[0].x, 100);
  assert.equal(pictures[0].y, 200);
  assert.equal(pictures[0].cx, 300);
  assert.equal(pictures[0].cy, 400);
  assert.equal(pictures[1].name, "grouped");
  assert.equal(pictures[1].x, 1200);
  assert.equal(pictures[1].y, 2200);
  assert.equal(pictures[1].cx, 200);
  assert.equal(pictures[1].cy, 200);
});

test("text boxes round-trip through the existing extract and inject", async () => {
  const zip = await buildZip();
  const before = await extractTextsFromZip(zip);
  assert.equal(before.uidCount, 1);

  const added = await addPictureTextBoxes(zip, [{
    slidePath: "ppt/slides/slide1.xml",
    x: 100,
    y: 200,
    cx: 300,
    cy: 400,
    text: "A & B <図>",
    fill: "112233",
    ink: "FFFFFF",
  }]);
  assert.equal(added.added, 1);
  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.match(xml, /<a:off x="100" y="200"\/>/);
  assert.match(xml, /lIns="0"/);
  assert.match(xml, /wrap="square"/);
  assert.match(xml, /algn="l"/);
  assert.match(xml, /<a:noAutofit\/>/);

  const extracted = await extractTextsFromZip(zip);
  assert.equal(extracted.uidCount, 2);
  assert.match(extracted.text, /uid_0001\nHello/);
  assert.match(extracted.text, /uid_0002\nA & B <図>/);

  const translations = parseUidDelimitedText([
    "uid_0001",
    "こんにちは",
    "uid_0002",
    "図の訳",
  ].join("\n"));
  const result = await injectTextsToZip(zip, translations, extracted.metadata, {
    bodyFont: "Noto Sans JP",
  });
  assert.equal(result.injected, 2);
  const again = await extractTextsFromZip(zip);
  assert.match(again.text, /uid_0002\n図の訳/);
});

test("replacing an image drops the previous text boxes", async () => {
  const zip = await buildZip();
  const first = await replacePictureTextBoxes(zip, "ppt/slides/slide1.xml", "imgtext-1-1-", [
    { x: 10, y: 20, cx: 100, cy: 40, text: "古い", fill: "FFFFFF", ink: "222222" },
    { x: 10, y: 80, cx: 100, cy: 40, text: "古い2", fill: "FFFFFF", ink: "222222" },
  ]);
  assert.equal(first.added, 2);
  assert.equal(first.removed, 0);

  const second = await replacePictureTextBoxes(zip, "ppt/slides/slide1.xml", "imgtext-1-1-", [
    { x: 10, y: 20, cx: 100, cy: 40, text: "新しい", fill: "FFFFFF", ink: "222222" },
  ]);
  assert.equal(second.removed, 2);
  assert.equal(second.added, 1);

  const extracted = await extractTextsFromZip(zip);
  assert.match(extracted.text, /新しい/);
  assert.doesNotMatch(extracted.text, /古い/);
  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.equal((xml.match(/name="imgtext-1-1-/g) || []).length, 1);
});

test("slider keeps more characters on the aggressive end", () => {
  const words = [
    { text: "組織", confidence: 96, bbox: { x0: 0, y0: 0, x1: 40, y1: 20 } },
    { text: "概要", confidence: 90, bbox: { x0: 42, y0: 0, x1: 80, y1: 20 } },
    { text: "??", confidence: 20, bbox: { x0: 200, y0: 0, x1: 230, y1: 20 } },
  ];
  const loose = clusterWords(words, 1, profileFromSlider(0));
  const tight = clusterWords(words, 1, profileFromSlider(100));
  const looseChars = loose.reduce((sum, seg) => sum + seg.text.length, 0);
  const tightChars = tight.reduce((sum, seg) => sum + seg.text.length, 0);
  assert.ok(looseChars > tightChars);
  assert.ok(tight.every((seg) => seg.confidence >= 70));
});

test("maps a segment onto the picture rectangle", () => {
  const box = segmentToEmu({
    slidePath: "ppt/slides/slide1.xml",
    x: 1000,
    y: 2000,
    cx: 1000,
    cy: 500,
    src: { l: 0, t: 0, visibleW: 1, visibleH: 1 },
  }, { width: 100, height: 50 }, { x0: 10, y0: 5, x1: 30, y1: 15 }, 0);
  assert.equal(box.x, 1100);
  assert.equal(box.y, 2050);
  assert.equal(box.cx, 200);
  assert.equal(box.cy, 100);
});

test("an upscaled bitmap is mapped back onto the original image", () => {
  const words = [{ bbox: { x0: 20, y0: 10, x1: 60, y1: 30 } }];
  assert.equal(resolveOcrScale(words, 100, 50, 2), 1);
  const scaled = [{ bbox: { x0: 40, y0: 20, x1: 190, y1: 90 } }];
  assert.equal(resolveOcrScale(scaled, 100, 50, 2), 2);
});

test("wrapped lines of one block become a single segment", () => {
  const words = [
    { text: "AI", confidence: 93, bbox: { x0: 100, y0: 40, x1: 160, y1: 100 } },
    { text: "observability", confidence: 93, bbox: { x0: 170, y0: 40, x1: 520, y1: 100 } },
    { text: "has", confidence: 93, bbox: { x0: 530, y0: 40, x1: 620, y1: 100 } },
    { text: "a", confidence: 93, bbox: { x0: 630, y0: 40, x1: 670, y1: 100 } },
    { text: "credibility", confidence: 90, bbox: { x0: 180, y0: 108, x1: 420, y1: 168 } },
    { text: "challenge", confidence: 90, bbox: { x0: 430, y0: 108, x1: 640, y1: 168 } },
    { text: "Low", confidence: 88, bbox: { x0: 120, y0: 210, x1: 180, y1: 240 } },
    { text: "accuracy", confidence: 88, bbox: { x0: 190, y0: 210, x1: 320, y1: 240 } },
    { text: "Gartner", confidence: 90, bbox: { x0: 0, y0: 300, x1: 80, y1: 340 } },
    { text: "IDC", confidence: 90, bbox: { x0: 220, y0: 300, x1: 280, y1: 340 } },
  ];
  const segs = clusterWords(words, 1, profileFromSlider(0));
  assert.equal(segs.length, 4);
  assert.equal(segs[0].text, "AI observability has a credibility challenge");
  assert.equal(segs[0].align, "ctr");
  assert.equal(segs[0].lines, 2);
  assert.ok(segs[0].lineH < segs[0].y1 - segs[0].y0);
  assert.equal(segs[1].text, "Low accuracy");
  assert.equal(segs[2].text, "Gartner");
  assert.equal(segs[3].text, "IDC");
});

test("a wrapped block uses the line height and can be centered", async () => {
  const zip = await buildZip();
  await addPictureTextBoxes(zip, [{
    slidePath: "ppt/slides/slide1.xml",
    x: 100,
    y: 200,
    cx: 900,
    cy: 1828800,
    lineCy: 914400,
    lines: 2,
    align: "ctr",
    text: "AI observability has a credibility challenge",
    fill: "111111",
    ink: "FFFFFF",
  }]);
  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.match(xml, /wrap="square"/);
  assert.match(xml, /anchor="t"/);
  assert.match(xml, /algn="ctr"/);
  assert.match(xml, /sz="6624"/);
});

test("a dragged region stays one block and replaces what it covers", () => {
  const words = [
    { text: "AI", confidence: 90, bbox: { x0: 0, y0: 0, x1: 40, y1: 20 } },
    { text: "observability", confidence: 90, bbox: { x0: 48, y0: 0, x1: 180, y1: 20 } },
    { text: "challenge", confidence: 90, bbox: { x0: 10, y0: 28, x1: 120, y1: 48 } },
  ];
  const built = textFromRegionWords(words, 1);
  assert.equal(built.text, "AI observability\nchallenge");
  assert.equal(built.lines, 2);
  const region = { x0: 0, y0: 0, x1: 200, y1: 60, text: built.text, manualId: "r1" };
  const merged = suppressOverlaps(clusterWords(words, 1, profileFromSlider(0)), [region]);
  assert.equal(merged.length, 1);
  assert.equal(merged[0].manualId, "r1");
});

test("a region keeps its line breaks as one extract entry", async () => {
  const zip = await buildZip();
  await addPictureTextBoxes(zip, [{
    slidePath: "ppt/slides/slide1.xml",
    x: 100,
    y: 200,
    cx: 400,
    cy: 200,
    text: "AI observability has a\ncredibility challenge",
    fill: "111111",
    ink: "FFFFFF",
  }]);
  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.match(xml, /<a:br\b/);
  const extracted = await extractTextsFromZip(zip);
  assert.match(extracted.text, /AI observability has a\ncredibility challenge/);
  const added = extracted.metadata.filter((item) => item.type === "shape" && item.slidePath === "ppt/slides/slide1.xml");
  assert.equal(added.length, 2);
});

test("slider movement preserves manually edited text on the same region", () => {
  const priors = [
    { x0: 0, y0: 0, x1: 40, y1: 20, edited: true, text: "組織（修正）" },
  ];
  const same = inheritEdit(priors, { x0: 2, y0: 1, x1: 38, y1: 18, text: "組織概要" });
  assert.equal(same.text, "組織（修正）");
  assert.equal(same.edited, true);
  const fresh = inheritEdit(priors, { x0: 80, y0: 40, x1: 120, y1: 60, text: "新規" });
  assert.equal(fresh.text, "新規");
  assert.equal(fresh.edited, false);
});
