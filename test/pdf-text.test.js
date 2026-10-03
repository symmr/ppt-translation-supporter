"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const PDFLib = require("pdf-lib");

const {
  itemsToLines,
  linesToBlocks,
  extractPdfTexts,
  wrapText,
  fitBlock,
  injectPdfTexts,
  sampleBlockColors,
  buildDocxFromPdf,
} = require("../docs/pdf-text.js");
const { extractDocxTexts } = require("../docs/docx-text.js");
const rewrite = require("../docs/pdf-rewrite.js");
const { collectPageSegments } = rewrite;

async function loadPdfjs() {
  return import("pdfjs-dist/legacy/build/pdf.mjs");
}

async function openPdf(bytes) {
  const pdfjs = await loadPdfjs();
  return pdfjs.getDocument({ data: new Uint8Array(bytes), verbosity: 0 }).promise;
}

async function buildPdf() {
  const doc = await PDFLib.PDFDocument.create();
  const font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  const page = doc.addPage([600, 400]);
  page.drawText("Quarterly Report", { x: 50, y: 350, size: 24, font });
  page.drawText("This is the first line of a paragraph that", { x: 50, y: 300, size: 12, font });
  page.drawText("continues on the second line.", { x: 50, y: 285, size: 12, font });
  page.drawText("Left cell", { x: 50, y: 200, size: 12, font });
  page.drawText("Right cell", { x: 300, y: 200, size: 12, font });
  const second = doc.addPage([600, 400]);
  second.drawText("Second page", { x: 50, y: 350, size: 12, font });
  return doc.save();
}

function item(str, x, y, size, width, extra) {
  return { str, transform: [size, 0, 0, size, x, y], width, height: size, fontName: "f", hasEOL: false, ...extra };
}

test("groups items into lines and splits wide gaps", () => {
  const lines = itemsToLines([
    item("Hello", 10, 100, 10, 25),
    item("world", 37, 100, 10, 25),
    item("Far", 300, 100, 10, 15),
    item("日本", 10, 80, 10, 20),
    item("語", 30, 80, 10, 10),
  ], { f: { ascent: 0.8, descent: -0.2 } });
  assert.deepEqual(lines.map((l) => l.text), ["Hello world", "Far", "日本語"]);
});

test("joins wrapped lines into one block and keeps headings apart", () => {
  const lines = itemsToLines([
    item("Title", 10, 200, 24, 60),
    item("first line of a para-", 10, 150, 12, 120),
    item("graph continues", 10, 136, 12, 90),
  ], {});
  const blocks = linesToBlocks(lines);
  assert.deepEqual(blocks.map((b) => b.text), ["Title", "first line of a paragraph continues"]);
  assert.equal(blocks[1].lineCount, 2);
  assert.equal(blocks[1].leading, 14);
});

test("wraps CJK by character and Latin by word, keeping punctuation off line starts", () => {
  const measure = (s) => s.length * 10;
  assert.deepEqual(wrapText("これは長い文です。", 30, measure), ["これは", "長い文", "です。"]);
  assert.deepEqual(wrapText("alpha beta gamma", 100, measure), ["alpha beta", "gamma"]);
  assert.deepEqual(wrapText("a\nb", 100, measure), ["a", "b"]);
});

test("fitBlock shrinks text until it fits the box", () => {
  const fit = fitBlock("あいうえおかきくけこ", { w: 50, h: 12, size: 10, leading: 0 }, (s, size) => s.length * size);
  assert.ok(fit.size < 10);
  assert.ok(fit.size >= 5);
});

test("extracts blocks per page with positions", async () => {
  const pdf = await openPdf(await buildPdf());
  const extracted = await extractPdfTexts(pdf);
  assert.deepEqual(extracted.lines, [
    "uid_0001", "Quarterly Report",
    "uid_0002", "This is the first line of a paragraph that continues on the second line.",
    "uid_0003", "Left cell",
    "uid_0004", "Right cell",
    "uid_0005", "Second page",
  ]);
  assert.equal(extracted.slideCount, 2);
  assert.equal(extracted.metadata[0].title, true);
  assert.equal(extracted.metadata[1].title, false);
  assert.equal(extracted.metadata[1].lines, 2);
  assert.equal(extracted.metadata[4].page, 1);
  assert.equal(extracted.metadata[2].x, 50);
});

test("overlays translations on the PDF", async () => {
  const bytes = await buildPdf();
  const extracted = await extractPdfTexts(await openPdf(bytes));
  const result = await injectPdfTexts(bytes, {
    uid_0001: "Rapport trimestriel",
    uid_0002: "Ceci est un paragraphe traduit.",
    uid_0003: "Gauche",
    uid_0005: "Deuxieme page",
  }, extracted.metadata, { PDFLib, colors: { uid_0003: { fill: "336699", ink: "FFFFFF" } } });
  assert.equal(result.injected, 4);
  assert.equal(result.missing, 1);
  const again = await extractPdfTexts(await openPdf(result.bytes));
  assert.match(again.text, /Rapport trimestriel/);
  assert.match(again.text, /Gauche/);
  assert.match(again.text, /Right cell/);
  assert.equal((await PDFLib.PDFDocument.load(result.bytes)).getPageCount(), 2);
});

test("builds a Word document from the PDF blocks", async () => {
  const extracted = await extractPdfTexts(await openPdf(await buildPdf()));
  const result = buildDocxFromPdf(JSZip, extracted, {
    uid_0001: "四半期報告",
    uid_0002: "段落 & <訳>\n二行目",
  }, { titleFont: "Meiryo", bodyFont: "Noto Sans JP" });
  assert.equal(result.injected, 2);
  assert.equal(result.missing, 3);
  const zip = await JSZip.loadAsync(await result.zip.generateAsync({ type: "uint8array" }));
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /w:eastAsia="Meiryo"/);
  assert.match(xml, /w:eastAsia="Noto Sans JP"/);
  assert.match(xml, /<w:pageBreakBefore\/>/);
  assert.match(xml, /段落 &amp; &lt;訳&gt;/);
  const docx = await extractDocxTexts(zip);
  assert.deepEqual(docx.lines.filter((_, i) => i % 2 === 1), [
    "四半期報告",
    "段落 & <訳>\n二行目",
    "Left cell",
    "Right cell",
    "Second page",
  ]);
});

test("samples background and text colors around a block", () => {
  const width = 20;
  const height = 10;
  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = (y * width + x) * 4;
      const ink = x >= 8 && x <= 12 && y >= 4 && y <= 6;
      data.set(ink ? [255, 255, 255, 255] : [0, 51, 102, 255], i);
    }
  }
  assert.deepEqual(sampleBlockColors({ data, width, height }, { x0: 5, y0: 3, x1: 15, y1: 7 }), {
    fill: "003366",
    ink: "FFFFFF",
  });
});

test("fitBlock widens toward the page edge before stacking a narrow column", () => {
  const measure = (s, size) => s.length * size;
  const block = { w: 40, h: 12, size: 10, leading: 0 };
  const narrow = fitBlock("あいうえおかきくけこさしすせそ".repeat(2), block, measure);
  const wide = fitBlock("あいうえおかきくけこさしすせそ".repeat(2), block, measure, 400);
  assert.ok(narrow.overflow);
  assert.ok(wide.lines.length < narrow.lines.length);
  assert.equal(wide.width, 400);
});

test("rewrites the page so the original text is gone, keeping the rest in place", async () => {
  const bytes = await buildPdf();
  const extracted = await extractPdfTexts(await openPdf(bytes));
  const result = await injectPdfTexts(bytes, {
    uid_0002: "Paragraphe traduit.",
    uid_0003: "Gauche",
  }, extracted.metadata, { PDFLib });
  assert.deepEqual(result.covered, []);
  const again = await extractPdfTexts(await openPdf(result.bytes));
  assert.doesNotMatch(again.text, /first line of a paragraph/);
  assert.doesNotMatch(again.text, /Left cell/);
  assert.match(again.text, /Paragraphe traduit\./);
  assert.match(again.text, /Gauche/);
  const right = again.metadata.find((m) => m.text === "Right cell");
  assert.equal(right.x, 300);
  assert.equal(right.firstBaseline, 200);
  assert.ok(again.metadata.some((m) => m.text === "Quarterly Report"));
});

test("text color is the most common ink, not a mix of a sentence's colors", () => {
  const width = 30;
  const height = 10;
  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const inside = y >= 3 && y <= 6 && x >= 3 && x <= 26;
      let color = [255, 255, 255];
      if (inside) color = x < 16 ? [20, 20, 20] : x < 21 ? [220, 38, 38] : [22, 163, 74];
      data.set([...color, 255], (y * width + x) * 4);
    }
  }
  assert.deepEqual(sampleBlockColors({ data, width, height }, { x0: 2, y0: 2, x1: 27, y1: 7 }), {
    fill: "FFFFFF",
    ink: "141414",
  });
});

// One line in three colors, a link with an underline, and an image.
async function buildStyledPdf() {
  const doc = await PDFLib.PDFDocument.create();
  const font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  const page = doc.addPage([600, 400]);
  const png = await doc.embedPng(Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg==",
    "base64"
  ));
  page.drawImage(png, { x: 400, y: 50, width: 100, height: 100 });
  const parts = [["Sync is ", [0, 0, 0]], ["twice as fast", [0.86, 0.15, 0.15]], [" now.", [0, 0, 0]]];
  let x = 50;
  for (const [text, [r, g, b]] of parts) {
    page.drawText(text, { x, y: 300, size: 14, font, color: PDFLib.rgb(r, g, b) });
    x += font.widthOfTextAtSize(text, 14);
  }
  page.drawText("Read the ", { x: 50, y: 250, size: 14, font });
  const linkX = 50 + font.widthOfTextAtSize("Read the ", 14);
  const linkW = font.widthOfTextAtSize("manual", 14);
  page.drawText("manual", { x: linkX, y: 250, size: 14, font, color: PDFLib.rgb(0, 0, 0.9) });
  page.drawRectangle({ x: linkX, y: 247, width: linkW, height: 1, color: PDFLib.rgb(0, 0, 0.9) });
  page.drawText(" first.", { x: linkX + linkW, y: 250, size: 14, font });
  const link = doc.context.register(doc.context.obj({
    Type: "Annot",
    Subtype: "Link",
    Rect: [linkX, 246, linkX + linkW, 264],
    Border: [0, 0, 0],
    A: { Type: "Action", S: "URI", URI: PDFLib.PDFString.of("https://example.com/manual") },
  }));
  page.node.set(PDFLib.PDFName.of("Annots"), doc.context.obj([link]));
  return doc.save();
}

async function extractStyled(bytes) {
  const doc = await PDFLib.PDFDocument.load(bytes);
  const pages = doc.getPages();
  return extractPdfTexts(await openPdf(bytes), null, {
    segmentsFor: (index) => collectPageSegments(PDFLib, pages[index]),
  });
}

test("tags color and link runs and keeps them through the rewrite", async () => {
  const bytes = await buildStyledPdf();
  const extracted = await extractStyled(bytes);
  assert.deepEqual(extracted.lines, [
    "uid_0001", "[0]Sync is [/0][1]twice as fast[/1][2] now.[/2]",
    "uid_0002", "[0]Read the [/0][1]manual[/1][2] first.[/2]",
  ]);
  const first = extracted.metadata[0].runs;
  assert.deepEqual(first.map((r) => r.color), ["000000", "DB2626", "000000"]);
  assert.equal(extracted.metadata[1].runs[1].link.url, "https://example.com/manual");

  const result = await injectPdfTexts(bytes, {
    uid_0001: "[0]Sync is now [/0][1]twice as fast[/1][2].[/2]",
    uid_0002: "[0]First read [/0][1]the guide[/1][2].[/2]",
  }, extracted.metadata, { PDFLib });
  assert.deepEqual(result.flattened, []);
  assert.deepEqual(result.covered, []);

  const again = await extractStyled(result.bytes);
  assert.deepEqual(again.lines, [
    "uid_0001", "[0]Sync is now [/0][1]twice as fast[/1][2].[/2]",
    "uid_0002", "[0]First read [/0][1]the guide[/1][2].[/2]",
  ]);
  assert.equal(again.metadata[0].runs[1].color, "DB2626");
  const out = await PDFLib.PDFDocument.load(result.bytes);
  const page = out.getPages()[0];
  const annots = page.node.lookup(PDFLib.PDFName.of("Annots"), PDFLib.PDFArray);
  assert.equal(annots.size(), 1);
  const rect = annots.lookup(0, PDFLib.PDFDict).lookup(PDFLib.PDFName.of("Rect"), PDFLib.PDFArray);
  const x0 = rect.lookup(0).asNumber();
  const helvetica = await out.embedFont(PDFLib.StandardFonts.Helvetica);
  assert.ok(Math.abs(x0 - (50 + helvetica.widthOfTextAtSize("First read ", 14))) < 1);
  // the old underline under "manual" is gone, one new one under "the guide"
  const content = rewrite.pageContentText(PDFLib, page);
  assert.equal((content.match(/(?:^|\s)(?:re|h)\s+f(?=\s|$)/g) || []).length, 1);
  assert.doesNotMatch(content, /45\.766\d* 1 l/);
  assert.equal(out.getPages()[0].node.Resources().lookup(PDFLib.PDFName.of("XObject"), PDFLib.PDFDict).keys().length, 1);
});

test("mismatched tags fall back to one run and are reported", async () => {
  const bytes = await buildStyledPdf();
  const extracted = await extractStyled(bytes);
  const result = await injectPdfTexts(bytes, {
    uid_0001: "[0]Sync is twice as fast now.[/0]",
  }, extracted.metadata, { PDFLib });
  assert.deepEqual(result.flattened, ["uid_0001"]);
  const again = await extractStyled(result.bytes);
  assert.ok(again.lines.includes("Sync is twice as fast now."));
});

test("Word output keeps run colors and hyperlinks", async () => {
  const extracted = await extractStyled(await buildStyledPdf());
  const result = buildDocxFromPdf(JSZip, extracted, {
    uid_0002: "[0]まず[/0][1]ガイド[/1][2]を読む。[/2]",
  }, {});
  const zip = await JSZip.loadAsync(await result.zip.generateAsync({ type: "uint8array" }));
  const xml = await zip.file("word/document.xml").async("string");
  const rels = await zip.file("word/_rels/document.xml.rels").async("string");
  assert.match(xml, /<w:color w:val="DB2626"\/>/);
  assert.match(xml, /<w:hyperlink r:id="rIdLink1">.*ガイド.*<\/w:hyperlink>/);
  assert.match(rels, /Target="https:\/\/example.com\/manual" TargetMode="External"/);
});
