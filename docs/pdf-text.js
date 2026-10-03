// PDF text extract and two kinds of write-back: cover each text block on the
// original pages with its translation (pdf-lib), or lay the translations out
// as a new Word document. pdf.js, pdf-lib and fontkit are passed in, so the
// browser can load them from the CDN only when a PDF is dropped.
(function () {
"use strict";

// Static TTF: fontkit subsets TrueType outlines reliably, CFF/OTF less so.
const PDF_FONT_URL = "https://cdn.jsdelivr.net/npm/@expo-google-fonts/noto-sans-jp@0.4.1/400Regular/NotoSansJP_400Regular.ttf";
const PDF_FONT_NAME = "Noto Sans JP";

const rewrite = typeof module !== "undefined" && module.exports
  ? require("./pdf-rewrite.js")
  : window;

const TAG_LIKE_RE = /(?:\[\/?\d+\]|⟦\/?\d+⟧)/;
// CJK, kana, full-width forms: no spaces between these when joining pieces.
const CJK_RE = /[⺀-鿿가-힯豈-﫿＀-￯]/;
// Never start a wrapped line with these; they stay on the previous line.
const NO_LINE_START_RE = /[、。，．,.:;!?！？）」』】〕〉》”’ー々ぁぃぅぇぉっゃゅょァィゥェォッャュョ・：；\])}]/;
const NO_LINE_END_RE = /[（「『【〔〈《“‘([{]/;
const MIN_SHRINK = 0.5;

function isCjk(ch) {
  return CJK_RE.test(ch || "");
}

function joinPieces(left, right, spaced) {
  if (!left) return right;
  if (!right) return left;
  if (!spaced || /\s$/.test(left) || /^\s/.test(right)) return left + right;
  if (isCjk(left.slice(-1)) || isCjk(right[0])) return left + right;
  return `${left} ${right}`;
}

function joinLines(prev, next) {
  if (/[A-Za-z]-$/.test(prev) && /^[a-z]/.test(next)) return prev.slice(0, -1) + next;
  if (isCjk(prev.slice(-1)) || isCjk(next[0])) return prev + next;
  return `${prev} ${next}`;
}

function horizontal(transform) {
  const [a, b, c, d] = transform;
  return a > 0 && d > 0 && Math.abs(b) <= a * 0.02 && Math.abs(c) <= d * 0.02;
}

// pdf.js text items to lines: items that sit on one baseline and follow each
// other closely. A wide gap (table columns, tab stops) starts a new line so
// the cells stay separate blocks.
function itemsToLines(items, styles) {
  const lines = [];
  let line = null;
  let pendingSpace = false;
  const close = () => {
    if (line && line.text.trim()) {
      line.text = line.text.replace(/\s+/g, " ").trim();
      lines.push(line);
    }
    line = null;
    pendingSpace = false;
  };
  for (const item of items || []) {
    const transform = item.transform || [];
    if (transform.length < 6 || !horizontal(transform)) {
      if (item.str && item.str.trim()) close();
      continue;
    }
    const size = item.height > 0 ? item.height : Math.abs(transform[3]);
    const x = transform[4];
    const y = transform[5];
    const str = String(item.str || "");
    if (!str.trim()) {
      if (line && item.width > 1.5 * line.size) close();
      else if (line && str) pendingSpace = true;
      if (item.hasEOL) close();
      continue;
    }
    if (!(size > 0)) continue;
    const style = (styles && styles[item.fontName]) || {};
    if (line) {
      const sameBaseline = Math.abs(y - line.y) < 0.35 * Math.max(size, line.size);
      const gap = x - line.x1;
      const near = gap > -0.5 * line.size && gap < 1.5 * Math.max(size, line.size);
      if (!sameBaseline || !near) close();
    }
    if (!line) {
      line = {
        x0: x,
        x1: x + (item.width || 0),
        y,
        size,
        ascent: typeof style.ascent === "number" && style.ascent > 0 ? style.ascent : 0.8,
        descent: typeof style.descent === "number" ? Math.abs(style.descent) : 0.2,
        text: str,
      };
    } else {
      const spaced = pendingSpace || x - line.x1 > 0.18 * line.size;
      line.text = joinPieces(line.text, str, spaced);
      line.x1 = Math.max(line.x1, x + (item.width || 0));
      line.size = Math.max(line.size, size);
    }
    pendingSpace = false;
    if (item.hasEOL) close();
  }
  close();
  return lines;
}

// Lines to paragraph-like blocks: next line below, similar size, overlapping
// horizontally, and (from the third line on) the same line pitch.
function linesToBlocks(lines) {
  const blocks = [];
  let block = null;
  for (const line of lines) {
    if (block) {
      const ratio = line.size / block.size;
      const dy = block.lastY - line.y;
      const overlaps = line.x0 < block.x1 && line.x1 > block.x0;
      const indentOk = line.x0 > block.x0 - line.size;
      const pitchOk = !block.leading || Math.abs(dy - block.leading) < 0.35 * line.size;
      const joins = ratio > 0.8 && ratio < 1.25 &&
        dy > 0.6 * block.size && dy < 1.9 * block.size &&
        overlaps && indentOk && pitchOk;
      if (joins) {
        block.leading = block.leading || dy;
        block.text = joinLines(block.text, line.text);
        block.x0 = Math.min(block.x0, line.x0);
        block.x1 = Math.max(block.x1, line.x1);
        block.bottom = Math.min(block.bottom, line.y - line.descent * line.size);
        block.lastY = line.y;
        block.lineCount += 1;
        continue;
      }
      blocks.push(block);
    }
    block = {
      x0: line.x0,
      x1: line.x1,
      top: line.y + line.ascent * line.size,
      bottom: line.y - line.descent * line.size,
      firstY: line.y,
      lastY: line.y,
      size: line.size,
      leading: 0,
      lineCount: 1,
      text: line.text,
    };
  }
  if (block) blocks.push(block);
  return blocks;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

// pdf: a pdf.js PDFDocumentProxy.
async function extractPdfTexts(pdf, onPage) {
  const texts = [];
  const metadata = [];
  const pages = [];
  let uidNum = 1;
  for (let index = 0; index < pdf.numPages; index += 1) {
    if (onPage) onPage(index + 1, pdf.numPages);
    const page = await pdf.getPage(index + 1);
    const content = await page.getTextContent();
    const [vx0, vy0, vx1, vy1] = page.view;
    pages.push({ width: vx1 - vx0, height: vy1 - vy0 });
    for (const block of linesToBlocks(itemsToLines(content.items, content.styles))) {
      const uid = `uid_${String(uidNum).padStart(4, "0")}`;
      uidNum += 1;
      texts.push(uid);
      texts.push(block.text);
      metadata.push({
        id: uid,
        type: "pdf",
        page: index,
        x: round2(block.x0),
        y: round2(block.bottom),
        w: round2(block.x1 - block.x0),
        h: round2(block.top - block.bottom),
        size: round2(block.size),
        leading: round2(block.leading),
        lines: block.lineCount,
        firstBaseline: round2(block.firstY),
        text: block.text,
        title: false,
      });
    }
    if (page.cleanup) page.cleanup();
  }
  // Titles for the Word output: clearly larger than the usual body size.
  const body = median(metadata.map((m) => m.size));
  for (const m of metadata) m.title = body > 0 && m.size >= body * 1.3;
  return {
    lines: texts,
    text: texts.join("\n"),
    metadata,
    pages,
    slideCount: pdf.numPages,
    uidCount: metadata.length,
    unitLabel: `${pdf.numPages} ページ`,
  };
}

// Break into wrap units: one CJK character, or a Latin word with the spaces
// after it. Closing punctuation sticks to the unit before it.
function wrapUnits(text) {
  const units = [];
  const re = /[A-Za-z0-9À-ɏ'’\-_.,:;/@#%&+=*!?()]+\s*|\s+|[\s\S]/g;
  let match;
  while ((match = re.exec(text)) !== null) {
    const unit = match[0];
    const prev = units[units.length - 1];
    if (prev && (NO_LINE_START_RE.test(unit[0]) || NO_LINE_END_RE.test(prev.slice(-1)))) {
      units[units.length - 1] = prev + unit;
    } else {
      units.push(unit);
    }
  }
  return units;
}

function wrapText(text, maxWidth, measure) {
  const out = [];
  for (const para of String(text).split(/\r?\n/)) {
    let line = "";
    for (const unit of wrapUnits(para)) {
      const candidate = line + unit;
      if (!line || measure(candidate.trimEnd()) <= maxWidth) {
        line = candidate;
        continue;
      }
      out.push(line.trimEnd());
      line = unit.trimStart();
      // a single unit wider than the box: split it by character
      while (line.length > 1 && measure(line) > maxWidth) {
        let cut = line.length - 1;
        while (cut > 1 && measure(line.slice(0, cut)) > maxWidth) cut -= 1;
        out.push(line.slice(0, cut));
        line = line.slice(cut);
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

function fitWithin(text, block, measure, width) {
  const lineHeight = (size) => (block.leading > 0 ? block.leading * (size / block.size) : size * 1.2);
  const limit = block.h * 1.05 + 0.5;
  let size = block.size;
  let lines = wrapText(text, width, (s) => measure(s, size));
  for (let step = 1; step <= 20; step += 1) {
    if (size + (lines.length - 1) * lineHeight(size) <= limit) break;
    const next = block.size * (1 - step * ((1 - MIN_SHRINK) / 20));
    if (next < block.size * MIN_SHRINK - 1e-9) break;
    size = next;
    lines = wrapText(text, width, (s) => measure(s, size));
  }
  const overflow = size + (lines.length - 1) * lineHeight(size) > limit;
  return { size, lineHeight: lineHeight(size), lines, width, overflow };
}

// Largest size (down to half the original) at which the wrapped text fits
// the block's box. When even that overflows, the text may run to maxWidth
// (the page's right margin) rather than stack up in a narrow column.
// measure(text, size) returns the width in points.
function fitBlock(text, block, measure, maxWidth) {
  const fit = fitWithin(text, block, measure, Math.max(block.w, block.size));
  if (!fit.overflow || !(maxWidth > fit.width)) return fit;
  const wide = fitWithin(text, block, measure, maxWidth);
  return wide.lines.length < fit.lines.length || !wide.overflow ? wide : fit;
}

// @pdf-lib/fontkit writes a short-format loca table for small subsets but
// does not pad glyphs to an even length, so every glyph after an odd-length
// one points at the wrong bytes and renders blank. Padding fixes the offsets.
function patchFontkitSubset(fontkit, fontBytes) {
  const subset = fontkit.create(fontBytes).createSubset();
  const proto = Object.getPrototypeOf(subset);
  if (!proto || typeof proto._addGlyph !== "function" || proto._addGlyphEvenPadded) return;
  const addGlyph = proto._addGlyph;
  proto._addGlyph = function addEvenGlyph(gid) {
    const index = addGlyph.call(this, gid);
    const buf = this.glyf && this.glyf[index];
    if (buf && buf.length % 2) {
      const Ctor = buf.constructor;
      const padded = Ctor && typeof Ctor.alloc === "function" ? Ctor.alloc(buf.length + 1) : new Uint8Array(buf.length + 1);
      padded.set(buf);
      this.glyf[index] = padded;
      this.offset += 1;
    }
    return index;
  };
  proto._addGlyphEvenPadded = true;
}

function hexToRgb(PDFLib, hex, fallback) {
  const value = /^[0-9a-f]{6}$/i.test(hex || "") ? hex : fallback;
  const n = parseInt(value, 16);
  return PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
}

// Removes the original text of each translated block from the page content
// and draws the translation in its place. Blocks whose text cannot be removed
// that way (form XObjects, invisible OCR text) are covered with their
// background color instead, and their original text stays in the file.
// deps: { PDFLib, fontkit, fontBytes, colors: { uid: { fill, ink } } }
async function injectPdfTexts(bytes, translations, metadata, deps) {
  const { PDFLib } = deps;
  let doc;
  try {
    doc = await PDFLib.PDFDocument.load(bytes, { updateMetadata: false });
  } catch (err) {
    if (err && /encrypt/i.test(err.message || err.name || "")) {
      throw new Error("暗号化された PDF には上書きできません。Word 出力を使ってください", { cause: err });
    }
    throw err;
  }
  let font;
  if (deps.fontBytes) {
    patchFontkitSubset(deps.fontkit, deps.fontBytes);
    doc.registerFontkit(deps.fontkit);
    font = await doc.embedFont(deps.fontBytes, { subset: true });
  } else {
    font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  }
  const pages = doc.getPages();
  const colors = deps.colors || {};
  let injected = 0;
  let missing = 0;
  const tagArtifacts = [];
  const overflowed = [];
  const targets = [];
  for (const meta of metadata || []) {
    if (translations[meta.id] === undefined || !pages[meta.page]) missing += 1;
    else targets.push(meta);
  }
  // Take the original text out of each page first; pdf-lib's own drawing is
  // appended afterwards as a separate content stream.
  const removed = new Set();
  const byPage = new Map();
  for (const meta of targets) {
    if (!byPage.has(meta.page)) byPage.set(meta.page, []);
    byPage.get(meta.page).push(meta);
  }
  for (const [index, metas] of byPage) {
    for (const id of rewrite.removePageText(PDFLib, pages[index], metas)) removed.add(id);
  }
  const covered = [];
  for (const meta of targets) {
    const page = pages[meta.page];
    const text = String(translations[meta.id]).replace(/\s+$/, "");
    if (TAG_LIKE_RE.test(text)) tagArtifacts.push(meta.id);
    const color = colors[meta.id] || {};
    if (!removed.has(meta.id)) {
      // text in a form XObject or an invisible OCR layer: cover it instead
      covered.push(meta.id);
      const pad = Math.max(1, meta.size * 0.08);
      page.drawRectangle({
        x: meta.x - pad,
        y: meta.y - pad,
        width: meta.w + pad * 2,
        height: meta.h + pad * 2,
        color: hexToRgb(PDFLib, color.fill, "FFFFFF"),
      });
    }
    const box = page.getMediaBox();
    const maxWidth = box.x + box.width - meta.x - Math.min(36, box.width * 0.05);
    const fit = fitBlock(text, meta, (s, size) => font.widthOfTextAtSize(s, size), maxWidth);
    if (fit.overflow) overflowed.push(meta.id);
    const ink = hexToRgb(PDFLib, color.ink, "1E1E1E");
    // first baseline where the original first line sat, scaled with the size
    let baseline = meta.firstBaseline
      ? meta.y + meta.h - (meta.y + meta.h - meta.firstBaseline) * (fit.size / meta.size)
      : meta.y + meta.h - fit.size * 0.88;
    for (const line of fit.lines) {
      if (line) page.drawText(line, { x: meta.x, y: baseline, size: fit.size, font, color: ink });
      baseline -= fit.lineHeight;
    }
    injected += 1;
  }
  const out = await doc.save();
  return { bytes: out, injected, missing, flattened: [], tagArtifacts, overflowed, covered };
}

// The most common text color. A per-channel median would mix a sentence's
// black, red and green words into a color none of them has.
function dominantColor(pixels, channel) {
  const bins = new Map();
  for (const p of pixels) {
    const key = (p[0] >> 5) * 64 + (p[1] >> 5) * 8 + (p[2] >> 5);
    if (!bins.has(key)) bins.set(key, []);
    bins.get(key).push(p);
  }
  let best = [];
  for (const members of bins.values()) if (members.length > best.length) best = members;
  return [0, 1, 2].map((i) => channel(best, i));
}

// Background and text colors of each block, from a rendered page.
// image: { data, width, height } (RGBA), rect: pixel box of the block.
function sampleBlockColors(image, rect) {
  const { data, width, height } = image;
  const x0 = Math.max(0, Math.floor(rect.x0) - 2);
  const y0 = Math.max(0, Math.floor(rect.y0) - 2);
  const x1 = Math.min(width - 1, Math.ceil(rect.x1) + 2);
  const y1 = Math.min(height - 1, Math.ceil(rect.y1) + 2);
  if (x1 <= x0 || y1 <= y0) return null;
  const at = (x, y) => {
    const i = (y * width + x) * 4;
    return [data[i], data[i + 1], data[i + 2]];
  };
  const ring = [];
  for (let x = x0; x <= x1; x += 1) ring.push(at(x, y0), at(x, y1));
  for (let y = y0; y <= y1; y += 1) ring.push(at(x0, y), at(x1, y));
  const channel = (pixels, i) => median(pixels.map((p) => p[i]));
  const fill = [0, 1, 2].map((i) => channel(ring, i));
  const inkPixels = [];
  for (let y = y0 + 1; y < y1; y += 1) {
    for (let x = x0 + 1; x < x1; x += 1) {
      const p = at(x, y);
      if (Math.hypot(p[0] - fill[0], p[1] - fill[1], p[2] - fill[2]) > 60) inkPixels.push(p);
    }
  }
  const lum = 0.2126 * fill[0] + 0.7152 * fill[1] + 0.0722 * fill[2];
  const ink = inkPixels.length < 4
    ? (lum < 140 ? [255, 255, 255] : [30, 30, 30])
    : dominantColor(inkPixels, channel);
  const hex = (p) => p.map((n) => Math.round(n).toString(16).padStart(2, "0")).join("").toUpperCase();
  return { fill: hex(fill), ink: hex(ink) };
}

// Browser only: renders each page that has blocks and samples their colors.
async function samplePdfColors(pdf, metadata, onPage) {
  const byPage = new Map();
  for (const meta of metadata) {
    if (!byPage.has(meta.page)) byPage.set(meta.page, []);
    byPage.get(meta.page).push(meta);
  }
  const colors = {};
  let done = 0;
  for (const [index, metas] of byPage) {
    done += 1;
    if (onPage) onPage(done, byPage.size);
    const page = await pdf.getPage(index + 1);
    const viewport = page.getViewport({ scale: 1.5 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
    for (const meta of metas) {
      const [ax, ay, bx, by] = viewport.convertToViewportRectangle([meta.x, meta.y, meta.x + meta.w, meta.y + meta.h]);
      const sampled = sampleBlockColors(image, {
        x0: Math.min(ax, bx),
        y0: Math.min(ay, by),
        x1: Math.max(ax, bx),
        y1: Math.max(ay, by),
      });
      if (sampled) colors[meta.id] = sampled;
    }
    canvas.width = 0;
    canvas.height = 0;
    if (page.cleanup) page.cleanup();
  }
  return colors;
}

function escapeXml(text) {
  return String(text)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runXml(text, rPr) {
  const parts = String(text).split(/\r?\n/);
  const body = parts
    .map((part, i) => `${i ? "<w:br/>" : ""}<w:t xml:space="preserve">${escapeXml(part)}</w:t>`)
    .join("");
  return `<w:r>${rPr}${body}</w:r>`;
}

function fontXml(font) {
  if (!font) return "";
  const f = escapeXml(font);
  return `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`;
}

// One paragraph per block, in page order, a page break between pages.
// Untranslated blocks keep the original text. Returns a JSZip.
function buildDocxFromPdf(JSZipCtor, extracted, translations, options) {
  const opts = options || {};
  const metadata = extracted.metadata || [];
  const firstPage = (extracted.pages && extracted.pages[0]) || { width: 595, height: 842 };
  const paras = [];
  let injected = 0;
  let missing = 0;
  const tagArtifacts = [];
  let lastPage = metadata.length ? metadata[0].page : 0;
  for (const meta of metadata) {
    const translated = translations[meta.id];
    let text = meta.text || "";
    if (translated === undefined) {
      missing += 1;
    } else {
      text = String(translated).replace(/\s+$/, "");
      if (TAG_LIKE_RE.test(text)) tagArtifacts.push(meta.id);
      injected += 1;
    }
    const pageBreak = meta.page !== lastPage;
    lastPage = meta.page;
    const halfPoints = Math.max(16, Math.min(80, Math.round(meta.size * 2)));
    const font = meta.title ? opts.titleFont : opts.bodyFont;
    const rPr = `<w:rPr>${fontXml(font)}${meta.title ? "<w:b/>" : ""}<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr>`;
    const pPr = `<w:pPr>${pageBreak ? "<w:pageBreakBefore/>" : ""}${meta.title ? '<w:outlineLvl w:val="0"/>' : ""}<w:spacing w:after="120"/></w:pPr>`;
    paras.push(`<w:p>${pPr}${runXml(text, rPr)}</w:p>`);
  }
  if (!paras.length) paras.push("<w:p/>");
  const twips = (pt) => Math.round(pt * 20);
  const sect = `<w:sectPr><w:pgSz w:w="${twips(firstPage.width)}" w:h="${twips(firstPage.height)}"${firstPage.width > firstPage.height ? ' w:orient="landscape"' : ""}/>` +
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>';
  const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${paras.join("")}${sect}</w:body></w:document>`;
  const zip = new JSZipCtor();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>");
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>");
  zip.file("word/document.xml", documentXml);
  return { zip, injected, missing, flattened: [], tagArtifacts };
}

const api = {
  PDF_FONT_URL,
  PDF_FONT_NAME,
  itemsToLines,
  linesToBlocks,
  extractPdfTexts,
  wrapText,
  fitBlock,
  injectPdfTexts,
  sampleBlockColors,
  samplePdfColors,
  buildDocxFromPdf,
};

if (typeof module !== "undefined" && module.exports) module.exports = api;
if (typeof window !== "undefined") Object.assign(window, api);
})();
