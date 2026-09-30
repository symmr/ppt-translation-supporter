// DOCX text extract / inject, and text boxes placed inside a picture's own
// coordinate space so they stay with the image. Body paragraphs, table cells
// and text boxes are in scope. Headers, footers and notes are not.
(function () {
"use strict";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_XML = "http://www.w3.org/XML/1998/namespace";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const GROUP_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const PART = "word/document.xml";
const RASTER_RE = /\.(png|jpe?g|webp)$/i;
const MARK_RE = /^imgtext-\d+-\d+-\d+$/;

function getDOMParser() {
  if (typeof DOMParser !== "undefined") return DOMParser;
  return require("@xmldom/xmldom").DOMParser;
}

function getXMLSerializer() {
  if (typeof XMLSerializer !== "undefined") return XMLSerializer;
  return require("@xmldom/xmldom").XMLSerializer;
}

function parseXml(xml) {
  if (/<!DOCTYPE/i.test(xml)) throw new Error("DOCTYPE を含む XML は扱えません");
  const Parser = getDOMParser();
  const doc = new Parser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (!root || root.localName === "parsererror" || root.nodeName === "parsererror") {
    throw new Error("XML の解析に失敗しました");
  }
  return doc;
}

function serializeXml(doc) {
  const Serializer = getXMLSerializer();
  let out = new Serializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) {
    out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
  }
  return out;
}

function elementChildren(el) {
  const out = [];
  if (!el || !el.childNodes) return out;
  for (let i = 0; i < el.childNodes.length; i += 1) {
    const child = el.childNodes[i];
    if (child.nodeType === 1) out.push(child);
  }
  return out;
}

function child(el, name) {
  return elementChildren(el).find((item) => item.localName === name) || null;
}

function attr(el, name) {
  if (!el) return "";
  return el.getAttributeNS(NS_W, name) || el.getAttribute(`w:${name}`) || el.getAttribute(name) || "";
}

function wSet(el, name, value) {
  el.setAttributeNS(NS_W, `w:${name}`, value);
}

function insideFallback(el) {
  let node = el;
  while (node) {
    if (node.localName === "Fallback") return true;
    node = node.parentNode;
  }
  return false;
}

function paragraphRuns(p) {
  const runs = [];
  for (const node of elementChildren(p)) {
    const host = node.localName === "hyperlink" || node.localName === "ins" || node.localName === "smartTag"
      ? elementChildren(node)
      : [node];
    for (const item of host) {
      if (item.localName === "r") runs.push(item);
    }
  }
  return runs;
}

function runHasDrawing(run) {
  return elementChildren(run).some((item) => item.localName === "drawing" || item.localName === "pict" || item.localName === "object");
}

function runText(run) {
  let text = "";
  for (const item of elementChildren(run)) {
    if (item.localName === "t") text += item.textContent || "";
    else if (item.localName === "tab") text += "\t";
    else if (item.localName === "br" || item.localName === "cr") text += "\n";
  }
  return text;
}

function textRuns(p) {
  return paragraphRuns(p).filter((run) => runText(run) !== "" || !runHasDrawing(run));
}

function paragraphPlain(p) {
  return textRuns(p).map(runText).join("");
}

function taggedText(p) {
  const runs = textRuns(p).filter((run) => !runHasDrawing(run));
  if (runs.length <= 1) return paragraphPlain(p);
  return runs.map((run, index) => `[${index}]${runText(run)}[/${index}]`).join("");
}

function isMarker(text) {
  return MARK_RE.test(String(text || "").trim());
}

function styleOf(p) {
  const pPr = child(p, "pPr");
  const pStyle = pPr && child(pPr, "pStyle");
  return attr(pStyle, "val");
}

function isTitleStyle(style) {
  return /^(Title|Subtitle|Heading\d+)$/i.test(style || "");
}

function blockChildren(el) {
  return elementChildren(el).filter((item) => item.localName === "p" || item.localName === "tbl");
}

function txbxOf(p) {
  const found = [];
  const walk = (el, inBox) => {
    for (const item of elementChildren(el)) {
      if (item.localName === "Fallback") continue;
      if (item.localName === "txbxContent") {
        if (!inBox) found.push(item);
        walk(item, true);
        continue;
      }
      walk(item, inBox);
    }
  };
  walk(p, false);
  return found;
}

function isVMergeContinue(tc) {
  const tcPr = child(tc, "tcPr");
  const merge = tcPr && child(tcPr, "vMerge");
  if (!merge) return false;
  return attr(merge, "val") !== "restart";
}

function pushParagraph(p, loc, texts, metadata, uidRef) {
  const plain = paragraphPlain(p);
  if (!plain.trim() || isMarker(plain)) return;
  const uid = `uid_${String(uidRef[0]).padStart(4, "0")}`;
  uidRef[0] += 1;
  const text = taggedText(p);
  texts.push(uid);
  texts.push(text);
  metadata.push({
    id: uid,
    type: loc.type,
    part: PART,
    path: loc.path,
    title: Boolean(loc.title),
  });
}

function walkBlocks(container, path, type, texts, metadata, uidRef) {
  blockChildren(container).forEach((block, index) => {
    if (block.localName === "tbl") {
      walkTable(block, path.concat(["tbl", index]), texts, metadata, uidRef);
      return;
    }
    const pPath = path.concat(["p", index]);
    pushParagraph(block, {
      type,
      path: pPath,
      title: type === "body" && isTitleStyle(styleOf(block)),
    }, texts, metadata, uidRef);
    txbxOf(block).forEach((box, boxIndex) => {
      walkBlocks(box, pPath.concat(["tb", boxIndex]), "textbox", texts, metadata, uidRef);
    });
  });
}

function walkTable(tbl, path, texts, metadata, uidRef) {
  elementChildren(tbl).filter((item) => item.localName === "tr").forEach((tr, row) => {
    elementChildren(tr).filter((item) => item.localName === "tc").forEach((tc, col) => {
      if (isVMergeContinue(tc)) return;
      walkBlocks(tc, path.concat(["tr", row, "tc", col]), "table", texts, metadata, uidRef);
    });
  });
}

async function extractDocxTexts(zip) {
  const entry = zip.file(PART);
  if (!entry) throw new Error("word/document.xml がありません");
  const doc = parseXml(await entry.async("string"));
  const body = doc.getElementsByTagNameNS(NS_W, "body")[0];
  const texts = [];
  const metadata = [];
  const uidRef = [1];
  if (body) walkBlocks(body, [], "body", texts, metadata, uidRef);
  return {
    lines: texts,
    text: texts.join("\n"),
    metadata,
    slideCount: 1,
    uidCount: metadata.length,
    unitLabel: "本文",
  };
}

function locate(doc, path) {
  let el = doc.getElementsByTagNameNS(NS_W, "body")[0];
  for (let i = 0; i < path.length; i += 2) {
    const kind = path[i];
    const index = path[i + 1];
    if (!el) return null;
    if (kind === "p" || kind === "tbl") el = blockChildren(el)[index];
    else if (kind === "tr") el = elementChildren(el).filter((item) => item.localName === "tr")[index];
    else if (kind === "tc") el = elementChildren(el).filter((item) => item.localName === "tc")[index];
    else if (kind === "tb") el = txbxOf(el)[index];
    else return null;
  }
  return el && el.localName === "p" ? el : null;
}

function clearRunText(run) {
  for (const item of [...elementChildren(run)]) {
    if (item.localName === "t" || item.localName === "br" || item.localName === "cr" || item.localName === "tab") {
      run.removeChild(item);
    }
  }
}

function writeRunText(run, text) {
  clearRunText(run);
  const doc = run.ownerDocument;
  String(text).split("\n").forEach((line, index) => {
    if (index) run.appendChild(doc.createElementNS(NS_W, "w:br"));
    const node = doc.createElementNS(NS_W, "w:t");
    if (/^\s|\s$|\s{2}/.test(line)) node.setAttributeNS(NS_XML, "xml:space", "preserve");
    node.textContent = line;
    run.appendChild(node);
  });
}

function ensureRPr(run) {
  let rPr = child(run, "rPr");
  if (rPr) return rPr;
  rPr = run.ownerDocument.createElementNS(NS_W, "w:rPr");
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function setDocxFont(run, font) {
  if (!font) return;
  const rPr = ensureRPr(run);
  let fonts = child(rPr, "rFonts");
  if (!fonts) {
    fonts = run.ownerDocument.createElementNS(NS_W, "w:rFonts");
    rPr.appendChild(fonts);
  }
  for (const name of ["ascii", "hAnsi", "eastAsia", "cs"]) wSet(fonts, name, font);
}

function setDocxParagraphText(p, translated, font) {
  const runs = textRuns(p).filter((run) => !runHasDrawing(run));
  const text = String(translated);
  const matches = [];
  const re = /\[(\d+)\]([\s\S]*?)\[\/\1\]/g;
  let found;
  while ((found = re.exec(text)) !== null) matches.push([Number(found[1]), found[2]]);
  const indexed = matches.length === runs.length
    && new Set(matches.map(([index]) => index)).size === runs.length
    && matches.every(([index]) => index >= 0 && index < runs.length);
  if (runs.length > 1 && indexed) {
    for (const [index, value] of matches) {
      writeRunText(runs[index], value);
      setDocxFont(runs[index], font);
    }
    return { mode: "runs" };
  }
  if (!runs.length) return { mode: "missing" };
  const plain = runs.length > 1 ? text.replace(/(?:\[\/?\d+\])/g, "") : text;
  writeRunText(runs[0], plain);
  setDocxFont(runs[0], font);
  for (let i = 1; i < runs.length; i += 1) {
    writeRunText(runs[i], "");
    setDocxFont(runs[i], font);
  }
  return { mode: runs.length > 1 ? "flattened" : "plain" };
}

async function injectDocxTexts(zip, translations, metadata, options) {
  const titleFont = (options && options.titleFont) || "";
  const bodyFont = (options && options.bodyFont) || "";
  const entry = zip.file(PART);
  if (!entry) return { injected: 0, missing: (metadata || []).length, flattened: [], tagArtifacts: [] };
  const doc = parseXml(await entry.async("string"));
  let injected = 0;
  let missing = 0;
  const flattened = [];
  for (const loc of metadata || []) {
    if (!Object.prototype.hasOwnProperty.call(translations, loc.id)) {
      missing += 1;
      continue;
    }
    const paragraph = locate(doc, loc.path || []);
    if (!paragraph) {
      missing += 1;
      continue;
    }
    const outcome = setDocxParagraphText(paragraph, translations[loc.id], loc.title ? titleFont : bodyFont);
    if (outcome.mode === "missing") {
      missing += 1;
      continue;
    }
    if (outcome.mode === "flattened") flattened.push(loc.id);
    injected += 1;
  }
  zip.file(PART, serializeXml(doc));
  return { injected, missing, flattened, tagArtifacts: [] };
}

function resolveTarget(target) {
  const clean = String(target || "").replace(/\\/g, "/");
  if (clean.startsWith("/")) return clean.replace(/^\/+/, "");
  const parts = ["word"];
  for (const piece of clean.split("/")) {
    if (!piece || piece === ".") continue;
    if (piece === "..") parts.pop();
    else parts.push(piece);
  }
  return parts.join("/");
}

async function listDocxRasterPictures(zip) {
  const entry = zip.file(PART);
  const rels = zip.file("word/_rels/document.xml.rels");
  if (!entry || !rels) return [];
  const media = await mediaMap(zip);
  const doc = parseXml(await entry.async("string"));
  const pictures = [];
  const nodes = doc.getElementsByTagNameNS(NS_PIC, "pic");
  for (let i = 0; i < nodes.length; i += 1) {
    const pic = nodes[i];
    if (insideFallback(pic)) continue;
    const blip = pic.getElementsByTagNameNS(NS_A, "blip")[0];
    const id = blip && (blip.getAttributeNS(NS_R, "embed") || blip.getAttribute("r:embed"));
    const path = media.get(id) || "";
    if (!RASTER_RE.test(path)) continue;
    const extent = displayedExtent(pic);
    if (!extent.cx || !extent.cy) continue;
    pictures.push({
      key: `${PART}:${pictures.length + 1}`,
      slideIndex: 1,
      slidePath: PART,
      label: "本文",
      name: `Picture ${pictures.length + 1}`,
      media: path,
      x: 0,
      y: 0,
      cx: extent.cx,
      cy: extent.cy,
      src: { l: 0, t: 0, visibleW: 1, visibleH: 1 },
    });
  }
  return pictures;
}

function displayedExtent(pic) {
  let node = pic;
  while (node) {
    if (node.localName === "inline" || node.localName === "anchor") {
      const extent = child(node, "extent");
      const cx = Number(extent && extent.getAttribute("cx")) || 0;
      const cy = Number(extent && extent.getAttribute("cy")) || 0;
      if (cx && cy) return { cx, cy };
    }
    node = node.parentNode;
  }
  const ext = pic.getElementsByTagNameNS(NS_A, "ext")[0];
  return {
    cx: Number(ext && ext.getAttribute("cx")) || 0,
    cy: Number(ext && ext.getAttribute("cy")) || 0,
  };
}

function rasterPics(doc, media) {
  const out = [];
  const nodes = doc.getElementsByTagNameNS(NS_PIC, "pic");
  for (let i = 0; i < nodes.length; i += 1) {
    const pic = nodes[i];
    if (insideFallback(pic)) continue;
    const blip = pic.getElementsByTagNameNS(NS_A, "blip")[0];
    const id = blip && (blip.getAttributeNS(NS_R, "embed") || blip.getAttribute("r:embed"));
    const path = media.get(id) || "";
    if (!RASTER_RE.test(path)) continue;
    out.push(pic);
  }
  return out;
}

async function mediaMap(zip) {
  const rels = zip.file("word/_rels/document.xml.rels");
  const media = new Map();
  if (!rels) return media;
  const relDoc = parseXml(await rels.async("string"));
  for (const rel of elementChildren(relDoc.documentElement)) {
    if (rel.getAttribute("Type") !== REL_IMAGE) continue;
    media.set(rel.getAttribute("Id"), resolveTarget(rel.getAttribute("Target")));
  }
  return media;
}

function occurrenceFromPrefix(prefix) {
  const match = String(prefix || "").match(/^imgtext-(\d+)-(\d+)-$/);
  return match ? Number(match[2]) : 0;
}

function markName(prefix, index) {
  return `${prefix}${index}`;
}

function ensureGroup(pic) {
  let data = pic.parentNode;
  while (data && data.localName !== "graphicData") data = data.parentNode;
  if (!data) return null;
  const uri = data.getAttribute("uri") || "";
  if (uri.indexOf("wordprocessingGroup") !== -1) {
    const groups = data.getElementsByTagNameNS(NS_WPG, "wgp");
    if (groups && groups[0]) return groups[0];
    return elementChildren(data).find((item) => (item.localName || String(item.nodeName || "").split(":").pop()) === "wgp") || null;
  }
  const doc = pic.ownerDocument;
  const extent = displayedExtent(pic);
  const wgp = doc.createElementNS(NS_WPG, "wpg:wgp");
  wgp.appendChild(doc.createElementNS(NS_WPG, "wpg:cNvGrpSpPr"));
  const grp = doc.createElementNS(NS_WPG, "wpg:grpSpPr");
  grp.appendChild(makeXfrm(doc, 0, 0, extent.cx, extent.cy, true));
  wgp.appendChild(grp);
  data.removeChild(pic);
  wgp.appendChild(pic);
  data.appendChild(wgp);
  data.setAttribute("uri", GROUP_URI);
  return wgp;
}

function makeXfrm(doc, x, y, cx, cy, child) {
  const xfrm = doc.createElementNS(NS_A, "a:xfrm");
  const off = doc.createElementNS(NS_A, "a:off");
  off.setAttribute("x", String(Math.round(x)));
  off.setAttribute("y", String(Math.round(y)));
  const ext = doc.createElementNS(NS_A, "a:ext");
  ext.setAttribute("cx", String(Math.max(1, Math.round(cx))));
  ext.setAttribute("cy", String(Math.max(1, Math.round(cy))));
  xfrm.appendChild(off);
  xfrm.appendChild(ext);
  if (child) {
    const chOff = doc.createElementNS(NS_A, "a:chOff");
    chOff.setAttribute("x", "0");
    chOff.setAttribute("y", "0");
    const chExt = doc.createElementNS(NS_A, "a:chExt");
    chExt.setAttribute("cx", String(Math.max(1, Math.round(cx))));
    chExt.setAttribute("cy", String(Math.max(1, Math.round(cy))));
    xfrm.appendChild(chOff);
    xfrm.appendChild(chExt);
  }
  return xfrm;
}

function maxBookmarkId(doc) {
  let max = 0;
  const nodes = doc.getElementsByTagNameNS(NS_W, "bookmarkStart");
  for (let i = 0; i < nodes.length; i += 1) {
    max = Math.max(max, Number(attr(nodes[i], "id")) || 0);
  }
  return max;
}

function buildDocxTextBox(doc, spec, bookmarkId) {
  const useLine = Boolean(spec.lineCy) && spec.cy > spec.lineCy * 1.35;
  const fontEmu = useLine ? spec.lineCy : spec.cy;
  const halfPoints = Math.max(2, Math.round(((fontEmu / 914400) * 72 * 0.92) * 2));
  const wsp = doc.createElementNS(NS_WPS, "wps:wsp");
  const cNv = doc.createElementNS(NS_WPS, "wps:cNvSpPr");
  cNv.setAttribute("txBox", "1");
  wsp.appendChild(cNv);

  const spPr = doc.createElementNS(NS_WPS, "wps:spPr");
  spPr.appendChild(makeXfrm(doc, spec.x, spec.y, spec.cx, spec.cy, false));
  const geom = doc.createElementNS(NS_A, "a:prstGeom");
  geom.setAttribute("prst", "rect");
  geom.appendChild(doc.createElementNS(NS_A, "a:avLst"));
  spPr.appendChild(geom);
  const fill = doc.createElementNS(NS_A, "a:solidFill");
  const fillClr = doc.createElementNS(NS_A, "a:srgbClr");
  fillClr.setAttribute("val", /^[0-9A-F]{6}$/i.test(spec.fill || "") ? spec.fill : "FFFFFF");
  fill.appendChild(fillClr);
  spPr.appendChild(fill);
  const ln = doc.createElementNS(NS_A, "a:ln");
  ln.appendChild(doc.createElementNS(NS_A, "a:noFill"));
  spPr.appendChild(ln);
  wsp.appendChild(spPr);

  const txbx = doc.createElementNS(NS_WPS, "wps:txbx");
  const content = doc.createElementNS(NS_W, "w:txbxContent");
  const p = doc.createElementNS(NS_W, "w:p");
  const pPr = doc.createElementNS(NS_W, "w:pPr");
  const spacing = doc.createElementNS(NS_W, "w:spacing");
  const lineEmu = useLine ? spec.lineCy : spec.cy;
  const lineTwips = Math.max(20, Math.round((lineEmu / 914400) * 72 * 20));
  wSet(spacing, "before", "0");
  wSet(spacing, "after", "0");
  wSet(spacing, "line", String(lineTwips));
  wSet(spacing, "lineRule", "exact");
  pPr.appendChild(spacing);
  const jc = doc.createElementNS(NS_W, "w:jc");
  wSet(jc, "val", spec.align === "ctr" ? "center" : "left");
  pPr.appendChild(jc);
  p.appendChild(pPr);
  const start = doc.createElementNS(NS_W, "w:bookmarkStart");
  wSet(start, "id", String(bookmarkId));
  wSet(start, "name", spec.name.replace(/-/g, "_"));
  p.appendChild(start);

  const run = doc.createElementNS(NS_W, "w:r");
  const rPr = doc.createElementNS(NS_W, "w:rPr");
  const fonts = doc.createElementNS(NS_W, "w:rFonts");
  for (const name of ["ascii", "hAnsi", "eastAsia", "cs"]) wSet(fonts, name, "Noto Sans JP");
  rPr.appendChild(fonts);
  const sz = doc.createElementNS(NS_W, "w:sz");
  wSet(sz, "val", String(halfPoints));
  rPr.appendChild(sz);
  const szCs = doc.createElementNS(NS_W, "w:szCs");
  wSet(szCs, "val", String(halfPoints));
  rPr.appendChild(szCs);
  const color = doc.createElementNS(NS_W, "w:color");
  wSet(color, "val", /^[0-9A-F]{6}$/i.test(spec.ink || "") ? spec.ink : "222222");
  rPr.appendChild(color);
  run.appendChild(rPr);
  writeRunText(run, spec.text);
  p.appendChild(run);
  const end = doc.createElementNS(NS_W, "w:bookmarkEnd");
  wSet(end, "id", String(bookmarkId));
  p.appendChild(end);
  content.appendChild(p);
  txbx.appendChild(content);
  wsp.appendChild(txbx);

  const bodyPr = doc.createElementNS(NS_WPS, "wps:bodyPr");
  bodyPr.setAttribute("wrap", "square");
  bodyPr.setAttribute("spcFirstLastPara", "0");
  bodyPr.setAttribute("anchor", useLine ? "t" : "ctr");
  for (const name of ["lIns", "tIns", "rIns", "bIns"]) bodyPr.setAttribute(name, "0");
  bodyPr.appendChild(doc.createElementNS(NS_A, "a:noAutofit"));
  wsp.appendChild(bodyPr);
  return wsp;
}

function ownedShape(wsp, prefix) {
  const needle = prefix.replace(/-/g, "_");
  const marks = wsp.getElementsByTagNameNS(NS_W, "bookmarkStart");
  for (let i = 0; i < marks.length; i += 1) {
    if ((attr(marks[i], "name") || "").startsWith(needle)) return true;
  }
  const nodes = wsp.getElementsByTagNameNS(NS_W, "t");
  for (let i = 0; i < nodes.length; i += 1) {
    if ((nodes[i].textContent || "").trim().startsWith(prefix)) return true;
  }
  return false;
}

async function replaceDocxPictureTextBoxes(zip, slidePath, namePrefix, boxes) {
  const entry = zip.file(slidePath || PART);
  if (!entry || !namePrefix) return { added: 0, removed: 0 };
  const doc = parseXml(await entry.async("string"));
  const pics = rasterPics(doc, await mediaMap(zip));
  const pic = pics[occurrenceFromPrefix(namePrefix) - 1];
  if (!pic) return { added: 0, removed: 0 };
  const group = ensureGroup(pic);
  if (!group) return { added: 0, removed: 0 };
  let removed = 0;
  for (const child of [...elementChildren(group)]) {
    if (child.localName !== "wsp") continue;
    if (!ownedShape(child, namePrefix)) continue;
    group.removeChild(child);
    removed += 1;
  }
  let bookmarkId = maxBookmarkId(doc);
  let added = 0;
  for (const item of boxes || []) {
    const text = String(item.text || "").trim();
    if (!text) continue;
    added += 1;
    bookmarkId += 1;
    group.appendChild(buildDocxTextBox(doc, {
      name: markName(namePrefix, added),
      x: item.x,
      y: item.y,
      cx: item.cx,
      cy: item.cy,
      lineCy: item.lineCy,
      text,
      fill: item.fill,
      ink: item.ink,
      align: item.align,
    }, bookmarkId));
  }
  if (added || removed) zip.file(slidePath || PART, serializeXml(doc));
  return { added, removed };
}

async function collectDocxFonts(zip) {
  const entry = zip.file(PART);
  if (!entry) return [];
  const xml = await entry.async("string");
  const counts = new Map();
  const re = /w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"/g;
  let match;
  while ((match = re.exec(xml)) !== null) {
    const name = match[1];
    if (!name) continue;
    counts.set(name, (counts.get(name) || 0) + 1);
  }
  return [...counts.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

const api = {
  extractDocxTexts,
  injectDocxTexts,
  listDocxRasterPictures,
  replaceDocxPictureTextBoxes,
  collectDocxFonts,
};

if (typeof module !== "undefined" && module.exports) module.exports = api;
if (typeof window !== "undefined") Object.assign(window, api);
})();
