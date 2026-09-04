// PPTX text extract / inject. DOMParser in the browser, @xmldom/xmldom in Node.
// No font replacement — that stays in PPT Finalizer.

const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_XML = "http://www.w3.org/XML/1998/namespace";
const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

const UID_PATTERN = /(uid_[0-9]+)/;
const TAG_RE = /\[(\d+)\]([\s\S]*?)\[\/\1\]/g;
const TAG_RE_LEGACY = /⟦(\d+)⟧([\s\S]*?)⟦\/\1⟧/g;
const STRIP_TAG_RE = /(?:\[\/?\d+\]|⟦\/?\d+⟧)/g;

const TRANSLATION_PROMPT = `あなたはプロのローカライズ翻訳者です。以下のスライド抽出テキストを日本語に翻訳してください。

# 厳守するルール
1. 「uid_0001」のような uid 行は、翻訳・改変・削除・並べ替えをせず、そのまま出力する。
2. 「[0]...[/0]」のようなタグは、番号・個数・順序・開閉をすべて維持する。タグの中のテキストだけを翻訳し、タグ自体は一切追加・削除・変更しない。空のタグ([2][/2])もそのまま残す。
3. 入力の行構造(uid 行 → 本文行)を厳密に保つ。
4. 出力は翻訳結果のテキストのみ。前置き・解説を付けない。
5. 製品名・固有名詞のみ原文のまま残す。それ以外の英文はすべて日本語に訳す。一般語・説明文・見出し・ラベルを英語のまま残さない。
6. 文体は常体(だ・である調)で統一する。見出しや短い名詞句は体言止めを適切に用いる。
7. 訳文が対応するタグの範囲からはみ出さないようにする(各タグ内は、その原文に対応する訳のみを入れる)。
8. 訳漏れを避ける。各 uid の本文行に英文が残る場合、製品名・固有名詞でない限り必ず訳す。
`;

function getDOMParser() {
  if (typeof DOMParser !== "undefined") return DOMParser;
  return require("@xmldom/xmldom").DOMParser;
}

function getXMLSerializer() {
  if (typeof XMLSerializer !== "undefined") return XMLSerializer;
  return require("@xmldom/xmldom").XMLSerializer;
}

function parseXml(xml) {
  if (/<!DOCTYPE/i.test(xml)) {
    throw new Error("DOCTYPE を含む XML は扱えません");
  }
  const Parser = getDOMParser();
  const doc = new Parser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (!root) {
    throw new Error("XML の解析に失敗しました（空のドキュメント）");
  }
  // Some parsers (e.g. Chromium) nest <parsererror> as a child of the root
  // instead of replacing it, so a root-only check misses those failures.
  const parserErrors = typeof doc.getElementsByTagName === "function"
    ? doc.getElementsByTagName("parsererror")
    : null;
  if (root.localName === "parsererror" || root.nodeName === "parsererror" || (parserErrors && parserErrors.length)) {
    const detail = parserErrors && parserErrors.length
      ? String(parserErrors[0].textContent || "").trim().replace(/\s+/g, " ").slice(0, 300)
      : "";
    throw new Error(`XML の解析に失敗しました${detail ? `: ${detail}` : ""}`);
  }
  return doc;
}

function parseXmlAt(xml, path) {
  try {
    return parseXml(xml);
  } catch (err) {
    throw new Error(`${path} の解析に失敗しました: ${err.message || err}`, { cause: err });
  }
}

function serializeXml(doc) {
  const Serializer = getXMLSerializer();
  let out = new Serializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) {
    out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
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

function firstChildLocal(el, localName) {
  return elementChildren(el).find((child) => child.localName === localName) || null;
}

function directTxBody(el) {
  return elementChildren(el).find((child) => child.localName === "txBody") || null;
}

function paragraphsOf(txBody) {
  if (!txBody) return [];
  return elementChildren(txBody).filter(
    (child) => child.localName === "p" && child.namespaceURI === NS_A
  );
}

function runsOf(paragraph) {
  return elementChildren(paragraph).filter((child) => child.localName === "r");
}

function collectTText(el) {
  let text = "";
  const nodes = el.getElementsByTagNameNS
    ? el.getElementsByTagNameNS(NS_A, "t")
    : [];
  if (nodes && nodes.length) {
    for (let i = 0; i < nodes.length; i += 1) {
      text += nodes[i].textContent || "";
    }
    return text;
  }
  for (const child of elementChildren(el)) {
    if (child.localName === "t") text += child.textContent || "";
  }
  return text;
}

function runText(run) {
  return collectTText(run);
}

function paragraphPlainText(paragraph) {
  let text = "";
  for (const child of elementChildren(paragraph)) {
    if (child.localName === "r" || child.localName === "fld") {
      text += collectTText(child);
    } else if (child.localName === "br") {
      text += "\n";
    }
  }
  return text;
}

function runsToTaggedText(paragraph) {
  const runs = runsOf(paragraph);
  if (runs.length <= 1) return paragraphPlainText(paragraph);
  return runs.map((run, i) => `[${i}]${runText(run)}[/${i}]`).join("");
}

function firstT(run) {
  for (const child of elementChildren(run)) {
    if (child.localName === "t") return child;
  }
  return null;
}

function setRunText(run, text) {
  const t = firstT(run);
  if (!t) return;
  t.textContent = text;
  t.setAttributeNS(NS_XML, "xml:space", "preserve");
}

function setParagraphText(paragraph, translatedText) {
  const runs = runsOf(paragraph);
  const matches = [];
  TAG_RE.lastIndex = 0;
  let match = TAG_RE.exec(translatedText);
  if (!match) {
    TAG_RE_LEGACY.lastIndex = 0;
    match = TAG_RE_LEGACY.exec(translatedText);
    const re = match ? TAG_RE_LEGACY : null;
    if (re) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(translatedText)) !== null) {
        matches.push([m[1], m[2]]);
      }
    }
  } else {
    TAG_RE.lastIndex = 0;
    let m;
    while ((m = TAG_RE.exec(translatedText)) !== null) {
      matches.push([m[1], m[2]]);
    }
  }

  const indices = matches.map(([i]) => Number(i));
  const unique = new Set(indices);
  if (
    matches.length &&
    runs.length &&
    indices.every((i) => i >= 0 && i < runs.length) &&
    unique.size === indices.length
  ) {
    const used = new Set();
    for (const [idxStr, seg] of matches) {
      const i = Number(idxStr);
      setRunText(runs[i], seg);
      used.add(i);
    }
    for (let i = 0; i < runs.length; i += 1) {
      if (!used.has(i)) setRunText(runs[i], "");
    }
    return { mode: "runs" };
  }

  // Fallback: the tags do not line up with the runs, so the paragraph is
  // flattened into the first run. With more than one run that silently drops
  // per-run formatting (colour, bold, ...), so the caller is told about it.
  const plain = String(translatedText).replace(STRIP_TAG_RE, "");
  if (runs.length) {
    setRunText(runs[0], plain);
    for (let i = 1; i < runs.length; i += 1) setRunText(runs[i], "");
  }
  return { mode: runs.length > 1 ? "flattened" : "plain" };
}

function expandShapeNode(el) {
  if (el.localName !== "AlternateContent") return [el];
  const choice = firstChildLocal(el, "Choice") || firstChildLocal(el, "Fallback");
  return choice ? elementChildren(choice) : [];
}

function shapeChildren(parent) {
  const out = [];
  for (const child of elementChildren(parent)) {
    if (child.localName === "nvGrpSpPr" || child.localName === "grpSpPr") continue;
    out.push(...expandShapeNode(child));
  }
  return out;
}

function getSpTree(doc) {
  const trees = doc.getElementsByTagNameNS(NS_P, "spTree");
  return trees && trees[0] ? trees[0] : null;
}

function findTable(el) {
  const tables = el.getElementsByTagNameNS(NS_A, "tbl");
  return tables && tables[0] ? tables[0] : null;
}

function tableRows(tbl) {
  return elementChildren(tbl).filter((child) => child.localName === "tr");
}

function tableCells(tr) {
  return elementChildren(tr).filter((child) => child.localName === "tc");
}

function isContinuationCell(tc) {
  const tcPr = firstChildLocal(tc, "tcPr");
  if (!tcPr) return false;
  for (const child of elementChildren(tcPr)) {
    if (child.localName !== "hMerge" && child.localName !== "vMerge") continue;
    const val = (child.getAttribute("val") || "").toLowerCase();
    if (val === "restart" || val === "0" || val === "false") continue;
    return true;
  }
  return false;
}

function parseRelationships(xml) {
  const rels = [];
  const re = /<Relationship\s+([^>]+?)\/?>/g;
  let match;
  while ((match = re.exec(xml)) !== null) {
    const attrs = match[1];
    const id = /(?:\bId|Id)="([^"]+)"/.exec(attrs)?.[1];
    const type = /Type="([^"]+)"/.exec(attrs)?.[1];
    const target = /Target="([^"]+)"/.exec(attrs)?.[1];
    if (type && target) rels.push({ id, type, target });
  }
  return rels;
}

function resolveZipPath(basePath, target) {
  const normalized = String(target || "").replace(/\\/g, "/");
  if (normalized.startsWith("/")) return normalized.replace(/^\//, "");
  const baseDir = basePath.includes("/")
    ? basePath.slice(0, basePath.lastIndexOf("/") + 1)
    : "";
  const parts = (baseDir + normalized).split("/");
  const stack = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function relsPathForPart(partPath) {
  const slash = partPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : partPath.slice(0, slash);
  const name = slash === -1 ? partPath : partPath.slice(slash + 1);
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

async function getPresentationSlideOrder(zip) {
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const presFile = zip.file("ppt/presentation.xml");
  if (!presRelsFile || !presFile) return [];

  const presRels = parseRelationships(await presRelsFile.async("string"));
  const rIdToSlide = new Map();
  for (const rel of presRels) {
    if (rel.type === REL_SLIDE) {
      rIdToSlide.set(rel.id, resolveZipPath("ppt/presentation.xml", rel.target));
    }
  }

  const presXml = await presFile.async("string");
  const slides = [];
  const sldIdRe = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
  let match;
  while ((match = sldIdRe.exec(presXml)) !== null) {
    const slidePath = rIdToSlide.get(match[1]);
    if (slidePath) slides.push(slidePath);
  }
  return slides;
}

async function notesPathForSlide(zip, slidePath) {
  const relsFile = zip.file(relsPathForPart(slidePath));
  if (!relsFile) return null;
  const rels = parseRelationships(await relsFile.async("string"));
  for (const rel of rels) {
    if (rel.type === REL_NOTES_SLIDE) {
      return resolveZipPath(slidePath, rel.target);
    }
  }
  return null;
}

function findNotesBodyTxBody(doc) {
  const phs = doc.getElementsByTagNameNS(NS_P, "ph");
  if (!phs) return null;
  for (let i = 0; i < phs.length; i += 1) {
    const ph = phs[i];
    const type = ph.getAttribute("type") || "";
    if (type !== "body" && type !== "nbody") continue;
    let el = ph.parentNode;
    while (el && el.localName !== "sp") el = el.parentNode;
    if (el) {
      const txBody = directTxBody(el);
      if (txBody) return txBody;
    }
  }
  return null;
}

function locateShape(doc, shapePath) {
  let node = getSpTree(doc);
  if (!node) return null;
  for (const index of shapePath) {
    if (node.localName !== "spTree" && node.localName !== "grpSp") return null;
    const next = shapeChildren(node)[index];
    if (!next) return null;
    node = next;
  }
  return node;
}

function walkShapes(parent, shapePath, onShape) {
  const kids = shapeChildren(parent);
  kids.forEach((el, index) => {
    const path = shapePath.concat(index);
    if (el.localName === "grpSp") {
      walkShapes(el, path, onShape);
      return;
    }
    onShape(el, path);
  });
}

function collectFromTxBody(txBody, extra, texts, metadata, uidRef) {
  paragraphsOf(txBody).forEach((paragraph, paraIdx) => {
    const tagged = runsToTaggedText(paragraph);
    if (!String(tagged).trim()) return;
    const uid = `uid_${String(uidRef[0]).padStart(4, "0")}`;
    uidRef[0] += 1;
    texts.push(uid);
    texts.push(tagged);
    metadata.push({ id: uid, paraIdx, ...extra });
  });
}

function collectFromShape(el, shapePath, slidePath, texts, metadata, uidRef) {
  const tbl = el.localName === "graphicFrame" ? findTable(el) : null;
  if (tbl) {
    tableRows(tbl).forEach((tr, rowIdx) => {
      tableCells(tr).forEach((tc, colIdx) => {
        if (isContinuationCell(tc)) return;
        collectFromTxBody(
          directTxBody(tc),
          { type: "table", slidePath, shapePath, rowIdx, colIdx },
          texts,
          metadata,
          uidRef
        );
      });
    });
    return;
  }
  const txBody = directTxBody(el);
  if (!txBody) return;
  collectFromTxBody(
    txBody,
    { type: "shape", slidePath, shapePath },
    texts,
    metadata,
    uidRef
  );
}

function collectFromNotes(doc, slidePath, notesPath, texts, metadata, uidRef) {
  const txBody = findNotesBodyTxBody(doc);
  if (!txBody) return;
  collectFromTxBody(
    txBody,
    { type: "notes", slidePath, notesPath },
    texts,
    metadata,
    uidRef
  );
}

async function extractTextsFromZip(zip) {
  const slides = await getPresentationSlideOrder(zip);
  const texts = [];
  const metadata = [];
  const uidRef = [1];

  for (let slideIdx = 0; slideIdx < slides.length; slideIdx += 1) {
    const slidePath = slides[slideIdx];
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;
    const doc = parseXmlAt(await slideFile.async("string"), slidePath);
    const tree = getSpTree(doc);
    if (tree) {
      walkShapes(tree, [], (el, shapePath) => {
        collectFromShape(el, shapePath, slidePath, texts, metadata, uidRef);
      });
    }

    const notesPath = await notesPathForSlide(zip, slidePath);
    if (!notesPath) continue;
    const notesFile = zip.file(notesPath);
    if (!notesFile) continue;
    try {
      const notesDoc = parseXmlAt(await notesFile.async("string"), notesPath);
      collectFromNotes(notesDoc, slidePath, notesPath, texts, metadata, uidRef);
    } catch (err) {
      // empty / malformed notes must not abort extract; log for diagnosis
      console.warn(`[pptx-text] notes の抽出をスキップしました (${notesPath}):`, err);
    }
  }

  return {
    lines: texts,
    text: texts.join("\n"),
    metadata,
    slideCount: slides.length,
    uidCount: metadata.length,
  };
}

function locateParagraph(doc, loc) {
  if (loc.type === "notes") {
    const txBody = findNotesBodyTxBody(doc);
    return paragraphsOf(txBody)[loc.paraIdx] || null;
  }

  const shape = locateShape(doc, loc.shapePath || []);
  if (!shape) return null;

  if (loc.type === "table") {
    const tbl = findTable(shape);
    if (!tbl) return null;
    const tr = tableRows(tbl)[loc.rowIdx];
    if (!tr) return null;
    const tc = tableCells(tr)[loc.colIdx];
    if (!tc || isContinuationCell(tc)) return null;
    return paragraphsOf(directTxBody(tc))[loc.paraIdx] || null;
  }

  return paragraphsOf(directTxBody(shape))[loc.paraIdx] || null;
}

async function injectTextsToZip(zip, translations, metadata) {
  const byPath = new Map();
  for (const loc of metadata) {
    const path = loc.type === "notes" ? loc.notesPath : loc.slidePath;
    if (!path) continue;
    if (!byPath.has(path)) byPath.set(path, []);
    byPath.get(path).push(loc);
  }

  let injected = 0;
  let missing = 0;
  const flattened = [];

  for (const [path, locs] of byPath) {
    const entry = zip.file(path);
    if (!entry) {
      missing += locs.length;
      continue;
    }
    const doc = parseXmlAt(await entry.async("string"), path);
    for (const loc of locs) {
      if (!Object.prototype.hasOwnProperty.call(translations, loc.id)) {
        missing += 1;
        continue;
      }
      const paragraph = locateParagraph(doc, loc);
      if (!paragraph) {
        missing += 1;
        continue;
      }
      const outcome = setParagraphText(paragraph, translations[loc.id]);
      if (outcome && outcome.mode === "flattened") flattened.push(loc.id);
      injected += 1;
    }
    zip.file(path, serializeXml(doc));
  }

  return { injected, missing, flattened };
}

function parseUidDelimitedText(source) {
  const translations = {};
  let currentUid = null;
  const currentLines = [];

  const flush = () => {
    if (currentUid && currentLines.length) {
      translations[currentUid] = currentLines.join("\n");
    }
  };

  String(source).replace(/^\uFEFF/, "").split(/\r?\n/).forEach((line) => {
    const match = UID_PATTERN.exec(line);
    if (match && line.trim() === match[1]) {
      flush();
      currentUid = match[1];
      currentLines.length = 0;
      return;
    }
    if (currentUid) currentLines.push(line);
  });
  flush();
  return translations;
}

function formatExtractFile(lines) {
  return `\uFEFF${lines.join("\n")}`;
}

function validateTranslations(metadata, translations) {
  const expected = metadata.map((item) => item.id);
  const got = Object.keys(translations);
  const expectedSet = new Set(expected);
  const gotSet = new Set(got);
  const missing = expected.filter((id) => !gotSet.has(id));
  const extra = got.filter((id) => !expectedSet.has(id));
  return {
    ok: missing.length === 0 && extra.length === 0,
    expected: expected.length,
    got: got.length,
    missing,
    extra,
  };
}

function translatedOutputName(originalName) {
  const base = String(originalName || "deck").replace(/\.pptx$/i, "");
  return `${base}_translated.pptx`;
}

function extractOutputName(originalName) {
  const base = String(originalName || "deck").replace(/\.pptx$/i, "");
  return `${base}_to_translate.txt`;
}

const api = {
  TRANSLATION_PROMPT,
  extractTextsFromZip,
  injectTextsToZip,
  parseUidDelimitedText,
  validateTranslations,
  formatExtractFile,
  translatedOutputName,
  extractOutputName,
  parseXml,
  runsToTaggedText,
  setParagraphText,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = api;
}

if (typeof window !== "undefined") {
  Object.assign(window, api);
}
