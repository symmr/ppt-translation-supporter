// Removes the original text of translated blocks from a page's content
// stream, so the rewritten PDF no longer carries it (not drawn, not found by
// search or copy). Each removed string is replaced by a TJ advance of the
// same width, so text that follows on the same line keeps its position.
// Text in form XObjects and invisible text (OCR layers) is left alone; the
// caller covers those blocks instead.
(function () {
"use strict";

const WS = new Set([0, 9, 10, 12, 13, 32]);
const DELIM = new Set([40, 41, 60, 62, 91, 93, 123, 125, 47, 37]);
const IDENTITY = [1, 0, 0, 1, 0, 0];

function mul(m, n) {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

function apply(m, x, y) {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function latin1(bytes, start, end) {
  let out = "";
  for (let i = start; i < end; i += 1) out += String.fromCharCode(bytes[i]);
  return out;
}

// Content stream lexer. Tokens keep their byte range so untouched operators
// are copied through byte for byte.
function tokenize(bytes) {
  const tokens = [];
  const n = bytes.length;
  let i = 0;
  while (i < n) {
    const c = bytes[i];
    if (WS.has(c)) { i += 1; continue; }
    const start = i;
    if (c === 37) { // % comment
      while (i < n && bytes[i] !== 10 && bytes[i] !== 13) i += 1;
      continue;
    }
    if (c === 40) { // ( literal string )
      const out = [];
      let depth = 1;
      i += 1;
      while (i < n && depth > 0) {
        const ch = bytes[i];
        if (ch === 92) {
          const nx = bytes[i + 1];
          i += 2;
          if (nx === 110) out.push(10);
          else if (nx === 114) out.push(13);
          else if (nx === 116) out.push(9);
          else if (nx === 98) out.push(8);
          else if (nx === 102) out.push(12);
          else if (nx === 13) { if (bytes[i] === 10) i += 1; }
          else if (nx === 10) { /* line continuation */ }
          else if (nx >= 48 && nx <= 55) {
            let v = nx - 48;
            for (let k = 0; k < 2 && bytes[i] >= 48 && bytes[i] <= 55; k += 1) {
              v = v * 8 + (bytes[i] - 48);
              i += 1;
            }
            out.push(v & 255);
          } else if (nx !== undefined) out.push(nx);
          continue;
        }
        if (ch === 40) depth += 1;
        if (ch === 41) {
          depth -= 1;
          if (depth === 0) { i += 1; break; }
        }
        out.push(ch);
        i += 1;
      }
      tokens.push({ type: "str", value: Uint8Array.from(out), start, end: i });
      continue;
    }
    if (c === 60 && bytes[i + 1] === 60) { tokens.push({ type: "<<", start, end: i + 2 }); i += 2; continue; }
    if (c === 62 && bytes[i + 1] === 62) { tokens.push({ type: ">>", start, end: i + 2 }); i += 2; continue; }
    if (c === 60) { // <hex string>
      i += 1;
      let hex = "";
      while (i < n && bytes[i] !== 62) {
        const ch = bytes[i];
        if (!WS.has(ch)) hex += String.fromCharCode(ch);
        i += 1;
      }
      i += 1;
      if (hex.length % 2) hex += "0";
      const out = new Uint8Array(hex.length / 2);
      for (let k = 0; k < out.length; k += 1) out[k] = parseInt(hex.substr(k * 2, 2), 16) || 0;
      tokens.push({ type: "str", value: out, start, end: i });
      continue;
    }
    if (c === 91 || c === 93 || c === 123 || c === 125) {
      tokens.push({ type: String.fromCharCode(c), start, end: i + 1 });
      i += 1;
      continue;
    }
    if (c === 47) { // /Name
      i += 1;
      while (i < n && !WS.has(bytes[i]) && !DELIM.has(bytes[i])) i += 1;
      const raw = latin1(bytes, start + 1, i).replace(/#([0-9a-fA-F]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
      tokens.push({ type: "name", value: raw, start, end: i });
      continue;
    }
    while (i < n && !WS.has(bytes[i]) && !DELIM.has(bytes[i])) i += 1;
    if (i === start) { i += 1; continue; } // stray delimiter such as ")"
    const word = latin1(bytes, start, i);
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) {
      tokens.push({ type: "num", value: parseFloat(word), start, end: i });
      continue;
    }
    tokens.push({ type: "op", value: word, start, end: i });
    if (word === "ID") {
      // inline image data runs to the next whitespace-delimited EI
      let j = i + 1;
      while (j < n) {
        if (bytes[j] === 69 && bytes[j + 1] === 73 && WS.has(bytes[j - 1]) && (j + 2 >= n || WS.has(bytes[j + 2]))) break;
        j += 1;
      }
      tokens.push({ type: "op", value: "EI", start: j, end: Math.min(n, j + 2) });
      i = Math.min(n, j + 2);
    }
  }
  return tokens;
}

// Operators with their operand values and the byte range they span.
function parseOps(bytes) {
  const tokens = tokenize(bytes);
  const ops = [];
  let operands = [];
  let operandStart = -1;
  const stack = [];
  const push = (value, start) => {
    if (stack.length) stack[stack.length - 1].push(value);
    else {
      if (operandStart < 0) operandStart = start;
      operands.push(value);
    }
  };
  for (const tok of tokens) {
    if (tok.type === "[" || tok.type === "<<" || tok.type === "{") {
      if (!stack.length && operandStart < 0) operandStart = tok.start;
      stack.push([]);
      continue;
    }
    if (tok.type === "]" || tok.type === ">>" || tok.type === "}") {
      const value = stack.pop() || [];
      if (tok.type !== "]") value.dict = true;
      push(value, tok.start);
      continue;
    }
    if (tok.type === "op" && !stack.length) {
      ops.push({ op: tok.value, operands, start: operandStart < 0 ? tok.start : operandStart, end: tok.end });
      operands = [];
      operandStart = -1;
      continue;
    }
    if (tok.type === "op") { push({ keyword: tok.value }, tok.start); continue; }
    push(tok.type === "name" ? { name: tok.value } : tok.value, tok.start);
  }
  return ops;
}

// Glyph widths per font resource, in 1/1000 text space units.
// fonts: { [resourceName]: { twoByte, width(code) } }
function codesOf(font, bytes) {
  const codes = [];
  if (font && font.twoByte) {
    for (let i = 0; i + 1 < bytes.length; i += 2) codes.push((bytes[i] << 8) | bytes[i + 1]);
  } else {
    for (const b of bytes) codes.push(b);
  }
  return codes;
}

function stringAdvance(state, bytes) {
  const font = state.font;
  let tx = 0;
  for (const code of codesOf(font, bytes)) {
    const w0 = font ? font.width(code) : 500;
    const wordSpace = (!font || !font.twoByte) && code === 32 ? state.Tw : 0;
    tx += ((w0 / 1000) * state.Tfs + state.Tc + wordSpace) * state.Th;
  }
  return tx;
}

function fmt(n) {
  if (!Number.isFinite(n)) return "0";
  const s = n.toFixed(3).replace(/\.?0+$/, "");
  return s === "-0" ? "0" : s;
}

function hexOf(bytes) {
  let s = "<";
  for (const b of bytes) s += b.toString(16).padStart(2, "0");
  return `${s}>`;
}

function inRect(rects, x, y) {
  for (let k = 0; k < rects.length; k += 1) {
    const r = rects[k];
    if (x >= r.x0 && x <= r.x1 && y >= r.y0 && y <= r.y1) return k;
  }
  return -1;
}

// bytes: decoded page content. rects: [{ x0, y0, x1, y1 }] in user space.
// Returns the new content and, per rect, how many visible strings went.
function removeTextInRects(bytes, rects, fonts) {
  const ops = parseOps(bytes);
  const hits = new Array(rects.length).fill(0);
  const edits = [];
  let gs = { ctm: IDENTITY.slice(), Tc: 0, Tw: 0, Th: 1, TL: 0, Tfs: 0, font: null, rise: 0, Tr: 0 };
  const saved = [];
  let tm = IDENTITY.slice();
  let tlm = IDENTITY.slice();
  const num = (v, d) => (typeof v === "number" ? v : d);

  // Shows one TJ-style array; returns the replacement elements, or null when
  // nothing in it was removed.
  const show = (elements) => {
    const out = [];
    let changed = false;
    for (const el of elements) {
      if (typeof el === "number") {
        out.push(el);
        tm = mul([1, 0, 0, 1, (-el / 1000) * gs.Tfs * gs.Th, 0], tm);
        continue;
      }
      if (!(el instanceof Uint8Array)) continue;
      const tx = stringAdvance(gs, el);
      const at = mul(tm, gs.ctm);
      const [px, py] = apply(at, tx / 2, gs.rise + gs.Tfs * 0.3);
      const hit = el.length ? inRect(rects, px, py) : -1;
      const scale = gs.Tfs * gs.Th;
      if (hit >= 0 && scale) {
        if (gs.Tr !== 3 && gs.Tr !== 7) hits[hit] += 1;
        out.push(-(tx / scale) * 1000);
        changed = true;
      } else {
        out.push(el);
      }
      tm = mul([1, 0, 0, 1, tx, 0], tm);
    }
    return changed ? out : null;
  };
  const serialize = (elements) => `[${elements.map((el) => (typeof el === "number" ? fmt(el) : hexOf(el))).join(" ")}] TJ`;
  const nextLine = () => {
    tlm = mul([1, 0, 0, 1, 0, -gs.TL], tlm);
    tm = tlm.slice();
  };

  for (const { op, operands: a, start, end } of ops) {
    switch (op) {
      case "q": saved.push({ ...gs, ctm: gs.ctm.slice() }); break;
      case "Q": if (saved.length) gs = saved.pop(); break;
      case "cm":
        if (a.length >= 6) gs.ctm = mul(a.slice(0, 6).map((v) => num(v, 0)), gs.ctm);
        break;
      case "BT": tm = IDENTITY.slice(); tlm = IDENTITY.slice(); break;
      case "Tc": gs.Tc = num(a[0], 0); break;
      case "Tw": gs.Tw = num(a[0], 0); break;
      case "Tz": gs.Th = num(a[0], 100) / 100; break;
      case "TL": gs.TL = num(a[0], 0); break;
      case "Ts": gs.rise = num(a[0], 0); break;
      case "Tr": gs.Tr = num(a[0], 0); break;
      case "Tf":
        gs.font = (a[0] && fonts[a[0].name]) || null;
        gs.Tfs = num(a[1], 0);
        break;
      case "Td":
      case "TD":
        if (op === "TD") gs.TL = -num(a[1], 0);
        tlm = mul([1, 0, 0, 1, num(a[0], 0), num(a[1], 0)], tlm);
        tm = tlm.slice();
        break;
      case "Tm":
        if (a.length >= 6) {
          tlm = a.slice(0, 6).map((v) => num(v, 0));
          tm = tlm.slice();
        }
        break;
      case "T*": nextLine(); break;
      case "Tj": {
        const out = a[0] instanceof Uint8Array ? show([a[0]]) : null;
        if (out) edits.push({ start, end, text: serialize(out) });
        break;
      }
      case "TJ": {
        const out = Array.isArray(a[0]) ? show(a[0]) : null;
        if (out) edits.push({ start, end, text: serialize(out) });
        break;
      }
      case "'": {
        nextLine();
        const out = a[0] instanceof Uint8Array ? show([a[0]]) : null;
        if (out) edits.push({ start, end, text: `T* ${serialize(out)}` });
        break;
      }
      case "\"": {
        gs.Tw = num(a[0], 0);
        gs.Tc = num(a[1], 0);
        nextLine();
        const out = a[2] instanceof Uint8Array ? show([a[2]]) : null;
        if (out) edits.push({ start, end, text: `${fmt(gs.Tw)} Tw ${fmt(gs.Tc)} Tc T* ${serialize(out)}` });
        break;
      }
      default:
        break;
    }
  }
  if (!edits.length) return { bytes, hits, removed: 0 };
  const parts = [];
  let pos = 0;
  let size = 0;
  for (const edit of edits) {
    const head = bytes.subarray(pos, edit.start);
    const body = Uint8Array.from(edit.text, (ch) => ch.charCodeAt(0));
    parts.push(head, body);
    size += head.length + body.length;
    pos = edit.end;
  }
  const tail = bytes.subarray(pos);
  parts.push(tail);
  size += tail.length;
  const out = new Uint8Array(size);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return { bytes: out, hits, removed: edits.length };
}

// --- pdf-lib side -----------------------------------------------------------

function numberOf(PDFLib, obj) {
  return obj instanceof PDFLib.PDFNumber ? obj.asNumber() : undefined;
}

function fontWidths(PDFLib, context, fontDict) {
  const { PDFName, PDFArray, PDFDict } = PDFLib;
  const subtype = fontDict.lookup(PDFName.of("Subtype"));
  if (subtype === PDFName.of("Type0")) {
    const descendants = fontDict.lookupMaybe(PDFName.of("DescendantFonts"), PDFArray);
    const cid = descendants && descendants.size() ? descendants.lookupMaybe(0, PDFDict) : null;
    const map = new Map();
    let dw = 1000;
    if (cid) {
      dw = numberOf(PDFLib, cid.lookup(PDFName.of("DW"))) ?? 1000;
      const w = cid.lookupMaybe(PDFName.of("W"), PDFArray);
      if (w) {
        let i = 0;
        while (i < w.size()) {
          const first = numberOf(PDFLib, w.lookup(i));
          const next = w.lookup(i + 1);
          if (first === undefined) break;
          if (next instanceof PDFArray) {
            for (let k = 0; k < next.size(); k += 1) map.set(first + k, numberOf(PDFLib, next.lookup(k)) ?? dw);
            i += 2;
          } else {
            const last = numberOf(PDFLib, next);
            const value = numberOf(PDFLib, w.lookup(i + 2)) ?? dw;
            if (last === undefined) break;
            for (let c = first; c <= last && c - first < 65536; c += 1) map.set(c, value);
            i += 3;
          }
        }
      }
    }
    // Identity-H and the other two-byte CMaps; CID = code for Identity
    return { twoByte: true, width: (code) => (map.has(code) ? map.get(code) : dw) };
  }
  const first = numberOf(PDFLib, fontDict.lookup(PDFName.of("FirstChar"))) ?? 0;
  const widths = fontDict.lookupMaybe(PDFName.of("Widths"), PDFArray);
  const descriptor = fontDict.lookupMaybe(PDFName.of("FontDescriptor"), PDFDict);
  const missing = (descriptor && numberOf(PDFLib, descriptor.lookup(PDFName.of("MissingWidth")))) || 500;
  const list = [];
  if (widths) for (let k = 0; k < widths.size(); k += 1) list.push(numberOf(PDFLib, widths.lookup(k)) ?? missing);
  return {
    twoByte: false,
    width: (code) => {
      const value = list[code - first];
      return value === undefined ? missing : value;
    },
  };
}

function pageFonts(PDFLib, page) {
  const { PDFName, PDFDict } = PDFLib;
  const fonts = {};
  const resources = page.node.Resources();
  const fontRes = resources && resources.lookupMaybe(PDFName.of("Font"), PDFDict);
  if (!fontRes) return fonts;
  for (const [key, value] of fontRes.entries()) {
    const dict = page.doc.context.lookupMaybe(value, PDFDict);
    if (!dict) continue;
    try {
      fonts[key.decodeText ? key.decodeText() : key.asString().slice(1)] = fontWidths(PDFLib, page.doc.context, dict);
    } catch (_) {
      // unreadable font dict: its text stays and the block gets covered
    }
  }
  return fonts;
}

function pageContentBytes(PDFLib, page) {
  const { PDFArray, PDFRawStream } = PDFLib;
  const contents = page.node.Contents();
  const streams = [];
  if (contents instanceof PDFArray) {
    for (let k = 0; k < contents.size(); k += 1) streams.push(contents.lookup(k));
  } else if (contents) {
    streams.push(contents);
  }
  const chunks = [];
  for (const stream of streams) {
    if (!(stream instanceof PDFRawStream)) return null;
    chunks.push(PDFLib.decodePDFRawStream(stream).decode());
  }
  let size = 0;
  for (const chunk of chunks) size += chunk.length + 1;
  const out = new Uint8Array(size);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
    out[at] = 10; // operators never continue across stream boundaries
    at += 1;
  }
  return out;
}

// Rewrites one page. blocks: metadata entries. Returns the ids whose original
// text was removed; the rest still need a cover.
function removePageText(PDFLib, page, blocks) {
  if (!blocks.length) return new Set();
  let content;
  try {
    content = pageContentBytes(PDFLib, page);
  } catch (_) {
    content = null;
  }
  if (!content) return new Set();
  const rects = blocks.map((b) => ({
    x0: b.x - b.size * 0.5,
    y0: b.y - b.size * 0.3,
    x1: b.x + b.w + b.size * 0.3,
    y1: b.y + b.h + b.size * 0.2,
  }));
  const result = removeTextInRects(content, rects, pageFonts(PDFLib, page));
  const done = new Set();
  result.hits.forEach((count, k) => { if (count > 0) done.add(blocks[k].id); });
  if (result.removed) {
    const context = page.doc.context;
    const ref = context.register(context.flateStream(result.bytes));
    page.node.set(PDFLib.PDFName.of("Contents"), ref);
  }
  return done;
}

const api = { tokenize, parseOps, removeTextInRects, removePageText };

if (typeof module !== "undefined" && module.exports) module.exports = api;
if (typeof window !== "undefined") Object.assign(window, api);
})();
