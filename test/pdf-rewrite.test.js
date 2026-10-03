"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const { parseOps, removeTextInRects } = require("../docs/pdf-rewrite.js");

const enc = (s) => Uint8Array.from(s, (c) => c.charCodeAt(0));
const dec = (b) => String.fromCharCode(...b);
const font = { F1: { twoByte: false, width: () => 500 } };

test("parses strings, arrays, dicts and inline images", () => {
  const ops = parseOps(enc("/P <</MCID 0>> BDC BT (a\\)b) Tj [(x) -20 <4142>] TJ ET EMC BI /W 1 ID \x00EI\x01 EI Q"));
  assert.deepEqual(ops.map((o) => o.op), ["BDC", "BT", "Tj", "TJ", "ET", "EMC", "BI", "ID", "EI", "Q"]);
  assert.equal(dec(ops[2].operands[0]), "a)b");
  assert.equal(dec(ops[3].operands[0][2]), "AB");
});

test("removes only the strings inside the rect and keeps the advance", () => {
  const src = "BT /F1 10 Tf 1 0 0 1 50 100 Tm (Left) Tj (Next) Tj 0 -20 Td (Below) Tj ET";
  const rects = [{ x0: 45, y0: 95, x1: 72, y1: 110 }];
  const out = removeTextInRects(enc(src), rects, font);
  assert.deepEqual(out.hits, [1]);
  assert.equal(dec(out.bytes), "BT /F1 10 Tf 1 0 0 1 50 100 Tm [-2000] TJ (Next) Tj 0 -20 Td (Below) Tj ET");
});

test("splits a TJ array and follows q/cm/Q", () => {
  const src = "q 1 0 0 1 100 0 cm BT /F1 10 Tf 0 0 Td [(AA) -3000 (BB)] TJ ET Q BT /F1 10 Tf 0 0 Td (CC) Tj ET";
  // second string starts at x = 100 + 10 + 30 = 140 in user space
  const out = removeTextInRects(enc(src), [{ x0: 135, y0: -5, x1: 160, y1: 10 }], font);
  assert.deepEqual(out.hits, [1]);
  assert.match(dec(out.bytes), /\[<4141> -3000 -1000\] TJ/);
  assert.match(dec(out.bytes), /\(CC\) Tj/);
});

test("invisible text is removed but not counted as visible", () => {
  const src = "BT 3 Tr /F1 10 Tf 10 10 Td (Ghost) Tj ET";
  const out = removeTextInRects(enc(src), [{ x0: 0, y0: 0, x1: 50, y1: 30 }], font);
  assert.deepEqual(out.hits, [0]);
  assert.equal(out.removed, 1);
});
