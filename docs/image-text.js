(function () {
"use strict";

// Word boxes come from one OCR pass. The page always uses the aggressive profile.

function profileFromSlider(t) {
  const u = Math.min(100, Math.max(0, Number(t))) / 100;
  return {
    minConfidence: 12 + 63 * u,
    minHeight: 12 + 12 * u,
    maxHeight: 190 - 60 * u,
    maxWidthFactor: 3.6 - 1.8 * u,
    minChars: u >= 0.7 ? 2 : 1,
    gapFactor: 0.95 - 0.5 * u,
  };
}

function median(values) {
  const sorted = values.slice().sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

function mergeSegment(words) {
  let text = "";
  words.forEach((word, index) => {
    if (index === 0) {
      text = word.text;
      return;
    }
    const prev = text.slice(-1);
    const next = word.text.slice(0, 1);
    const space = prev.charCodeAt(0) < 128 && next.charCodeAt(0) < 128;
    text += (space ? " " : "") + word.text;
  });
  const confidence = words.reduce((sum, word) => sum + word.confidence, 0) / words.length;
  return {
    text,
    confidence,
    x0: Math.min(...words.map((word) => word.x0)),
    y0: Math.min(...words.map((word) => word.y0)),
    x1: Math.max(...words.map((word) => word.x1)),
    y1: Math.max(...words.map((word) => word.y1)),
  };
}

function iou(a, b) {
  const ix0 = Math.max(a.x0, b.x0);
  const iy0 = Math.max(a.y0, b.y0);
  const ix1 = Math.min(a.x1, b.x1);
  const iy1 = Math.min(a.y1, b.y1);
  const inter = Math.max(0, ix1 - ix0) * Math.max(0, iy1 - iy0);
  if (inter <= 0) return 0;
  const areaA = (a.x1 - a.x0) * (a.y1 - a.y0);
  const areaB = (b.x1 - b.x0) * (b.y1 - b.y0);
  return inter / Math.max(1, areaA + areaB - inter);
}

function clusterWords(words, scale, profile, options) {
  const kept = [];
  (words || []).forEach((word) => {
    const text = String(word.text || "").trim();
    if (!text || word.confidence < profile.minConfidence) return;
    const box = word.bbox;
    if (!box) return;
    const w = box.x1 - box.x0;
    const h = box.y1 - box.y0;
    if (h < profile.minHeight || h > profile.maxHeight || w < 8) return;
    if (w > Math.max(90, text.length * h * profile.maxWidthFactor)) return;
    kept.push({
      text,
      confidence: word.confidence,
      x0: box.x0 / scale,
      y0: box.y0 / scale,
      x1: box.x1 / scale,
      y1: box.y1 / scale,
    });
  });
  kept.sort((a, b) => (a.y0 + a.y1) / 2 - (b.y0 + b.y1) / 2 || a.x0 - b.x0);

  const lines = [];
  kept.forEach((word) => {
    const cy = (word.y0 + word.y1) / 2;
    const h = word.y1 - word.y0;
    const line = lines.find((candidate) => {
      const lcy = median(candidate.map((item) => (item.y0 + item.y1) / 2));
      const lh = median(candidate.map((item) => item.y1 - item.y0));
      return Math.abs(cy - lcy) <= 0.45 * Math.max(h, lh);
    });
    if (line) line.push(word);
    else lines.push([word]);
  });

  const segments = [];
  lines.forEach((line) => {
    line.sort((a, b) => a.x0 - b.x0);
    const gapLimit = Math.max(8, profile.gapFactor * median(line.map((item) => item.y1 - item.y0)));
    let current = [line[0]];
    for (let i = 1; i < line.length; i += 1) {
      const word = line[i];
      if (word.x0 - current[current.length - 1].x1 > gapLimit) {
        segments.push(mergeSegment(current));
        current = [word];
      } else {
        current.push(word);
      }
    }
    segments.push(mergeSegment(current));
  });

  const visible = segments.filter((seg) => seg.text.length >= profile.minChars || seg.confidence >= 90);
  visible.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const deduped = [];
  visible.forEach((seg) => {
    if (deduped.some((prev) => iou(seg, prev) > 0.55)) return;
    deduped.push(seg);
  });
  if (options && options.wrap === false) return deduped;
  return groupWrappedLines(deduped);
}

function joinText(a, b) {
  const prev = a.slice(-1);
  const next = b.slice(0, 1);
  if (!prev || !next) return a + b;
  const space = prev.charCodeAt(0) < 128 && next.charCodeAt(0) < 128;
  return a + (space ? " " : "") + b;
}

// A wrapped title or sentence is several OCR lines of the same size, stacked
// tightly, sharing the same horizontal span. A shorter centered second line
// still overlaps the first. A label in another column does not.
function continuesBlock(block, next) {
  const prev = block.members[block.members.length - 1];
  const ph = prev.lineH;
  const nh = next.y1 - next.y0;
  const taller = Math.max(ph, nh);
  const shorter = Math.max(1, Math.min(ph, nh));
  if (taller > shorter * 1.45) return false;
  const yOverlap = Math.min(prev.y1, next.y1) - Math.max(prev.y0, next.y0);
  if (yOverlap >= shorter * 0.55) {
    if (taller > shorter * 2.2) return false;
    if (next.x0 < prev.x1 - 4) return false;
    return next.x0 <= block.x1 + shorter * 0.8;
  }
  const gap = next.y0 - prev.y1;
  if (gap > shorter * 0.55 || gap < -shorter * 0.25) return false;
  const overlap = Math.min(prev.x1, next.x1) - Math.max(prev.x0, next.x0);
  const narrow = Math.min(prev.x1 - prev.x0, next.x1 - next.x0);
  return narrow > 0 && overlap >= narrow * 0.45;
}

function blockAlign(members) {
  if (members.length < 2) return "l";
  const x0 = Math.min(...members.map((line) => line.x0));
  const x1 = Math.max(...members.map((line) => line.x1));
  const width = x1 - x0;
  if (width <= 0) return "l";
  const mid = (x0 + x1) / 2;
  const centered = members.every((line) => Math.abs((line.x0 + line.x1) / 2 - mid) <= width * 0.12);
  return centered ? "ctr" : "l";
}

function groupWrappedLines(segments) {
  const blocks = [];
  for (const seg of segments) {
    const lineH = seg.y1 - seg.y0;
    const prev = blocks[blocks.length - 1];
    if (prev && continuesBlock(prev, seg)) {
      prev.text = joinText(prev.text, seg.text);
      prev.confidence = (prev.confidence * prev.lines + seg.confidence) / (prev.lines + 1);
      prev.x0 = Math.min(prev.x0, seg.x0);
      prev.y0 = Math.min(prev.y0, seg.y0);
      prev.x1 = Math.max(prev.x1, seg.x1);
      prev.y1 = Math.max(prev.y1, seg.y1);
      prev.lineH = (prev.lineH * prev.lines + lineH) / (prev.lines + 1);
      prev.lines += 1;
      prev.members.push({ ...seg, lineH });
      continue;
    }
    blocks.push({
      ...seg,
      lineH,
      lines: 1,
      members: [{ ...seg, lineH }],
    });
  }
  return blocks.map((block) => {
    const align = blockAlign(block.members);
    return {
      text: block.text,
      confidence: block.confidence,
      x0: block.x0,
      y0: block.y0,
      x1: block.x1,
      y1: block.y1,
      lineH: block.lineH,
      lines: block.lines,
      align,
    };
  });
}

function suppressOverlaps(segments, regions) {
  const extra = regions || [];
  if (!extra.length) return segments || [];
  const kept = (segments || []).filter((seg) => !extra.some((region) => overlapRatio(seg, region) >= 0.5));
  return kept.concat(extra);
}

// A dragged rectangle is one block even when the lines inside differ in size.
function textFromRegionWords(words, scale) {
  const lines = clusterWords(words, scale, profileFromSlider(0), { wrap: false });
  const heights = lines.map((line) => line.y1 - line.y0).filter((height) => height > 0);
  const confidence = lines.length
    ? lines.reduce((sum, line) => sum + line.confidence, 0) / lines.length
    : 0;
  return {
    text: lines.map((line) => line.text).filter(Boolean).join("\n"),
    lineH: heights.length ? median(heights) : 0,
    lines: Math.max(1, lines.length),
    confidence,
  };
}

function segmentKey(seg) {
  if (seg.manualId) return `region:${seg.manualId}`;
  const round = (value) => Math.round(value / 4) * 4;
  return [round(seg.x0), round(seg.y0), round(seg.x1), round(seg.y1)].join(",");
}

function overlapRatio(a, b) {
  const ix0 = Math.max(a.x0, b.x0);
  const iy0 = Math.max(a.y0, b.y0);
  const ix1 = Math.min(a.x1, b.x1);
  const iy1 = Math.min(a.y1, b.y1);
  const inter = Math.max(0, ix1 - ix0) * Math.max(0, iy1 - iy0);
  if (inter <= 0) return 0;
  const areaA = Math.max(1, (a.x1 - a.x0) * (a.y1 - a.y0));
  const areaB = Math.max(1, (b.x1 - b.x0) * (b.y1 - b.y0));
  return inter / Math.min(areaA, areaB);
}

// Slider moves rebuild the rows. Preserve manually edited text when the new
// segment still covers substantially the same image area.
function inheritEdit(priors, seg) {
  const edited = (priors || []).find(
    (edit) => edit.edited && overlapRatio(edit, seg) >= 0.8
  );
  return {
    text: edited ? edited.text : seg.text,
    edited: Boolean(edited),
  };
}

// Tesseract boxes follow the bitmap it actually saw. An upscaled canvas reports
// coordinates past the original image; a pass that already used image pixels does not.
function resolveOcrScale(words, width, height, upscale) {
  let maxX = 0;
  let maxY = 0;
  for (const word of words || []) {
    const box = word && word.bbox;
    if (!box) continue;
    maxX = Math.max(maxX, Number(box.x1) || 0);
    maxY = Math.max(maxY, Number(box.y1) || 0);
  }
  const pastWidth = width > 0 && maxX > width + 2;
  const pastHeight = height > 0 && maxY > height + 2;
  if (!pastWidth && !pastHeight) return 1;
  const detected = Math.max(width > 0 ? maxX / width : 1, height > 0 ? maxY / height : 1);
  if (upscale > 1 && Math.abs(detected - upscale) / upscale < 0.25) return upscale;
  return detected > 0 ? detected : 1;
}

const SRC_RECT_BASE = 100000;

function segmentToEmu(pic, imageSize, seg, pad) {
  const imgW = imageSize.width;
  const imgH = imageSize.height;
  const src = pic.src || { l: 0, t: 0, visibleW: 1, visibleH: 1 };
  if (!(src.visibleW > 0) || !(src.visibleH > 0) || !imgW || !imgH) return null;
  const cropL = (src.l || 0) / SRC_RECT_BASE;
  const cropT = (src.t || 0) / SRC_RECT_BASE;
  const mapX = (px) => (px / imgW - cropL) / src.visibleW;
  const mapY = (py) => (py / imgH - cropT) / src.visibleH;
  const x0 = Math.max(0, seg.x0 - pad);
  const y0 = Math.max(0, seg.y0 - pad);
  const x1 = Math.min(imgW, seg.x1 + pad);
  const y1 = Math.min(imgH, seg.y1 + pad);
  let rx0 = mapX(x0);
  let ry0 = mapY(y0);
  let rx1 = mapX(x1);
  let ry1 = mapY(y1);
  if (rx1 < 0 || ry1 < 0 || rx0 > 1 || ry0 > 1) return null;
  rx0 = Math.max(0, rx0);
  ry0 = Math.max(0, ry0);
  rx1 = Math.min(1, rx1);
  ry1 = Math.min(1, ry1);
  if (rx1 - rx0 < 0.002 || ry1 - ry0 < 0.002) return null;
  const lineH = seg.lineH || (y1 - y0);
  const lineCy = Math.max(1, Math.round((mapY(Math.min(imgH, y0 + lineH)) - mapY(y0)) * pic.cy));
  return {
    slidePath: pic.slidePath,
    x: Math.round(pic.x + rx0 * pic.cx),
    y: Math.round(pic.y + ry0 * pic.cy),
    cx: Math.max(1, Math.round((rx1 - rx0) * pic.cx)),
    cy: Math.max(1, Math.round((ry1 - ry0) * pic.cy)),
    lineCy,
    lines: seg.lines || 1,
    align: seg.align === "ctr" ? "ctr" : "l",
  };
}

const imageTextApi = {
  profileFromSlider,
  clusterWords,
  suppressOverlaps,
  textFromRegionWords,
  overlapRatio,
  segmentKey,
  segmentToEmu,
  inheritEdit,
  resolveOcrScale,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = imageTextApi;
}

if (typeof window !== "undefined") {
  Object.assign(window, imageTextApi);
}
})();
