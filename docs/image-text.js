(function () {
"use strict";

// Word boxes come from one OCR pass. The slider only changes which of those
// boxes are kept and how far apart words may be and still form one line.

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

function clusterWords(words, scale, profile) {
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
  return deduped;
}

function segmentKey(seg) {
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

// Slider moves rebuild the rows. A new row stays unchecked unless it still
// sits on a row the user left checked. seedChecks is only for the first pass.
function inheritEdit(priors, seg, seedChecks) {
  const hits = (priors || []).filter((edit) => overlapRatio(edit, seg) >= 0.5);
  const edited = hits.find((edit) => edit.edited && overlapRatio(edit, seg) >= 0.8);
  return {
    text: edited ? edited.text : seg.text,
    checked: hits.length ? hits.every((edit) => edit.checked) : Boolean(seedChecks),
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
  return {
    slidePath: pic.slidePath,
    x: Math.round(pic.x + rx0 * pic.cx),
    y: Math.round(pic.y + ry0 * pic.cy),
    cx: Math.max(1, Math.round((rx1 - rx0) * pic.cx)),
    cy: Math.max(1, Math.round((ry1 - ry0) * pic.cy)),
  };
}

const imageTextApi = {
  profileFromSlider,
  clusterWords,
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
