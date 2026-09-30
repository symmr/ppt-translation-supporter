"use strict";

let hooks = null;
let pictures = [];
let states = new Map();
let urls = [];
let openKey = "";
let workerPromise = null;
let ocrToken = 0;
let imagePage = 0;
const thumbUrls = new Map();
const PAGE_SIZE = 10;

function bindImageTools(next) {
  hooks = next;
}

function resetImageTools() {
  ocrToken += 1;
  openKey = "";
  pictures = [];
  states = new Map();
  for (const url of urls) URL.revokeObjectURL(url);
  urls = [];
  imagePage = 0;
  thumbUrls.clear();
  const panel = document.getElementById("imagePanel");
  const list = document.getElementById("imageList");
  const pager = document.getElementById("imagePager");
  if (panel) panel.hidden = true;
  if (list) list.replaceChildren();
  if (pager) pager.hidden = true;
  closeOcrDialog();
  if (workerPromise) {
    const pending = workerPromise;
    workerPromise = null;
    pending.then((worker) => worker.terminate()).catch(() => {});
  }
}

function mimeFor(path) {
  const ext = String(path.split(".").pop() || "").toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "webp") return "image/webp";
  return "image/jpeg";
}

function fileName(path) {
  return path.split("/").pop() || path;
}

function boxPrefix(pic) {
  const occurrence = String(pic.key).split(":").pop();
  return `imgtext-${pic.slideIndex}-${occurrence}-`;
}

const PICK_PROFILES = [0, 50, 100];

function profileFor(step) {
  const index = Math.min(PICK_PROFILES.length - 1, Math.max(0, Math.round(Number(step) || 0)));
  return profileFromSlider(PICK_PROFILES[index]);
}

function setOcrStatus(text) {
  const row = document.getElementById("ocrStatusRow");
  const el = document.getElementById("ocrStatus");
  if (el) el.textContent = text || "";
  if (row) row.hidden = !text;
}

// The page-level message area sits under the dialog backdrop, so anything
// said while the dialog is open has to be shown inside it.
function setOcrNote(text, kind = "ok") {
  const el = document.getElementById("ocrNote");
  if (!el) return;
  el.textContent = text || "";
  el.className = `msg ${kind}`;
  el.hidden = !text;
}

function stopOcr() {
  if (!hooks || !hooks.isBusy()) return;
  ocrToken += 1;
  const pending = workerPromise;
  workerPromise = null;
  if (pending) pending.then((worker) => worker.terminate()).catch(() => {});
  setOcrStatus("");
  showOcrReady(false);
  hooks.setBusy(false);
  syncImagePager();
  setOcrNote("読み込みを停止しました。");
}

function showSource(pic, state) {
  const img = document.getElementById("ocrSource");
  if (!img) return;
  if (state && state.image) {
    img.src = state.image.src;
    return;
  }
  const card = document.querySelector(`.image-card[data-key="${CSS.escape(pic.key)}"] img`);
  if (card && card.src) img.src = card.src;
}

function openOcrDialog(pic, state) {
  document.getElementById("ocrTitle").textContent =
    `スライド ${pic.slideIndex}  ${fileName(pic.media)}`;
  showSource(pic, state);
  const dialog = document.getElementById("ocrDialog");
  if (!dialog.open) dialog.showModal();
}

function closeOcrDialog() {
  setOcrNote("");
  const dialog = document.getElementById("ocrDialog");
  if (dialog && dialog.open) dialog.close();
}

function medianChannel(pixels, index) {
  const values = pixels.map((pixel) => pixel[index]).sort((a, b) => a - b);
  return values[Math.floor(values.length / 2)] || 0;
}

function rgbHex(pixel) {
  return pixel.map((n) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function sampleCanvas(state) {
  if (state.sampleCanvas) return state.sampleCanvas;
  const canvas = document.createElement("canvas");
  canvas.width = state.image.naturalWidth;
  canvas.height = state.image.naturalHeight;
  canvas.getContext("2d").drawImage(state.image, 0, 0);
  state.sampleCanvas = canvas;
  return canvas;
}

function sampleColors(state, seg) {
  const canvas = sampleCanvas(state);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const pad = 2;
  const x0 = Math.max(0, Math.floor(seg.x0) - pad);
  const y0 = Math.max(0, Math.floor(seg.y0) - pad);
  const x1 = Math.min(canvas.width - 1, Math.ceil(seg.x1) + pad);
  const y1 = Math.min(canvas.height - 1, Math.ceil(seg.y1) + pad);
  const rw = Math.max(1, x1 - x0 + 1);
  const rh = Math.max(1, y1 - y0 + 1);
  const data = ctx.getImageData(x0, y0, rw, rh).data;
  const at = (x, y) => {
    const i = ((y - y0) * rw + (x - x0)) * 4;
    return [data[i], data[i + 1], data[i + 2]];
  };
  const ring = [];
  for (let x = x0; x <= x1; x += 1) ring.push(at(x, y0), at(x, y1));
  for (let y = y0; y <= y1; y += 1) ring.push(at(x0, y), at(x1, y));
  const fill = [0, 1, 2].map((index) => medianChannel(ring, index));
  const inkPixels = [];
  for (let y = y0; y <= y1; y += 2) {
    for (let x = x0; x <= x1; x += 2) {
      const color = at(x, y);
      const dist = Math.hypot(color[0] - fill[0], color[1] - fill[1], color[2] - fill[2]);
      if (dist > 50) inkPixels.push(color);
    }
  }
  const lum = 0.2126 * fill[0] + 0.7152 * fill[1] + 0.0722 * fill[2];
  const ink = inkPixels.length < 8
    ? (lum < 150 ? [255, 255, 255] : [30, 30, 30])
    : [0, 1, 2].map((index) => medianChannel(inkPixels, index));
  return { fill: rgbHex(fill), ink: rgbHex(ink) };
}

function currentSegments(state) {
  const auto = state.words ? clusterWords(state.words, state.scale, profileFor(state.slider)) : [];
  return suppressOverlaps(auto, state.regions);
}

function editFor(state, seg) {
  const key = segmentKey(seg);
  if (state.edits.has(key)) return state.edits.get(key);
  const inherited = inheritEdit(state.priorEdits, seg, state.seedChecks !== false);
  const record = {
    text: inherited.text,
    checked: inherited.checked,
    edited: inherited.edited,
    x0: seg.x0,
    y0: seg.y0,
    x1: seg.x1,
    y1: seg.y1,
  };
  state.edits.set(key, record);
  return record;
}

function countLine(state) {
  const segments = currentSegments(state);
  let adopted = 0;
  let chars = 0;
  for (const seg of segments) {
    const edit = editFor(state, seg);
    const text = edit.text.trim();
    if (!edit.checked || !text) continue;
    adopted += 1;
    chars += text.length;
  }
  return `表示 ${segments.length} 行 / 抽出する ${adopted} 行 / ${chars} 文字`;
}

function drawCrop(canvas, image, seg) {
  const ctx = canvas.getContext("2d");
  const pad = 6;
  const sx = Math.max(0, seg.x0 - pad);
  const sy = Math.max(0, seg.y0 - pad);
  const sw = Math.max(1, Math.min(image.naturalWidth, seg.x1 + pad) - sx);
  const sh = Math.max(1, Math.min(image.naturalHeight, seg.y1 + pad) - sy);
  ctx.fillStyle = "#f5f5f4";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  const scale = Math.min(canvas.width / sw, canvas.height / sh);
  const dw = sw * scale;
  const dh = sh * scale;
  ctx.drawImage(image, sx, sy, sw, sh, (canvas.width - dw) / 2, (canvas.height - dh) / 2, dw, dh);
}

function renderCandidates() {
  const state = states.get(openKey);
  const list = document.getElementById("ocrList");
  const count = document.getElementById("ocrCount");
  const thresholds = document.getElementById("ocrThresholds");
  const slider = document.getElementById("imagePick");
  if (!state || (!state.words && !(state.regions && state.regions.length))) {
    list.replaceChildren();
    count.textContent = "";
    thresholds.textContent = "";
    paintOverlay();
    return;
  }
  state.priorEdits = [...state.edits.values()];
  slider.value = String(state.slider);
  const profile = profileFor(state.slider);
  thresholds.textContent =
    `信頼度 ${Math.round(profile.minConfidence)}% 以上、近い単語の結合 ${profile.gapFactor.toFixed(2)}`;
  const segments = currentSegments(state);
  list.replaceChildren();
  segments.forEach((seg, index) => {
    const edit = editFor(state, seg);
    const row = document.createElement("li");
    row.className = edit.checked ? "ocr-row" : "ocr-row is-off";

    const label = document.createElement("label");
    label.className = "check";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = edit.checked;
    box.addEventListener("change", () => {
      edit.checked = box.checked;
      renderCandidates();
    });
    const caption = document.createElement("span");
    caption.textContent = "抽出する";
    label.append(box, caption);

    const crop = document.createElement("canvas");
    crop.className = "crop";
    crop.width = 160;
    crop.height = 44;
    drawCrop(crop, state.image, seg);

    const field = document.createElement(seg.manualId ? "textarea" : "input");
    if (!seg.manualId) field.type = "text";
    field.value = edit.text;
    field.setAttribute("aria-label", `原文 ${index + 1}`);
    field.addEventListener("input", () => {
      edit.text = field.value;
      edit.edited = true;
      count.textContent = countLine(state);
    });

    const conf = document.createElement("span");
    conf.className = "conf";
    if (seg.manualId) {
      const drop = document.createElement("button");
      drop.type = "button";
      drop.textContent = "消す";
      drop.addEventListener("click", () => removeRegion(seg.manualId));
      conf.append(drop);
    } else {
      conf.textContent = `${Math.round(seg.confidence)}%`;
    }
    row.append(label, crop, field, conf);
    list.append(row);
  });
  count.textContent = countLine(state);
  markCards();
  paintOverlay();
}

function markCards() {
  for (const card of document.querySelectorAll(".image-card")) {
    const key = card.getAttribute("data-key");
    card.classList.toggle("is-on", key === openKey);
    const state = states.get(key);
    const badge = card.querySelector(".badge");
    if (!badge) continue;
    if (!state || !state.words) {
      badge.textContent = "";
      continue;
    }
    const segments = currentSegments(state);
    let adopted = 0;
    for (const seg of segments) {
      const edit = state.edits.get(segmentKey(seg));
      if (edit && edit.checked && edit.text.trim()) adopted += 1;
    }
    badge.textContent = `抽出 ${adopted}`;
  }
}

function wordsFromResult(data) {
  if (Array.isArray(data.words) && data.words.length) return data.words;
  const words = [];
  for (const block of data.blocks || []) {
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        for (const word of line.words || []) words.push(word);
      }
    }
  }
  return words;
}

function getWorker() {
  if (!workerPromise) {
    if (typeof Tesseract === "undefined") {
      return Promise.reject(new Error("OCR エンジンを読み込めませんでした"));
    }
    workerPromise = Tesseract.createWorker("jpn+eng", 1, {
      logger(message) {
        if (message.status === "recognizing text") {
          const pct = Math.round((message.progress || 0) * 100);
          setOcrStatus(`画像を読み込み中… ${pct}%`);
        } else if (message.status) {
          setOcrStatus(`画像を読み込み中… ${message.status}`);
        }
      },
    }).then(async (worker) => {
      await worker.setParameters({ tessedit_pageseg_mode: "11" });
      return worker;
    });
  }
  return workerPromise;
}

async function loadImage(pic) {
  const entry = hooks.getZip().file(pic.media);
  if (!entry) throw new Error(`${pic.media} が PPTX 内にありません`);
  const bytes = await entry.async("uint8array");
  const url = URL.createObjectURL(new Blob([bytes], { type: mimeFor(pic.media) }));
  urls.push(url);
  const image = new Image();
  image.src = url;
  await image.decode();
  return image;
}

async function selectPicture(pic) {
  if (!hooks || hooks.isBusy()) return;
  openKey = pic.key;
  let state = states.get(pic.key);
  if (!state) {
    state = { slider: 0, edits: new Map(), words: null, scale: 1, image: null, seedChecks: true, regions: [] };
    states.set(pic.key, state);
  }
  const index = pictures.findIndex((item) => item.key === pic.key);
  const page = index < 0 ? imagePage : Math.floor(index / PAGE_SIZE);
  if (index >= 0 && page !== imagePage) {
    imagePage = page;
    await renderImagePage();
  } else {
    markCards();
  }
  openOcrDialog(pic, state);
  setOcrNote("");
  document.getElementById("imagePick").value = String(state.slider);
  syncOcrNav();
  if ((state.words && state.image) || (state.regions || []).length) {
    setOcrStatus("");
    showOcrReady(true);
    renderCandidates();
    return;
  }
  document.getElementById("ocrList").replaceChildren();
  document.getElementById("ocrCount").textContent = "";
  document.getElementById("ocrThresholds").textContent = "";
  setOcrStatus("");
  showOcrReady(false);
}

async function startOcr() {
  if (!hooks || hooks.isBusy()) return;
  const pic = pictures.find((item) => item.key === openKey);
  const state = pic && states.get(pic.key);
  if (!pic || !state) return;
  if (state.words && state.image) {
    showOcrReady(true);
    renderCandidates();
    return;
  }
  hooks.clearMessages();
  setOcrNote("");
  const token = ocrToken;
  let outcome = "error";
  try {
    hooks.setBusy(true, `${fileName(pic.media)} を読み取り中…`);
    setOcrStatus("画像を読み込み中…");
    const image = await loadImage(pic);
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    const scale = Math.max(1, Math.min(2, 4000 / Math.max(image.naturalWidth, image.naturalHeight)));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(image.naturalWidth * scale);
    canvas.height = Math.round(image.naturalHeight * scale);
    canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
    const worker = await getWorker();
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    const { data } = await worker.recognize(canvas);
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    const words = wordsFromResult(data);
    state.image = image;
    state.scale = resolveOcrScale(
      words,
      image.naturalWidth,
      image.naturalHeight,
      canvas.width / image.naturalWidth
    );
    state.words = words;
    state.pic = pic;
    showSource(pic, state);
    setOcrStatus("");
    showOcrReady(true);
    renderCandidates();
    outcome = "ok";
    setOcrNote(`${fileName(pic.media)} を読みました。不要な行はチェックを外すか、一括除外を押してください。`);
  } catch (err) {
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    console.error("[ppt-translation-supporter] 画像 OCR:", err);
    setOcrNote(`画像の読み取りに失敗しました: ${err.message || err}`, "error");
    showOcrReady(false);
  } finally {
    if (token !== ocrToken) return;
    hooks.setBusy(false);
    syncImagePager();
    if (outcome === "cancel") {
      setOcrStatus("");
      showOcrReady(false);
      setOcrNote("読み込みを停止しました。");
    }
  }
}

async function stepPicture(delta) {
  if (!hooks || hooks.isBusy()) return;
  const index = pictures.findIndex((item) => item.key === openKey);
  const pic = pictures[index + delta];
  if (!pic) return;
  await selectPicture(pic);
}

async function openPanel() {
  if (!hooks || hooks.isBusy()) return;
  const panel = document.getElementById("imagePanel");
  panel.hidden = false;
  if (pictures.length) return;
  hooks.clearMessages();
  try {
    hooks.setBusy(true, "画像を列挙中…");
    pictures = await listSlideRasterPictures(hooks.getZip());
    const list = document.getElementById("imageList");
    list.replaceChildren();
    if (!pictures.length) {
      const empty = document.createElement("p");
      empty.className = "meta";
      empty.textContent = "スライドに貼られた png / jpeg / webp はありません。";
      list.append(empty);
      return;
    }
    imagePage = 0;
    await renderImagePage();
  } catch (err) {
    console.error("[ppt-translation-supporter] 画像一覧:", err);
    hooks.showError(`画像一覧の取得に失敗しました: ${err.message || err}`);
  } finally {
    hooks.setBusy(false);
    syncImagePager();
  }
}

function syncImagePager() {
  const pager = document.getElementById("imagePager");
  if (!pager) return;
  const total = pictures.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (imagePage > pages - 1) imagePage = Math.max(0, pages - 1);
  pager.hidden = total <= PAGE_SIZE;
  const start = imagePage * PAGE_SIZE;
  const end = Math.min(total, start + PAGE_SIZE);
  const label = document.getElementById("imagePageLabel");
  if (label) label.textContent = total ? `${start + 1}–${end} / ${total}` : "";
  const prev = document.getElementById("imagePrev");
  const next = document.getElementById("imageNext");
  if (prev) prev.disabled = imagePage <= 0;
  if (next) next.disabled = imagePage >= pages - 1;
  syncOcrNav();
}

function syncOcrNav() {
  const index = pictures.findIndex((item) => item.key === openKey);
  const prev = document.getElementById("ocrPrev");
  const next = document.getElementById("ocrNext");
  if (prev) prev.disabled = index <= 0;
  if (next) next.disabled = index < 0 || index >= pictures.length - 1;
}

function showOcrReady(ready) {
  const panel = document.getElementById("ocrPanel");
  const start = document.getElementById("startOcrBtn");
  const state = states.get(openKey);
  if (panel) panel.hidden = !ready;
  if (start) start.hidden = Boolean(ready && state && state.words);
}

async function thumbUrl(pic) {
  if (thumbUrls.has(pic.key)) return thumbUrls.get(pic.key);
  const entry = hooks.getZip().file(pic.media);
  if (!entry) return "";
  const url = URL.createObjectURL(new Blob([await entry.async("uint8array")], { type: mimeFor(pic.media) }));
  urls.push(url);
  thumbUrls.set(pic.key, url);
  return url;
}

async function renderImagePage() {
  const list = document.getElementById("imageList");
  list.replaceChildren();
  const slice = pictures.slice(imagePage * PAGE_SIZE, imagePage * PAGE_SIZE + PAGE_SIZE);
  for (const pic of slice) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "image-card";
    card.setAttribute("data-key", pic.key);
    const img = document.createElement("img");
    img.alt = "";
    img.src = await thumbUrl(pic);
    const caption = document.createElement("span");
    caption.textContent = `スライド ${pic.slideIndex}  ${fileName(pic.media)}`;
    const badge = document.createElement("span");
    badge.className = "badge";
    card.append(img, caption, badge);
    card.addEventListener("click", () => selectPicture(pic));
    list.append(card);
  }
  markCards();
  syncImagePager();
}

async function changeImagePage(delta) {
  if (!hooks || hooks.isBusy()) return;
  const pages = Math.max(1, Math.ceil(pictures.length / PAGE_SIZE));
  const next = Math.min(pages - 1, Math.max(0, imagePage + delta));
  if (next === imagePage) return;
  imagePage = next;
  await renderImagePage();
}

function excludeAll() {
  const state = states.get(openKey);
  if (!state || (!state.words && !(state.regions || []).length)) return;
  state.seedChecks = false;
  for (const seg of currentSegments(state)) {
    const edit = editFor(state, seg);
    edit.checked = false;
  }
  renderCandidates();
}

function boxesFor(pic, state) {
  const boxes = [];
  for (const seg of currentSegments(state)) {
    const edit = editFor(state, seg);
    const text = edit.text.trim();
    if (!edit.checked || !text) continue;
    const emu = segmentToEmu(pic, {
      width: state.image.naturalWidth,
      height: state.image.naturalHeight,
    }, seg, 0);
    if (!emu) continue;
    const colors = sampleColors(state, seg);
    boxes.push({ ...emu, text, fill: colors.fill, ink: colors.ink });
  }
  return boxes;
}

async function applyBoxes() {
  if (!hooks || hooks.isBusy()) return;
  const pic = pictures.find((item) => item.key === openKey);
  const state = pic && states.get(pic.key);
  if (!pic || !state || !state.image) {
    setOcrNote("載せる行がありません。抽出する行にチェックを入れてください。", "error");
    return;
  }
  const boxes = boxesFor(pic, state);
  const zip = hooks.getZip();
  const already = await zip.file(pic.slidePath).async("string");
  const prefix = boxPrefix(pic);
  const hadBoxes = already.includes(prefix);
  if (!boxes.length && !hadBoxes) {
    setOcrNote("載せる行がありません。抽出する行にチェックを入れてください。", "error");
    return;
  }
  hooks.clearMessages();
  try {
    hooks.setBusy(true, "テキストボックスを追加中…");
    setOcrNote("テキストボックスを追加中…", "busy");
    const result = await replacePictureTextBoxes(zip, pic.slidePath, prefix, boxes);
    setOcrNote("テキストを抽出し直しています…", "busy");
    hooks.markDeckEdited();
    const extracted = await extractTextsFromZip(zip);
    const hadPaste = hooks.consumePaste();
    hooks.applyExtract(extracted);
    markCards();
    closeOcrDialog();
    const note = hadPaste ? " 以前貼った訳文は uid が変わるので消しました。" : "";
    if (result.added && result.removed) {
      hooks.showOk(`テキストボックスを ${result.added} 個に差し替え、抽出テキストを更新しました。${note}`);
    } else if (result.added) {
      hooks.showOk(`テキストボックスを ${result.added} 個追加し、抽出テキストを更新しました。${note}`);
    } else {
      hooks.showOk(`この画像のテキストボックスを外し、抽出テキストを更新しました。${note}`);
    }
  } catch (err) {
    console.error("[ppt-translation-supporter] テキストボックス追加:", err);
    setOcrNote(`テキストボックスの追加に失敗しました: ${err.message || err}`, "error");
  } finally {
    hooks.setBusy(false);
    syncImagePager();
  }
}

let regionDrag = null;

function imageFrame(img) {
  const nw = img.naturalWidth;
  const nh = img.naturalHeight;
  if (!nw || !nh) return null;
  const rect = img.getBoundingClientRect();
  const scale = Math.min(rect.width / nw, rect.height / nh);
  return {
    scale,
    offsetX: (rect.width - nw * scale) / 2,
    offsetY: (rect.height - nh * scale) / 2,
    width: rect.width,
    height: rect.height,
  };
}

function eventToImagePx(event, img) {
  const frame = imageFrame(img);
  if (!frame) return null;
  const rect = img.getBoundingClientRect();
  const x = (event.clientX - rect.left - frame.offsetX) / frame.scale;
  const y = (event.clientY - rect.top - frame.offsetY) / frame.scale;
  return {
    x: Math.max(0, Math.min(img.naturalWidth, x)),
    y: Math.max(0, Math.min(img.naturalHeight, y)),
  };
}

function normalizeRect(a, b) {
  if (!a || !b) return null;
  const x0 = Math.min(a.x, b.x);
  const y0 = Math.min(a.y, b.y);
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  if (x1 - x0 < 8 || y1 - y0 < 8) return null;
  return { x0, y0, x1, y1 };
}

function paintOverlay() {
  const img = document.getElementById("ocrSource");
  const canvas = document.getElementById("ocrOverlay");
  if (!img || !canvas) return;
  const frame = imageFrame(img);
  if (!frame) return;
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.round(frame.width * dpr));
  canvas.height = Math.max(1, Math.round(frame.height * dpr));
  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, frame.width, frame.height);
  const state = states.get(openKey);
  const rects = state && state.regions ? state.regions.slice() : [];
  if (regionDrag) {
    const live = normalizeRect(regionDrag.start, regionDrag.current);
    if (live) rects.push(live);
  }
  for (const rect of rects) {
    const x = frame.offsetX + rect.x0 * frame.scale;
    const y = frame.offsetY + rect.y0 * frame.scale;
    const w = (rect.x1 - rect.x0) * frame.scale;
    const h = (rect.y1 - rect.y0) * frame.scale;
    ctx.fillStyle = "rgba(37, 99, 235, 0.18)";
    ctx.strokeStyle = "#2563eb";
    ctx.lineWidth = 2;
    ctx.fillRect(x, y, w, h);
    ctx.strokeRect(x, y, w, h);
  }
}

function removeRegion(id) {
  const state = states.get(openKey);
  if (!state || !state.regions) return;
  state.regions = state.regions.filter((region) => region.manualId !== id);
  state.edits.delete(`region:${id}`);
  renderCandidates();
}

async function ocrRegion(rect) {
  if (!hooks || hooks.isBusy()) return;
  const pic = pictures.find((item) => item.key === openKey);
  const state = pic && states.get(pic.key);
  if (!pic || !state) return;
  const token = ocrToken;
  let outcome = "error";
  try {
    hooks.setBusy(true, "範囲を読み取り中…");
    setOcrNote("範囲を読み取り中…", "busy");
    setOcrStatus("範囲を読み取り中…");
    if (!state.image) {
      state.image = await loadImage(pic);
      showSource(pic, state);
    }
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    const w = Math.round(rect.x1 - rect.x0);
    const h = Math.round(rect.y1 - rect.y0);
    const scale = Math.max(1, Math.min(2, 4000 / Math.max(w, h)));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    canvas.getContext("2d").drawImage(
      state.image,
      rect.x0,
      rect.y0,
      w,
      h,
      0,
      0,
      canvas.width,
      canvas.height
    );
    const worker = await getWorker();
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    await worker.setParameters({ tessedit_pageseg_mode: "6" });
    let data;
    try {
      ({ data } = await worker.recognize(canvas));
    } finally {
      await worker.setParameters({ tessedit_pageseg_mode: "11" }).catch(() => {});
    }
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    const words = wordsFromResult(data);
    const ocrScale = resolveOcrScale(words, w, h, scale);
    const built = textFromRegionWords(words, ocrScale);
    if (!built.text.trim()) {
      setOcrNote("範囲内から文字を読めませんでした。", "error");
      outcome = "empty";
      return;
    }
    const region = {
      manualId: `r${Date.now().toString(36)}`,
      text: built.text,
      confidence: built.confidence,
      x0: rect.x0,
      y0: rect.y0,
      x1: rect.x1,
      y1: rect.y1,
      lineH: built.lineH || h / Math.max(1, built.lines),
      lines: built.lines,
      align: "l",
    };
    state.regions = (state.regions || []).filter((item) => overlapRatio(item, region) < 0.5);
    state.regions.push(region);
    state.edits.set(segmentKey(region), {
      text: region.text,
      checked: true,
      edited: false,
      x0: region.x0,
      y0: region.y0,
      x1: region.x1,
      y1: region.y1,
    });
    setOcrStatus("");
    showOcrReady(true);
    renderCandidates();
    outcome = "ok";
    setOcrNote("範囲を1件にしました。重なっていた行は置き換えています。");
  } catch (err) {
    if (token !== ocrToken) {
      outcome = "cancel";
      return;
    }
    console.error("[ppt-translation-supporter] 範囲 OCR:", err);
    setOcrNote(`範囲の読み取りに失敗しました: ${err.message || err}`, "error");
  } finally {
    if (token !== ocrToken) return;
    hooks.setBusy(false);
    syncImagePager();
    if (outcome === "cancel") {
      setOcrStatus("");
      setOcrNote("読み込みを停止しました。");
    }
    paintOverlay();
  }
}

function bindRegionDraw() {
  const canvas = document.getElementById("ocrOverlay");
  const img = document.getElementById("ocrSource");
  if (!canvas || canvas.dataset.bound) return;
  canvas.dataset.bound = "1";
  img.addEventListener("load", paintOverlay);
  window.addEventListener("resize", paintOverlay);
  canvas.addEventListener("pointerdown", (event) => {
    if (!hooks || hooks.isBusy()) return;
    if (!img.naturalWidth) return;
    const start = eventToImagePx(event, img);
    if (!start) return;
    regionDrag = { start, current: start, pointerId: event.pointerId };
    canvas.setPointerCapture(event.pointerId);
    paintOverlay();
  });
  canvas.addEventListener("pointermove", (event) => {
    if (!regionDrag || event.pointerId !== regionDrag.pointerId) return;
    regionDrag.current = eventToImagePx(event, img);
    paintOverlay();
  });
  const finish = (event) => {
    if (!regionDrag || event.pointerId !== regionDrag.pointerId) return;
    const rect = normalizeRect(regionDrag.start, regionDrag.current);
    regionDrag = null;
    paintOverlay();
    if (rect) ocrRegion(rect);
  };
  canvas.addEventListener("pointerup", finish);
  canvas.addEventListener("pointercancel", () => {
    regionDrag = null;
    paintOverlay();
  });
}

document.getElementById("openImageBtn").addEventListener("click", openPanel);
document.getElementById("imagePrev").addEventListener("click", () => changeImagePage(-1));
document.getElementById("imageNext").addEventListener("click", () => changeImagePage(1));
document.getElementById("ocrPrev").addEventListener("click", () => stepPicture(-1));
document.getElementById("ocrNext").addEventListener("click", () => stepPicture(1));
document.getElementById("startOcrBtn").addEventListener("click", startOcr);
document.getElementById("stopOcrBtn").addEventListener("click", stopOcr);
document.getElementById("excludeAllBtn").addEventListener("click", excludeAll);
document.getElementById("applyImageBtn").addEventListener("click", applyBoxes);
document.getElementById("closeOcrBtn").addEventListener("click", closeOcrDialog);
document.getElementById("ocrDialog").addEventListener("click", (event) => {
  if (event.target === event.currentTarget) closeOcrDialog();
});
bindRegionDraw();
document.getElementById("imagePick").addEventListener("input", () => {
  const state = states.get(openKey);
  if (!state) return;
  state.slider = Number(document.getElementById("imagePick").value);
  state.seedChecks = false;
  renderCandidates();
});
