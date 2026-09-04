const pptxDrop = document.getElementById("pptxDrop");
const txtDrop = document.getElementById("txtDrop");
const fileInput = document.getElementById("fileInput");
const txtInput = document.getElementById("txtInput");
const step1 = document.getElementById("step1");
const step2 = document.getElementById("step2");
const step3 = document.getElementById("step3");
const step4 = document.getElementById("step4");
const pptxMeta = document.getElementById("pptxMeta");
const extractMeta = document.getElementById("extractMeta");
const txtMeta = document.getElementById("txtMeta");
const resultMeta = document.getElementById("resultMeta");
const promptBox = document.getElementById("promptBox");
const extractBox = document.getElementById("extractBox");
const errorMsg = document.getElementById("errorMsg");
const warnMsg = document.getElementById("warnMsg");
const okMsg = document.getElementById("okMsg");
const busyMsg = document.getElementById("busyMsg");
const titleFontSelect = document.getElementById("titleFontSelect");
const bodyFontSelect = document.getElementById("bodyFontSelect");
const titleCustomFont = document.getElementById("titleCustomFont");
const bodyCustomFont = document.getElementById("bodyCustomFont");
const titleFontPreview = document.getElementById("titleFontPreview");
const bodyFontPreview = document.getElementById("bodyFontPreview");

const KEEP_VALUE = "";
const CUSTOM_VALUE = "__custom__";

const PROMPT_STORAGE_KEY = "ppt-translation-supporter:prompt";

let sourceFile = null;
let extracted = null;
let sourceZip = null;
let resultBlob = null;
let resultName = "";
let busy = false;
let deckFonts = [];

function show(el, text, className) {
  if (!text) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.textContent = text;
  if (className) el.className = `msg ${className}`;
}

function logError(context, err) {
  console.error(`[ppt-translation-supporter] ${context}:`, err);
  if (err && err.cause) console.error(`[ppt-translation-supporter] ${context} (cause):`, err.cause);
}

function clearMessages() {
  show(errorMsg, "");
  show(warnMsg, "");
  show(okMsg, "");
}

function setStepState(el, state) {
  el.classList.remove("is-wait", "is-done");
  if (state) el.classList.add(state);
}

// Long PPTX reads and zip generation block for seconds on a real deck. Without
// a visible busy state users click again, which used to start concurrent runs
// and download the file several times over.
function setBusy(on, label) {
  busy = on;
  document.body.classList.toggle("is-busy", on);
  for (const el of document.querySelectorAll("button, .opt input, .opt select")) el.disabled = on;
  show(busyMsg, on ? label || "処理中…" : "", "busy");
}

function setProgress(label) {
  if (busy) show(busyMsg, label, "busy");
}

function addOption(parent, value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  parent.appendChild(option);
}

// Same grouping as PPT Finalizer's picker: presets, then whatever the deck
// already uses, then free text. "指定しない" keeps the deck's fonts untouched
// and stays the default, since this tool only rewrites translated runs.
function populateFontSelect(selectEl, customInput, deckFonts) {
  const previous = selectEl.value;
  selectEl.innerHTML = "";
  addOption(selectEl, KEEP_VALUE, "指定しない（元のまま）");

  const presets = document.createElement("optgroup");
  presets.label = "おすすめ";
  for (const name of PRESET_FONTS) addOption(presets, name, name);
  selectEl.appendChild(presets);

  const presetSet = new Set(PRESET_FONTS);
  const inDeck = deckFonts.filter(({ name }) => !presetSet.has(name));
  if (inDeck.length) {
    const group = document.createElement("optgroup");
    group.label = "この PPTX 内";
    for (const { name, count } of inDeck) addOption(group, name, `${name}（${count} 箇所）`);
    selectEl.appendChild(group);
  }

  addOption(selectEl, CUSTOM_VALUE, "その他（自由入力）");

  const known = new Set([KEEP_VALUE, CUSTOM_VALUE, ...PRESET_FONTS, ...inDeck.map((f) => f.name)]);
  selectEl.value = known.has(previous) ? previous : KEEP_VALUE;
  customInput.hidden = selectEl.value !== CUSTOM_VALUE;
}

function fontFromPicker(selectEl, customInput) {
  if (selectEl.value === CUSTOM_VALUE) return customInput.value.trim();
  return selectEl.value;
}

function updateFontPreview(selectEl, customInput, previewEl) {
  const font = fontFromPicker(selectEl, customInput);
  previewEl.style.fontFamily = font ? `"${font}", sans-serif` : "";
  previewEl.hidden = !font;
}

function bindFontPicker(selectEl, customInput, previewEl) {
  selectEl.addEventListener("change", () => {
    customInput.hidden = selectEl.value !== CUSTOM_VALUE;
    if (!customInput.hidden) customInput.focus();
    updateFontPreview(selectEl, customInput, previewEl);
  });
  customInput.addEventListener("input", () => updateFontPreview(selectEl, customInput, previewEl));
  updateFontPreview(selectEl, customInput, previewEl);
}

function defaultPrompt() {
  return TRANSLATION_PROMPT.trim() + "\n";
}

function loadStoredPrompt() {
  try {
    return localStorage.getItem(PROMPT_STORAGE_KEY) || "";
  } catch (_) {
    return "";
  }
}

function storePrompt(value) {
  try {
    if (value && value.trim() && value !== defaultPrompt()) {
      localStorage.setItem(PROMPT_STORAGE_KEY, value);
    } else {
      localStorage.removeItem(PROMPT_STORAGE_KEY);
    }
  } catch (_) {
    // a browser with storage blocked still works, the prompt just won't persist
  }
}

function bindDrop(zone, onFile, acceptTest) {
  zone.addEventListener("click", () => zone.dispatchEvent(new CustomEvent("pick")));
  zone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      zone.click();
    }
  });
  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("is-drag");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("is-drag"));
  zone.addEventListener("drop", async (event) => {
    event.preventDefault();
    zone.classList.remove("is-drag");
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    if (acceptTest && !acceptTest(file)) {
      show(errorMsg, "対応していないファイルです。");
      return;
    }
    await onFile(file);
  });
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    show(okMsg, "コピーしました。");
  } catch (err) {
    logError("クリップボードへのコピー", err);
    show(errorMsg, `コピーに失敗しました（${err.name || "Error"}）。テキストを選択して手動でコピーしてください。`);
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function handlePptx(file) {
  if (busy) return;
  clearMessages();
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    show(errorMsg, ".pptx を置いてください。");
    return;
  }
  sourceFile = file;
  resultBlob = null;
  resultName = "";
  try {
    setBusy(true, `${file.name} を読み込み中…`);
    sourceZip = await JSZip.loadAsync(await file.arrayBuffer());
    setProgress("テキストを抽出中…");
    extracted = await extractTextsFromZip(sourceZip);
    deckFonts = await collectFontsFromZip(sourceZip);
  } catch (err) {
    logError("PPTX の抽出", err);
    show(errorMsg, `抽出に失敗しました: ${err.message || err}`);
    return;
  } finally {
    setBusy(false);
  }

  pptxMeta.hidden = false;
  pptxMeta.textContent = `${file.name} · ${(file.size / 1024).toFixed(1)} KB`;
  extractMeta.textContent = `${extracted.slideCount} 枚 · ${extracted.uidCount} 件のテキスト`;
  promptBox.value = loadStoredPrompt() || defaultPrompt();
  extractBox.value = extracted.text;
  populateFontSelect(titleFontSelect, titleCustomFont, deckFonts);
  populateFontSelect(bodyFontSelect, bodyCustomFont, deckFonts);
  updateFontPreview(titleFontSelect, titleCustomFont, titleFontPreview);
  updateFontPreview(bodyFontSelect, bodyCustomFont, bodyFontPreview);
  setStepState(step1, "is-done");
  step2.hidden = false;
  setStepState(step2, "");
  step3.hidden = false;
  setStepState(step3, "");
  step4.hidden = true;
  setStepState(step4, "is-wait");
  txtMeta.hidden = true;
  show(okMsg, "抽出しました。プロンプトと本文をコピーして翻訳し、訳文をドロップまたは貼り付けてください。");
}

async function handleTxt(file) {
  if (busy) return;
  let source;
  try {
    source = await file.text();
  } catch (err) {
    logError("翻訳テキストファイルの読み込み", err);
    show(errorMsg, `テキストの読み込みに失敗しました: ${err.message || err}`);
    return;
  }
  await applyTranslation(source, file.name);
}

async function applyTranslation(source, sourceLabel) {
  if (busy) return;
  clearMessages();
  if (!sourceZip || !extracted) {
    show(errorMsg, "先に PPTX を置いてください。");
    return;
  }
  let text = String(source || "").replace(/^\uFEFF/, "").trim();
  if (text.startsWith("```")) {
    text = text.replace(/^```[a-zA-Z]*\r?\n/, "").replace(/\r?\n```\s*$/, "");
  }
  if (!text.trim()) {
    show(errorMsg, "翻訳テキストが空です。");
    return;
  }

  const translations = parseUidDelimitedText(text);
  const report = validateTranslations(extracted.metadata, translations);
  const matched = report.expected - report.missing.length;
  txtMeta.hidden = false;
  txtMeta.textContent = report.ok
    ? `${sourceLabel} · 翻訳 ${report.got} 件 / 抽出 ${report.expected} 件`
    : `${sourceLabel} · 一致 ${matched} 件 / 抽出 ${report.expected} 件`;

  let mismatchNote = "";
  if (!report.ok) {
    const bits = [];
    if (report.missing.length) bits.push(`不足 ${report.missing.slice(0, 8).join(", ")}${report.missing.length > 8 ? " …" : ""}`);
    if (report.extra.length) bits.push(`余分 ${report.extra.slice(0, 5).join(", ")}`);
    mismatchNote = `uid が一致しません。見つかった分だけ書き戻します。${bits.join(" / ")}`;
  }

  // A new attempt invalidates the previous result, so a failure below can
  // never leave the old file downloadable behind an error message.
  resultBlob = null;
  resultName = "";
  resultMeta.textContent = "";
  step4.hidden = true;
  setStepState(step4, "is-wait");

  try {
    setBusy(true, "PPTX を読み込み中…");
    const zip = await JSZip.loadAsync(await sourceFile.arrayBuffer());
    setProgress("訳文を書き戻し中…");
    const fonts = {
      titleFont: fontFromPicker(titleFontSelect, titleCustomFont),
      bodyFont: fontFromPicker(bodyFontSelect, bodyCustomFont),
    };
    const result = await injectTextsToZip(zip, translations, extracted.metadata, fonts);
    const blob = await zip.generateAsync(
      {
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      },
      (meta) => setProgress(`PPTX を生成中… ${Math.round(meta.percent)}%`)
    );
    resultBlob = blob;
    resultName = translatedOutputName(sourceFile.name);
    const flattened = result.flattened || [];
    resultMeta.textContent = `${resultName} · 書き戻し ${result.injected} 件` +
      (result.missing ? ` · スキップ ${result.missing} 件` : "") +
      (flattened.length ? ` · 書式維持できず ${flattened.length} 件` : "");
    step4.hidden = false;
    setStepState(step3, "is-done");
    setStepState(step4, "");
    downloadBlob(resultBlob, resultName);

    const notes = [];
    if (mismatchNote) notes.push(mismatchNote);
    if (flattened.length) {
      notes.push(
        `タグ（[0]...[/0]）が原文と一致しないため、${flattened.length} 件は段落内の書式（色分けなど）を維持できませんでした: ` +
        `${flattened.slice(0, 8).join(", ")}${flattened.length > 8 ? " …" : ""}`
      );
    }
    const artifacts = result.tagArtifacts || [];
    if (artifacts.length) {
      notes.push(
        `${artifacts.length} 件の訳文にタグ記法が残っています。そのまま書き戻したので、原文由来の表記ならそのままで問題ありません: ` +
        `${artifacts.slice(0, 8).join(", ")}${artifacts.length > 8 ? " …" : ""}`
      );
    }
    if (notes.length) {
      show(warnMsg, `${notes.join("。")}。書き戻しは完了しました。ダウンロードが始まらない場合はボタンを押してください。`);
    } else {
      show(okMsg, "書き戻しました。ダウンロードが始まらない場合はボタンを押してください。");
    }
  } catch (err) {
    logError("PPTX への書き戻し", err);
    if (mismatchNote) show(warnMsg, mismatchNote);
    show(errorMsg, `書き戻しに失敗しました: ${err.message || err}`);
  } finally {
    setBusy(false);
  }
}

function resetAll() {
  sourceFile = null;
  extracted = null;
  sourceZip = null;
  resultBlob = null;
  resultName = "";
  fileInput.value = "";
  txtInput.value = "";
  pptxMeta.hidden = true;
  txtMeta.hidden = true;
  extractBox.value = "";
  promptBox.value = loadStoredPrompt() || defaultPrompt();
  document.getElementById("pasteBox").value = "";
  deckFonts = [];
  for (const [sel, input, preview] of [
    [titleFontSelect, titleCustomFont, titleFontPreview],
    [bodyFontSelect, bodyCustomFont, bodyFontPreview],
  ]) {
    sel.value = KEEP_VALUE;
    input.value = "";
    input.hidden = true;
    updateFontPreview(sel, input, preview);
  }
  step2.hidden = true;
  step3.hidden = true;
  step4.hidden = true;
  setStepState(step1, "");
  setStepState(step2, "is-wait");
  setStepState(step3, "is-wait");
  setStepState(step4, "is-wait");
  clearMessages();
}

bindDrop(pptxDrop, handlePptx, (file) => file.name.toLowerCase().endsWith(".pptx"));
bindDrop(txtDrop, handleTxt, (file) => file.name.toLowerCase().endsWith(".txt"));
pptxDrop.addEventListener("pick", () => fileInput.click());
txtDrop.addEventListener("pick", () => txtInput.click());
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) handlePptx(file);
});
txtInput.addEventListener("change", () => {
  const file = txtInput.files?.[0];
  if (file) handleTxt(file);
});
document.getElementById("applyPasteBtn").addEventListener("click", () => {
  applyTranslation(document.getElementById("pasteBox").value, "貼り付け");
});

document.getElementById("copyPromptBtn").addEventListener("click", () => {
  copyText(promptBox.value);
});
document.getElementById("copyAllBtn").addEventListener("click", () => {
  copyText(`${promptBox.value.trim()}\n\n${extractBox.value}`);
});
document.getElementById("downloadTxtBtn").addEventListener("click", () => {
  if (!extracted || !sourceFile) return;
  const blob = new Blob([formatExtractFile(extracted.lines)], { type: "text/plain;charset=utf-8" });
  downloadBlob(blob, extractOutputName(sourceFile.name));
});
document.getElementById("downloadPptxBtn").addEventListener("click", () => {
  if (!resultBlob) return;
  downloadBlob(resultBlob, resultName);
});
document.getElementById("resetBtn").addEventListener("click", resetAll);

bindFontPicker(titleFontSelect, titleCustomFont, titleFontPreview);
bindFontPicker(bodyFontSelect, bodyCustomFont, bodyFontPreview);

promptBox.addEventListener("input", () => storePrompt(promptBox.value));
document.getElementById("resetPromptBtn").addEventListener("click", () => {
  promptBox.value = defaultPrompt();
  storePrompt(promptBox.value);
  show(okMsg, "プロンプトを既定に戻しました。");
});

(async function showVersion() {
  try {
    const res = await fetch("version.json", { cache: "no-store" });
    if (!res.ok) return;
    const data = await res.json();
    if (!data.version) return;
    document.getElementById("appVersion").textContent = `v${data.version}`;
    document.getElementById("appVersionLine").hidden = false;
  } catch (_) {}
})();
