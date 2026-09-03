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

let sourceFile = null;
let extracted = null;
let sourceZip = null;
let resultBlob = null;
let resultName = "";

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

function clearMessages() {
  show(errorMsg, "");
  show(warnMsg, "");
  show(okMsg, "");
}

function setStepState(el, state) {
  el.classList.remove("is-wait", "is-done");
  if (state) el.classList.add(state);
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
  await navigator.clipboard.writeText(text);
  show(okMsg, "コピーしました。");
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
  clearMessages();
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    show(errorMsg, ".pptx を置いてください。");
    return;
  }
  sourceFile = file;
  resultBlob = null;
  try {
    sourceZip = await JSZip.loadAsync(await file.arrayBuffer());
    extracted = await extractTextsFromZip(sourceZip);
  } catch (err) {
    show(errorMsg, `抽出に失敗しました: ${err.message || err}`);
    return;
  }

  pptxMeta.hidden = false;
  pptxMeta.textContent = `${file.name} · ${(file.size / 1024).toFixed(1)} KB`;
  extractMeta.textContent = `${extracted.slideCount} 枚 · ${extracted.uidCount} 件のテキスト`;
  promptBox.value = TRANSLATION_PROMPT.trim() + "\n";
  extractBox.value = extracted.text;
  setStepState(step1, "is-done");
  step2.hidden = false;
  setStepState(step2, "");
  step3.hidden = false;
  setStepState(step3, "");
  step4.hidden = true;
  setStepState(step4, "is-wait");
  txtMeta.hidden = true;
  show(okMsg, "抽出しました。プロンプトと本文をコピーして翻訳し、訳文 txt をドロップしてください。");
}

async function handleTxt(file) {
  clearMessages();
  if (!sourceZip || !extracted) {
    show(errorMsg, "先に PPTX を置いてください。");
    return;
  }
  let source;
  try {
    source = await file.text();
  } catch (err) {
    show(errorMsg, `テキストの読み込みに失敗しました: ${err.message || err}`);
    return;
  }

  const translations = parseUidDelimitedText(source);
  const report = validateTranslations(extracted.metadata, translations);
  txtMeta.hidden = false;
  txtMeta.textContent = `${file.name} · 翻訳 ${report.got} 件 / 抽出 ${report.expected} 件`;

  if (!report.ok) {
    const bits = [];
    if (report.missing.length) bits.push(`不足 ${report.missing.slice(0, 8).join(", ")}${report.missing.length > 8 ? " …" : ""}`);
    if (report.extra.length) bits.push(`余分 ${report.extra.slice(0, 5).join(", ")}`);
    show(warnMsg, `uid が一致しません。見つかった分だけ書き戻します。${bits.join(" / ")}`);
  }

  try {
    const zip = await JSZip.loadAsync(await sourceFile.arrayBuffer());
    const result = await injectTextsToZip(zip, translations, extracted.metadata);
    resultBlob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    resultName = translatedOutputName(sourceFile.name);
    resultMeta.textContent = `${resultName} · 書き戻し ${result.injected} 件` +
      (result.missing ? ` · スキップ ${result.missing} 件` : "");
    step4.hidden = false;
    setStepState(step3, "is-done");
    setStepState(step4, "");
    downloadBlob(resultBlob, resultName);
    show(okMsg, "書き戻しました。ダウンロードが始まらない場合はボタンを押してください。");
  } catch (err) {
    show(errorMsg, `書き戻しに失敗しました: ${err.message || err}`);
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
  promptBox.value = "";
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
