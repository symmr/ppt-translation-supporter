// Prompt options: the writing-style, amount, exchange-rate and glossary choices
// turn into a "表記ルール" section that is appended to the editable base prompt.
// Pure functions only, so they run in the browser and in Node tests.

const PROMPT_STYLES = [
  { id: "plain", label: "常体（だ・である）" },
  { id: "polite", label: "敬体（です・ます）" },
];

// Used for English to Japanese only. Going into English, amounts are always
// written with $ and K / M / B.
const PROMPT_AMOUNT_UNITS = [
  { id: "kmb", label: "K / M / B（原文のまま）" },
  { id: "man", label: "万・億（日本語の表記）" },
];

const DEFAULT_PROMPT_OPTIONS = { style: "plain", amountUnit: "kmb" };

const RULES_HEADER = "# 表記ルール（上のルールと食い違う場合は、こちらを優先する）";

// Exchange-rate sources that need no key and allow calls from a browser page.
const FX_SOURCES = [
  {
    id: "frankfurter",
    name: "欧州中央銀行の参考レート（Frankfurter）",
    site: "https://frankfurter.dev",
    url: "https://api.frankfurter.dev/v1/latest?base=USD&symbols=JPY",
  },
  {
    // the free tier asks for a visible credit with a link back
    id: "er-api",
    name: "Rates By Exchange Rate API",
    site: "https://www.exchangerate-api.com",
    url: "https://open.er-api.com/v6/latest/USD",
  },
];

function validFxRate(value) {
  const rate = Number(value);
  return Number.isFinite(rate) && rate > 0 && rate < 100000 ? rate : null;
}

// JPY per 1 USD from each source's JSON, or null when the reply is not usable.
function parseFxResponse(sourceId, json) {
  if (!json || typeof json !== "object") return null;
  if (sourceId === "frankfurter") {
    const rate = validFxRate(json.rates && json.rates.JPY);
    const date = typeof json.date === "string" ? json.date : "";
    return rate && date ? { rate, date } : null;
  }
  if (sourceId === "er-api") {
    const rate = validFxRate(json.rates && json.rates.JPY);
    const unix = Number(json.time_last_update_unix);
    if (!rate || json.result !== "success" || !Number.isFinite(unix)) return null;
    return { rate, date: new Date(unix * 1000).toISOString().slice(0, 10) };
  }
  return null;
}

function styleRule(style, titleIds) {
  if (style === "polite") {
    const lines = [
      "本文は敬体(です・ます調)で統一する。タイトル・見出し・短い名詞句は、「です・ます」を付けず、体言止めを中心にした常体にする。",
    ];
    if (titleIds && titleIds.length) {
      lines.push(`   次の uid はタイトルなので、常体(体言止め中心)で訳す: ${titleIds.join(", ")}`);
    }
    return lines.join("\n");
  }
  return "文体は常体(だ・である調)で統一する。見出しや短い名詞句は体言止めを適切に用いる。";
}

function amountRule(direction, amountUnit) {
  if (direction === "ja-en") {
    return "金額は英語圏の表記にする。「ドル」は「$」にし、「千・万・億」などの単位は K / M / B にする(例: 250万ドル → $2.5M、10万ドル → $100K、1億ドル → $100M、10億ドル → $1B)。円など他の通貨は、通貨を変えずに同じ要領で書く(例: 3億円 → ¥300M)。";
  }
  if (amountUnit === "man") {
    return `金額の単位は日本語の表記にする。K は千、M は百万、B は十億として、万・億にそろえる。通貨記号は「ドル」「円」などの通貨名にする(例: $250K → 25万ドル、$1.5M → 150万ドル、$2B → 20億ドル)。${approxRule("~$100K → 約10万ドル")}`;
  }
  return `金額は原文の表記のまま残す。「$」などの通貨記号・通貨コードと、K / M / B などの略記を、円・万・億などの日本語の表記に直さない(例: $1.5M、$250K はそのまま)。${approxRule("~$100K → 約$100K")}`;
}

// The tilde in "~$100K" means "approximately". Without this the model sometimes
// carries it into the Japanese as a tilde.
function approxRule(example) {
  return `概数を表す原文の「~」「≈」、approx.、about、around は、いつも「約」にする。「~」を訳文に残さない(例: ${example})。数値と数値の間にある「~」(例: 2~4)は範囲なので、「～」にする。`;
}

function fxRule(direction, fx) {
  const rate = fx && fx.enabled ? validFxRate(fx.rate) : null;
  if (!rate) return "";
  const basis = [fx.date, fx.source].filter(Boolean).join("、");
  const rateLine = `   - レート: 1 USD = ${rate} JPY${basis ? `(${basis})` : ""}。このレートだけを使う。`;
  if (direction === "ja-en") {
    return [
      "為替換算: 円(円、¥、JPY)の金額は、米ドルに換算して書く。上の金額の表記のルールより優先する。",
      rateLine,
      `   - 円の金額を ${rate} で割ってドルにし、「$」と K / M / B で書く。`,
      "   - 換算した金額の直前に、「約」(approximately)の意味で半角のチルダ「~」を付ける(例: ~$1.9M)。全角の「〜」にしない。範囲(から、未満)の意味では使わない。",
      "   - 有効数字は原文と同じ桁数を目安にし、最大でも 3 桁に丸める。桁を数え直し、桁違いにしない。",
      "   - 円以外の通貨と、金額でない数値は換算しない。",
      "   - 換算前の金額や、換算したことの注記は書かない(出力は訳文のみ)。",
    ].join("\n");
  }
  return [
    "為替換算: 米ドル($、USD、ドル)の金額は、円に換算して書く。上の金額の表記のルールより優先する。",
    rateLine,
    `   - ドルの金額に ${rate} を掛けて円にし、「万円」「億円」の単位でまとめる。`,
    "   - 換算した金額には「約」を付ける。",
    "   - 有効数字は原文と同じ桁数を目安にし、最大でも 3 桁に丸める。桁を数え直し、桁違いにしない。",
    "   - ドル以外の通貨(€、£ など)と、金額でない数値は換算しない。",
    "   - 換算前の金額や、換算したことの注記は書かない(出力は訳文のみ)。",
  ].join("\n");
}

// Glossary: a two-column CSV, English then Japanese. One file serves both
// directions; the source column follows the translation direction. Only the
// terms found in the extracted text go into the prompt, so a long glossary
// does not crowd it.
const GLOSSARY_TEMPLATE = [
  "英語,日本語",
  "Notable Event,注目すべきイベント",
  "risk-based alerting,リスクベースアラート",
  "Splunk,",
].join("\r\n") + "\r\n";

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  const src = String(text || "").replace(/^\uFEFF/, "");
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && src[i + 1] === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += ch;
    }
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

const GLOSSARY_HEADER_EN = /^(英語|english|en)$/i;
const GLOSSARY_HEADER_JA = /^(日本語|japanese|ja|jp)$/i;

// Returns [{ en, ja }]. A header row ("英語,日本語") is optional; blank rows and
// rows starting with # are skipped.
function parseGlossaryCsv(text) {
  const entries = [];
  const rows = parseCsvRows(text);
  rows.forEach((cells, index) => {
    const en = (cells[0] || "").trim();
    const ja = (cells[1] || "").trim();
    if (!en && !ja) return;
    if (en.startsWith("#")) return;
    if (index === 0 && GLOSSARY_HEADER_EN.test(en) && GLOSSARY_HEADER_JA.test(ja)) return;
    entries.push({ en, ja });
  });
  return entries;
}

// Folds the spellings that differ without changing the word: full and half
// width, case, spaces, the middle dot, and a long-vowel mark closing a
// katakana word (サーバー and サーバ).
function normalizeJa(text) {
  return String(text || "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s・]+/g, "")
    .replace(/([ァ-ヴ])ー+(?![ァ-ヴー])/g, "$1");
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// English terms match whole words, in any case, with an optional plural s/es.
function englishTermPattern(term) {
  const body = escapeRegExp(term).replace(/\s+/g, "\\s+");
  return new RegExp(`(?<![A-Za-z0-9])${body}(?:e?s)?(?![A-Za-z0-9])`, "i");
}

// Entries whose source-side term appears in the text. English to Japanese
// looks for the English column, Japanese to English for the Japanese one.
function matchGlossary(entries, text, direction) {
  const list = entries || [];
  if (direction === "ja-en") {
    const haystack = normalizeJa(text);
    return list.filter((e) => e.ja && haystack.includes(normalizeJa(e.ja)));
  }
  const haystack = String(text || "");
  return list.filter((e) => e.en && englishTermPattern(e.en).test(haystack));
}

function glossaryRule(matched, direction) {
  if (!matched || !matched.length) return "";
  const toJa = direction !== "ja-en";
  const targets = new Map();
  for (const e of matched) {
    const source = toJa ? e.en : e.ja;
    const target = toJa ? e.ja : e.en;
    if (!targets.has(source)) targets.set(source, []);
    if (target && !targets.get(source).includes(target)) targets.get(source).push(target);
  }
  const lines = ["用語集: 次の語は、指定のとおりに訳す。"];
  for (const [source, list] of targets) {
    if (!list.length) lines.push(`   - ${source}(訳さず、原文のまま残す)`);
    else if (list.length === 1) lines.push(`   - ${source} → ${list[0]}`);
    else lines.push(`   - ${source} → ${list.join(" / ")}(文脈に合うものを使う)`);
  }
  return lines.join("\n");
}

// options: { direction, style, amountUnit, titleIds, fx: { enabled, rate, date,
// source }, glossary: matched entries }. Returns the "表記ルール" section, or "" when empty.
function buildPromptRules(options) {
  const opts = options || {};
  const direction = opts.direction === "ja-en" ? "ja-en" : "en-ja";
  const rules = [];
  if (direction === "en-ja") rules.push(styleRule(opts.style, opts.titleIds));
  rules.push(amountRule(direction, opts.amountUnit));
  const fx = fxRule(direction, opts.fx);
  if (fx) rules.push(fx);
  const glossary = glossaryRule(opts.glossary, direction);
  if (glossary) rules.push(glossary);
  return `${RULES_HEADER}\n${rules.map((rule, i) => `${i + 1}. ${rule}`).join("\n")}\n`;
}

function buildFullPrompt(basePrompt, rules) {
  const base = String(basePrompt || "").trim();
  const extra = String(rules || "").trim();
  return extra ? `${base}\n\n${extra}\n` : `${base}\n`;
}

// The prompt box shows the base prompt and the rules as one text. Everything
// before the rules header is the user's base prompt; without the header the
// whole text is.
function basePromptOf(fullText) {
  const text = String(fullText || "");
  const at = text.indexOf(RULES_HEADER);
  return `${(at < 0 ? text : text.slice(0, at)).trim()}\n`;
}

const promptOptionsApi = {
  PROMPT_STYLES,
  PROMPT_AMOUNT_UNITS,
  DEFAULT_PROMPT_OPTIONS,
  FX_SOURCES,
  validFxRate,
  parseFxResponse,
  buildPromptRules,
  buildFullPrompt,
  basePromptOf,
  GLOSSARY_TEMPLATE,
  parseGlossaryCsv,
  normalizeJa,
  matchGlossary,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = promptOptionsApi;
}

if (typeof window !== "undefined") {
  Object.assign(window, promptOptionsApi);
}
