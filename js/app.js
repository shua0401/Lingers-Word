/**
 * Lingers Word — 間隔反復（SM-2系）+ 習慣カレンダー
 * 素材: Googleスプレッドシート（「リンクを知っている全員が閲覧可」推奨）
 */

const LS = "lingers_word_v1";
/** メインと同内容のミラー（書き込み事故時の復旧用）。カレンダー・XP・SRS すべて含む */
const LS_BACKUP = "lingers_word_v1_backup";
const LS_URL = "lingers_word_csv_url";
const LS_SYNC_URL = "lingers_word_supabase_url";
const LS_SYNC_KEY = "lingers_word_supabase_anon";
const LS_SYNC_ON = "lingers_word_sync_on";
const LS_SYNC_LAST_OK = "lingers_word_sync_last_ok";
const CLOUD_TABLE = "lingers_word_cloud";
const CLOUD_ROW_ID = "me";

/** スマホの貼り付けで混ざりやすいゼロ幅・NBSPなどを除く（401 の主因になりやすい） */
function sanitizeSupabaseKey(raw) {
  if (raw == null || typeof raw !== "string") return "";
  return raw
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/^["'`]+|["'`]+$/g, "")
    .split(/\r?\n/)
    .join("")
    .trim();
}

function sanitizeSupabaseUrl(raw) {
  if (raw == null || typeof raw !== "string") return "";
  return raw
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/\/+$/, "")
    .replace(/\s+/g, "");
}

/** 「英会話」既定シート。入力欄が空のときもここから毎回取得します */
const DEFAULT_SHEET_ID = "1r9_TB-w8X1A2I0WrVjvlY3eX4epyr_ceS1ke_AwenUs";

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

/**
 * @param {string} raw
 * @returns {{ id: string, gid: number } | null}
 */
function extractSheetParams(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const idM = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (idM) {
    let gid = 0;
    const gidM = s.match(/[#&?]gid=(\d+)/);
    if (gidM) gid = parseInt(gidM[1], 10);
    return { id: idM[1], gid };
  }
  if (/^[a-zA-Z0-9-_]{30,}$/.test(s)) return { id: s, gid: 0 };
  return null;
}

/**
 * @param {string} id
 * @param {number} gid
 * @returns {Promise<string | null>}
 */
async function fetchGoogleSheetCsv(id, gid) {
  const urls = [
    `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&gid=${gid}`,
    `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`,
  ];
  for (const base of urls) {
    try {
      const sep = base.includes("?") ? "&" : "?";
      const url = `${base}${sep}_cb=${Date.now()}`;
      const res = await fetch(url, { cache: "no-store", mode: "cors" });
      if (!res.ok) continue;
      const text = await res.text();
      if (!text || /^\s*</.test(text)) continue;
      if (text.replace(/\s/g, "").length > 10) return text;
    } catch {
      /* next */
    }
  }
  return null;
}

/** @typedef {{ wordEn: string, wordJa: string, wordNote: string, sentenceEn: string, sentenceJa: string, sentenceNote: string }} Entry */

function todayYMD() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addCalendarDays(ymd, days) {
  const [y, m, d] = ymd.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

/** SM-2風スケジュール（quality 0–5） */
function scheduleReview(prev, quality) {
  let reps = prev.reps ?? 0;
  let ef = prev.ef ?? 2.5;
  let interval = prev.interval ?? 0;

  if (quality < 3) {
    reps = 0;
    interval = 1;
  } else {
    if (reps === 0) interval = 1;
    else if (reps === 1) interval = 6;
    else interval = Math.max(1, Math.round(interval * ef));
    reps += 1;
  }
  ef = ef + (0.1 - (5 - quality) * (0.08 + (5 - quality) * 0.02));
  if (ef < 1.3) ef = 1.3;

  const next = addCalendarDays(todayYMD(), interval);
  return { reps, ef, interval, next };
}

function loadPersisted() {
  const empty = { srs: {}, habit: {}, meta: { xp: 0 } };
  /** @param {string | null} raw */
  function parse(raw) {
    if (!raw) return null;
    const j = JSON.parse(raw);
    return {
      srs: j.srs || {},
      habit: j.habit || {},
      meta: j.meta || { xp: 0 },
    };
  }
  for (const key of [LS, LS_BACKUP]) {
    try {
      const p = parse(localStorage.getItem(key));
      if (p) return p;
    } catch {
      /* try backup key */
    }
  }
  return empty;
}

function cloneData(/** @type {unknown} */ x) {
  try {
    return JSON.parse(JSON.stringify(x));
  } catch {
    return x;
  }
}

function mergeHabits(/** @type {Record<string, { count?: number, marked?: boolean }>} */ h1, h2) {
  const out = /** @type {Record<string, { count: number, marked: boolean }>} */ ({});
  const keys = new Set([...Object.keys(h1 || {}), ...Object.keys(h2 || {})]);
  for (const k of keys) {
    const a = h1[k] || { count: 0, marked: false };
    const b = h2[k] || { count: 0, marked: false };
    out[k] = {
      count: Math.max(Number(a.count) || 0, Number(b.count) || 0),
      marked: !!(a.marked || b.marked),
    };
  }
  return out;
}

/** PC/スマホなど別端末のデータをまとめる（丸は両方の max、SRSは新しい方のスナップショット、XPは max） */
function mergePersistCrossDevice(
  /** @type {{ srs?: object, habit?: object, meta?: { xp?: number, syncAt?: number } }} */ local,
  remote
) {
  const tLocal = local.meta?.syncAt || 0;
  const tRemote = remote.meta?.syncAt || 0;
  const srs = tLocal >= tRemote ? cloneData(local.srs) : cloneData(remote.srs);
  const habit = mergeHabits(local.habit || {}, remote.habit || {});
  const xp = Math.max(local.meta?.xp || 0, remote.meta?.xp || 0);
  return {
    srs: srs || {},
    habit,
    meta: { xp, syncAt: Date.now() },
  };
}

function supaHeaders() {
  const key = sanitizeSupabaseKey(localStorage.getItem(LS_SYNC_KEY) || "");
  return {
    apikey: key,
    Authorization: `Bearer ${key}`,
    "Content-Type": "application/json",
  };
}

/** クラウドの学習データ。undefined=通信/HTTP失敗、null=まだ行がない、object=中身 */
async function fetchCloudPersist() {
  const base = sanitizeSupabaseUrl(localStorage.getItem(LS_SYNC_URL) || "");
  const key = sanitizeSupabaseKey(localStorage.getItem(LS_SYNC_KEY) || "");
  if (!base || !key) return undefined;
  const url = `${base}/rest/v1/${CLOUD_TABLE}?id=eq.${encodeURIComponent(CLOUD_ROW_ID)}&select=body`;
  const res = await fetch(url, { headers: { apikey: key, Authorization: `Bearer ${key}` } });
  if (!res.ok) return undefined;
  const rows = await res.json();
  if (!Array.isArray(rows) || rows.length === 0) return null;
  const body = rows[0].body;
  if (typeof body === "string") {
    try {
      return JSON.parse(body);
    } catch {
      return null;
    }
  }
  return body && typeof body === "object" ? body : null;
}

function recordSyncSuccess() {
  localStorage.setItem(LS_SYNC_LAST_OK, new Date().toISOString());
  updateSyncStatusLine();
}

function updateSyncStatusLine() {
  const el = document.getElementById("syncStatusLine");
  if (!el) return;
  const raw = localStorage.getItem(LS_SYNC_LAST_OK);
  if (!raw) {
    el.textContent =
      "最終クラウド通信: まだ記録なし →「接続・同期を確認」で Supabase に届いているか試せます。";
    return;
  }
  const d = new Date(raw);
  el.textContent = `最終クラウド通信: ${d.toLocaleString("ja-JP")} （この時刻以降に保存・同期が成功しています）`;
}

async function upsertCloudPersist(/** @type {{ srs: object, habit: object, meta: object }} */ payload) {
  const base = sanitizeSupabaseUrl(localStorage.getItem(LS_SYNC_URL) || "");
  const key = sanitizeSupabaseKey(localStorage.getItem(LS_SYNC_KEY) || "");
  if (!base || !key) return false;
  const row = { id: CLOUD_ROW_ID, body: payload };
  const res = await fetch(`${base}/rest/v1/${CLOUD_TABLE}`, {
    method: "POST",
    headers: { ...supaHeaders(), Prefer: "resolution=merge-duplicates" },
    body: JSON.stringify([row]),
  });
  if (res.ok) recordSyncSuccess();
  return res.ok;
}

let cloudPushTimer = 0;
function queueCloudPush(/** @type {{ srs: object, habit: object, meta: object }} */ data) {
  if (localStorage.getItem(LS_SYNC_ON) !== "1") return;
  window.clearTimeout(cloudPushTimer);
  cloudPushTimer = window.setTimeout(async () => {
    try {
      await upsertCloudPersist(cloneData(data));
    } catch {
      /* オフライン時などは静かに失敗。「接続・同期を確認」で調べられます。 */
    }
  }, 800);
}

/** 今の設定で Supabase に届くか調べる（読み取り） */
async function testCloudSync() {
  const on = localStorage.getItem(LS_SYNC_ON) === "1";
  const base = sanitizeSupabaseUrl(
    $("#supabaseUrlInput")?.value || localStorage.getItem(LS_SYNC_URL) || ""
  );
  const key = sanitizeSupabaseKey(
    $("#supabaseKeyInput")?.value || localStorage.getItem(LS_SYNC_KEY) || ""
  );
  if (!on) {
    toast("「クラウドと同期する」にチェックを入れ、「設定を保存してクラウドと取り込み」を押してください。");
    return;
  }
  if (!base || !key) {
    toast("Project URL と API キー（Publishable / anon）を入れてから保存してください。");
    return;
  }
  const url = `${base}/rest/v1/${CLOUD_TABLE}?id=eq.${encodeURIComponent(CLOUD_ROW_ID)}&select=body`;
  try {
    const res = await fetch(url, { headers: { apikey: key, Authorization: `Bearer ${key}` } });
    if (!res.ok) {
      let hint = "";
      try {
        const errJson = await res.json();
        if (errJson?.message) hint = ` — ${String(errJson.message).slice(0, 120)}`;
      } catch {
        /* ignore */
      }
      toast(`接続失敗（HTTP ${res.status}）。setup.sql でテーブル作成・URL・Publishable キーを確認${hint}`);
      return;
    }
    const rows = await res.json();
    recordSyncSuccess();
    if (!Array.isArray(rows)) {
      toast("接続はできましたが応答の形が想定外です。");
      return;
    }
    if (rows.length === 0) {
      toast(
        "接続OK。クラウドにまだ行がありません。練習して保存するか「設定を保存～」でこの端末のデータを上げます。"
      );
      return;
    }
    const body = rows[0].body;
    const habit = body && typeof body.habit === "object" ? body.habit : {};
    const nDays = Object.keys(habit).length;
    const xp = body?.meta?.xp;
    toast(
      `接続OK。クラウドにデータあり（習慣記録の日: ${nDays} 日分${xp != null ? `、XP ${xp}` : ""}）。他端末でも同じキーで見えます。`
    );
  } catch {
    toast("接続失敗（ネットワーク等）。URL が https://…supabase.co か確認してください。");
  }
}

/**
 * @param {{ srs: object, habit: object, meta: object }} data
 * @param {{ skipCloud?: boolean, skipSyncBump?: boolean }} [opts]
 */
function savePersisted(data, opts = {}) {
  data.meta = data.meta || {};
  if (!opts.skipSyncBump) data.meta.syncAt = Date.now();
  const s = JSON.stringify(data);
  try {
    localStorage.setItem(LS, s);
    localStorage.setItem(LS_BACKUP, s);
  } catch {
    try {
      localStorage.setItem(LS_BACKUP, s);
    } catch {
      /* 容量不足など */
    }
  }
  if (!opts.skipCloud && localStorage.getItem(LS_SYNC_ON) === "1") {
    queueCloudPush(JSON.parse(s));
  }
}

async function syncPullMerge() {
  if (localStorage.getItem(LS_SYNC_ON) !== "1") return;
  let remote;
  try {
    remote = await fetchCloudPersist();
  } catch {
    return;
  }
  if (remote === undefined) return;
  if (remote === null) {
    await upsertCloudPersist(persist);
    return;
  }
  if (typeof remote.habit !== "object") return;
  const merged = mergePersistCrossDevice(persist, remote);
  persist = merged;
  savePersisted(persist, { skipCloud: true });
  const ok = await upsertCloudPersist(persist);
  if (!ok) toast("クラウドへのマージ反映に失敗（Project URL・キー・SQL を確認）");
}

function parseCSV(text) {
  const rows = [];
  let i = 0;
  let field = "";
  let row = [];
  let inQuotes = false;
  while (i < text.length) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += c;
      i++;
      continue;
    }
    if (c === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (c === ",") {
      row.push(field);
      field = "";
      i++;
      continue;
    }
    if (c === "\r") {
      i++;
      continue;
    }
    if (c === "\n") {
      row.push(field);
      if (row.some((cell) => String(cell).trim() !== "")) rows.push(row);
      row = [];
      field = "";
      i++;
      continue;
    }
    field += c;
    i++;
  }
  if (field.length || row.length) {
    row.push(field);
    if (row.some((cell) => String(cell).trim() !== "")) rows.push(row);
  }
  return rows;
}

/** @returns {Entry[]} */
function rowsToEntries(rows) {
  if (!rows.length) return [];
  const norm = (x) => String(x ?? "").trim().toLowerCase().replace(/^["']|["']$/g, "");
  const h = rows[0].map(norm);
  const idx = (names) => {
    for (const n of names) {
      const j = h.indexOf(n.toLowerCase());
      if (j !== -1) return j;
    }
    return -1;
  };
  let cWordEn = idx(["word_en", "英語", "english"]);
  let cWordJa = idx(["word_ja", "日本語", "japanese", "意味"]);
  let cWordNote = idx(["word_note", "単語説明", "word_explanation"]);
  let cSentEn = idx(["sentence_en", "英文"]);
  let cSentJa = idx(["sentence_ja", "日本文"]);
  let cSentNote = idx(["sentence_note", "文説明", "sentence_explanation"]);

  const sampleRow = rows.length > 1 ? rows[1] : [];
  const colCount = Math.max(sampleRow.length, h.length);

  if (cWordEn < 0 && cWordJa < 0) {
    if (colCount >= 7) {
      cWordEn = 1;
      cWordJa = 2;
      cWordNote = 3;
      cSentEn = 4;
      cSentJa = 5;
      cSentNote = 6;
    } else {
      cWordEn = 0;
      cWordJa = 1;
      cWordNote = 2;
      cSentEn = 3;
      cSentJa = 4;
      cSentNote = 5;
    }
  }

  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const g = (ci) => (ci >= 0 && row[ci] != null ? String(row[ci]).trim() : "");
    const wordEn = g(cWordEn);
    const wordJa = g(cWordJa);
    const sentenceEn = g(cSentEn);
    const sentenceJa = g(cSentJa);
    if (!sentenceEn) continue;
    out.push({
      wordEn,
      wordJa,
      wordNote: cWordNote >= 0 ? g(cWordNote) : "",
      sentenceEn,
      sentenceJa,
      sentenceNote: cSentNote >= 0 ? g(cSentNote) : "",
    });
  }
  return out;
}

function normalizeAnswer(s, isEnSide) {
  let t = String(s ?? "")
    .trim()
    .replace(/\s+/g, " ");
  if (isEnSide) t = t.toLowerCase();
  return t;
}

/** 日→英の一致用。文の末尾にある . / ． は無くても正解 */
function normalizeEnglishAnswerForMatch(s) {
  return normalizeAnswer(s, true).replace(/[.．]+$/u, "");
}

/** 英文の「だいたい一致」判定用（句読点・空白・記号を除いて比較） */
function stripEnCompare(s) {
  return String(s ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000]+/g, "")
    .replace(/[.,!?;:·…'"「」『』()（）\[\]|\\/@#%^&*+=<>{}~`]/g, "");
}

function enKeyWords(expectedRaw) {
  const raw =
    String(expectedRaw ?? "")
      .toLowerCase()
      .match(/[a-z]{4,}/g) || [];
  const stop = new Set([
    "that",
    "this",
    "with",
    "from",
    "have",
    "been",
    "will",
    "your",
    "what",
    "when",
    "where",
    "which",
    "their",
    "there",
    "would",
    "could",
    "should",
    "about",
    "after",
    "before",
    "because",
    "while",
    "those",
    "these",
    "other",
    "into",
    "than",
    "then",
    "them",
    "some",
    "very",
    "just",
    "also",
    "only",
    "even",
    "much",
    "such",
    "here",
    "more",
    "most",
  ]);
  return [...new Set(raw.filter((w) => !stop.has(w)))];
}

/** 日→英・英文モード用。単語モードの英訳は evaluateAnswer 側で厳密一致。 */
function enSentenceAnswerGrade(userRaw, expectedRaw) {
  const u = stripEnCompare(userRaw);
  const exp = stripEnCompare(expectedRaw);
  if (!u || !exp) return null;
  if (u === exp) return "exact";

  const uLoose = normalizeEnglishAnswerForMatch(userRaw);
  const expLoose = normalizeEnglishAnswerForMatch(expectedRaw);
  if (uLoose && expLoose && uLoose === expLoose) return "exact";

  if (u.includes(exp) || exp.includes(u)) return "close";

  const words = enKeyWords(expectedRaw);
  if (words.length >= 1) {
    let hit = 0;
    for (const w of words) {
      if (u.includes(w)) hit++;
    }
    const need = Math.max(1, Math.ceil(words.length * 0.5));
    if (hit >= need) return "close";
  }

  const mx = Math.max(u.length, exp.length);
  if (mx <= 1) return u === exp ? "exact" : null;
  const ratio = 1 - levenshtein(u, exp) / mx;
  if (ratio >= 0.64) return "close";
  return null;
}

/** 日本語正解判定用（句読点・空白を除く） */
function stripJPCompare(s) {
  return String(s ?? "")
    .normalize("NFKC")
    .trim()
    .replace(/[\s\u3000]+/g, "")
    .replace(/[。、，．・，､｡,!！?？'"「」『』()（）\[\]]/g, "");
}

/** 重要語: 漢字2文字以上 / かな・カナ3文字以上（汎用助動詞っぽいものは除外） */
function jpKeyChunks(expectedStripped) {
  const s = expectedStripped;
  const kanji = s.match(/[\u4e00-\u9fff]{2,}/g) || [];
  const kana = s.match(/[\u3040-\u309f\u30a0-\u30ff]{3,}/g) || [];
  const stop = new Set([
    "であり",
    "です",
    "ます",
    "ません",
    "でした",
    "しまし",
    "という",
    "ように",
    "される",
    "できる",
    "なりま",
  ]);
  return [...new Set([...kanji, ...kana].filter((c) => c.length >= 2 && !stop.has(c)))];
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  if (!m) return n;
  if (!n) return m;
  /** @type {number[]} */
  let dp = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const cur = dp[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + cost);
      prev = cur;
    }
  }
  return dp[n];
}

/**
 * @param {string} userRaw
 * @param {string} expectedRaw
 * @param {{ sentenceRelaxed?: boolean }} [opts] 文モードは重要語＋類似度8割で合格しやすく
 * @returns {'exact'|'close'|null}
 */
function jpAnswerGrade(userRaw, expectedRaw, opts) {
  const relaxed = !!opts?.sentenceRelaxed;
  const u = stripJPCompare(userRaw);
  const exp = stripJPCompare(expectedRaw);
  if (!u || !exp) return null;
  if (u === exp) return "exact";
  if (u.includes(exp) || exp.includes(u)) return "close";

  const chunks = jpKeyChunks(exp);
  let chunkHit = 0;
  if (chunks.length >= 1) {
    for (const c of chunks) {
      if (u.includes(c)) chunkHit++;
    }
    const need = Math.max(1, Math.ceil(chunks.length * 0.5));
    if (chunkHit >= need && relaxed) {
      const mx = Math.max(u.length, exp.length);
      const ratio = mx <= 1 ? (u === exp ? 1 : 0) : 1 - levenshtein(u, exp) / mx;
      if (ratio >= 0.8) return "exact";
      if (ratio >= 0.72) return "close";
    }
    if (chunkHit >= need) return "close";
  }

  const mx = Math.max(u.length, exp.length);
  if (mx <= 1) return u === exp ? "exact" : null;
  const ratio = 1 - levenshtein(u, exp) / mx;
  if (ratio >= (relaxed ? 0.72 : 0.64)) return "close";
  return null;
}

/** 英単語トークン（省みない・語形はそのまま一致が必要） */
function englishTokens(s) {
  return String(s)
    .toLowerCase()
    .match(/[a-z]+(?:'[a-z]+)?/g) || [];
}

/** 日→英: 回答に英単語マスタの語が「トークン一致」で含まれるか（語形違いは不可） */
function sentenceHasKeyEnglishWord(userRaw, wordEn) {
  const w = String(wordEn ?? "").trim();
  if (!w) return true;
  const need = englishTokens(w);
  if (!need.length) return true;
  const have = new Set(englishTokens(userRaw));
  return need.every((t) => have.has(t));
}

/** 英→日: 回答に単語マスタの日本語が含まれるか */
function sentenceHasKeyJapanese(userRaw, wordJa) {
  const w = String(wordJa ?? "").trim();
  if (!w) return true;
  const ws = stripJPCompare(w);
  const us = stripJPCompare(userRaw);
  if (!ws) return true;
  return us.includes(ws);
}

/**
 * 英文モード・英→日: 単語列が本文とずれていても、模範訳の「重要語」が拾えていれば通過。
 * （単語列が空／英語のまま等で旧ゲートだけだと、句読点以外まで一致していても落ちるのを防ぐ）
 */
function sentenceAllowsJapaneseKeywords(userRaw, wordJa, targets) {
  const us = stripJPCompare(userRaw);
  if (!us) return false;
  if (sentenceHasKeyJapanese(userRaw, wordJa)) return true;
  for (const t of targets) {
    const exp = stripJPCompare(t);
    if (!exp) continue;
    const chunks = jpKeyChunks(exp);
    if (!chunks.length) return true;
    let hit = 0;
    for (const c of chunks) {
      if (us.includes(c)) hit++;
    }
    const need = Math.max(1, Math.ceil(chunks.length * 0.5));
    if (hit >= need) return true;
  }
  return false;
}

/**
 * @param {unknown} entry rows の1行
 * @returns {{ grade: 'exact' | 'close' | 'wrong' }}
 */
function evaluateAnswer(user, expected, isEnExpected, entry) {
  const parts = String(expected)
    .split(/[／\/|｜]/g)
    .map((p) => String(p).trim())
    .filter(Boolean);
  const targets = parts.length ? parts : [String(expected).trim()];
  if (!targets.some(Boolean)) return { grade: "wrong" };

  if (isEnExpected) {
    if (studyMode === "sentence") {
      const uRaw = String(user ?? "").trim();
      if (!uRaw) return { grade: "wrong" };
      if (!sentenceHasKeyEnglishWord(uRaw, entry?.wordEn)) return { grade: "wrong" };
      let best = /** @type {'exact'|'close'|'wrong'} */ ("wrong");
      for (const t of targets) {
        const su = stripEnCompare(userRaw);
        const st = stripEnCompare(t);
        if (su && st && su === st) return { grade: "exact" };
        const g = enSentenceAnswerGrade(uRaw, t);
        if (g === "exact") return { grade: "exact" };
        if (g === "close") best = "close";
      }
      if (best === "close") return { grade: "close" };
      return { grade: "wrong" };
    }
    const u = normalizeEnglishAnswerForMatch(user);
    if (!u) return { grade: "wrong" };
    const hit = targets.some((t) => {
      const tn = normalizeEnglishAnswerForMatch(t);
      return !!tn && u === tn;
    });
    if (hit) return { grade: "exact" };
    const uRaw = String(user ?? "").trim();
    let hasClose = false;
    for (const t of targets) {
      const g = enSentenceAnswerGrade(uRaw, t);
      if (g === "close") hasClose = true;
    }
    if (hasClose) return { grade: "close" };
    return { grade: "wrong" };
  }

  const uRaw = String(user ?? "").trim();
  if (!uRaw) return { grade: "wrong" };

  for (const t of targets) {
    if (stripJPCompare(uRaw) === stripJPCompare(t)) return { grade: "exact" };
  }

  if (studyMode === "sentence" && !sentenceAllowsJapaneseKeywords(uRaw, entry?.wordJa, targets)) {
    return { grade: "wrong" };
  }

  const relaxed = studyMode === "sentence";
  let hasClose = false;
  for (const t of targets) {
    const g = jpAnswerGrade(uRaw, t, { sentenceRelaxed: relaxed });
    if (g === "exact") return { grade: "exact" };
    if (g === "close") hasClose = true;
  }
  if (hasClose) return { grade: "close" };
  return { grade: "wrong" };
}

function srsSlot(mode, dir) {
  return mode === "word" ? (dir === "en_ja" ? "w_ej" : "w_je") : dir === "en_ja" ? "s_ej" : "s_je";
}

function ensureCardState(persist, id, slot) {
  if (!persist.srs[id]) persist.srs[id] = {};
  if (!persist.srs[id][slot]) persist.srs[id][slot] = {};
  return persist.srs[id][slot];
}

function pickIds(entries, mode, dir, persist, take) {
  const t = todayYMD();
  const ids = entries
    .map((_, i) => i)
    .filter((i) => {
      const e = entries[i];
      if (mode === "sentence") return !!(e.sentenceEn && e.sentenceJa);
      return !!(e.wordEn && e.wordJa);
    });

  const scored = ids.map((id) => {
    const slot = srsSlot(mode, dir);
    const st = ensureCardState(persist, id, slot);
    const next = st.next;
    const due = !next || next <= t;
    const overdueDays = next && next < t ? Math.floor((+new Date(t) - +new Date(next)) / 86400000) : 0;
    return { id, due, overdueDays, reps: st.reps ?? 0 };
  });

  const preferSentenceRows =
    mode === "sentence" && lastWordRoundIds.length ? new Set(lastWordRoundIds) : null;

  scored.sort((a, b) => {
    if (preferSentenceRows) {
      const pa = preferSentenceRows.has(a.id) ? 0 : 1;
      const pb = preferSentenceRows.has(b.id) ? 0 : 1;
      if (pa !== pb) return pa - pb;
    }
    if (a.due !== b.due) return a.due ? -1 : 1;
    if (a.overdueDays !== b.overdueDays) return b.overdueDays - a.overdueDays;
    return a.reps - b.reps;
  });

  const chosen = [];
  for (const s of scored) {
    chosen.push(s.id);
    if (chosen.length >= take) break;
  }
  if (chosen.length < take && ids.length) {
    let k = 0;
    while (chosen.length < take) {
      chosen.push(ids[k % ids.length]);
      k++;
    }
  }
  return chosen;
}

function consecutiveStreak(habit, ymd) {
  let streak = 0;
  let d = ymd;
  for (;;) {
    const h = habit[d];
    if (h && h.count >= 5) {
      streak++;
      d = addCalendarDays(d, -1);
    } else break;
  }
  return streak;
}

// ---- UI state ----

/** @type {Entry[]} */
let entries = [];
let persist = loadPersisted();

let studyMode = "word"; // word | sentence
/** 直近「単語モード」で出した5問の行番号（英文モードで同じ語句を優先） */
let lastWordRoundIds = [];
let flow = "idle";
/** いま表示中の問題のヒント本文（空なら未登録） */
let currentHintPlain = "";
/** @type {'en_ja' | 'ja_en'} */
let direction = "en_ja";
let sessionIds = [];
let mainCursor = 0;
/** @type {number | null} */
let remedId = null;
/** @type {Set<number>} */
const wrongEnJa = new Set();
/** @type {Set<number>} */
const wrongJaEn = new Set();

/** 「次の問題」押下まで保留する採点結果 */
let pendingAdvance =
  /** @type {null | { uid: number, grade: "exact" | "close" | "wrong", flow: "main_en" | "remed_en" | "main_ja" | "remed_ja" }} */ (
    null
  );

/** カレンダー表示中の年・月（月は 0–11）。過去・未来を閲覧する */
let calendarViewYear = new Date().getFullYear();
let calendarViewMonth = new Date().getMonth();

// ---- DOM ----
const $ = (sel) => document.querySelector(sel);

const elCard = $("#card");
const elCardLabel = $("#cardLabel");
const elCardText = $("#cardText");
const elCardHint = $("#cardHint");
const elAnswer = $("#answerInput");
const elFeedback = $("#feedback");
const elFeedbackMain = $("#feedbackMain");
const elFeedbackModel = $("#feedbackModel");
const elBtnSubmit = $("#btnSubmit");
const elBtnNext = $("#btnNextQuestion");
const elPhaseBadge = $("#phaseBadge");
const elSessionMeta = $("#sessionMeta");
const elProgressDots = $("#progressDots");
const elXpVal = $("#xpVal");
const elStreakVal = $("#streakVal");
const elCalendar = $("#calendar");
const elCalMonthLabel = $("#calMonthLabel");
const elToastLayer = $("#toastLayer");
const elModal = $("#celebrateModal");
const elCelebrateBody = $("#celebrateBody");
const elCelebrateOk = $("#celebrateOk");

function toast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  elToastLayer.appendChild(t);
  setTimeout(() => t.remove(), 3200);
}

function showCelebrate(streak) {
  elCelebrateBody.textContent = `今日の5問クリア！ナイスログイン — 継続 ${streak}日目です。忘却曲線どおり、次回もサクッと復習しましょう。`;
  elModal.classList.remove("hidden");
}

elCelebrateOk.addEventListener("click", () => elModal.classList.add("hidden"));

function bumpHabitAndMaybeCelebrate() {
  const d = todayYMD();
  if (!persist.habit[d]) persist.habit[d] = { count: 0, marked: false };
  persist.habit[d].count += 1;
  const h = persist.habit[d];
  const streak = consecutiveStreak(persist.habit, d);

  if (h.count >= 5 && !h.marked) {
    h.marked = true;
    showCelebrate(streak);
    toast(`カレンダーに丸をつけました（継続 ${streak} 日）`);
  }
  savePersisted(persist);
  renderCalendar();
  elStreakVal.textContent = String(streak);
}

function xpDisplay(xp) {
  const x = xp || 0;
  const lv = Math.floor(x / 100) + 1;
  return `${x} · Lv${lv}`;
}

function addXp(n) {
  persist.meta.xp = (persist.meta.xp || 0) + n;
  savePersisted(persist);
  elXpVal.textContent = xpDisplay(persist.meta.xp);
}

function renderProgress() {
  elProgressDots.innerHTML = "";
  const total = 5;
  const filled = Math.min(mainCursor, total);
  for (let i = 0; i < total; i++) {
    const s = document.createElement("span");
    if (i < filled) s.classList.add("on");
    elProgressDots.appendChild(s);
  }
}

function getPromptExpectedHint(entry) {
  const isWord = studyMode === "word";
  if (direction === "en_ja") {
    const q = isWord ? entry.wordEn : entry.sentenceEn;
    const a = isWord ? entry.wordJa : entry.sentenceJa;
    const hint = isWord ? entry.wordNote : entry.sentenceNote;
    return {
      q,
      a,
      hint,
      expectEn: false,
    };
  }
  const q = isWord ? entry.wordJa : entry.sentenceJa;
  const a = isWord ? entry.wordEn : entry.sentenceEn;
  const hint = isWord ? entry.wordNote : entry.sentenceNote;
  return { q, a, hint, expectEn: true };
}

function flashCard() {
  elCard.classList.remove("is-flip");
  void elCard.offsetWidth;
  elCard.classList.add("is-flip");
}

function resetQuestionChrome() {
  elBtnSubmit.disabled = false;
  elBtnNext.classList.add("hidden");
}

function flushPendingAdvance() {
  if (!pendingAdvance) return;
  const { uid, grade, flow: fl } = pendingAdvance;
  pendingAdvance = null;
  elFeedback.classList.add("hidden");
  elFeedbackModel.textContent = "";
  elFeedbackModel.classList.add("hidden");
  elBtnSubmit.disabled = false;
  elBtnNext.classList.add("hidden");
  if (fl === "main_en") advanceAfterMainEn(uid, grade);
  else if (fl === "remed_en") advanceRemedEn(uid, grade);
  else if (fl === "main_ja") advanceMainJa(uid, grade);
  else if (fl === "remed_ja") advanceRemedJa(uid, grade);
}

/**
 * @param {{ focusInput?: boolean }} [opts] バックグラウンド復帰時の再読込などでは false（IMEが勝手にかなになるのを防ぐ）
 */
function showCurrent(opts = {}) {
  if (pendingAdvance) return;

  const shouldFocus = opts.focusInput !== false;

  resetQuestionChrome();
  elFeedback.classList.add("hidden");
  elFeedback.classList.remove("ok", "ok-soft", "bad");
  elFeedbackModel.textContent = "";
  elFeedbackModel.classList.add("hidden");
  elAnswer.value = "";
  elCardHint.hidden = true;
  $("#btnShowHint").disabled = true;

  if (!entries.length) {
    elCardLabel.textContent = "データなし";
    elCardText.textContent = "CSVを配置するか URL を設定してください。";
    elSessionMeta.textContent = "";
    elAnswer.blur();
    return;
  }

  if (flow === "idle") {
    elPhaseBadge.textContent = studyMode === "word" ? "単語" : "英語の文";
    elCardLabel.textContent = "今日のセッション";
    elCardText.textContent =
      (studyMode === "word" ? "単語" : "英文") + "モード — 下に答えを入力するか、そのまま「判定」で5問スタート";
    elSessionMeta.textContent = "忘却曲線ベースの復習キューから出題";
    elProgressDots.innerHTML = "";
    for (let i = 0; i < 5; i++) {
      const s = document.createElement("span");
      elProgressDots.appendChild(s);
    }
    elAnswer.setAttribute("lang", "ja");
    elAnswer.blur();
    return;
  }

  let entryId = null;

  if (flow === "main_en" && mainCursor < sessionIds.length) {
    entryId = sessionIds[mainCursor];
  } else if (flow === "remed_en" && remedId != null) {
    entryId = remedId;
  } else if (flow === "main_ja" && mainCursor < sessionIds.length) {
    entryId = sessionIds[mainCursor];
  } else if (flow === "remed_ja" && remedId != null) {
    entryId = remedId;
  }

  if (entryId == null) {
    elCardLabel.textContent = "完了";
    elCardText.textContent = "このセッション完了！また明日も5問いきましょう。";
    elPhaseBadge.textContent = direction === "en_ja" ? "英 → 日" : "日 → 英";
    elSessionMeta.textContent = "";
    flow = "idle";
    elAnswer.blur();
    return;
  }

  const entry = entries[entryId];
  const pe = getPromptExpectedHint(entry);

  elAnswer.setAttribute("lang", pe.expectEn ? "en" : "ja");

  elPhaseBadge.textContent = direction === "en_ja" ? "英 → 日" : "日 → 英";
  elCardLabel.textContent = studyMode === "word" ? "単語" : "英文";
  elCardText.textContent = pe.q;
  currentHintPlain = String(pe.hint ?? "").trim();
  elCardHint.textContent = "";
  elCardHint.hidden = true;
  $("#btnShowHint").disabled = false;

  const phaseLabel =
    flow === "main_en" || flow === "main_ja"
      ? `本番 ${Math.min(mainCursor + 1, 5)} / 5`
      : flow === "remed_en"
        ? "復習（英→日）覚えるまで"
        : flow === "remed_ja"
          ? "復習（日→英）覚えるまで"
          : "";
  elSessionMeta.textContent = phaseLabel;

  flashCard();
  if (shouldFocus) elAnswer.focus();
  else elAnswer.blur();
  renderProgress();
}

function applySrs(id, quality) {
  const slot = srsSlot(studyMode, direction);
  const st = ensureCardState(persist, id, slot);
  const nextSt = scheduleReview(st, quality);
  Object.assign(st, nextSt);
  savePersisted(persist);
}

function afterAnswerGrade(/** @type {'exact'|'close'|'wrong'} */ grade) {
  if (grade === "exact") {
    bumpHabitAndMaybeCelebrate();
    addXp(8);
  } else if (grade === "close") {
    addXp(4);
  } else {
    addXp(1);
  }
}

function advanceAfterMainEn(/** @type {number} */ id, /** @type {'exact'|'close'|'wrong'} */ grade) {
  if (grade === "wrong") {
    applySrs(id, 1);
    wrongEnJa.add(id);
  } else if (grade === "close") {
    applySrs(id, 3);
  } else {
    applySrs(id, 4);
  }
  mainCursor++;
  if (mainCursor < sessionIds.length) {
    showCurrent();
    return;
  }
  if (wrongEnJa.size) {
    flow = "remed_en";
    remedId = wrongEnJa.values().next().value;
    direction = "en_ja";
    mainCursor = 0;
    showCurrent();
    return;
  }
  startJaPhase();
}

function startJaPhase() {
  direction = "ja_en";
  flow = "main_ja";
  mainCursor = 0;
  wrongJaEn.clear();
  showCurrent();
}

function advanceRemedEn(id, grade) {
  if (grade === "wrong") {
    applySrs(id, 1);
  } else if (grade === "close") {
    applySrs(id, 3);
    wrongEnJa.delete(id);
    if (!wrongEnJa.size) {
      remedId = null;
      startJaPhase();
      return;
    }
    remedId = wrongEnJa.values().next().value;
  } else {
    applySrs(id, 4);
    wrongEnJa.delete(id);
    if (!wrongEnJa.size) {
      remedId = null;
      startJaPhase();
      return;
    }
    remedId = wrongEnJa.values().next().value;
  }
  showCurrent();
}

function advanceMainJa(id, grade) {
  if (grade === "wrong") {
    applySrs(id, 1);
    wrongJaEn.add(id);
  } else if (grade === "close") {
    applySrs(id, 3);
  } else {
    applySrs(id, 4);
  }
  mainCursor++;
  if (mainCursor < sessionIds.length) {
    showCurrent();
    return;
  }
  if (wrongJaEn.size) {
    flow = "remed_ja";
    remedId = wrongJaEn.values().next().value;
    showCurrent();
    return;
  }
  flow = "idle";
  direction = "en_ja";
  toast("このラウンド終了。お疲れさまでした！");
  showCurrent();
}

function advanceRemedJa(id, grade) {
  if (grade === "wrong") {
    applySrs(id, 1);
  } else if (grade === "close") {
    applySrs(id, 3);
    wrongJaEn.delete(id);
    if (!wrongJaEn.size) {
      remedId = null;
      flow = "idle";
      direction = "en_ja";
      toast("日→英もクリア。完璧です！");
      showCurrent();
      return;
    }
    remedId = wrongJaEn.values().next().value;
  } else {
    applySrs(id, 4);
    wrongJaEn.delete(id);
    if (!wrongJaEn.size) {
      remedId = null;
      flow = "idle";
      direction = "en_ja";
      toast("日→英もクリア。完璧です！");
      showCurrent();
      return;
    }
    remedId = wrongJaEn.values().next().value;
  }
  showCurrent();
}

function startSession() {
  if (!entries.length) {
    toast("単語データがありません");
    return;
  }
  direction = "en_ja";
  flow = "main_en";
  mainCursor = 0;
  remedId = null;
  wrongEnJa.clear();
  wrongJaEn.clear();
  sessionIds = pickIds(entries, studyMode, direction, persist, 5);
  if (!sessionIds.length) {
    toast(
      studyMode === "word"
        ? "単語モード用の行がありません（英語・日本語の両方がある行だけ出題されます）"
        : "英文モード用の行がありません（英文・日本文の両方がある行だけ出題されます）"
    );
    flow = "idle";
    showCurrent();
    return;
  }
  if (studyMode === "word") lastWordRoundIds = sessionIds.slice();
  showCurrent();
}

function onSubmitAnswer(e) {
  e.preventDefault();
  if (pendingAdvance) return;
  if (flow === "idle") {
    if (entries.length) startSession();
    return;
  }

  const uid =
    flow === "main_en" || flow === "main_ja"
      ? sessionIds[mainCursor]
      : flow === "remed_en" || flow === "remed_ja"
        ? remedId
        : null;
  if (uid == null) return;

  const fl =
    flow === "main_en" || flow === "remed_en" || flow === "main_ja" || flow === "remed_ja"
      ? flow
      : null;
  if (!fl) return;

  const entry = entries[uid];
  const pe = getPromptExpectedHint(entry);
  const { grade } = evaluateAnswer(elAnswer.value, pe.a, pe.expectEn, entry);

  elFeedback.classList.remove("hidden", "ok", "ok-soft", "bad");
  elFeedbackModel.textContent = "";
  elFeedbackModel.classList.add("hidden");
  if (grade === "wrong") {
    elFeedback.classList.add("bad");
    elFeedbackMain.textContent = `不正解… 正: ${pe.a}`;
  } else if (grade === "exact") {
    elFeedback.classList.add("ok");
    elFeedbackMain.textContent = "正解！";
  } else {
    elFeedback.classList.add("ok-soft");
    elFeedbackMain.textContent = "正解（スペル・表記のゆれあり。完全一致は緑のみ）";
    elFeedbackModel.textContent = `データの表記: ${pe.a}`;
    elFeedbackModel.classList.remove("hidden");
  }

  afterAnswerGrade(grade);

  pendingAdvance = { uid, grade, flow: fl };
  elBtnSubmit.disabled = true;
  elBtnNext.classList.remove("hidden");
  elBtnNext.focus();
}

$("#answerForm").addEventListener("submit", onSubmitAnswer);
$("#btnNextQuestion").addEventListener("click", () => flushPendingAdvance());

$("#btnShowHint").addEventListener("click", (ev) => {
  ev.preventDefault();
  if ($("#btnShowHint").disabled) return;
  if (!elCardHint.hidden) {
    elCardHint.hidden = true;
    return;
  }
  if (currentHintPlain) {
    elCardHint.textContent = `ヒント: ${currentHintPlain}`;
  } else {
    elCardHint.textContent =
      "（ヒント未登録）スプレッドシートの「単語説明」または「文説明」に書くと、ここに表示されます。";
  }
  elCardHint.hidden = false;
});

$("#btnModeWord").addEventListener("click", () => {
  if (pendingAdvance) flushPendingAdvance();
  studyMode = "word";
  $("#btnModeWord").classList.add("active");
  $("#btnModeSentence").classList.remove("active");
  $("#btnModeWord").setAttribute("aria-pressed", "true");
  $("#btnModeSentence").setAttribute("aria-pressed", "false");
  flow = "idle";
  showCurrent();
});

$("#btnModeSentence").addEventListener("click", () => {
  if (pendingAdvance) flushPendingAdvance();
  studyMode = "sentence";
  $("#btnModeSentence").classList.add("active");
  $("#btnModeWord").classList.remove("active");
  $("#btnModeSentence").setAttribute("aria-pressed", "true");
  $("#btnModeWord").setAttribute("aria-pressed", "false");
  flow = "idle";
  showCurrent();
});

function renderCalendar() {
  elCalendar.innerHTML = "";
  const y = calendarViewYear;
  const m = calendarViewMonth;
  elCalMonthLabel.textContent = `${y}年 ${m + 1}月`;

  const first = new Date(y, m, 1);
  const startPad = first.getDay();
  const lastDay = new Date(y, m + 1, 0).getDate();
  const tstr = todayYMD();

  for (let i = 0; i < 7; i++) {
    const w = document.createElement("div");
    w.className = "cal-weekday";
    w.textContent = WEEKDAYS[i];
    elCalendar.appendChild(w);
  }
  for (let i = 0; i < startPad; i++) {
    const e = document.createElement("div");
    e.className = "cal-day";
    elCalendar.appendChild(e);
  }
  for (let d = 1; d <= lastDay; d++) {
    const mm = String(m + 1).padStart(2, "0");
    const dd = String(d).padStart(2, "0");
    const key = `${y}-${mm}-${dd}`;
    const cell = document.createElement("div");
    cell.className = "cal-day";
    const num = document.createElement("span");
    num.className = "d-num";
    num.textContent = String(d);
    cell.appendChild(num);
    if (key === tstr) cell.classList.add("today");
    const h = persist.habit[key];
    if (h && h.count >= 5) cell.classList.add("done");
    elCalendar.appendChild(cell);
  }
}

$("#btnCalPrev").addEventListener("click", () => {
  calendarViewMonth -= 1;
  if (calendarViewMonth < 0) {
    calendarViewMonth = 11;
    calendarViewYear -= 1;
  }
  renderCalendar();
});

$("#btnCalNext").addEventListener("click", () => {
  calendarViewMonth += 1;
  if (calendarViewMonth > 11) {
    calendarViewMonth = 0;
    calendarViewYear += 1;
  }
  renderCalendar();
});

$("#btnCalToday").addEventListener("click", () => {
  const n = new Date();
  calendarViewYear = n.getFullYear();
  calendarViewMonth = n.getMonth();
  renderCalendar();
});

/**
 * @param {{ silent?: boolean }} [opts]
 */
async function loadData(opts = {}) {
  const silent = !!opts.silent;
  const savedUrl = localStorage.getItem(LS_URL) || "";
  $("#csvUrlInput").value = savedUrl;

  const tryFetchCsv = async (url) => {
    const sep = url.includes("?") ? "&" : "?";
    const busted = `${url}${sep}_cb=${Date.now()}`;
    const res = await fetch(busted, { cache: "no-store", mode: "cors" });
    if (!res.ok) throw new Error(String(res.status));
    return await res.text();
  };

  let text = null;
  let sourceLabel = "";

  const sheetFromSaved = extractSheetParams(savedUrl);
  const sheetFromDefault = { id: DEFAULT_SHEET_ID, gid: 0 };

  if (savedUrl.trim()) {
    const isGoogleEditLink =
      savedUrl.includes("docs.google.com/spreadsheets/d/") &&
      !/tqx=out:csv|format=csv|output=csv/i.test(savedUrl);

    if (isGoogleEditLink || /^[a-zA-Z0-9-_]{30,}$/.test(savedUrl.trim())) {
      const p = sheetFromSaved || sheetFromDefault;
      text = await fetchGoogleSheetCsv(p.id, p.gid);
      if (text) sourceLabel = `スプレッドシート (${p.id.slice(0, 8)}…)`;
    } else {
      try {
        text = await tryFetchCsv(savedUrl.trim());
        if (text) sourceLabel = "保存したCSV URL";
      } catch {
        toast("保存URLの取得に失敗。既定シートを試します");
      }
    }
  }

  if (!text) {
    const p = sheetFromDefault;
    text = await fetchGoogleSheetCsv(p.id, p.gid);
    if (text) sourceLabel = "既定「英会話」シート";
  }

  if (!text) {
    try {
      text = await tryFetchCsv("data/words.csv");
      if (text) sourceLabel = "data/words.csv";
    } catch {
      entries = [];
      persist = loadPersisted();
      await syncPullMerge();
      if (!pendingAdvance) showCurrent();
      renderCalendar();
      toast(
        "シートを取得できませんでした。共有は「リンクを知っている全員が閲覧可」にしてください。"
      );
      return;
    }
  }

  const rows = parseCSV(text);
  entries = rowsToEntries(rows);
  persist = loadPersisted();
  await syncPullMerge();
  elXpVal.textContent = xpDisplay(persist.meta.xp);
  elStreakVal.textContent = String(consecutiveStreak(persist.habit, todayYMD()));
  if (!pendingAdvance) showCurrent({ focusInput: !silent });
  renderCalendar();
  if (!silent) toast(`読み込み ${entries.length} 件（${sourceLabel}）`);
}

let lastBgReload = 0;
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") {
    savePersisted(persist);
    return;
  }
  const t = Date.now();
  if (t - lastBgReload < 45000) return;
  lastBgReload = t;
  loadData({ silent: true });
});

window.addEventListener("pagehide", () => {
  savePersisted(persist);
});

$("#btnSaveUrl").addEventListener("click", async () => {
  const u = $("#csvUrlInput").value.trim();
  localStorage.setItem(LS_URL, u);
  await loadData();
});

$("#btnReloadData").addEventListener("click", async () => {
  await loadData();
});

$("#btnDownloadTemplate").addEventListener("click", () => {
  const header =
    "No.,英語,日本語,単語説明,英文,日本文,文説明\n" +
    "1,apple,りんご,,I like apples.,私はリンゴが好きです。,\n" +
    "2,,,,Sentence only is OK.,文だけの行の例。,\n";
  const bom = "\ufeff";
  const blob = new Blob([bom + header], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "words_template.csv";
  a.click();
  URL.revokeObjectURL(a.href);
});

const BACKUP_FORMAT = "lingers_word_backup";

function exportBackupJson() {
  savePersisted(persist);
  const payload = {
    format: BACKUP_FORMAT,
    version: 1,
    exportedAt: new Date().toISOString(),
    srs: persist.srs,
    habit: persist.habit,
    meta: persist.meta,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `lingers_word_backup_${todayYMD()}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  toast("JSONを保存しました（公開先・別端末で「読み込み」）");
}

function exportHabitCsv() {
  const lines = ["date,question_count,circle_5plus"];
  const dates = Object.keys(persist.habit).sort();
  for (const date of dates) {
    const h = persist.habit[date];
    lines.push(`${date},${h.count ?? 0},${h.marked ? "yes" : "no"}`);
  }
  const bom = "\ufeff";
  const blob = new Blob([bom + lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `lingers_word_habit_${todayYMD()}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
  toast("習慣ログCSVを保存しました");
}

function applyImportedPersist(/** @type {{ srs?: object, habit?: object, meta?: object }} */ data) {
  if (pendingAdvance) flushPendingAdvance();
  persist = {
    srs: data.srs || {},
    habit: data.habit || {},
    meta: data.meta || { xp: 0 },
  };
  savePersisted(persist);
  elXpVal.textContent = xpDisplay(persist.meta.xp);
  elStreakVal.textContent = String(consecutiveStreak(persist.habit, todayYMD()));
  renderCalendar();
  showCurrent();
}

$("#btnExportJson").addEventListener("click", exportBackupJson);
$("#btnExportHabitCsv").addEventListener("click", exportHabitCsv);

const elImportJsonInput = $("#importJsonInput");
$("#btnImportJson").addEventListener("click", () => elImportJsonInput.click());
function fillSyncForm() {
  const rawU = localStorage.getItem(LS_SYNC_URL) || "";
  const rawK = localStorage.getItem(LS_SYNC_KEY) || "";
  const url = sanitizeSupabaseUrl(rawU);
  const key = sanitizeSupabaseKey(rawK);
  if (rawU) localStorage.setItem(LS_SYNC_URL, url);
  if (rawK) localStorage.setItem(LS_SYNC_KEY, key);
  $("#supabaseUrlInput").value = url;
  $("#supabaseKeyInput").value = key;
  $("#syncEnabledInput").checked = localStorage.getItem(LS_SYNC_ON) === "1";
  updateSyncStatusLine();
}

$("#btnTestSync").addEventListener("click", () => {
  testCloudSync();
});

$("#btnSaveSync").addEventListener("click", async () => {
  if (pendingAdvance) flushPendingAdvance();
  const url = sanitizeSupabaseUrl($("#supabaseUrlInput").value);
  const key = sanitizeSupabaseKey($("#supabaseKeyInput").value);
  localStorage.setItem(LS_SYNC_URL, url);
  localStorage.setItem(LS_SYNC_KEY, key);
  $("#supabaseUrlInput").value = url;
  $("#supabaseKeyInput").value = key;
  localStorage.setItem(LS_SYNC_ON, $("#syncEnabledInput").checked ? "1" : "0");
  persist = loadPersisted();
  await syncPullMerge();
  elXpVal.textContent = xpDisplay(persist.meta.xp);
  elStreakVal.textContent = String(consecutiveStreak(persist.habit, todayYMD()));
  renderCalendar();
  showCurrent({ focusInput: false });
  toast("同期設定を保存しました");
});

elImportJsonInput.addEventListener("change", async (ev) => {
  const input = ev.target;
  const f = input.files?.[0];
  input.value = "";
  if (!f) return;
  try {
    const text = await f.text();
    const data = JSON.parse(text);
    if (
      data.format !== BACKUP_FORMAT ||
      !data.srs ||
      !data.habit ||
      typeof data.srs !== "object" ||
      typeof data.habit !== "object"
    ) {
      toast("このアプリが書き出した JSON を選んでください");
      return;
    }
    if (!window.confirm("このブラウザの学習記録を、ファイルの内容で置き換えますか？")) return;
    applyImportedPersist(data);
    toast("学習データを読み込みました");
  } catch {
    toast("JSONの読み込みに失敗しました");
  }
});

elXpVal.textContent = xpDisplay(persist.meta.xp);
elStreakVal.textContent = String(consecutiveStreak(persist.habit, todayYMD()));
renderCalendar();
fillSyncForm();
loadData();
