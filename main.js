const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =========================
   初期ルール（デフォルト）
   - 正規表現は使わない（完全一致の部分文字列マスク）
========================= */
const DEFAULT_MASK_RULES = [
    { value: "株式会社コンフィック", enabled: true },
    { value: "コンフィック", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true },
    { value: "daichi@conphic.co.jp", enabled: false },
];

/* =========================
   初期化
========================= */
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
});

/* =========================
   イベント登録
========================= */
function bindEvents() {
    document.getElementById("addRowBtn")?.addEventListener("click", addRuleRow);
    document.getElementById("runBtn")?.addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

/* =========================
   行追加
========================= */
function addRuleRow() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const tr = document.createElement("tr");
    tr.innerHTML = `
    <td><input type="checkbox" class="mask-enable" checked></td>
    <td><input type="text" class="mask-word" value=""></td>
  `;
    tbody.appendChild(tr);
    tr.querySelector(".mask-word")?.focus();
    saveRules();
}

/* =========================
   マスキング実行
========================= */
async function runMasking() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput?.files?.[0];
    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

    const rules = getEnabledRules(); // [{ value }]
    if (rules.length === 0) {
        alert("マスキング対象がありません");
        return;
    }

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
        alert("document.xml が見つかりません");
        return;
    }

    let xml = await docXmlFile.async("string");
    xml = maskWordXmlByLiterals(xml, rules.map(r => r.value));

    zip.file("word/document.xml", xml);

    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(/\.docx$/i, "_マスキング.docx"));
}

/* =========================
   Word XML マスキング本体（固定文字列）
   - <w:t>を抽出して連結
   - 固定文字列でマスク（長さ維持）
   - 元の<w:t>長で再分配
   - <w:t>内側だけ関数replaceで差し替え（$問題回避）
========================= */
function maskWordXmlByLiterals(xml, literalWords) {
    const wtRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    const nodes = [];

    let m;
    while ((m = wtRegex.exec(xml)) !== null) {
        nodes.push({
            full: m[0],
            text: m[1],
            start: m.index,
            end: wtRegex.lastIndex,
            len: m[1].length,
        });
    }

    const joined = nodes.map(n => n.text).join("");
    let masked = joined;

    // 長い語から先に置換（部分一致の副作用を減らす）
    const words = literalWords
        .map(w => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    for (const word of words) {
        masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
    }

    // 再分配
    let cursor = 0;
    const parts = nodes.map(n => {
        const part = masked.slice(cursor, cursor + n.len);
        cursor += n.len;
        return part;
    });

    // XMLへ書き戻し
    let offset = 0;
    nodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        const replaced = replaceWtInner(node.full, parts[i]);
        xml = before + replaced + after;

        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =========================
   固定文字列を全部置換（同じ長さの■にする）
========================= */
function replaceAllLiteralKeepingLength(text, word, maskChar) {
    let result = text;
    let idx = 0;

    // ループで indexOf を使う（正規表現なし）
    while (true) {
        const found = result.indexOf(word, idx);
        if (found === -1) break;

        const mask = maskChar.repeat(word.length);
        result = result.slice(0, found) + mask + result.slice(found + word.length);

        // 無限ループ防止：置換後は次の位置へ
        idx = found + mask.length;
    }

    return result;
}

/* =========================
   <w:t> の内側だけ安全に差し替える（$問題回避）
========================= */
function replaceWtInner(fullWt, newInner) {
    return fullWt.replace(
        /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/,
        (_, p1, _old, p3) => p1 + newInner + p3
    );
}

/* =========================
   ルール取得
========================= */
function getEnabledRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return [];

    return Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? { value: word.value.trim() } : null;
        })
        .filter(Boolean)
        .filter(r => r.value.length > 0);
}

/* =========================
   localStorage
========================= */
function saveRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const rules = Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return { enabled: enable.checked, value: word.value };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    const savedRaw = localStorage.getItem(STORAGE_KEY);
    const saved = savedRaw ? safeJsonParse(savedRaw, []) : [];

    // DEFAULT + saved をマージ（value重複排除、saved優先）
    const map = new Map();
    DEFAULT_MASK_RULES.forEach(r => map.set(r.value, { ...r }));
    saved.forEach(r => {
        if (!r || !r.value) return;
        map.set(r.value, { value: r.value, enabled: !!r.enabled });
    });

    const merged = Array.from(map.values());

    merged.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
      <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
      <td><input type="text" class="mask-word" value="${escapeHtml(rule.value)}"></td>
    `;
        tbody.appendChild(tr);
    });
}

function safeJsonParse(text, fallback) {
    try { return JSON.parse(text); } catch { return fallback; }
}

/* =========================
   ダウンロード補助
========================= */
function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}

/* =========================
   HTML属性用エスケープ（value="" 崩れ防止）
========================= */
function escapeHtml(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}
