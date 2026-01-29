const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =========================
   初期ルール（デフォルト）
   ※ 全て「正規表現として扱う」前提
========================= */
const DEFAULT_MASK_RULES = [
    { value: "(株式会社)?コンフィック", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "\\d{3}[-‐-–−ー]\\d{4}", enabled: true }, // 郵便番号：ハイフン揺れ吸収
    {
        value:
            "[0-9０-９]{2,4}[-‐-–−ー]?[0-9０-９]{2,4}[-‐-–−ー]?[0-9０-９]{4}",
        enabled: true,
    }, // 電話番号：全角数字/ハイフン揺れ/ハイフン省略を吸収
    { value: "[a-zA-Z0-9._%+-]+@conphic\\.co\\.jp", enabled: true }, // メール
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
    document.getElementById("addRowBtn").addEventListener("click", addRuleRow);
    document.getElementById("runBtn").addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

/* =========================
   行追加
========================= */
function addRuleRow() {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");
    tr.innerHTML = `
    <td><input type="checkbox" class="mask-enable" checked></td>
    <td><input type="text" class="mask-word"></td>
  `;
    tbody.appendChild(tr);
    tr.querySelector(".mask-word").focus();
    saveRules();
}

/* =========================
   マスキング実行
========================= */
async function runMasking() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];
    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

    const rules = getEnabledRules(); // [{value}]
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
    xml = maskWordXml(xml, rules);

    zip.file("word/document.xml", xml);

    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(/\.docx$/i, "_masked.docx"));
}

/* =========================
   Word XML マスキング本体
   - <w:t> の連結 → 正規表現でマスク → 元の分割に再分配
   - XMLへ戻すときに「追加エスケープはしない」
     （Word XMLの中身は元々エスケープ済みで、追加するのは ■ なので安全）
   - $問題を避けるため replace は関数形式で差し替え
========================= */
function maskWordXml(xml, rules) {
    const textNodes = [];
    const wtRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;

    let match;
    while ((match = wtRegex.exec(xml)) !== null) {
        textNodes.push({
            full: match[0],      // <w:t ...>TEXT</w:t> 全体
            text: match[1],      // TEXT（XML上の文字列。既に &amp; 等を含み得る）
            start: match.index,
            end: wtRegex.lastIndex,
        });
    }

    // 連結（XML上の文字列として連結）
    const joined = textNodes.map((n) => n.text).join("");
    let masked = joined;

    // 置換（全て正規表現として扱う）
    for (const rule of rules) {
        const pattern = rule.value;
        if (!pattern) continue;

        try {
            const re = new RegExp(pattern, "g");
            masked = masked.replace(re, (m) => MASK_CHAR.repeat(m.length));
        } catch (e) {
            console.warn("Invalid regex skipped:", pattern, e);
        }
    }

    // 元の <w:t> の長さで再分配（XML上の長さ）
    let cursor = 0;
    const replacedTexts = textNodes.map((n) => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    // XMLへ書き戻し（offset考慮）
    let offset = 0;
    textNodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        // <w:t>の内側だけ差し替え（置換文字列ではなく関数で。$混入でも壊れない）
        const replaced = replaceWtInner(node.full, replacedTexts[i]);

        xml = before + replaced + after;
        offset += replaced.length - node.full.length;
    });

    return xml;
}

// <w:t> の内側だけ安全に差し替える（$問題回避のため関数形式）
function replaceWtInner(fullWt, newInner) {
    return fullWt.replace(
        /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/,
        (_, p1, _oldInner, p3) => p1 + newInner + p3
    );
}

/* =========================
   ルール取得（enabled のみ）
   - UIは「正規表現ON/OFFなし」
========================= */
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map((tr) => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? word.value.trim() : null;
        })
        .filter(Boolean)
        .map((value) => ({ value }));
}

/* =========================
   localStorage
========================= */
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map((tr) => {
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
    tbody.innerHTML = "";

    const savedRaw = localStorage.getItem(STORAGE_KEY);
    const saved = savedRaw ? safeJsonParse(savedRaw, []) : [];

    // DEFAULT + saved をマージ（valueで重複排除、savedでenabled/値を上書き）
    const map = new Map();
    DEFAULT_MASK_RULES.forEach((r) => map.set(r.value, { ...r }));
    saved.forEach((r) => {
        if (!r || !r.value) return;
        map.set(r.value, { value: r.value, enabled: !!r.enabled });
    });

    const merged = Array.from(map.values());

    merged.forEach((rule) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
      <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
      <td><input type="text" class="mask-word" value="${escapeHtml(rule.value)}"></td>
    `;
        tbody.appendChild(tr);
    });
}

function safeJsonParse(text, fallback) {
    try {
        return JSON.parse(text);
    } catch {
        return fallback;
    }
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
   HTML属性用の最低限エスケープ（value="" 崩れ防止）
========================= */
function escapeHtml(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}
