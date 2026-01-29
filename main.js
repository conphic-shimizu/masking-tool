const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =========================
   初期ルール（デフォルト）
   ※ 全て「正規表現」として扱う
========================= */
const DEFAULT_MASK_RULES = [
    { value: "(株式会社)?コンフィック", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "\\d{3}[-‐–−ー]\\d{4}", enabled: true }, // 郵便（ハイフン揺れ吸収）
    {
        value:
            "[0-9０-９]{2,4}[-‐–−ー]?[0-9０-９]{2,4}[-‐–−ー]?[0-9０-９]{4}",
        enabled: true,
    }, // 電話（全角・ハイフン揺れ・ハイフン省略吸収）
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
    <td><input type="text" class="mask-word"></td>
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
   Word XML マスキング本体（破損対策版）
   - <w:t> を抽出
   - 各 <w:t> の中身を「デコード」して実文字列に戻す
   - 連結して正規表現でマスク（実文字列に対して）
   - 元の分割長（実文字列の長さ）で再分配
   - <w:t> に戻すときは escapeXml で「再エンコード」
   ★これで &#8211; 等の数値参照を壊さない
========================= */
function maskWordXml(xml, rules) {
    // <w:t ...>...</w:t>
    const wtRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    const nodes = [];

    let m;
    while ((m = wtRegex.exec(xml)) !== null) {
        const full = m[0];
        const innerXmlText = m[1];              // XML内文字列（&amp; や &#8211; を含む）
        const decodedText = decodeXmlText(innerXmlText); // 実文字列に戻す

        nodes.push({
            full,
            innerXmlText,
            decodedText,
            decodedLen: decodedText.length,
            start: m.index,
            end: wtRegex.lastIndex,
        });
    }

    // 実文字列として連結 → マスク
    const joined = nodes.map((n) => n.decodedText).join("");
    let masked = joined;

    for (const rule of rules) {
        const pattern = rule.value;
        if (!pattern) continue;

        try {
            const re = new RegExp(pattern, "g");
            masked = masked.replace(re, (hit) => MASK_CHAR.repeat(hit.length));
        } catch (e) {
            console.warn("Invalid regex skipped:", pattern, e);
        }
    }

    // 元の <w:t> の「実文字列長」で再分配
    let cursor = 0;
    const parts = nodes.map((n) => {
        const part = masked.slice(cursor, cursor + n.decodedLen);
        cursor += n.decodedLen;
        return part;
    });

    // XMLへ書き戻し：<w:t>内側だけ差し替え（関数形式で $ 問題回避）
    let offset = 0;
    nodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        const newInner = escapeXml(parts[i]); // 実文字列 → XML用にエスケープ
        const replaced = replaceWtInner(node.full, newInner);

        xml = before + replaced + after;
        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =========================
   <w:t> 内側だけ安全に差し替え（$問題回避）
========================= */
function replaceWtInner(fullWt, newInnerXmlEscaped) {
    return fullWt.replace(
        /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/,
        (_, p1, _old, p3) => p1 + newInnerXmlEscaped + p3
    );
}

/* =========================
   ルール取得（enabled のみ）
========================= */
function getEnabledRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return [];

    return Array.from(tbody.querySelectorAll("tr"))
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
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const rules = Array.from(tbody.querySelectorAll("tr"))
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
    if (!tbody) return;

    tbody.innerHTML = "";

    const savedRaw = localStorage.getItem(STORAGE_KEY);
    const saved = savedRaw ? safeJsonParse(savedRaw, []) : [];

    // DEFAULT + saved をマージ（value重複排除、saved優先）
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
   XMLデコード/エンコード
   - decode: &amp; や &#8211; を実文字に戻す
   - escape: 実文字をXML内に戻せるようにする（& < > " '）
========================= */
function decodeXmlText(xmlEscapedText) {
    // DOMParserで textContent を使うのが一番安全（数値参照も復元される）
    const parser = new DOMParser();
    const doc = parser.parseFromString(`<r>${xmlEscapedText}</r>`, "application/xml");

    // パース失敗時はフォールバック（壊れたテキストをそのまま扱う）
    const parseError = doc.getElementsByTagName("parsererror")[0];
    if (parseError) return xmlEscapedText;

    return doc.documentElement.textContent ?? "";
}

function escapeXml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
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
