const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =========================
   初期ルール（デフォルト）※全て正規表現として扱う
========================= */
const DEFAULT_MASK_RULES = [
    { value: "(株式会社)?コンフィック", enabled: true },
    { value: "\\d{3}[--]\\d{4}", enabled: true },                 // 郵便番号
    { value: "[0-9０-９]{2,4}[0-9０-９]{2,4}[0-9０-９]{4}", enabled: true },
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
    if (!file) { alert("Wordファイルを選択してください"); return; }

    const rules = getEnabledRules();
    if (rules.length === 0) { alert("マスキング対象がありません"); return; }

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) { alert("document.xml が見つかりません"); return; }

    let xml = await docXmlFile.async("string");
    xml = maskWordXml(xml, rules);

    zip.file("word/document.xml", xml);

    // ここは「開けた」版と合わせて blob のまま
    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(".docx", "_masked.docx"));
}

/* =========================
   Word XML マスキング本体
   - <w:t> の連結 → 正規表現でマスク → 元の分割に再分配
   - ただし escapeXml はしない（破損要因になりやすいので）
========================= */
function maskWordXml(xml, rules) {
    const textNodes = [];
    const regex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let match;

    while ((match = regex.exec(xml)) !== null) {
        textNodes.push({
            full: match[0],
            text: match[1],
            start: match.index,
            end: regex.lastIndex
        });
    }

    // 連結テキスト
    const joined = textNodes.map(n => n.text).join("");
    let masked = joined;

    // マスク適用（全て正規表現として扱う）
    rules.forEach(rule => {
        try {
            const re = new RegExp(rule.value, "g");
            masked = masked.replace(re, m => MASK_CHAR.repeat(m.length));
        } catch (e) {
            console.warn("Invalid regex:", rule.value, e);
            // 無効な正規表現はスキップ
        }
    });

    // 再分配
    let cursor = 0;
    const replacedNodes = textNodes.map(n => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    // XML 書き戻し（offset 考慮）
    let offset = 0;
    textNodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        // full から text 部分だけ置換（タグ構造はそのまま）
        const replaced = node.full.replace(node.text, replacedNodes[i]);

        xml = before + replaced + after;
        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =========================
   ルール取得
   - enabled のものだけ
   - isRegex は廃止
========================= */
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? word.value.trim() : null;
        })
        .filter(Boolean)
        .map(value => ({ value }));
}

/* =========================
   localStorage
========================= */
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr"))
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
    tbody.innerHTML = "";

    const saved = localStorage.getItem(STORAGE_KEY);
    const rules = saved ? JSON.parse(saved) : DEFAULT_MASK_RULES;

    rules.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
            <td><input type="text" class="mask-word" value="${escapeHtml(rule.value)}"></td>
        `;
        tbody.appendChild(tr);
    });
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
   ※ XML に入れるエスケープではない
========================= */
function escapeHtml(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}
