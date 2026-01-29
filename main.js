const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

// デフォルトマスキングルール
const DEFAULT_MASK_RULES = [
    { value: "株式会社コンフィック", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "住所", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true }
];

// DOM読み込み時に初期化
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
});

// =========================
// イベント登録
// =========================
function bindEvents() {
    document.getElementById("addRowBtn").addEventListener("click", addRuleRow);
    document.getElementById("runBtn").addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

// =========================
// ルール行追加
// =========================
function addRuleRow(value = "", enabled = true) {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");

    tr.innerHTML = `
        <td><input type="checkbox" class="mask-enable" ${enabled ? "checked" : ""}></td>
        <td><input type="text" class="mask-word" value="${value}"></td>
    `;
    tbody.appendChild(tr);
    tr.querySelector(".mask-word").focus();
    saveRules();
}

// =========================
// マスキング実行
// =========================
async function runMasking() {
    const file = getSelectedFile();
    if (!file) return;

    const rules = getEnabledRules();
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
    download(blob, file.name.replace(".docx", "_masked.docx"));
}

// =========================
// Word XML マスキング本体
// =========================
function maskWordXml(xml, words) {
    // <w:t> をすべて抽出
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

    // 連結テキスト → マスク
    const joined = textNodes.map(n => n.text).join("");
    let masked = joined;

    words.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(escaped, "g");
        masked = masked.replace(re, MASK_CHAR.repeat(word.length));
    });

    // 元の分割単位に再分配
    let cursor = 0;
    const replacedTexts = textNodes.map(n => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    // XML書き戻し
    let offset = 0;
    textNodes.forEach((node, i) => {
        const replaced = node.full.replace(node.text, escapeXml(replacedTexts[i]));
        xml = xml.slice(0, node.start + offset) + replaced + xml.slice(node.end + offset);
        offset += replaced.length - node.full.length;
    });

    return xml;
}

// =========================
// 補助関数
// =========================
function getSelectedFile() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) alert("Wordファイルを選択してください");
    return file || null;
}

function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .filter(tr => tr.querySelector(".mask-enable").checked)
        .map(tr => tr.querySelector(".mask-word").value.trim())
        .filter(Boolean);
}

function escapeXml(str) {
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}

// =========================
// localStorage
// =========================
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr")).map(tr => ({
        enabled: tr.querySelector(".mask-enable").checked,
        value: tr.querySelector(".mask-word").value
    }));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

// =========================
// ルール読み込み
// =========================
function loadRules() {
    const saved = localStorage.getItem(STORAGE_KEY);
    const rules = saved ? JSON.parse(saved) : DEFAULT_MASK_RULES;

    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    rules.forEach(rule => addRuleRow(rule.value, rule.enabled));
}
