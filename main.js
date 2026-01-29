/* =====================================================
   定数・バージョン
===================================================== */
const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";
const APP_VERSION = "1.0.0"; // ←ここを更新すればフッターに反映

const DEFAULT_MASK_RULES = [
    { value: "コンフィック", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true },
    { value: "@conphic.co.jp", enabled: true }
];

/* =====================================================
   初期化
===================================================== */
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
    displayVersion();
});

/* =====================================================
   フッターにバージョン表示
===================================================== */
function displayVersion() {
    const footer = document.getElementById("footerVersion");
    footer.textContent = `Version ${APP_VERSION}`;
}

/* =====================================================
   イベント登録
===================================================== */
function bindEvents() {
    document.getElementById("addRowBtn").addEventListener("click", addRuleRow);
    document.getElementById("deleteRowBtn").addEventListener("click", deleteSelectedRows);
    document.getElementById("runBtn").addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

/* =====================================================
   ルール行操作
===================================================== */
function addRuleRow(rule) {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");

    const enabled = rule ? rule.enabled : true;
    const value = rule ? rule.value : "";

    tr.innerHTML = `
        <td><input type="checkbox" class="mask-enable" ${enabled ? "checked" : ""}></td>
        <td><input type="text" class="mask-word" value="${value}"></td>
    `;

    tbody.appendChild(tr);
    tr.querySelector(".mask-word").focus();
    saveRules();
}

function deleteSelectedRows() {
    const tbody = document.querySelector("#maskTable tbody");
    const rows = Array.from(tbody.querySelectorAll("tr"));

    rows.forEach(tr => {
        const checkbox = tr.querySelector(".mask-enable");
        if (checkbox && checkbox.checked) {
            tbody.removeChild(tr);
        }
    });

    saveRules();
}

/* =====================================================
   マスキング実行
===================================================== */
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

/* =====================================================
   Word XML マスキング本体
===================================================== */
function maskWordXml(xml, words) {
    const regex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    const textNodes = [];
    let match;

    while ((match = regex.exec(xml)) !== null) {
        textNodes.push({ full: match[0], text: match[1], start: match.index, end: regex.lastIndex });
    }

    const joined = textNodes.map(n => n.text).join("");

    // マスク適用
    let masked = joined;
    words.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(escaped, "g");
        masked = masked.replace(re, m => MASK_CHAR.repeat(m.length));
    });

    // 元の分割単位に再分配
    let cursor = 0;
    const replacedTexts = textNodes.map(n => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    let offset = 0;
    textNodes.forEach((node, i) => {
        const replaced = node.full.replace(node.text, escapeXml(replacedTexts[i]));
        xml = xml.slice(0, node.start + offset) + replaced + xml.slice(node.end + offset);
        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =====================================================
   ルール操作・保存
===================================================== */
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? word.value.trim() : null;
        })
        .filter(Boolean);
}

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
    let rules;

    if (saved) {
        rules = JSON.parse(saved);
    } else {
        rules = DEFAULT_MASK_RULES;
    }

    rules.forEach(rule => addRuleRow(rule));
}

/* =====================================================
   補助関数
===================================================== */
function getSelectedFile() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) alert("Wordファイルを選択してください");
    return file || null;
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
