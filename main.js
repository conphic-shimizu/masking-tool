const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

// デフォルトルール
const DEFAULT_MASK_RULES = [
    { value: "社名", enabled: true, isRegex: false },
    { value: "郵便番号", enabled: true, isRegex: false },
    { value: "住所", enabled: true, isRegex: false },
    { value: "電話番号", enabled: true, isRegex: false },
    { value: "FAX番号", enabled: true, isRegex: false },
    { value: "@gmail.com", enabled: true, isRegex: true }
];

// 初期化
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
});

// イベント登録
function bindEvents() {
    document.getElementById("addRowBtn").addEventListener("click", addRuleRow);
    document.getElementById("runBtn").addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

// 行追加
function addRuleRow(rule = { value: "", enabled: true, isRegex: false }) {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
        <td><input type="text" class="mask-word" value="${rule.value}"></td>
        <td><input type="checkbox" class="mask-regex" ${rule.isRegex ? "checked" : ""}></td>
    `;
    tbody.appendChild(tr);
    saveRules();
}

// マスキング実行
async function runMasking() {
    const file = getSelectedFile();
    if (!file) return;

    const rules = getEnabledRules();
    if (rules.length === 0) { alert("マスキング対象がありません"); return; }

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) { alert("document.xml が見つかりません"); return; }

    let xml = await docXmlFile.async("string");
    xml = maskWordXml(xml, rules);

    zip.file("word/document.xml", xml);

    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(".docx", "_masked.docx"));
}

// Word XMLマスキング
function maskWordXml(xml, rules) {
    const textNodes = [];
    const regex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let match;
    while ((match = regex.exec(xml)) !== null) {
        textNodes.push({ full: match[0], text: match[1] });
    }

    textNodes.forEach(node => {
        rules.forEach(rule => {
            if (!rule.value) return;
            try {
                if (rule.isRegex) {
                    const re = new RegExp(rule.value, "g");
                    node.text = node.text.replace(re, m => MASK_CHAR.repeat(m.length));
                } else {
                    const escaped = rule.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                    const re = new RegExp(escaped, "g");
                    node.text = node.text.replace(re, m => MASK_CHAR.repeat(m.length));
                }
            } catch (e) { console.warn("無効な正規表現:", rule.value); }
        });
    });

    textNodes.forEach(node => {
        xml = xml.replace(node.full, node.full.replace(/>([\s\S]*?)<\/w:t>/, `>${escapeXml(node.text)}</w:t>`));
    });

    return xml;
}

// 有効なルール取得
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            const regex = tr.querySelector(".mask-regex");
            if (!enable || !word || !regex) return null;
            return enable.checked ? { value: word.value.trim(), isRegex: regex.checked } : null;
        })
        .filter(Boolean);
}

// 補助関数
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

function getSelectedFile() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) alert("Wordファイルを選択してください");
    return file;
}

// localStorage
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr")).map(tr => ({
        value: tr.querySelector(".mask-word").value,
        enabled: tr.querySelector(".mask-enable").checked,
        isRegex: tr.querySelector(".mask-regex").checked
    }));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

// ロード
function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";
    const saved = localStorage.getItem(STORAGE_KEY);
    const rules = saved ? JSON.parse(saved) : DEFAULT_MASK_RULES;
    rules.forEach(addRuleRow);
}
