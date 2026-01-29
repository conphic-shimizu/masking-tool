const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

// デフォルトマスキングルール
const DEFAULT_MASK_RULES = [
    { value: "コンフィック", enabled: true, isRegex: true },
    { value: "190-0022", enabled: true, isRegex: false },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true, isRegex: false },
    { value: "042-595-7557", enabled: true, isRegex: false },
    { value: "042-595-7558", enabled: true, isRegex: false },
    { value: "@conphic.co.jp", enabled: true, isRegex: true }
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

// 有効ルール取得
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

// マスキング実行
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
    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(".docx", "_masked.docx"));
}

// Word XMLマスキング本体
function maskWordXml(xml, rules) {
    const textNodes = [];
    const regexNode = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let match;
    while ((match = regexNode.exec(xml)) !== null) {
        textNodes.push({ full: match[0], text: match[1], start: match.index, end: regexNode.lastIndex });
    }

    const joined = textNodes.map(n => n.text).join("");
    let masked = joined;

    rules.forEach(rule => {
        if (rule.isRegex) {
            try {
                const re = new RegExp(rule.value, "g");
                masked = masked.replace(re, m => MASK_CHAR.repeat(m.length));
            } catch (e) {
                console.warn("無効な正規表現:", rule.value);
            }
        } else {
            const escaped = rule.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const re = new RegExp(escaped, "g");
            masked = masked.replace(re, m => MASK_CHAR.repeat(m.length));
        }
    });

    // 再分配
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

// XMLエスケープ
function escapeXml(str) {
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ダウンロード
function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}

// localStorage
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            const regex = tr.querySelector(".mask-regex");
            if (!enable || !word || !regex) return null;
            return { enabled: enable.checked, value: word.value, isRegex: regex.checked };
        })
        .filter(Boolean);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

// 初期ルール読み込み
function loadRules() {
    const saved = localStorage.getItem(STORAGE_KEY);
    const rules = saved ? JSON.parse(saved) : DEFAULT_MASK_RULES;
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";
    rules.forEach(rule => addRuleRow(rule));
}
