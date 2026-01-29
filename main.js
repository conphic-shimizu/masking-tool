/* =====================================================
   定数
===================================================== */
const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =====================================================
   デフォルトマスキングルール
===================================================== */
const DEFAULT_MASK_RULES = [
    { value: "コンフィック", enabled: true, isRegex: false },
    { value: "190-0022", enabled: true, isRegex: false },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true, isRegex: false },
    { value: "042-595-7557", enabled: true, isRegex: false },
    { value: "042-595-7558", enabled: true, isRegex: false },
    { value: "@conphic.co.jp", enabled: true, isRegex: false }
];

/* =====================================================
   初期化
===================================================== */
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
});

/* =====================================================
   イベント登録
===================================================== */
function bindEvents() {
    document.getElementById("addRowBtn").addEventListener("click", addRuleRow);
    document.getElementById("runBtn").addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

/* =====================================================
   ルール行操作
===================================================== */
function addRuleRow() {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");

    tr.innerHTML = `
    <td><input type="checkbox" class="mask-enable" checked></td>
    <td><input type="text" class="mask-word"></td>
    <td><input type="checkbox" class="mask-regex"></td>
  `;

    tbody.appendChild(tr);
    tr.querySelector(".mask-word").focus();
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
   Word XML マスキング本体（安全）
===================================================== */
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

    // 各 <w:t> 内で安全に置換
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
            } catch (e) {
                console.warn("無効な正規表現:", rule.value);
            }
        });
    });

    // XMLに書き戻す
    let offset = 0;
    textNodes.forEach((node, i) => {
        const replaced = node.full.replace(node.text, escapeXml(node.text));
        xml = xml.slice(0, node.start + offset) + replaced + xml.slice(node.end + offset);
        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =====================================================
   ルール取得
===================================================== */
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            const regex = tr.querySelector(".mask-regex");
            if (!enable || !word || !regex) return null;
            return enable.checked
                ? { value: word.value.trim(), isRegex: regex.checked }
                : null;
        })
        .filter(Boolean);
}

/* =====================================================
   localStorage
===================================================== */
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

function loadRules() {
    const saved = localStorage.getItem(STORAGE_KEY);
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    let rules;
    if (saved) {
        rules = JSON.parse(saved);
    } else {
        rules = DEFAULT_MASK_RULES.map(r => ({ ...r, isRegex: false }));
    }

    rules.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
      <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
      <td><input type="text" class="mask-word" value="${rule.value}"></td>
      <td><input type="checkbox" class="mask-regex" ${rule.isRegex ? "checked" : ""}></td>
    `;
        tbody.appendChild(tr);
    });
}

/* =====================================================
   補助関数
===================================================== */
function getSelectedFile() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) {
        alert("Wordファイルを選択してください");
        return null;
    }
    return file;
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
