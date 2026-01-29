const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

const DEFAULT_MASK_RULES = [
    { value: "株式会社コンフィック", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "東京都立川市", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true },
    { value: "@conphic.co.jp", enabled: true }
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
function addRuleRow(rule = { value: "", enabled: true }) {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
        <td><input type="text" class="mask-word" value="${rule.value}"></td>
        <td><button class="delete-btn">×</button></td>
    `;
    tbody.appendChild(tr);

    // 削除ボタン
    tr.querySelector(".delete-btn").addEventListener("click", () => {
        tr.remove();
        saveRules();
    });

    tr.querySelector(".mask-word").focus();
    saveRules();
}

// マスキング実行
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

// Word XML マスキング本体
function maskWordXml(xml, words) {
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

    const joined = textNodes.map(n => n.text).join("");
    const masked = applyMask(joined, words);

    // 再分配
    let cursor = 0;
    const replacedTexts = textNodes.map(n => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    // XML書き戻し
    let offset = 0;
    textNodes.forEach((node, i) => {
        const replaced = node.full.replace(node.text, replacedTexts[i]);
        xml =
            xml.slice(0, node.start + offset) +
            replaced +
            xml.slice(node.end + offset);
        offset += replaced.length - node.full.length;
    });

    return xml;
}

// マスク処理（完全一致のみ）
function applyMask(text, words) {
    let result = text;
    words.forEach(rule => {
        if (!rule) return;
        const escaped = rule.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(escaped, "g");
        result = result.replace(re, m => MASK_CHAR.repeat(m.length));
    });
    return result;
}

// 有効ルール取得
function getEnabledRules() {
    return Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? { value: word.value.trim() } : null;
        })
        .filter(Boolean);
}

// ファイル選択チェック
function getSelectedFile() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) alert("Wordファイルを選択してください");
    return file;
}

// localStorage 保存
function saveRules() {
    const rules = Array.from(document.querySelectorAll("#maskTable tbody tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return {
                enabled: enable.checked,
                value: word.value
            };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

// localStorage からロード + デフォルトも表示
function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    const combined = [...DEFAULT_MASK_RULES, ...saved];

    combined.forEach(rule => addRuleRow(rule));
}

// ダウンロード補助
function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}
