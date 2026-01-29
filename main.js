/* =====================================================
   定数
===================================================== */
const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* ==============================
   デフォルトマスキングプリセット
   - 通常文字列 or 正規表現対応
============================== */
const DEFAULT_MASK_RULES = [
    { value: "コンフィック", enabled: true, isRegex: false },
    { value: "\\d{3}-\\d{4}", enabled: true, isRegex: true }, // 郵便番号
    { value: "042-595-7557", enabled: true, isRegex: false },
    { value: "042-595-7558", enabled: true, isRegex: false },
    { value: "@\\w+\\.\\w+", enabled: true, isRegex: true } // メールアドレス
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
   Word XML マスキング本体（破損防止版）
===================================================== */
function maskWordXml(xml, rules) {
    const textNodes = [];
    const regex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let match;

    // <w:t> をすべて取得
    while ((match = regex.exec(xml)) !== null) {
        textNodes.push({
            full: match[0],
            text: match[1],
            start: match.index,
            end: regex.lastIndex
        });
    }

    if (textNodes.length === 0) return xml;

    // 連結してマスク
    const joined = textNodes.map(n => n.text).join("");
    const masked = applyMask(joined, rules);

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
        const safeText = escapeXml(replacedTexts[i]);
        const replaced = node.full.replace(node.text, safeText);

        xml =
            xml.slice(0, node.start + offset) +
            replaced +
            xml.slice(node.end + offset);

        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =====================================================
   マスキング共通処理
===================================================== */
function applyMask(text, rules) {
    let result = text;

    rules.forEach(rule => {
        if (!rule || !rule.value) return;

        if (rule.isRegex) {
            try {
                const re = new RegExp(rule.value, "g");
                result = result.replace(re, m => MASK_CHAR.repeat(m.length));
            } catch (e) {
                console.warn("無効な正規表現:", rule.value);
            }
        } else {
            const escaped = rule.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const re = new RegExp(escaped, "g");
            result = result.replace(re, m => MASK_CHAR.repeat(m.length));
        }
    });

    return result;
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
            return {
                enabled: enable.checked,
                value: word.value,
                isRegex: regex.checked
            };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    // ① デフォルトルールを先に表示
    DEFAULT_MASK_RULES.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
            <td><input type="text" class="mask-word" value="${rule.value}" readonly></td>
            <td><input type="checkbox" class="mask-regex" ${rule.isRegex ? "checked" : ""} disabled></td>
        `;
        tbody.appendChild(tr);
    });

    // ② ローカルストレージに保存されたルールを表示
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
        const userRules = JSON.parse(saved);
        userRules.forEach(rule => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
                <td><input type="text" class="mask-word" value="${rule.value}"></td>
                <td><input type="checkbox" class="mask-regex" ${rule.isRegex ? "checked" : ""}></td>
            `;
            tbody.appendChild(tr);
        });
    }
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

function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}
