/* =====================================================
   定数
===================================================== */
const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* ==============================
   デフォルトのマスキングプリセット
============================== */
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
   Word XML マスキング本体
===================================================== */
function maskWordXml(xml, words) {
    // <w:t> ノードを全取得
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

    // 連結 → マスク
    const joined = textNodes.map(n => n.text).join("");
    let masked = applyMask(joined, words);

    // 元の分割単位に再分配
    let cursor = 0;
    const replacedTexts = textNodes.map(n => {
        const part = masked.slice(cursor, cursor + n.text.length);
        cursor += n.text.length;
        return part;
    });

    // XMLへ書き戻し（offset考慮）
    let offset = 0;
    textNodes.forEach((node, i) => {
        const replaced = node.full.replace(
            node.text,
            escapeXml(replacedTexts[i])
        );

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
function applyMask(text, words) {
    let result = text;

    words.forEach(rule => {
        if (!rule) return;

        if (rule.isRegex) {
            // 正規表現モード
            try {
                const re = new RegExp(rule.value, "g");
                result = result.replace(re, m => "■".repeat(m.length));
            } catch (e) {
                console.warn("無効な正規表現:", rule.value);
            }
        } else {
            // 通常文字列モード（完全一致）
            const escaped = rule.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const re = new RegExp(escaped, "g");
            result = result.replace(re, m => "■".repeat(m.length));
        }
    });

    return result;
}

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
   ルール取得
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

/* ==============================
   ルール読み込み
============================== */
function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    const saved = localStorage.getItem(STORAGE_KEY);
    let rules;

    if (saved) {
        rules = JSON.parse(saved);
    } else {
        rules = [
            { value: "社名", enabled: true, isRegex: false },
            { value: "郵便番号", enabled: true, isRegex: false },
            { value: "住所", enabled: true, isRegex: false },
            { value: "電話番号", enabled: true, isRegex: false },
            { value: "FAX番号", enabled: true, isRegex: false },
            { value: "メールアドレス", enabled: true, isRegex: false },
        ];
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

function loadDocumentXml(file) {
    return file.arrayBuffer()
        .then(buf => JSZip.loadAsync(buf))
        .then(zip => {
            const doc = zip.file("word/document.xml");
            if (!doc) {
                alert("document.xml が見つかりません");
                return null;
            }
            return doc.async("string");
        });
}

function escapeXml(str) {
    return str
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}
