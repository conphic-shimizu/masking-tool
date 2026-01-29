const STORAGE_KEY = "word-mask-rules";

/* =========================
   行追加
========================= */
document.getElementById("addRowBtn").addEventListener("click", () => {
    const tbody = document.querySelector("#maskTable tbody");
    const tr = document.createElement("tr");

    tr.innerHTML = `
        <td><input type="checkbox" class="mask-enable" checked></td>
        <td><input type="text" class="mask-word"></td>
    `;

    tbody.appendChild(tr);
    saveRules();
});

/* =========================
   プレビュー
========================= */
document.getElementById("previewBtn").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

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

    const xml = await docXmlFile.async("string");

    const { joined, masked } = buildPreview(xml, rules);

    const area = document.getElementById("previewArea");
    const text = document.getElementById("previewText");

    text.textContent =
        "【元のテキスト】\n" +
        joined +
        "\n\n【マスキング後】\n" +
        masked;

    area.style.display = "block";
});

/* =========================
   マスキング実行
========================= */
document.getElementById("runBtn").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

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
});

/* =========================
   Word XML マスキング本体
========================= */
function maskWordXml(xml, words) {
    // <w:t>...</w:t> を全部拾う
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

    const maskChar = "■";

    words.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(escaped, "g");
        masked = masked.replace(re, m => maskChar.repeat(m.length));
    });

    // 再分配
    let cursor = 0;
    const replacedNodes = textNodes.map(n => {
        const len = n.text.length;
        const part = masked.slice(cursor, cursor + len);
        cursor += len;
        return part;
    });

    // XML 書き戻し
    let offset = 0;
    textNodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        const replaced = node.full.replace(
            node.text,
            escapeXml(replacedNodes[i])
        );

        xml = before + replaced + after;
        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =========================
   補助関数
========================= */
function getEnabledRules() {
    return Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    )
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return { enable, word };
        })
        .filter(v => v && v.enable.checked)
        .map(v => v.word.value.trim())
        .filter(Boolean);
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

/* =========================
   localStorage
========================= */
function saveRules() {
    const rules = Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    )
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

function loadRules() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return;

    const rules = JSON.parse(saved);
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    rules.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>
                <input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}>
            </td>
            <td>
                <input type="text" class="mask-word" value="${rule.value}">
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function buildPreview(xml, words) {
    const regex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    const textNodes = [];

    let match;
    while ((match = regex.exec(xml)) !== null) {
        textNodes.push(match[1]);
    }

    const joined = textNodes.join("");
    let masked = joined;

    const maskChar = "■";

    words.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(escaped, "g");
        masked = masked.replace(re, m => maskChar.repeat(m.length));
    });

    return { joined, masked };
}

document
    .querySelector("#maskTable tbody")
    .addEventListener("input", saveRules);

document
    .querySelector("#maskTable tbody")
    .addEventListener("change", saveRules);

document.addEventListener("DOMContentLoaded", loadRules);

