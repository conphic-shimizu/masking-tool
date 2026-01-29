const STORAGE_KEY = "word-mask-rules";

/* =========================
   デフォルトのマスキングルール
========================= */
const DEFAULT_MASK_RULES = [
    "コンフィック",
    "190-0022",
    "東京都立川市錦町1-4-4立川サニーハイツ303",
    "042-595-7557",
    "042-595-7558",
    "@conphic.co.jp"
];

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
   マスキング実行
========================= */
document.getElementById("runBtn").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

    // 有効なルールを取得
    const rules = getEnabledRules();
    if (rules.length === 0) {
        alert("マスキング対象がありません");
        return;
    }

    // ZIPとして読み込み
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
        alert("document.xml が見つかりません");
        return;
    }

    let xml = await docXmlFile.async("string");

    // マスキング
    rules.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const mask = "■".repeat(word.length);
        const re = new RegExp(escaped, "g");
        xml = xml.replace(re, mask);
    });

    // 書き戻し
    zip.file("word/document.xml", xml);

    // 新しいdocx生成
    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(".docx", "_masked.docx"));
});

/* =========================
   ルール操作
========================= */
function getEnabledRules() {
    return Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    )
        .filter(tr => tr.querySelector(".mask-enable").checked)
        .map(tr => tr.querySelector(".mask-word").value.trim())
        .filter(Boolean);
}

/* =========================
   localStorage
========================= */
function saveRules() {
    const rules = Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    ).map(tr => ({
        enabled: tr.querySelector(".mask-enable").checked,
        value: tr.querySelector(".mask-word").value
    }));

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

/* =========================
   ルール読み込み
========================= */
function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    tbody.innerHTML = "";

    const saved = localStorage.getItem(STORAGE_KEY);
    let rules;

    if (saved) {
        rules = JSON.parse(saved);
    } else {
        rules = DEFAULT_MASK_RULES.map(value => ({ enabled: true, value }));
    }

    rules.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
            <td><input type="text" class="mask-word" value="${rule.value}"></td>
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
   初期化
========================= */
document
    .querySelector("#maskTable tbody")
    .addEventListener("input", saveRules);
document
    .querySelector("#maskTable tbody")
    .addEventListener("change", saveRules);

document.addEventListener("DOMContentLoaded", loadRules);
