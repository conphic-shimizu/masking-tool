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
   マスキング実行
========================= */
document.getElementById("runBtn").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

    const maskWords = Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    )
        .filter(tr => tr.querySelector(".mask-enable").checked)
        .map(tr => tr.querySelector(".mask-word").value.trim())
        .filter(Boolean);

    if (maskWords.length === 0) {
        alert("マスキング対象の文字列がありません");
        return;
    }

    // ① WordファイルをZIPとして読み込む
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    // ② document.xml を取得
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
        alert("document.xml が見つかりません");
        return;
    }

    let xml = await docXmlFile.async("string");

    // ③ マスキング処理
    const maskChar = "■";

    maskWords.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const mask = maskChar.repeat(word.length);
        const regex = new RegExp(escaped, "g");
        xml = xml.replace(regex, mask);
    });

    // ④ 書き戻す
    zip.file("word/document.xml", xml);

    // ⑤ 新しいdocxを生成
    const blob = await zip.generateAsync({ type: "blob" });
    download(blob, file.name.replace(".docx", "_masked.docx"));
});

/* =========================
   ダウンロード
========================= */
function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}

/* =========================
   ルール保存
========================= */
function saveRules() {
    const rules = Array.from(
        document.querySelectorAll("#maskTable tbody tr")
    ).map(tr => {
        return {
            enabled: tr.querySelector(".mask-enable").checked,
            value: tr.querySelector(".mask-word").value
        };
    });

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

/* =========================
   ルール復元
========================= */
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

/* =========================
   入力変更時に自動保存
========================= */
document
    .querySelector("#maskTable tbody")
    .addEventListener("input", saveRules);

document
    .querySelector("#maskTable tbody")
    .addEventListener("change", saveRules);

/* =========================
   初期化
========================= */
document.addEventListener("DOMContentLoaded", () => {
    loadRules();
});
