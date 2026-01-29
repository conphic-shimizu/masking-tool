document.getElementById("runBtn").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
        alert("Wordファイルを選択してください");
        return;
    }

    const maskWords = document
        .getElementById("maskWords")
        .value.split("\n")
        .map(w => w.trim())
        .filter(Boolean);

    if (maskWords.length === 0) {
        alert("マスキングする文字列を入力してください");
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
    maskWords.forEach(word => {
        const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const mask = "■".repeat(word.length);
        const regex = new RegExp(escaped, "g");
        xml = xml.replace(regex, mask);
    });

    // ④ 書き戻す
    zip.file("word/document.xml", xml);

    // ⑤ 新しいdocxを生成
    const blob = await zip.generateAsync({ type: "blob" });

    download(blob, file.name.replace(".docx", "_masked.docx"));
});

function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}
