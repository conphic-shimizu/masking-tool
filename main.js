const STORAGE_KEY = "word-mask-rules";
const MASK_CHAR = "■";

/* =========================
   初期ルール（デフォルト）
   - 正規表現は使わない（完全一致の部分文字列マスク）
========================= */
const DEFAULT_MASK_RULES = [
    { value: "株式会社コンフィック", enabled: true },
    { value: "コンフィック", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true },
    { value: "daichi@conphic.co.jp", enabled: false },
];

/* =========================
   初期化
========================= */
document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadRules();
});

/* =========================
   イベント登録
========================= */
function bindEvents() {
    document.getElementById("addRowBtn")?.addEventListener("click", addRuleRow);
    document.getElementById("runBtn")?.addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);
}

/* =========================
   行追加
========================= */
function addRuleRow() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const tr = document.createElement("tr");
    tr.innerHTML = `
    <td><input type="checkbox" class="mask-enable" checked></td>
    <td><input type="text" class="mask-word" value=""></td>
  `;
    tbody.appendChild(tr);
    tr.querySelector(".mask-word")?.focus();
    saveRules();
}

/* =========================
   マスキング実行（Word/PPTX 両対応）
========================= */
async function runMasking() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput?.files?.[0];
    if (!file) {
        alert("ファイル（.docx / .pptx）を選択してください");
        return;
    }

    const rules = getEnabledRules(); // [{ value }]
    if (rules.length === 0) {
        alert("マスキング対象がありません");
        return;
    }

    const ext = getLowerExt(file.name);
    if (ext !== "docx" && ext !== "pptx") {
        alert("対応形式は .docx / .pptx です");
        return;
    }

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    if (ext === "docx") {
        await maskDocx(zip, rules.map(r => r.value));
        const blob = await zip.generateAsync({ type: "blob" });
        download(blob, file.name.replace(/\.docx$/i, "_マスキング.docx"));
        return;
    }

    if (ext === "pptx") {
        const result = await maskPptx(zip, rules.map(r => r.value));
        if (result.slideFiles === 0) {
            alert("ppt/slides/slide*.xml が見つかりません（PPTX形式でない可能性）");
            return;
        }
        const blob = await zip.generateAsync({ type: "blob" });
        download(blob, file.name.replace(/\.pptx$/i, "_マスキング.pptx"));
        return;
    }
}

/* =========================
   DOCX
========================= */
async function maskDocx(zip, literalWords) {
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
        alert("document.xml が見つかりません");
        throw new Error("word/document.xml not found");
    }

    let xml = await docXmlFile.async("string");
    xml = maskWordXmlByLiterals(xml, literalWords);
    zip.file("word/document.xml", xml);
}

/* =========================
   PPTX（slide*.xml の <a:t> を段落 <a:p> 単位で）
========================= */
async function maskPptx(zip, literalWords) {
    const slidePaths = getPptxSlidePaths(zip);
    let totalHits = 0;

    for (const path of slidePaths) {
        const f = zip.file(path);
        if (!f) continue;

        const xmlText = await f.async("string");
        const { xml, hits } = maskPptxSlideXmlByLiterals(xmlText, literalWords, path);

        if (hits > 0) {
            zip.file(path, xml);
            totalHits += hits;
        }
    }

    return { slideFiles: slidePaths.length, totalHits };
}

function getPptxSlidePaths(zip) {
    const re = /^ppt\/slides\/slide\d+\.xml$/;
    const paths = [];
    zip.forEach((relativePath, file) => {
        if (!file.dir && re.test(relativePath)) paths.push(relativePath);
    });

    // slide1, slide2... の順に
    paths.sort((a, b) => {
        const na = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] ?? "0", 10);
        const nb = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] ?? "0", 10);
        return na - nb;
    });

    return paths;
}

/* =========================
   PPTX スライドXML マスキング本体（固定文字列）
   - <a:p>（段落）単位で
   - 段落内の <a:t> を連結
   - 固定文字列でマスク（長さ維持）
   - 元の <a:t> 長で再分配
   - DOMで安全に置換（タグ構造を壊しにくい）
========================= */
function maskPptxSlideXmlByLiterals(xmlText, literalWords, debugName = "") {
    const doc = parseXml(xmlText, debugName);

    // a:p（namespace問わず localName === "p" を拾う）
    const paragraphs = getElementsByLocalName(doc, "p");

    let hits = 0;

    for (const p of paragraphs) {
        // この段落配下の a:t を拾う
        const ts = [];
        const walker = doc.createTreeWalker(
            p,
            NodeFilter.SHOW_ELEMENT,
            {
                acceptNode(node) {
                    return node.localName === "t"
                        ? NodeFilter.FILTER_ACCEPT
                        : NodeFilter.FILTER_SKIP;
                }
            }
        );
        while (walker.nextNode()) ts.push(walker.currentNode);

        if (ts.length === 0) continue;

        const texts = ts.map(n => n.textContent ?? "");
        const lens = texts.map(s => s.length);
        const joined = texts.join("");
        if (joined.length === 0) continue;

        // マスク
        let masked = joined;

        const words = literalWords
            .map(w => (w ?? "").trim())
            .filter(Boolean)
            .sort((a, b) => b.length - a.length);

        let localHits = 0;
        for (const word of words) {
            const before = masked;
            masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
            if (masked !== before) {
                // 差分回数を厳密に数えるのはコスト高いので、
                // “変化あり”を1カウント（目安）にしておく（必要なら改良可）
                localHits++;
            }
        }

        if (localHits === 0) continue;

        // 再分配（元の <a:t> の長さに合わせてslice）
        let cursor = 0;
        for (let i = 0; i < ts.length; i++) {
            const part = masked.slice(cursor, cursor + lens[i]);
            cursor += lens[i];
            ts[i].textContent = part;
        }

        hits += localHits;
    }

    const out = new XMLSerializer().serializeToString(doc);
    return { xml: out, hits };
}

function parseXml(xmlText, fileName) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "application/xml");
    const err = doc.getElementsByTagName("parsererror")[0];
    if (err) {
        throw new Error(`XML parse error: ${fileName || "(unknown)"}`);
    }
    return doc;
}

function getElementsByLocalName(root, localName) {
    const all = root.getElementsByTagName("*");
    const arr = [];
    for (let i = 0; i < all.length; i++) {
        if (all[i].localName === localName) arr.push(all[i]);
    }
    return arr;
}

/* =========================
   Word XML マスキング本体（固定文字列）
   - <w:t>を抽出して連結
   - 固定文字列でマスク（長さ維持）
   - 元の<w:t>長で再分配
   - <w:t>内側だけ関数replaceで差し替え（$問題回避）
========================= */
function maskWordXmlByLiterals(xml, literalWords) {
    const wtRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    const nodes = [];

    let m;
    while ((m = wtRegex.exec(xml)) !== null) {
        nodes.push({
            full: m[0],
            text: m[1],
            start: m.index,
            end: wtRegex.lastIndex,
            len: m[1].length,
        });
    }

    const joined = nodes.map(n => n.text).join("");
    let masked = joined;

    // 長い語から先に置換（部分一致の副作用を減らす）
    const words = literalWords
        .map(w => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    for (const word of words) {
        masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
    }

    // 再分配
    let cursor = 0;
    const parts = nodes.map(n => {
        const part = masked.slice(cursor, cursor + n.len);
        cursor += n.len;
        return part;
    });

    // XMLへ書き戻し
    let offset = 0;
    nodes.forEach((node, i) => {
        const before = xml.slice(0, node.start + offset);
        const after = xml.slice(node.end + offset);

        const replaced = replaceWtInner(node.full, parts[i]);
        xml = before + replaced + after;

        offset += replaced.length - node.full.length;
    });

    return xml;
}

/* =========================
   固定文字列を全部置換（同じ長さの■にする）
========================= */
function replaceAllLiteralKeepingLength(text, word, maskChar) {
    let result = text;
    let idx = 0;

    while (true) {
        const found = result.indexOf(word, idx);
        if (found === -1) break;

        const mask = maskChar.repeat(word.length);
        result = result.slice(0, found) + mask + result.slice(found + word.length);

        idx = found + mask.length;
    }

    return result;
}

/* =========================
   <w:t> の内側だけ安全に差し替える（$問題回避）
========================= */
function replaceWtInner(fullWt, newInner) {
    return fullWt.replace(
        /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/,
        (_, p1, _old, p3) => p1 + newInner + p3
    );
}

/* =========================
   ルール取得
========================= */
function getEnabledRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return [];

    return Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return enable.checked ? { value: word.value.trim() } : null;
        })
        .filter(Boolean)
        .filter(r => r.value.length > 0);
}

/* =========================
   localStorage
========================= */
function saveRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const rules = Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return { enabled: enable.checked, value: word.value };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    const savedRaw = localStorage.getItem(STORAGE_KEY);
    const saved = savedRaw ? safeJsonParse(savedRaw, []) : [];

    // DEFAULT + saved をマージ（value重複排除、saved優先）
    const map = new Map();
    DEFAULT_MASK_RULES.forEach(r => map.set(r.value, { ...r }));
    saved.forEach(r => {
        if (!r || !r.value) return;
        map.set(r.value, { value: r.value, enabled: !!r.enabled });
    });

    const merged = Array.from(map.values());

    merged.forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
      <td><input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""}></td>
      <td><input type="text" class="mask-word" value="${escapeHtml(rule.value)}"></td>
    `;
        tbody.appendChild(tr);
    });
}

function safeJsonParse(text, fallback) {
    try { return JSON.parse(text); } catch { return fallback; }
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
   HTML属性用エスケープ（value="" 崩れ防止）
========================= */
function escapeHtml(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

/* =========================
   拡張子取得
========================= */
function getLowerExt(name) {
    const m = String(name).toLowerCase().match(/\.([a-z0-9]+)$/);
    return m ? m[1] : "";
}
