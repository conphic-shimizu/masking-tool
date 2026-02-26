// ===== Storage Keys =====
const STORAGE_KEY = "word-mask-rules";
const THEME_KEY = "ui-theme";

// ===== Mask settings =====
const MASK_CHAR = "■";

// ===== Default rules =====
const DEFAULT_MASK_RULES = [
    { value: "株式会社hogehoge", enabled: true },
    { value: "東京都テスト区テスト町1-2-3", enabled: true },
    { value: "100-0022", enabled: true },
    { value: "012-345-6789", enabled: true },
    { value: "sample@hogehoge.co.jp", enabled: true },
];

document.addEventListener("DOMContentLoaded", () => {
    applyThemeFromStorage();
    bindEvents();
    loadRules();
    syncSelectedFileName();
});

// =========================
// Theme
// =========================
function applyThemeFromStorage() {
    const t = localStorage.getItem(THEME_KEY) || "dark";
    document.documentElement.dataset.theme = t;
}

function toggleTheme() {
    const cur = document.documentElement.dataset.theme || "dark";
    const next = cur === "dark" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    localStorage.setItem(THEME_KEY, next);
}

// =========================
// README dialog
// =========================
function openReadmeDialog() {
    const dlg = document.getElementById("readmeDialog");
    const body = document.getElementById("readmeBody");
    const src = document.getElementById("readmeContent");

    if (!dlg || !body || !src) return;

    const md = (src.textContent || "").trim();

    // markedが使えるならMarkdown→HTML化して表示
    if (window.marked && typeof window.marked.parse === "function") {
        // breaks:true = Markdownの改行を<br>として扱う（README向けに読みやすい）
        body.innerHTML = window.marked.parse(md, { gfm: true, breaks: true });
    } else {
        // フォールバック：プレーン表示
        body.textContent = md;
    }

    if (typeof dlg.showModal === "function") dlg.showModal();
    else dlg.setAttribute("open", "");
}

function closeReadmeDialog() {
    const dlg = document.getElementById("readmeDialog");
    if (!dlg) return;

    if (typeof dlg.close === "function") dlg.close();
    else dlg.removeAttribute("open");
}

async function copyReadmeToClipboard() {
    const src = document.getElementById("readmeContent");
    const text = (src?.textContent || "").trim();
    if (!text) return;

    try {
        await navigator.clipboard.writeText(text);
        alert("READMEをコピーしました");
    } catch {
        // clipboardが使えない環境用フォールバック
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        ta.remove();
        alert("READMEをコピーしました");
    }
}

// =========================
// Events
// =========================
function bindEvents() {
    document.getElementById("themeToggle")?.addEventListener("click", toggleTheme);

    // README
    document.getElementById("readmeBtn")?.addEventListener("click", openReadmeDialog);
    document.getElementById("readmeCloseBtn")?.addEventListener("click", closeReadmeDialog);
    document.getElementById("readmeCopyBtn")?.addEventListener("click", copyReadmeToClipboard);

    // ダイアログ外側クリックで閉じる
    const dlg = document.getElementById("readmeDialog");
    dlg?.addEventListener("click", (e) => {
        if (e.target === dlg) closeReadmeDialog();
    });

    // 「同じファイルを選び直す」と change が発火しないことがあるので、事前に value を空にする
    document.getElementById("filePickBtn")?.addEventListener("click", () => {
        const input = document.getElementById("fileInput");
        if (input) input.value = "";
    });

    document.getElementById("fileInput")?.addEventListener("change", syncSelectedFileName);

    document.getElementById("addRowBtn")?.addEventListener("click", () => {
        addRuleRow({ value: "", enabled: true });
        saveRules();
    });

    document.getElementById("resetRulesBtn")?.addEventListener("click", () => {
        const ok = confirm("ルールを初期化します。よろしいですか？（localStorageの保存も消えます）");
        if (!ok) return;
        localStorage.removeItem(STORAGE_KEY);
        loadRules();
    });

    document.getElementById("runBtn")?.addEventListener("click", runMasking);

    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    // 入力/チェック変更で保存
    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);

    // 削除
    tbody.addEventListener("click", (e) => {
        const btn = e.target?.closest?.("button[data-action='delete']");
        if (!btn) return;
        const tr = btn.closest("tr");
        tr?.remove();
        saveRules();
    });

    // DnD 並べ替え
    setupDragSort(tbody);
}

function syncSelectedFileName() {
    const input = document.getElementById("fileInput");
    const nameEl = document.getElementById("fileName");
    if (!nameEl) return;

    const f = input?.files?.[0];
    nameEl.textContent = f ? f.name : "未選択";
}

// =========================
// Drag & Drop sorting (HTML5 DnD)
// =========================
function setupDragSort(tbody) {
    let draggingRow = null;

    tbody.addEventListener("dragstart", (e) => {
        const handle = e.target?.closest?.(".dragHandle");
        if (!handle) {
            e.preventDefault();
            return;
        }

        const tr = handle.closest("tr");
        if (!tr) return;

        draggingRow = tr;
        tr.classList.add("is-dragging");

        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", "move"); // Firefox対策
    });

    tbody.addEventListener("dragover", (e) => {
        if (!draggingRow) return;
        e.preventDefault();

        const target = e.target?.closest?.("tr");
        if (!target || target === draggingRow) return;

        clearDragOverClasses(tbody);

        const rect = target.getBoundingClientRect();
        const isAfter = (e.clientY - rect.top) > (rect.height / 2);

        target.classList.add(isAfter ? "drag-over-bottom" : "drag-over-top");
        tbody.insertBefore(draggingRow, isAfter ? target.nextSibling : target);
    });

    tbody.addEventListener("drop", (e) => {
        if (!draggingRow) return;
        e.preventDefault();
    });

    tbody.addEventListener("dragend", () => {
        if (!draggingRow) return;
        draggingRow.classList.remove("is-dragging");
        draggingRow = null;
        clearDragOverClasses(tbody);
        saveRules(); // 並び順を保存
    });

    function clearDragOverClasses(tbodyEl) {
        tbodyEl.querySelectorAll("tr.drag-over-top, tr.drag-over-bottom")
            .forEach(tr => tr.classList.remove("drag-over-top", "drag-over-bottom"));
    }
}

// =========================
// Rules table
// =========================
function loadRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    const savedRaw = localStorage.getItem(STORAGE_KEY);
    const parsed = savedRaw ? safeJsonParse(savedRaw, []) : [];
    const saved = Array.isArray(parsed) ? parsed : [];

    // DEFAULT + saved をマージ（saved優先、重複はvalueキーで上書き）
    const map = new Map();
    DEFAULT_MASK_RULES.forEach(r => map.set(r.value, { value: r.value, enabled: !!r.enabled }));
    saved.forEach(r => {
        if (!r || typeof r.value !== "string") return;
        map.set(r.value, { value: r.value, enabled: !!r.enabled });
    });

    Array.from(map.values()).forEach(rule => addRuleRow(rule));
}

function addRuleRow(rule) {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const tr = document.createElement("tr");
    tr.innerHTML = `
    <td>
      <span class="dragHandle" draggable="true" title="ドラッグして並べ替え">⠿</span>
    </td>
    <td style="text-align:center;">
      <input type="checkbox" class="mask-enable" ${rule.enabled ? "checked" : ""} />
    </td>
    <td>
      <input type="text" class="mask-word"
        value="${escapeHtmlAttr(rule.value ?? "")}"
        placeholder="例: 田中太郎 / 03-1234-5678 / 顧客ID12345" />
    </td>
    <td>
      <button type="button" class="btn danger" data-action="delete">削除</button>
    </td>
  `;
    tbody.appendChild(tr);
}

function getEnabledRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return [];

    return Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            if (!enable.checked) return null;
            const v = (word.value ?? "").trim();
            return v ? v : null;
        })
        .filter(Boolean);
}

function saveRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const rules = Array.from(tbody.querySelectorAll("tr"))
        .map(tr => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return { enabled: !!enable.checked, value: word.value ?? "" };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

function safeJsonParse(text, fallback) {
    try { return JSON.parse(text); } catch { return fallback; }
}

function escapeHtmlAttr(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

// =========================
// Main
// =========================
async function runMasking() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput?.files?.[0];

    if (!file) {
        alert("ファイル（.docx / .pptx / .xlsx）を選択してください");
        return;
    }

    const words = getEnabledRules();
    if (words.length === 0) {
        alert("有効なマスキング対象（空でない固定文字列）がありません");
        return;
    }

    const ext = getLowerExt(file.name);
    if (!["docx", "pptx", "xlsx"].includes(ext)) {
        alert("対応形式は .docx / .pptx / .xlsx です");
        return;
    }

    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    try {
        if (ext === "docx") {
            await maskDocx(zip, words);
            const blob = await zip.generateAsync({ type: "blob" });

            const outName = buildMaskedOutputFilename(file.name, words, "_マスキング");
            download(blob, outName);
            return;
        }

        if (ext === "pptx") {
            const result = await maskPptx(zip, words);
            if (result.slideFiles === 0) {
                alert("ppt/slides/slide*.xml が見つかりません（PPTX形式でない可能性）");
                return;
            }
            const blob = await zip.generateAsync({ type: "blob" });

            const outName = buildMaskedOutputFilename(file.name, words, "_マスキング");
            download(blob, outName);
            return;
        }

        if (ext === "xlsx") {
            const result = await maskXlsx(zip, words);
            if (!result.sharedStringsFound && result.inlineCells === 0) {
                alert("対象（xl/sharedStrings.xml または inlineStr）が見つかりませんでした（形式差異の可能性）");
            }
            const blob = await zip.generateAsync({ type: "blob" });

            const outName = buildMaskedOutputFilename(file.name, words, "_マスキング");
            download(blob, outName);
            return;
        }
    } catch (e) {
        console.error(e);
        alert(`エラー: ${e?.message ?? e}`);
    }
}

// =========================
// DOCX
// =========================
async function maskDocx(zip, literalWords) {
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) throw new Error("word/document.xml が見つかりません");

    let xml = await docXmlFile.async("string");
    xml = maskWordXmlByLiterals(xml, literalWords);
    zip.file("word/document.xml", xml);
}

// Word: <w:t> 連結→マスク→再分配（固定文字列）
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

    if (nodes.length === 0) return xml;

    const joined = nodes.map(n => n.text).join("");
    let masked = joined;

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

// <w:t> の内側だけ安全に差し替える
function replaceWtInner(fullWt, newInner) {
    return fullWt.replace(
        /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/,
        (_, p1, _old, p3) => p1 + newInner + p3
    );
}

// =========================
// PPTX
// =========================
async function maskPptx(zip, literalWords) {
    const slidePaths = getPptxSlidePaths(zip);
    let totalChanged = 0;

    for (const path of slidePaths) {
        const f = zip.file(path);
        if (!f) continue;

        const xmlText = await f.async("string");
        const { xml, changed } = maskPptxSlideXmlByLiterals(xmlText, literalWords, path);
        if (changed) {
            zip.file(path, xml);
            totalChanged++;
        }
    }

    return { slideFiles: slidePaths.length, changedSlides: totalChanged };
}

function getPptxSlidePaths(zip) {
    const re = /^ppt\/slides\/slide\d+\.xml$/;
    const paths = [];
    zip.forEach((relativePath, file) => {
        if (!file.dir && re.test(relativePath)) paths.push(relativePath);
    });

    paths.sort((a, b) => {
        const na = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] ?? "0", 10);
        const nb = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] ?? "0", 10);
        return na - nb;
    });

    return paths;
}

/**
 * PPTX:
 * - これまで：<a:p>単位で処理（段落をまたぐ一致は不可）
 * - 改善後：同一 <p:txBody>（同一テキストボックス）単位で、
 *   <a:p> / <a:br> / <a:tab> / テキスト中の \r\n を “マッチング上は無視” して連結→マスク→再分配
 */
function maskPptxSlideXmlByLiterals(xmlText, literalWords, debugName = "") {
    const doc = parseXml(xmlText, debugName);

    const words = literalWords
        .map(w => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    // テキストボックス単位（txBody）で処理
    const txBodies = getElementsByLocalName(doc, "txBody");
    if (txBodies.length === 0) return { xml: xmlText, changed: false };

    let changed = false;

    for (const txBody of txBodies) {
        // txBody配下の段落 <a:p> を順番に集める
        const paragraphs = [];
        for (let i = 0; i < txBody.childNodes.length; i++) {
            const n = txBody.childNodes[i];
            if (n.nodeType === 1 && n.localName === "p") paragraphs.push(n);
        }
        if (paragraphs.length === 0) continue;

        // ① txBody内を表示順に見て、<a:t>（文字）と区切り（段落/改行/タブ）を扱う
        //    マッチング用の flat（区切り無視）と、flatの各文字→元ノード位置対応表を作る
        let flat = "";
        const indexMap = []; // indexMap[i] = { node: <a:t>, offset: number }

        // txBody全体の <a:t> を順に辿る。ただし段落境界も保持したいので段落ごとに回す
        for (const p of paragraphs) {
            // 段落内を走査して <a:t>, <a:br>, <a:tab> を拾う
            const walker = doc.createTreeWalker(
                p,
                NodeFilter.SHOW_ELEMENT,
                {
                    acceptNode(node) {
                        if (node.localName === "t" || node.localName === "br" || node.localName === "tab") {
                            return NodeFilter.FILTER_ACCEPT;
                        }
                        return NodeFilter.FILTER_SKIP;
                    }
                }
            );

            while (walker.nextNode()) {
                const node = walker.currentNode;

                if (node.localName === "t") {
                    const text = node.textContent ?? "";
                    // text中の \r \n も “区切り扱い” としてマッチングから除外する（重要）
                    for (let j = 0; j < text.length; j++) {
                        const ch = text[j];
                        if (ch === "\n" || ch === "\r") {
                            // ここは区切り扱い：flatに入れない・indexMapも作らない
                            continue;
                        }
                        flat += ch;
                        indexMap.push({ node, offset: j });
                    }
                } else {
                    // <a:br/> / <a:tab/> は区切り扱い：flatに入れない
                    continue;
                }
            }

            // ★段落境界も “区切り扱い” として無視（flatに入れない）
            // 何もしない＝無視
        }

        if (!flat) continue;

        // ② flat上で固定文字列をマスク
        let maskedFlat = flat;
        const beforeAll = maskedFlat;

        for (const w of words) {
            maskedFlat = replaceAllLiteralKeepingLength(maskedFlat, w, MASK_CHAR);
        }
        if (maskedFlat === beforeAll) continue;

        // ③ 元の <a:t> を文字配列として用意し、indexMapに従って差し替える
        //    ※改行(\n\r)や区切りはindexMapに含めていないので保持される
        const touched = new Map(); // node -> charArray
        for (const m of indexMap) {
            if (!touched.has(m.node)) {
                touched.set(m.node, (m.node.textContent ?? "").split(""));
            }
        }

        for (let i = 0; i < maskedFlat.length; i++) {
            const m = indexMap[i];
            if (!m) continue;
            const arr = touched.get(m.node);
            if (!arr) continue;
            arr[m.offset] = maskedFlat[i];
        }

        for (const [node, arr] of touched.entries()) {
            node.textContent = arr.join("");
        }

        changed = true;
    }

    const out = new XMLSerializer().serializeToString(doc);
    return { xml: out, changed };
}

// =========================
// XLSX (sharedStrings + inlineStr)
// =========================
async function maskXlsx(zip, literalWords) {
    const words = literalWords
        .map(w => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    let sharedStringsFound = false;
    let inlineCells = 0;

    // --- 1) xl/sharedStrings.xml ---
    const sstPath = "xl/sharedStrings.xml";
    const sstFile = zip.file(sstPath);
    if (sstFile) {
        sharedStringsFound = true;
        const xmlText = await sstFile.async("string");
        const doc = parseXml(xmlText, sstPath);

        const siNodes = getElementsByLocalName(doc, "si");
        for (const si of siNodes) {
            const tNodes = collectDescendantTNodes(doc, si);
            if (tNodes.length === 0) continue;
            maskAndRedistributeTextNodes(tNodes, words);
        }

        zip.file(sstPath, new XMLSerializer().serializeToString(doc));
    }

    // --- 2) xl/worksheets/*.xml : inlineStr ---
    const sheetPaths = [];
    zip.forEach((relativePath, file) => {
        if (!file.dir && /^xl\/worksheets\/sheet\d+\.xml$/.test(relativePath)) sheetPaths.push(relativePath);
    });

    sheetPaths.sort((a, b) => {
        const na = parseInt(a.match(/sheet(\d+)\.xml$/)?.[1] ?? "0", 10);
        const nb = parseInt(b.match(/sheet(\d+)\.xml$/)?.[1] ?? "0", 10);
        return na - nb;
    });

    for (const path of sheetPaths) {
        const f = zip.file(path);
        if (!f) continue;

        const xmlText = await f.async("string");
        const doc = parseXml(xmlText, path);

        // c[@t="inlineStr"] の <is> 配下の <t> を対象
        const cellNodes = getElementsByLocalName(doc, "c");
        let changed = false;

        for (const c of cellNodes) {
            const tAttr = c.getAttribute("t");
            if (tAttr !== "inlineStr") continue;

            const isNodes = Array.from(c.children).filter(n => n.localName === "is");
            if (isNodes.length === 0) continue;

            const tNodes = collectDescendantTNodes(doc, isNodes[0]);
            if (tNodes.length === 0) continue;

            const did = maskAndRedistributeTextNodes(tNodes, words);
            if (did) {
                inlineCells++;
                changed = true;
            }
        }

        if (changed) {
            zip.file(path, new XMLSerializer().serializeToString(doc));
        }
    }

    return { sharedStringsFound, inlineCells };
}

// =========================
// Shared helpers (XML)
// =========================
function parseXml(xmlText, fileName = "") {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "application/xml");
    const err = doc.getElementsByTagName("parsererror")[0];
    if (err) throw new Error(`XML parse error: ${fileName || "(unknown)"}`);
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

function collectDescendantTNodes(doc, rootNode) {
    const out = [];
    const walker = doc.createTreeWalker(
        rootNode,
        NodeFilter.SHOW_ELEMENT,
        {
            acceptNode(node) {
                return node.localName === "t"
                    ? NodeFilter.FILTER_ACCEPT
                    : NodeFilter.FILTER_SKIP;
            }
        }
    );
    while (walker.nextNode()) out.push(walker.currentNode);
    return out;
}

/**
 * 与えられた <t> ノード群を「連結→マスク→再分配」する
 * @returns {boolean} changed
 */
function maskAndRedistributeTextNodes(tNodes, words) {
    const texts = tNodes.map(n => n.textContent ?? "");
    const lens = texts.map(s => s.length);
    const joined = texts.join("");
    if (!joined) return false;

    let masked = joined;
    const beforeAll = masked;

    for (const w of words) {
        masked = replaceAllLiteralKeepingLength(masked, w, MASK_CHAR);
    }
    if (masked === beforeAll) return false;

    let cursor = 0;
    for (let i = 0; i < tNodes.length; i++) {
        const part = masked.slice(cursor, cursor + lens[i]);
        cursor += lens[i];
        tNodes[i].textContent = part;
    }
    return true;
}

// =========================
// Replace: fixed string → mask (same length)
// =========================
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

// =========================
// Output filename (remove rule words in filename by exact match)
// =========================
function sanitizeFilenameForOS(name) {
    if (!name) return "";
    return String(name)
        // Windows禁止文字
        .replace(/[\\\/:*?"<>|]/g, "")
        // 制御文字
        .replace(/[\u0000-\u001f\u007f]/g, "")
        // 末尾のドット/スペースはOSで問題になりやすい
        .replace(/[. ]+$/g, "")
        .trim();
}

/**
 * ファイル名（拡張子除く）を区切り文字で分割し、
 * トークンがルールと「完全一致」した場合のみ削除する。
 */
function removeRuleWordsFromBaseNameExact(baseName, ruleWords) {
    const words = Array.from(
        new Set((ruleWords || []).map(w => (w ?? "").toString().trim()).filter(Boolean))
    );
    if (!baseName || words.length === 0) return baseName;

    // トークン分割に使う区切り（広め）
    const sepRe = /([ _\-.・、,，\(\)\[\]\{\}【】（）「」『』]+)/;

    // 区切りを保持したまま分割
    const parts = String(baseName).split(sepRe);

    const removed = parts.map((part) => {
        if (!part) return "";
        if (sepRe.test(part)) return part; // 区切りは保持
        const token = part.trim();
        if (!token) return "";
        return words.includes(token) ? "" : part;
    });

    // 再結合→見た目を整える（連続区切りは "_" に寄せる）
    let joined = removed.join("");
    joined = joined
        .replace(/[ _\-.・、,，\(\)\[\]\{\}【】（）「」『』]+/g, "_")
        .replace(/^_+|_+$/g, "");

    return joined;
}

function buildMaskedOutputFilename(inputFilename, ruleWords, suffix = "_マスキング") {
    const name = String(inputFilename || "");
    const dot = name.lastIndexOf(".");
    const hasExt = dot > 0 && dot < name.length - 1;

    const base = hasExt ? name.slice(0, dot) : name;
    const ext = hasExt ? name.slice(dot) : "";

    const stripped = removeRuleWordsFromBaseNameExact(base, ruleWords);
    const safeBase = sanitizeFilenameForOS(stripped) || "masked";
    const outBase = sanitizeFilenameForOS(`${safeBase}${suffix || ""}`) || `masked${suffix || ""}`;

    return `${outBase}${ext}`;
}

// =========================
// Utils
// =========================
function getLowerExt(name) {
    const m = String(name).toLowerCase().match(/\.([a-z0-9]+)$/);
    return m ? m[1] : "";
}

function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}
