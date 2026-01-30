// ===== Storage Keys =====
const STORAGE_KEY = "word-mask-rules";
const THEME_KEY = "ui-theme";

// ===== Mask settings =====
const MASK_CHAR = "■";

// ===== Default rules =====
const DEFAULT_MASK_RULES = [
    { value: "株式会社コンフィック", enabled: true },
    { value: "コンフィック", enabled: true },
    { value: "東京都立川市錦町1-4-4立川サニーハイツ303", enabled: true },
    { value: "190-0022", enabled: true },
    { value: "042-595-7557", enabled: true },
    { value: "042-595-7558", enabled: true },
    { value: "daichi@conphic.co.jp", enabled: true },
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
// Events
// =========================
function bindEvents() {
    document.getElementById("themeToggle")?.addEventListener("click", toggleTheme);

    document.getElementById("filePickBtn")?.addEventListener("click", () => {
        const input = document.getElementById("fileInput");
        if (input) input.value = ""; // 同じファイルでもchangeを出す
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

    tbody.addEventListener("input", saveRules);
    tbody.addEventListener("change", saveRules);

    tbody.addEventListener("click", (e) => {
        const btn = e.target?.closest?.("button[data-action='delete']");
        if (!btn) return;
        const tr = btn.closest("tr");
        tr?.remove();
        saveRules();
    });

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
// Drag & Drop sorting
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
        e.dataTransfer.setData("text/plain", "move");
    });

    tbody.addEventListener("dragover", (e) => {
        if (!draggingRow) return;
        e.preventDefault();

        const target = e.target?.closest?.("tr");
        if (!target || target === draggingRow) return;

        clearDragOverClasses(tbody);

        const rect = target.getBoundingClientRect();
        const isAfter = e.clientY - rect.top > rect.height / 2;

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
        saveRules();
    });

    function clearDragOverClasses(tbodyEl) {
        tbodyEl
            .querySelectorAll("tr.drag-over-top, tr.drag-over-bottom")
            .forEach((tr) => tr.classList.remove("drag-over-top", "drag-over-bottom"));
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

    const map = new Map();
    DEFAULT_MASK_RULES.forEach((r) => map.set(r.value, { value: r.value, enabled: !!r.enabled }));

    saved.forEach((r) => {
        if (!r || typeof r.value !== "string") return;
        map.set(r.value, { value: r.value, enabled: !!r.enabled });
    });

    const merged = Array.from(map.values());
    merged.forEach((rule) => addRuleRow(rule));
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
        .map((tr) => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            if (!enable.checked) return null;
            const v = (word.value ?? "").trim();
            if (!v) return null;
            return v;
        })
        .filter(Boolean);
}

function saveRules() {
    const tbody = document.querySelector("#maskTable tbody");
    if (!tbody) return;

    const rules = Array.from(tbody.querySelectorAll("tr"))
        .map((tr) => {
            const enable = tr.querySelector(".mask-enable");
            const word = tr.querySelector(".mask-word");
            if (!enable || !word) return null;
            return { enabled: !!enable.checked, value: word.value ?? "" };
        })
        .filter(Boolean);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

function safeJsonParse(text, fallback) {
    try {
        return JSON.parse(text);
    } catch {
        return fallback;
    }
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
    if (ext !== "docx" && ext !== "pptx" && ext !== "xlsx") {
        alert("対応形式は .docx / .pptx / .xlsx です");
        return;
    }

    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    try {
        if (ext === "docx") {
            await maskDocx(zip, words);
            const blob = await zip.generateAsync({ type: "blob" });
            download(blob, file.name.replace(/\.docx$/i, "_マスキング.docx"));
            return;
        }

        if (ext === "pptx") {
            const result = await maskPptx(zip, words);
            if (result.slideFiles === 0) {
                alert("ppt/slides/slide*.xml が見つかりません（PPTX形式でない可能性）");
                return;
            }
            const blob = await zip.generateAsync({ type: "blob" });
            download(blob, file.name.replace(/\.pptx$/i, "_マスキング.pptx"));
            return;
        }

        if (ext === "xlsx") {
            const result = await maskXlsx(zip, words);
            if (!result.sharedStringsFound && result.inlineCells === 0) {
                alert("対象（xl/sharedStrings.xml または inlineStr）が見つかりませんでした（形式差異の可能性）");
            }
            const blob = await zip.generateAsync({ type: "blob" });
            download(blob, file.name.replace(/\.xlsx$/i, "_マスキング.xlsx"));
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

    const joined = nodes.map((n) => n.text).join("");
    let masked = joined;

    const words = literalWords
        .map((w) => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    for (const word of words) {
        masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
    }

    let cursor = 0;
    const parts = nodes.map((n) => {
        const part = masked.slice(cursor, cursor + n.len);
        cursor += n.len;
        return part;
    });

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

function maskPptxSlideXmlByLiterals(xmlText, literalWords, debugName = "") {
    const doc = parseXml(xmlText, debugName);

    const paragraphs = getElementsByLocalName(doc, "p");
    if (paragraphs.length === 0) return { xml: xmlText, changed: false };

    const words = literalWords
        .map((w) => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    let changed = false;

    for (const p of paragraphs) {
        const ts = [];

        const walker = doc.createTreeWalker(
            p,
            NodeFilter.SHOW_ELEMENT,
            {
                acceptNode(node) {
                    return node.localName === "t" ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
                },
            }
        );

        while (walker.nextNode()) ts.push(walker.currentNode);
        if (ts.length === 0) continue;

        const texts = ts.map((n) => n.textContent ?? "");
        const lens = texts.map((s) => s.length);
        const joined = texts.join("");
        if (!joined) continue;

        let masked = joined;
        const beforeAll = masked;

        for (const word of words) {
            masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
        }
        if (masked === beforeAll) continue;

        let cursor = 0;
        for (let i = 0; i < ts.length; i++) {
            const part = masked.slice(cursor, cursor + lens[i]);
            cursor += lens[i];
            ts[i].textContent = part;
        }

        changed = true;
    }

    const out = new XMLSerializer().serializeToString(doc);
    return { xml: out, changed };
}

// =========================
// XLSX (sharedStrings + inlineStr) - 色は触らない版
// =========================
async function maskXlsx(zip, literalWords) {
    const words = literalWords
        .map((w) => (w ?? "").trim())
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);

    let sharedStringsFound = false;
    let sharedStringsSi = 0;
    let inlineCells = 0;

    // --- 1) xl/sharedStrings.xml ---
    const sstPath = "xl/sharedStrings.xml";
    const sstFile = zip.file(sstPath);
    if (sstFile) {
        sharedStringsFound = true;
        const xmlText = await sstFile.async("string");
        const doc = parseXml(xmlText, sstPath);

        // <si> 単位：配下の <t> を全部集めて連結→マスク→再分配
        const siNodes = getElementsByLocalName(doc, "si");
        for (const si of siNodes) {
            const tNodes = collectDescendantTNodes(doc, si);
            if (tNodes.length === 0) continue;

            const changed = maskAndRedistributeTextNodes(tNodes, words, true);
            if (changed) sharedStringsSi++;
        }

        zip.file(sstPath, new XMLSerializer().serializeToString(doc));
    }

    // --- 2) xl/worksheets/*.xml : inlineStr ---
    const worksheetPaths = [];
    zip.forEach((relativePath, file) => {
        if (!file.dir && relativePath.startsWith("xl/worksheets/") && relativePath.endsWith(".xml")) {
            worksheetPaths.push(relativePath);
        }
    });

    worksheetPaths.sort((a, b) => {
        const na = parseInt(a.match(/sheet(\d+)\.xml$/)?.[1] ?? "999999", 10);
        const nb = parseInt(b.match(/sheet(\d+)\.xml$/)?.[1] ?? "999999", 10);
        if (na !== nb) return na - nb;
        return a.localeCompare(b);
    });

    for (const path of worksheetPaths) {
        const f = zip.file(path);
        if (!f) continue;

        const xmlText = await f.async("string");
        const doc = parseXml(xmlText, path);

        const cNodes = getElementsByLocalName(doc, "c");
        let changedAny = false;

        for (const c of cNodes) {
            if (c.getAttribute("t") !== "inlineStr") continue;

            // <c t="inlineStr"><is>...</is>
            const isNodes = [];
            for (let i = 0; i < c.childNodes.length; i++) {
                const n = c.childNodes[i];
                if (n?.nodeType === 1 && n.localName === "is") isNodes.push(n);
            }

            for (const isNode of isNodes) {
                const tNodes = collectDescendantTNodes(doc, isNode);
                if (tNodes.length === 0) continue;

                const changed = maskAndRedistributeTextNodes(tNodes, words, true);
                if (changed) {
                    changedAny = true;
                    inlineCells++;
                }
            }
        }

        if (changedAny) {
            zip.file(path, new XMLSerializer().serializeToString(doc));
        }
    }

    return {
        sharedStringsFound,
        sharedStringsSi,
        inlineCells,
    };
}

// =========================
// XML helpers
// =========================
function parseXml(xmlText, fileName) {
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

function collectDescendantTNodes(doc, rootEl) {
    const ts = [];
    const walker = doc.createTreeWalker(
        rootEl,
        NodeFilter.SHOW_ELEMENT,
        {
            acceptNode(node) {
                return node.localName === "t" ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
            },
        }
    );
    while (walker.nextNode()) ts.push(walker.currentNode);
    return ts;
}

// 連結→固定文字列マスク→再分配（色や構造は触らない）
function maskAndRedistributeTextNodes(tNodes, words, preserveSpace) {
    const texts = tNodes.map((n) => n.textContent ?? "");
    const lens = texts.map((s) => s.length);
    const joined = texts.join("");
    if (!joined) return false;

    let masked = joined;
    const before = masked;

    for (const word of words) {
        masked = replaceAllLiteralKeepingLength(masked, word, MASK_CHAR);
    }

    if (masked === before) return false;

    let cursor = 0;
    for (let i = 0; i < tNodes.length; i++) {
        const part = masked.slice(cursor, cursor + lens[i]);
        cursor += lens[i];
        tNodes[i].textContent = part;

        if (preserveSpace) ensureXmlSpacePreserveIfNeeded(tNodes[i]);
    }

    // 念のため余りがあれば最後に足す（通常は起きない）
    if (cursor < masked.length && tNodes.length > 0) {
        tNodes[tNodes.length - 1].textContent += masked.slice(cursor);
        if (preserveSpace) ensureXmlSpacePreserveIfNeeded(tNodes[tNodes.length - 1]);
    }

    return true;
}

function ensureXmlSpacePreserveIfNeeded(tNode) {
    const text = tNode.textContent ?? "";
    if (/^\s|\s$/.test(text)) {
        tNode.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    }
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
