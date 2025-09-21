// ===== ユーティリティ（ログ） =====
const $ = (q) => document.querySelector(q);
const logBox = $("#log");
function log(msg, cls = "") {
  const d = document.createElement("div");
  if (cls) d.className = cls;
  d.textContent = msg;
  logBox.appendChild(d);
  logBox.scrollTop = logBox.scrollHeight;
}
function clearLog() {
  logBox.textContent = "";
}

// ===== 状態 =====
let masterOfficial = [];              // 正式名称（出力用）
let masterLooseSet = new Set();       // 正規化キー（official/別名すべて）
let aliasToOfficial = new Map();      // 正規化キー -> official名のSet
let extracted = [];                   // 抽出結果（正式名称のみ）
let lastTextSample = "";

// ===== 正規化 =====
function kanaToHira(s) {
  return s.replace(/[\u30a1-\u30f6]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0x60)
  );
}
function normalizeJa(s) {
  if (!s) return "";
  let t = s.normalize("NFKC")
    .replace(/[‐-‒–—―－ｰー]/g, "-")
    .replace(/[・･․‧•]/g, "")
    .replace(/[＿_]/g, "")
    .replace(/[（）\(\)\[\]【】｛｝\{\}]/g, " ")
    .replace(/[。、，、,/]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  t = kanaToHira(t);
  return t.replace(/[\s-]/g, "").toLowerCase();
}
const makeLooseKey = (s) => normalizeJa(s);

// ===== マスターCSV 読み込み =====
const csvInput = $("#csvFile");
csvInput.addEventListener("change", async (ev) => {
  masterOfficial = [];
  masterLooseSet = new Set();
  aliasToOfficial = new Map();
  $("#btnExtract").disabled = true;

  const file = ev.target.files?.[0];
  if (!file) return;

  clearLog();
  log(`マスターCSV読込中: ${file.name} ...`);
  try {
    await new Promise((resolve, reject) => {
      Papa.parse(file, {
        worker: false,
        skipEmptyLines: true,
        complete: (res) => {
          try {
            applyMasterRows(res.data);
            log(`マスター読込完了: ${masterOfficial.length}件`, "ok");
            if ($("#pdfFile").files?.[0]) $("#btnExtract").disabled = false;
            resolve();
          } catch (e) {
            reject(e);
          }
        },
        error: (err) => reject(err),
      });
    });
  } catch (e) {
    log(`マスターCSV読込エラー: ${e.message}`, "err");
  }
});

function applyMasterRows(rows) {
  const seenOfficial = new Set();

  for (const r of rows) {
    const arr = Array.isArray(r) ? r : [r];
    const official = String(arr[0] ?? "").trim();
    if (!official) continue;

    if (!seenOfficial.has(official)) {
      masterOfficial.push(official);
      seenOfficial.add(official);
    }

    // official自身＋別名すべてを同じグループとして登録
    for (const name of arr) {
      const lk = makeLooseKey(String(name || "").trim());
      if (!lk) continue;

      masterLooseSet.add(lk);
      if (!aliasToOfficial.has(lk)) aliasToOfficial.set(lk, new Set());
      aliasToOfficial.get(lk).add(official);
    }
  }
}

// ===== PDF → テキスト抽出（pdf.js） =====
async function extractTextFromPDF(pdfFile) {
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let texts = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent({ normalizeWhitespace: true });
    const txt = content.items.map(it => ("str" in it ? it.str : "")).join("\n");
    texts.push(txt);
  }
  return { text: texts.join("\n").trim(), pages: pdf.numPages };
}

// ===== OCR フォールバック（Tesseract） =====
async function ocrPdfWithTesseract(pdfFile, lang = "jpn+eng") {
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const scale = 2;
  let out = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;
    log(`OCR中: ${i}/${pdf.numPages} ...`);
    const { data: { text } } = await Tesseract.recognize(canvas, lang);
    out.push(text);
  }
  return { text: out.join("\n").trim(), pages: pdf.numPages };
}

// ===== テキストを行に分割 =====
function splitToLines(raw) {
  return raw
    .replace(/\r/g, "")
    .replace(/[、，。]/g, "\n")
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);
}

// ===== 照合：正規化後の「完全一致」のみ =====
function matchLinesToMaster_STRICT(lines) {
  const hits = new Set();

  for (const raw of lines) {
    const lk = makeLooseKey(raw);
    if (!lk) continue;
    if (aliasToOfficial.has(lk)) {
      for (const off of aliasToOfficial.get(lk)) hits.add(off);
    }
  }

  return Array.from(hits).sort((a, b) => a.localeCompare(b, "ja"));
}

// ===== Excel出力（機種名1列のみ） =====
function downloadXlsx(names, filename = "extracted.xlsx") {
  const rows = names.map((nm) => ({ 機種名: nm }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "抽出結果");
  XLSX.writeFile(wb, filename);
}

// ===== メインUI =====
const pdfInput = $("#pdfFile");
const btnExtract = $("#btnExtract");
const btnDownload = $("#btnDownload");
const preview = $("#preview");

pdfInput.addEventListener("change", () => {
  if (pdfInput.files?.[0] && masterOfficial.length > 0) btnExtract.disabled = false;
});

// 抽出実行
btnExtract.addEventListener("click", async () => {
  const pdfFile = pdfInput.files?.[0];
  if (!pdfFile) { alert("PDFを選択してください"); return; }
  if (masterOfficial.length === 0) { alert("マスターCSVを読み込んでください"); return; }

  extracted = [];
  preview.textContent = "";
  clearLog();
  log(`PDF解析開始: ${pdfFile.name}`);

  let textRes = { text: "", pages: 0 };
  try {
    textRes = await extractTextFromPDF(pdfFile);
    log(`テキスト抽出: ${textRes.pages}ページ`, "ok");
  } catch (e) {
    log(`pdf.jsテキスト抽出エラー: ${e.message}`, "err");
  }
  lastTextSample = textRes.text.slice(0, 400);

  // テキストが少なければOCRに自動切替
  if (textRes.text.replace(/\s/g, "").length < 50) {
    log("テキスト層が少ないためOCRにフォールバックします…", "warn");
    try {
      textRes = await ocrPdfWithTesseract(pdfFile, "jpn+eng");
      log(`OCR完了: ${textRes.pages}ページ`, "ok");
    } catch (e) {
      log(`OCRエラー: ${e.message}`, "err");
      return;
    }
  }

  const lines = splitToLines(textRes.text);
  log(`行分割: ${lines.length}行`);

  // ★完全一致のみで照合
  const hits = matchLinesToMaster_STRICT(lines);
  log(`ヒット機種（完全一致）: ${hits.length}件`, hits.length ? "ok" : "warn");

  extracted = hits;

  // プレビュー
  preview.innerHTML = hits.map((m, i) => `${i + 1}. ${m}`).join("<br>");
  btnDownload.disabled = hits.length === 0;
});

// Excel出力
btnDownload.addEventListener("click", () => {
  if (extracted.length === 0) return;
  downloadXlsx(extracted, "extracted.xlsx");
  log("Excelを書き出しました。", "ok");
});
