// ==============================
// ユーティリティ（ログ）
// ==============================
const $ = (q) => document.querySelector(q);
const logBox = $("#log");
function log(msg, cls = "") {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = msg;
  logBox.appendChild(line);
  logBox.scrollTop = logBox.scrollHeight;
}
function clearLog(){ logBox.textContent = ""; }

// ==============================
// 状態
// ==============================
let masterNames = [];          // 正式名称の配列
let masterLooseSet = new Set(); // ルーズキーの集合
let extractedMachines = [];    // 照合後のヒット配列（正式名称）
let lastTextSampling = "";     // デバッグ用（生テキストのサンプル）

// ==============================
// 正規化＆ルーズキー（日本語向け）
// ==============================
// - 全角/半角統一
// - ダッシュ類/長音/中黒の統一
// - 記号/空白の除去
// - カタカナ→ひらがな（or ひら→カタ）どちらかに寄せる
// ここでは「カタカナをひらがな」に寄せる実装
function toZenkakuAscii(s){
  // 半角英数字→全角にしない方がマッチ強くなるケース多いので今回は未使用
  return s;
}
function kanaToHira(s){
  return s.replace(/[\u30a1-\u30f6]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}
function normalizeJa(s){
  if(!s) return "";
  let t = s;

  // Unicode正規化
  t = t.normalize("NFKC");

  // よくあるダッシュ/長音/中黒の揺れを統一
  t = t
    .replace(/[‐-‒–—―－ｰー]/g, "-")   // 全てハイフンに寄せる
    .replace(/[・･․‧•]/g, "")        // 中黒系は除去
    .replace(/[＿_]/g, "")            // アンダー系除去
    .replace(/[（）\(\)\[\]【】｛｝\{\}]/g, " ") // 括弧類はスペースに
    .replace(/[。、，、,/]/g, " ")    // 句読点やスラッシュはスペースに
    .replace(/\s+/g, " ")             // 連続空白圧縮
    .trim();

  // カタカナ→ひらがな
  t = kanaToHira(t);

  // スペース・ハイフンを除去してキー化
  t = t.replace(/[\s-]/g, "").toLowerCase();

  return t;
}

function makeLooseKey(raw){ return normalizeJa(raw); }

// ==============================
// マスターCSV読み込み
// ==============================
const csvInput = $("#csvFile");
csvInput.addEventListener("change", async (ev) => {
  masterNames = [];
  masterLooseSet = new Set();
  $("#btnExtract").disabled = true;

  const file = ev.target.files?.[0];
  if(!file){ return; }

  clearLog();
  log(`マスターCSV読込中: ${file.name} ...`);

  await new Promise((resolve, reject)=>{
    Papa.parse(file, {
      worker: false, // 安定優先
      skipEmptyLines: true,
      complete: (res) => {
        try{
          const rows = res.data;
          for(const r of rows){
            const name = String((Array.isArray(r)? r[0] : r) ?? "").trim();
            if(!name) continue;
            const loose = makeLooseKey(name);
            if(!loose || masterLooseSet.has(loose)) continue;
            masterLooseSet.add(loose);
            masterNames.push(name);
          }
          log(`マスター読込完了: ${masterNames.length}件`, "ok");
          if ($("#pdfFile").files?.[0]) $("#btnExtract").disabled = false;
          resolve();
        }catch(e){ reject(e); }
      },
      error: (err) => reject(err)
    });
  }).catch(err=>{
    log(`マスターCSV読込エラー: ${err.message}`, "err");
  });
});

// ==============================
// PDF → テキスト抽出（pdf.js）
// テキストが少なければ OCR フォールバック
// ==============================
async function extractTextFromPDF(pdfFile){
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  let texts = [];
  for(let i=1; i<=pdf.numPages; i++){
    const page = await pdf.getPage(i);
    const content = await page.getTextContent({ normalizeWhitespace: true });
    const pageText = content.items.map(it => ("str" in it ? it.str : "")).join("\n");
    texts.push(pageText);
  }
  const all = texts.join("\n").trim();
  return { text: all, pages: pdf.numPages };
}

// OCR（画像レンダ→Tesseract）
async function ocrPdfWithTesseract(pdfFile, lang="jpn+eng"){
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  const scale = 2; // 画質と速度のバランス
  let ocrText = [];

  for(let i=1; i<=pdf.numPages; i++){
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;

    log(`OCR実行中: ${i}/${pdf.numPages} ...`);

    const { data: { text } } = await Tesseract.recognize(canvas, lang, {
      // 進捗は必要あれば onProgress で拾える
    });
    ocrText.push(text);
  }
  return { text: ocrText.join("\n").trim(), pages: pdf.numPages };
}

// ==============================
// テキスト → 行分割 → マスター照合
// ==============================
function splitToLines(rawText){
  // 改行＋句読点などでほどよく分割
  const tmp = rawText.replace(/\r/g,"")
    .replace(/[、，。]/g, "\n");
  const lines = tmp.split("\n")
    .map(s => s.trim())
    .filter(Boolean);
  return lines;
}

function matchLinesToMaster(lines){
  const hits = [];
  // ルーズキーの逆引き辞書（= loose -> 正式名称群）
  const looseToName = new Map();
  for(const name of masterNames){
    const lk = makeLooseKey(name);
    if(!looseToName.has(lk)) looseToName.set(lk, []);
    looseToName.get(lk).push(name);
  }

  for(const raw of lines){
    const lk = makeLooseKey(raw);
    if(!lk) continue;

    // 完全一致（ルーズキー）
    if (masterLooseSet.has(lk)) {
      // 同じルーズキーに複数の正式名称が結びつく可能性も一応考慮
      const names = looseToName.get(lk) || [];
      for (const nm of names) hits.push(nm);
      continue;
    }

    // サブストリング準一致（PDFの分割や余計な語が混入した場合）
    // 例：raw の中に正式名称ルーズキーが含まれている
    for (const [mk, arr] of looseToName.entries()){
      if (lk.includes(mk) || mk.includes(lk)) {
        for(const nm of arr) hits.push(nm);
      }
    }
  }

  // 重複除外はここではしない（1台1行にする要件を満たすため）
  return hits;
}

// ==============================
// 出力（Excel）
// ==============================
function toToday() {
  const d = new Date();
  const m = (""+(d.getMonth()+1)).padStart(2,"0");
  const day = (""+d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${m}-${day}`;
}

function buildRows(machines, aggregate){
  const date = toToday();
  if(!aggregate){
    // 1台1行
    return machines.map(nm => ({ 日付: date, 機種名: nm, 台数: 1 }));
  }
  // 集計
  const map = new Map();
  for(const nm of machines){
    map.set(nm, (map.get(nm)||0)+1);
  }
  return Array.from(map.entries()).map(([nm, cnt]) => ({ 日付: date, 機種名: nm, 台数: cnt }));
}

function downloadXlsx(rows, filename="extracted.xlsx"){
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "抽出結果");
  XLSX.writeFile(wb, filename);
}

// ==============================
// メインフロー
// ==============================
const pdfInput = $("#pdfFile");
const btnExtract = $("#btnExtract");
const btnDownload = $("#btnDownload");
const preview = $("#preview");

pdfInput.addEventListener("change", ()=>{
  if (pdfInput.files?.[0] && masterNames.length > 0) {
    btnExtract.disabled = false;
  }
});

btnExtract.addEventListener("click", async ()=>{
  const pdfFile = pdfInput.files?.[0];
  if(!pdfFile){ alert("PDFファイルを選択してください"); return; }
  if(masterNames.length===0){ alert("マスターCSVを読み込んでください"); return; }

  extractedMachines = [];
  preview.textContent = "";
  clearLog();

  log(`PDF解析開始: ${pdfFile.name}`);

  // まずテキスト層から抽出
  let textRes;
  try{
    textRes = await extractTextFromPDF(pdfFile);
    log(`テキスト抽出: ${textRes.pages}ページ / ${Math.min(textRes.text.length, 10000)}文字`, "ok");
  }catch(e){
    log(`pdf.jsテキスト抽出エラー: ${e.message}`, "err");
    textRes = { text: "", pages: 0 };
  }

  lastTextSampling = textRes.text.slice(0, 400);

  // テキストがほぼ無いPDF（画像ベース）ならOCRへ
  const NEED_OCR = textRes.text.replace(/\s/g,"").length < 50;
  if (NEED_OCR){
    log("テキスト層が見つからないため、OCRにフォールバックします…", "warn");
    try{
      // 日本語＋英語（型番アルファベット混在想定）
      textRes = await ocrPdfWithTesseract(pdfFile, "jpn+eng");
      log(`OCR完了: ${textRes.pages}ページ / ${Math.min(textRes.text.length, 10000)}文字`, "ok");
    }catch(e){
      log(`OCRエラー: ${e.message}`, "err");
      return;
    }
  }

  // 分割＆照合
  const lines = splitToLines(textRes.text);
  log(`行分割: ${lines.length}行`);

  const hits = matchLinesToMaster(lines);
  log(`照合ヒット: ${hits.length}件`, hits.length ? "ok" : "warn");

  extractedMachines = hits;

  // プレビュー（最大200件）
  const cap = 200;
  const show = hits.slice(0, cap);
  preview.innerHTML = show.map((nm, i)=>`${i+1}. ${nm}`).join("<br>");
  if(hits.length > cap){
    const rest = hits.length - cap;
    preview.innerHTML += `<br>…ほか ${rest} 件`;
  }

  btnDownload.disabled = extractedMachines.length === 0;
});

btnDownload.addEventListener("click", ()=>{
  if(extractedMachines.length===0){ return; }
  const aggregate = $("#chkAgg").checked;
  const rows = buildRows(extractedMachines, aggregate);
  downloadXlsx(rows, aggregate ? "extracted_agg.xlsx" : "extracted_raw.xlsx");
  log("Excelを書き出しました。", "ok");
});

// ==============================
// デバッグ用: グローバルに一部公開（任意）
// ==============================
// window.__DEBUG = { normalizeJa, makeLooseKey, lastText: ()=>lastTextSampling };
