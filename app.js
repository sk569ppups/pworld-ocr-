// ===== ユーティリティ（ログ） =====
const $ = (q)=>document.querySelector(q);
const logBox = $("#log");
function log(msg, cls=""){ const d=document.createElement("div"); if(cls) d.className=cls; d.textContent=msg; logBox.appendChild(d); logBox.scrollTop=logBox.scrollHeight; }
function clearLog(){ logBox.textContent=""; }

// ===== 状態 =====
let masterOfficial = [];        // 正式名称（出力に使う）
let masterLooseSet = new Set(); // すべての別名を含むルーズキー集合
let extracted = [];             // 抽出結果（正式名称）
let lastTextSample = "";

// ===== 正規化 =====
function kanaToHira(s){ return s.replace(/[\u30a1-\u30f6]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0x60)); }
function normalizeJa(s){
  if(!s) return "";
  let t = s.normalize("NFKC")
    .replace(/[‐-‒–—―－ｰー]/g,"-")
    .replace(/[・･․‧•]/g,"")
    .replace(/[＿_]/g,"")
    .replace(/[（）\(\)\[\]【】｛｝\{\}]/g," ")
    .replace(/[。、，、,/]/g," ")
    .replace(/\s+/g," ")
    .trim();
  t = kanaToHira(t);
  return t.replace(/[\s-]/g,"").toLowerCase();
}
const makeLooseKey = (s)=>normalizeJa(s);

// ===== マスターCSV 読み込み =====
const csvInput = $("#csvFile");
csvInput.addEventListener("change", async (ev)=>{
  masterOfficial = []; masterLooseSet = new Set(); $("#btnExtract").disabled = true;

  const file = ev.target.files?.[0]; if(!file) return;

  clearLog(); log(`マスターCSV読込中: ${file.name} ...`);
  try {
    await new Promise((resolve,reject)=>{
      Papa.parse(file, {
        worker:false, skipEmptyLines:true,
        complete: (res)=>{
          try{
            applyMasterRows(res.data);
            log(`マスター読込完了: ${masterOfficial.length}件`, "ok");
            if($("#pdfFile").files?.[0]) $("#btnExtract").disabled = false;
            resolve();
          }catch(e){ reject(e); }
        },
        error: (err)=>reject(err)
      });
    });
  } catch(e){
    log(`マスターCSV読込エラー: ${e.message}`, "err");
  }
});

function applyMasterRows(rows){
  // 1列目: official、2列目以降: 別名
  const seenOfficial = new Set();
  for(const r of rows){
    const arr = Array.isArray(r)? r : [r];
    const official = String(arr[0] ?? "").trim();
    if(!official) continue;
    if(!seenOfficial.has(official)){ masterOfficial.push(official); seenOfficial.add(official); }
    for(const name of arr){
      const lk = makeLooseKey(String(name||"").trim());
      if(lk) masterLooseSet.add(lk);
    }
  }
}

// ===== PDF → テキスト抽出 =====
async function extractTextFromPDF(pdfFile){
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let texts = [];
  for(let i=1;i<=pdf.numPages;i++){
    const page = await pdf.getPage(i);
    const content = await page.getTextContent({ normalizeWhitespace:true });
    const txt = content.items.map(it => ("str" in it ? it.str : "")).join("\n");
    texts.push(txt);
  }
  return { text: texts.join("\n").trim(), pages: pdf.numPages };
}

// OCR フォールバック
async function ocrPdfWithTesseract(pdfFile, lang="jpn+eng"){
  const buf = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const scale = 2;
  let out = [];
  for(let i=1;i<=pdf.numPages;i++){
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = viewport.width; canvas.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;
    log(`OCR中: ${i}/${pdf.numPages} ...`);
    const { data:{ text } } = await Tesseract.recognize(canvas, lang);
    out.push(text);
  }
  return { text: out.join("\n").trim(), pages: pdf.numPages };
}

// ===== 照合 =====
function splitToLines(raw){
  return raw.replace(/\r/g,"").replace(/[、，。]/g,"\n")
           .split("\n").map(s=>s.trim()).filter(Boolean);
}
function matchLinesToMaster(lines){
  // ルーズキー -> official一覧
  const map = new Map();
  for(const off of masterOfficial){
    const lk = makeLooseKey(off);
    if(!map.has(lk)) map.set(lk, []);
    map.get(lk).push(off);
  }

  const hits = [];
  for(const raw of lines){
    const lk = makeLooseKey(raw);
    if(!lk) continue;

    if(masterLooseSet.has(lk)){
      // officialの候補（同キー複数想定）
      const offs = map.get(lk) || [];
      for(const o of offs) hits.push(o);
      continue;
    }
    // サブストリング準一致
    for(const [mk, offs] of map.entries()){
      if(lk.includes(mk) || mk.includes(lk)){
        for(const o of offs) hits.push(o);
      }
    }
  }
  return hits;
}

// ===== 出力（機種名1列だけ） =====
function downloadXlsx(names, filename="extracted.xlsx"){
  const uniqueSorted = Array.from(new Set(names)).sort((a,b)=>a.localeCompare(b,'ja'));
  const rows = uniqueSorted.map(nm => ({ 機種名: nm }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "抽出結果");
  XLSX.writeFile(wb, filename);
}

// ===== メインフロー =====
const pdfInput = $("#pdfFile");
const btnExtract = $("#btnExtract");
const btnDownload = $("#btnDownload");
const preview = $("#preview");

pdfInput.addEventListener("change", ()=>{
  if(pdfInput.files?.[0] && masterOfficial.length>0) btnExtract.disabled = false;
});

btnExtract.addEventListener("click", async ()=>{
  const pdfFile = pdfInput.files?.[0];
  if(!pdfFile){ alert("PDFを選択してください"); return; }
  if(masterOfficial.length===0){ alert("マスターCSVを読み込んでください"); return; }

  extracted = []; preview.textContent = ""; clearLog();
  log(`PDF解析開始: ${pdfFile.name}`);

  let textRes = { text:"", pages:0 };
  try{
    textRes = await extractTextFromPDF(pdfFile);
    log(`テキスト抽出: ${textRes.pages}ページ`, "ok");
  }catch(e){
    log(`pdf.jsテキスト抽出エラー: ${e.message}`, "err");
  }
  lastTextSample = textRes.text.slice(0,400);

  if(textRes.text.replace(/\s/g,"").length < 50){
    log("テキスト層が少ないためOCRにフォールバックします…","warn");
    try{
      textRes = await ocrPdfWithTesseract(pdfFile, "jpn+eng");
      log(`OCR完了: ${textRes.pages}ページ`, "ok");
    }catch(e){
      log(`OCRエラー: ${e.message}`, "err");
      return;
    }
  }

  const lines = splitToLines(textRes.text);
  log(`行分割: ${lines.length}行`);

  const hits = matchLinesToMaster(lines);
  log(`照合ヒット: ${hits.length}件`, hits.length? "ok":"warn");

  const lines = splitToLines(textRes.text);
log(`行分割: ${lines.length}行`);
function matchLinesToMaster_STRICT(lines){
  const looseToOfficial = new Map();
  for (const off of masterOfficial){
    const lk = makeLooseKey(off);
    if(!lk) continue;
    if(!looseToOfficial.has(lk)) looseToOfficial.set(lk, new Set());
    looseToOfficial.get(lk).add(off);
  }

  const hits = [];
  for (const raw of lines){
    const lk = makeLooseKey(raw);
    if(!lk) continue;
    if (masterLooseSet.has(lk)){
      const offs = looseToOfficial.get(lk) || new Set();
      for (const o of offs) hits.push(o);
    }
  }
  return Array.from(new Set(hits)).sort((a,b)=>a.localeCompare(b,'ja'));
}

// ★新しい関数で完全一致のみチェック
const hits = matchLinesToMaster_STRICT(lines);
log(`ヒット機種（完全一致）: ${hits.length}件`, hits.length ? "ok" : "warn");

extracted = hits;

// プレビュー出力
preview.innerHTML = hits.map((nm,i)=>`${i+1}. ${nm}`).join("<br>");
btnDownload.disabled = hits.length===0;

btnDownload.addEventListener("click", ()=>{
  if(extracted.length===0) return;
  downloadXlsx(extracted, "extracted.xlsx");
  log("Excelを書き出しました。", "ok");
});

