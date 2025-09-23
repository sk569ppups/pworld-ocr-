/* global pdfjsLib, Tesseract, NameNormalizer, MasterIndex, toCsv */
(() => {

const els = {
  pdfFile:    document.getElementById('pdfFile'),
  masterFile: document.getElementById('masterFile'),
  runBtn:     document.getElementById('runBtn'),
  log:        document.getElementById('log'),
  progress:   document.getElementById('progress'),
  dlOfficial: document.getElementById('dlOfficial'),
  dlDetailed: document.getElementById('dlDetailed'),
  summary:    document.getElementById('summary'),
  canvas:     document.getElementById('workCanvas'),
};

let master = null;
let pdfArrayBuffer = null;

// ===== UI helpers =====
function log(msg, cls=''){ 
  const d=document.createElement('div'); 
  if(cls) d.className=cls; 
  d.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  els.log.appendChild(d); 
  els.log.scrollTop = els.log.scrollHeight;
}
function setProgress(msg){
  els.progress.textContent = msg;
}
function saveAs(filename, text){
  const blob = new Blob([text], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

}

// ===== pdf.js ローダ（グローバル未定義なら動的に読み込む） =====

  if (window.pdfjsLib) return;

  // 1) まず legacy UMD を動的ロード（グローバル pdfjsLib を期待）
  try {
    await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/legacy/build/pdf.min.js';
      s.defer = true;
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
    // worker も読み込み（失敗しても後段でセットするので続行）
    const s2 = document.createElement('script');
    s2.src = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/legacy/build/pdf.worker.min.js';
    s2.defer = true;
    document.head.appendChild(s2);

    if (window.pdfjsLib) return; // ここで取れればOK
  } catch (_) {
    // 何もしない（次のESMへ）
  }

  // 2) それでもダメなら ESM を import して window に載せる
  const mod = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/build/pdf.mjs');
  window.pdfjsLib = mod;

  // ESM 用 worker の場所を明示（mjs）
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/build/pdf.worker.min.mjs';
}

// ===== Inputs =====
els.masterFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f){ master=null; els.runBtn.disabled = true; return; }
  const txt = await f.text();
  try{
    const idx = new MasterIndex();
    idx.loadFromCsv(txt);
    master = idx;
    log(`マスター読込: ${f.name}（公式 ${master.officials.length} 件）`, 'ok');
  }catch(err){
    log(`マスター読込エラー: ${err.message||err}`, 'err');
    master=null;
  }
  els.runBtn.disabled = !(master && pdfArrayBuffer);
});

els.pdfFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f){ pdfArrayBuffer=null; els.runBtn.disabled = true; return; }
  pdfArrayBuffer = await f.arrayBuffer();
  log(`PDF読込: ${f.name}（${(pdfArrayBuffer.byteLength/1024/1024).toFixed(2)}MB）`, 'ok');
  els.runBtn.disabled = !(master && pdfArrayBuffer);
});

// ===== Core run =====
els.runBtn.addEventListener('click', async ()=>{
  if(!master || !pdfArrayBuffer){ return; }
  els.runBtn.disabled = true;
  els.dlOfficial.disabled = true;
  els.dlDetailed.disabled = true;
  els.summary.textContent = '';
  setProgress('処理開始…');
  log('PDF解析開始');

  try{
    const {lines, usedTextLayerPages, usedOcrPages} = await extractLinesFromPdf(pdfArrayBuffer);
    log(`抽出行数: ${lines.length}（textLayer:${usedTextLayerPages} / OCR:${usedOcrPages}）`, 'ok');

    const filtered = filterCandidateLines(lines);
    log(`候補行数（フィルタ後）: ${filtered.length}`);

    // ルーズ重複を削除して順序維持
    const uniqLoose = new Set();
    const results = [];
    for(const raw of filtered){
      const m = master.match(raw);
      if(m.type==='skip') continue;

      // 横並び同セルの「区切れ」対策：スペーサや「・」「/」「｜」などを再分割
      // ただし既に公式一致した場合はそのまま採用
      const needSplit = (m.type==='unmatched' || m.type==='fuzzy') && /[／\/\|｜・、,，]/.test(m.visible);
      if(needSplit){
        for(const frag of m.visible.split(/[／\/\|｜・、,，]/)){
          const t = frag.trim();
          if(!t) continue;
          const mm = master.match(t);
          if(mm.type!=='skip'){
            if(!uniqLoose.has(mm.loose)){
              uniqLoose.add(mm.loose);
              results.push(mm);
            }
          }
        }
      }else{
        if(!uniqLoose.has(m.loose)){
          uniqLoose.add(m.loose);
          results.push(m);
        }
      }
    }

    // 公式名のみ／詳細 の2種CSVを生成
    const officialRows = [['official']];
    const detailedRows = [['raw','visible','loose','type','official','score']];
    let exact=0, fuzzy=0, unmatched=0;

    for(const r of results){
      if(r.type==='exact') exact++;
      else if(r.type==='fuzzy') fuzzy++;
      else if(r.type==='unmatched') unmatched++;

      const officialName = r.official || r.visible; // unmatched は正規化名を暫定
      officialRows.push([officialName]);
      detailedRows.push([
        r.raw ?? '',
        r.visible ?? '',
        r.loose ?? '',
        r.type ?? '',
        r.official ?? '',
        r.score!=null ? String(r.score.toFixed(3)) : ''
      ]);
    }

    const base = (els.pdfFile.files?.[0]?.name || 'output').replace(/\.pdf$/i,'');
    const officialCsv = toCsv(officialRows);
    const detailedCsv = toCsv(detailedRows);

    // ダウンロードボタン有効化
    els.dlOfficial.onclick = ()=> saveAs(`${base}_machines_official.csv`, officialCsv);
    els.dlDetailed.onclick = ()=> saveAs(`${base}_machines_detailed.csv`, detailedCsv);
    els.dlOfficial.disabled = false;
    els.dlDetailed.disabled = false;

    log('CSVを準備しました（公式名のみ／詳細の2種）','ok');
    setProgress('完了');
    els.summary.textContent = `一致: exact=${exact}, fuzzy=${fuzzy}, unmatched=${unmatched} / 合計=${results.length}`;

  }catch(err){
    console.error(err);
    log(`エラー: ${err.message||err}`, 'err');
    setProgress('エラー');
  }finally{
    els.runBtn.disabled = !(master && pdfArrayBuffer);
  }
});

// ===== PDF → 行テキスト抽出 =====
// 1) textLayer抽出（成功すれば最速・高精度）
// 2) 失敗または文字数が少ないページのみ Tesseract OCR
async function extractLinesFromPdf(arrayBuffer){
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsLib.GlobalWorkerOptions.workerSrc || 'https://unpkg.com/pdfjs-dist@4.7.76/build/pdf.worker.min.js';

  const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
  const N = pdf.numPages;
  const lines = [];
  let usedTextLayerPages=0, usedOcrPages=0;

  // OCRエンジン準備（必要な時だけ）
  let ocrWorker = null;
  async function ensureOcr(){
    if(ocrWorker) return;
    ocrWorker = await Tesseract.createWorker('jpn', 1, { logger: m => {
      if(m.status==='recognizing text'){
        setProgress(`OCR中… ${Math.round(m.progress*100)}%`);
      }
    }});
  }

  for(let p=1;p<=N;p++){
    setProgress(`ページ ${p}/${N} 処理中`);
    const page = await pdf.getPage(p);

    // まず textLayer を試す
    try{
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(it=>it.str).join('\n');
      const visible = NameNormalizer.normalizeVisible(pageText);
      if(visible.replace(/\s/g,'').length >= 40){ // 十分な量の文字があれば採用
        usedTextLayerPages++;
        lines.push(...splitToLines(visible));
        continue;
      }
    }catch(_){ /* text layer だめでもフォールバックへ */ }

    // 画像OCRへ（フォールバック）
    await ensureOcr();
    const canvas = els.canvas;
    const viewport = page.getViewport({scale: 2.0}); // 解像度を上げて精度確保
    const ctx = canvas.getContext('2d');
    canvas.width = viewport.width|0;
    canvas.height = viewport.height|0;
    await page.render({canvasContext:ctx, viewport}).promise;
    const { data: { text } } = await ocrWorker.recognize(canvas);
    usedOcrPages++;
    const visible = NameNormalizer.normalizeVisible(text||'');
    lines.push(...splitToLines(visible));
  }

  if(ocrWorker){ await ocrWorker.terminate(); }

  return {lines, usedTextLayerPages, usedOcrPages};
}

function splitToLines(visibleText){
  // 行ベースに分解し、極端に短い行は除外
  return (visibleText||'').split(/\r?\n/)
    .map(s=>s.trim())
    .filter(s=>s.length>=2);
}

// ===== 機種候補の行をざっくり抽出（ノイズ行を除外） =====
const STOP_WORDS = [
  '新台','新装','おすすめ','導入','近日','増台','減台','推し','コーナー','本日','明日','休止',
  '営業時間','交換率','イベント','開店','閉店','店内','案内','抽選','整理券','並び','注意',
  'スロットコーナー','パチンココーナー','機種一覧','店舗','設置','台','予定','ご案内','ご確認',
  '会員','特典','貸玉','貯玉','景品','メダル','サービス','スタッフ','詳細','pdf','ご利用'
];

function filterCandidateLines(lines){
  const out = [];
  for(const raw of lines){
    const vis = NameNormalizer.normalizeVisible(raw);
    if(!vis) continue;

    // ひら/カナ/漢字/英数字の混成を含むこと（純ノイズ除外）
    if(!/[\u3041-\u309F\u30A1-\u30FA\u4E00-\u9FFFA-Za-z0-9]/.test(vis)) continue;

    // ストップ語含む行を除外
    if(STOP_WORDS.some(w=>vis.includes(w))) continue;

    // 1語だけ極端に短い等を除外
    if(vis.replace(/\s+/g,'').length <= 3) continue;

    out.push(raw);
  }
  return out;
}

})();





