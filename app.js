/*
 * app.js（NGフィルタ＋縦1列CSVメイン・緩和マッチ版）
 * 2025-09-22
 *
 * 変更点（前版からの追加）
 *  - NGワード/フレーズを除外（「新台」「おすすめ」など運用系語句）
 *  - 出力のメインを「公式名だけの縦1列CSV（*_machines_official.csv）」に変更
 *  - 詳細CSV（match種別/ソース付き）はサブ出力（*_machines_detailed.csv）
 *  - CSVはUTF-8 BOM付きで文字化け防止
 *
 * 依存: pdf.js / （必要時）tesseract.js
 */

(() => {
  'use strict';

  // ===== 設定 =====
  const CONFIG = {
    FUZZY_BASE_MIN: 2,     // レーベンシュタイン最小許容距離
    FUZZY_RATE: 0.20,      // 文字長に対する許容割合（20%）
    PARTIAL_MIN: 4,        // 部分一致を許容する最小キー長
    SOURCE_CUT: 120,       // 詳細CSVのsource_textを切り詰める長さ
    NG_WORDS: [
      // 運用系・訴求語句
      '新台','おすすめ','オススメ','人気','注目','最新','機種入替','導入','増台','減台','設置','台数','配置','コーナー','ラインナップ','ラインアップ','入替','入替え','リニューアル',
      // 共通UI／見出し系
      '一覧','ご案内','お知らせ','営業時間','抽選','整理券','SNS','X','Twitter','LINE','YouTube','休業','開店','閉店',
      // 雑多
      'バラエティ','スロットコーナー','パチンココーナー','甘デジ','ライトミドル','ミドル','高確率','低確率','右打ち','へそ','電サポ','遊タイム',
    ],
  };

  const $ = (q) => document.querySelector(q);
  const logBox = () => $('#log');
  function log(msg, cls = ''){
    const box = logBox(); if(!box) return;
    const d = document.createElement('div'); if(cls) d.className = cls;
    d.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
    box.appendChild(d); box.scrollTop = box.scrollHeight;
  }
  function setProgress(label, value, max=100){ const p=$('#progress'); if(!p) return; if('value' in p){ p.max=max; p.value=value; p.setAttribute('data-label',label);} else { p.textContent=`${label} ${Math.round((value/max)*100)}%`; } }
  function saveAs(filename, text, addBOM=true){ const BOM = addBOM ? '\uFEFF' : ''; const blob = new Blob([BOM, text], { type: 'text/csv;charset=utf-8' }); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }

  // ===== 正規化 =====
  const NameNormalizer = {
    toNFKC:s=>s.normalize('NFKC'),
    kataToHira(s){ return s.replace(/[\u30A1-\u30F6]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0x60)); },
    toHalfWidthAscii(s){ return s.replace(/[！-～]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); },
    unifyDashes(s){ return s.replace(/[‐‑‒–—―ー－]/g,'-'); },
    stripSymbols(s){ return s.replace(/[\u200B\u200C\u200D\uFEFF]/g,'').replace(/[\t\r\n]/g,' ').replace(/[・~〜~◆☆★♪☓×✕✖︎△○●◇◆◎◯\+\*＝=≠≒≈≡≪≫＜＞<>\(\)\[\]{}\|\\/,:;!?！？。、「」『』【】：]/g,' ').replace(/\s{2,}/g,' '); },
    trimSpaces:s=>s.trim(),
    toLower:s=>s.toLowerCase(),
    makeLooseKey(raw){ if(!raw) return ''; let s=this.toNFKC(String(raw)); s=this.toHalfWidthAscii(s); s=this.kataToHira(s); s=this.unifyDashes(s); s=this.stripSymbols(s); s=this.toLower(s); s=s.replace(/\s+/g,''); return s; },
    normalizeDisplay(raw){ let s=this.toNFKC(String(raw)); s=this.unifyDashes(s); s=this.stripSymbols(s); return this.trimSpaces(s); }
  };

  // ===== Levenshtein =====
  function levenshtein(a,b){ if(a===b) return 0; const m=a.length,n=b.length; if(m===0) return n; if(n===0) return m; const dp=new Array(n+1); for(let j=0;j<=n;j++) dp[j]=j; for(let i=1;i<=m;i++){ let prev=dp[0]; dp[0]=i; for(let j=1;j<=n;j++){ const temp=dp[j]; const cost=a[i-1]===b[j-1]?0:1; dp[j]=Math.min(dp[j]+1, dp[j-1]+1, prev+cost); prev=temp; } } return dp[n]; }

  // ===== マスター =====
  const Master={ entries:[], looseSet:new Set(), byLoose:new Map(), clear(){ this.entries=[]; this.looseSet=new Set(); this.byLoose=new Map(); }, add(official){ const norm=NameNormalizer.normalizeDisplay(official); const loose=NameNormalizer.makeLooseKey(norm); if(!loose||this.looseSet.has(loose)) return; this.looseSet.add(loose); this.entries.push({official:norm, loose}); this.byLoose.set(loose,norm); } };
  async function loadMasterCSV(file){ Master.clear(); const text=await file.text(); const lines=text.split(/\r?\n/).map(s=>s.replace(/^\ufeff/,'')); for(let i=0;i<lines.length;i++){ const line=lines[i]; if(!line) continue; const [name]=line.split(','); if(!name) continue; if(i===0 && /機種|正式|name/i.test(name)) continue; Master.add(name); } log(`マスター読込: ${Master.entries.length}件`); }

  // ===== PDF抽出 =====
  async function extractTextFromPDF(file, ocrFallback=true){ if(!('pdfjsLib' in window)) throw new Error('pdf.js が読み込まれていません。'); const pdfjs=window.pdfjsLib; const arrayBuffer=await file.arrayBuffer(); const pdf=await pdfjs.getDocument({data:arrayBuffer}).promise; const allLines=[]; for(let p=1;p<=pdf.numPages;p++){ setProgress(`PDFテキスト抽出中 (page ${p}/${pdf.numPages})`,p,pdf.numPages); const page=await pdf.getPage(p); const content=await page.getTextContent(); const text=content.items.map(it=> (it.str||'')).join('\n'); const pageText=NameNormalizer.normalizeDisplay(text); const lines=pageText.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); if(ocrFallback && lines.join(' ').length<20 && 'Tesseract' in window){ log(`page ${p}: テキスト弱→OCR`); const ocrLines=await ocrPageToLines(page); allLines.push(...ocrLines); } else { allLines.push(...lines); } } return allLines; }
  async function ocrPageToLines(page){ const viewport=page.getViewport({scale:2.0}); const canvas=document.createElement('canvas'); const ctx=canvas.getContext('2d'); canvas.width=viewport.width; canvas.height=viewport.height; const renderTask=page.render({canvasContext:ctx,viewport}); await renderTask.promise; const {data:{text}}=await window.Tesseract.recognize(canvas,'jpn+eng',{logger:(m)=>{ if(m.status&&m.progress!=null){ setProgress(`OCR ${m.status}`, Math.round(m.progress*100),100); } }}); const normalized=NameNormalizer.normalizeDisplay(text); return normalized.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); }

// ===== ノイズ判定（NGワード） =====
const NG = (()=>{
  // ★ 長さ3以上だけをNGに採用（単文字・2文字は誤爆が多い）
  const ngLoose = new Set(
    CONFIG.NG_WORDS
      .filter(w => (w || '').length >= 3)   // ここを追加
      .map(w => NameNormalizer.makeLooseKey(w))
      .filter(Boolean)
  );

  function isNoiseLine(s){
    if(!s) return true;
    const disp = NameNormalizer.normalizeDisplay(s);
    if(!disp) return true;
    const loose = NameNormalizer.makeLooseKey(disp);
    if(!loose) return true;

    // 数字や記号だけ
    if(/^[0-9\-_. ]+$/.test(loose)) return true;

    // NGワードを含む（※3文字以上のみ）
    for(const ng of ngLoose){
      if(loose.includes(ng)) return true;
    }
    return false;
  }
  return { isNoiseLine };
})();


  // ===== 照合・抽出 =====
  const Extractor={
    extract(lines){
      const uniq = new Set();
      const out = [];
      for(const raw of lines){
        if(NG.isNoiseLine(raw)) continue; // NGフィルタ
        const disp = NameNormalizer.normalizeDisplay(raw);
        if(!disp || disp.length < 3) continue;
        const loose = NameNormalizer.makeLooseKey(disp);
        if(!loose || uniq.has(loose)) continue;

        // 完全一致
        if(Master.looseSet.has(loose)){
          uniq.add(loose);
          out.push({ match:'exact', official: Master.byLoose.get(loose), source: disp });
          continue;
        }
        // 部分一致（緩め）
        const partial = this.partialMatch(loose);
        if(partial){
          uniq.add(loose);
          out.push({ match:'partial', official: partial.official, source: disp });
          continue;
        }
        // あいまい一致（可変閾値）
        const fuzzy = this.fuzzyMatch(loose);
        if(fuzzy){
          uniq.add(loose);
          out.push({ match:'fuzzy', official: fuzzy.official, source: disp, dist: fuzzy.dist });
        }
      }
      // 公式名でユニーク化
      const seen = new Set();
      const deduped = [];
      for(const r of out){ if(seen.has(r.official)) continue; seen.add(r.official); deduped.push(r); }
      return deduped;
    },
    getFuzzyMax(a,b){ const L=Math.max(a.length,b.length); return Math.max(CONFIG.FUZZY_BASE_MIN, Math.ceil(L*CONFIG.FUZZY_RATE)); },
    partialMatch(loose){ for(const e of Master.entries){ if(e.loose.length < CONFIG.PARTIAL_MIN) continue; if(loose.includes(e.loose) || e.loose.includes(loose)) return { official: e.official }; } return null; },
    fuzzyMatch(loose){ let best=null, bestDist=Infinity; for(const e of Master.entries){ const d=levenshtein(loose,e.loose); const maxD=this.getFuzzyMax(loose,e.loose); if(d<=maxD && d<bestDist){ bestDist=d; best=e; } } if(best) return { official: best.official, dist: bestDist }; return null; }
  };

  // ===== CSV出力 =====
  function toDetailedCsv(records){
    const header=['official_name','match_type','source_text'];
    const lines=[header.join(',')];
    let exact=0,partial=0,fuzzy=0;
    const cut=(s,max=CONFIG.SOURCE_CUT)=>{ const str=String(s||''); return str.length>max? str.slice(0,max)+'…' : str; };
    for(const r of records){ if(r.match==='exact') exact++; else if(r.match==='partial') partial++; else if(r.match==='fuzzy') fuzzy++; lines.push([csvEscape(r.official), r.match, csvEscape(cut(r.source))].join(',')); }
    log(`内訳: 完全一致 ${exact} / 部分一致 ${partial} / あいまい ${fuzzy}`);
    return lines.join('\r\n');
  }
  function toOfficialOnlyCsv(records){
    const uniques=[...new Set(records.map(r=>r.official))].sort((a,b)=>a.localeCompare(b,'ja'));
    return ['official_name'].concat(uniques.map(csvEscape)).join('\r\n');
  }
  function csvEscape(s){ const needs=/[",\n\r]/.test(String(s)); let out=String(s).replace(/"/g,'""'); return needs?`"${out}"`:out; }

  // ===== メイン =====
  async function startPipeline(){
    try{
      const pdfFile=$('#pdfInput')?.files?.[0];
      const masterFile=$('#masterInput')?.files?.[0];
      if(!masterFile) throw new Error('マスターCSVを選択してください。');
      if(!pdfFile) throw new Error('PDFファイルを選択してください。');
      log('マスターCSV読み込み開始'); await loadMasterCSV(masterFile);
      log('PDF抽出開始'); const lines=await extractTextFromPDF(pdfFile,true); log(`PDF生テキスト行数: ${lines.length}`);
      log('マスター照合・抽出'); window.__lastLines=lines.slice(); const records=Extractor.extract(lines); log(`抽出結果: ${records.length} 機種`);

      const base=(pdfFile.name||'result').replace(/\.pdf$/i,'');
      // メイン: 公式名だけの縦1列CSV
      saveAs(`${base}_machines_official.csv`, toOfficialOnlyCsv(records));
      // サブ: 詳細CSV
      saveAs(`${base}_machines_detailed.csv`, toDetailedCsv(records));
      log('CSVをダウンロードしました（公式名のみ／詳細の2種）');
    }catch(err){
      console.error(err); log(`エラー: ${err.message||err}`,'error'); alert(err.message||String(err));
    } finally { setProgress('完了',100,100); }
  }

  // ===== 補助 =====
  function getUnmatched(allLines){ const results=new Set(Extractor.extract(allLines).map(r=>r.source)); const out=[]; for(const raw of allLines){ const disp=NameNormalizer.normalizeDisplay(raw); if(!disp) continue; if(!results.has(disp) && !NG.isNoiseLine(disp)) out.push(disp); } return out; }

  function bindEvents(){
    const startBtn=$('#startBtn'); if(startBtn) startBtn.addEventListener('click', startPipeline);
    const downloadBtn=$('#downloadCsvBtn'); if(downloadBtn) downloadBtn.addEventListener('click',()=>{ const box=logBox(); if(!box) return; const lines=Array.from(box.children).map(el=>el.textContent||''); saveAs(`log_${Date.now()}.csv`, lines.join('\n')); });
    const dumpBtn=$('#dumpUnmatchedBtn'); if(dumpBtn) dumpBtn.addEventListener('click',()=>{ if(!window.__lastLines) return alert('抽出後に実行してください'); const unmatched=getUnmatched(window.__lastLines); const csv=['source_text'].concat(unmatched.map(csvEscape)).join('\r\n'); saveAs(`unmatched_${Date.now()}.csv`, csv); });
  }

  function checkDeps(){ if(!('pdfjsLib' in window)){ log('pdf.js が未ロードです（致命的）','error'); } else { log('pdf.js OK'); } if('Tesseract' in window){ log('Tesseract.js OK (OCR利用可)'); } else { log('Tesseract.js 未ロード（必要時のみ導入してください）'); } }

  window.addEventListener('DOMContentLoaded',()=>{ bindEvents(); checkDeps(); log('app.js 初期化完了（NGフィルタ＋縦1列CSVメイン）'); });
})();

