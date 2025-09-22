// ===== 文字正規化＆マッチング補助 =====
const NameNormalizer = (() => {
  // 全角→半角（数字・英字・記号）、カナ→ひら
  function toHalfWidth(str){
    return str.replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
              .replace(/\u3000/g, ' ');
  }
  function kanaToHira(s){
    return s.replace(/[\u30A1-\u30FA]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
  }
  // ダッシュ類統一
  function normalizeDashes(s){
    return s.replace(/[‐－ー–—]/g, '-');
  }
  // よく混入する括弧・記号・絵文字を除去 or スペースに
  function stripNoise(s){
    return s
      .replace(/[【】〔〕［］「」『』（）\(\)\[\]\{\}]/g, ' ')
      .replace(/[☆★●○◇◆♦︎■□▲△▼▽※♪♪♫♡♥︎❤︎◆▶︎▶▶️➡️⟁■・◆]/g, ' ')
      .replace(/[^\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}A-Za-z0-9\-\s]/gu, ' ');
  }
  function collapseSpaces(s){ return s.trim().replace(/\s+/g,' '); }

  function normalizeVisible(s){
    if(!s) return '';
    return collapseSpaces(stripNoise(normalizeDashes(kanaToHira(toHalfWidth(s)))));
  }

  // ルーズキー：小文字化・スペース&記号除去
  function looseKey(s){
    return normalizeVisible(s)
      .toLowerCase()
      .replace(/\s+/g,'')
      .replace(/[^a-z0-9\u3040-\u309f\u4e00-\u9faf-]/g,''); // hira/kanji/latin数字/-
  }

  // レーベンシュタイン距離（短め・高速）
  function levenshtein(a,b){
    if(a===b) return 0;
    const m=a.length, n=b.length;
    if(!m) return n; if(!n) return m;
    const dp = new Uint16Array(n+1);
    for(let j=0;j<=n;j++) dp[j]=j;
    for(let i=1;i<=m;i++){
      let prev=dp[0], tmp; dp[0]=i;
      for(let j=1;j<=n;j++){
        tmp=dp[j];
        const cost = a[i-1]===b[j-1] ? 0 : 1;
        dp[j] = Math.min(dp[j]+1, dp[j-1]+1, prev+cost);
        prev=tmp;
      }
    }
    return dp[n];
  }

  return { normalizeVisible, looseKey, levenshtein };
})();

// ===== マスターシート管理 =====
class MasterIndex {
  constructor(){
    this.officials = [];             // オブジェクト配列 {official, loose}
    this.mapLooseToOfficial = new Map(); // ルーズキー → 公式名
  }

  loadFromCsv(csvText){
    const rows = parseCsv(csvText);
    const head = rows[0]?.map(h=>h.trim().toLowerCase()) || [];
    const colOfficial = head.indexOf('official');
    const colAliases  = head.indexOf('aliases'); // 任意
    if(colOfficial===-1) throw new Error('CSVに official 列がありません。');

    for(let i=1;i<rows.length;i++){
      const row = rows[i];
      if(!row || row.length===0) continue;
      const official = (row[colOfficial]||'').trim();
      if(!official) continue;

      const allNames = [official];
      if(colAliases>-1 && row[colAliases]){
        // 区切り：| / , / ／ の混在を許容
        String(row[colAliases]).split(/[|／,]/).forEach(a=>{
          const t=(a||'').trim(); if(t) allNames.push(t);
        });
      }

      for(const name of allNames){
        const loose = NameNormalizer.looseKey(name);
        if(!loose) continue;
        // 既存があっても公式名は上書きしない（最初優先）
        if(!this.mapLooseToOfficial.has(loose)){
          this.mapLooseToOfficial.set(loose, official);
        }
      }
      this.officials.push({official, loose: NameNormalizer.looseKey(official)});
    }
  }

  // 厳格（完全一致）→ 別名一致 → あいまい（距離比）
  match(raw){
    const visible = NameNormalizer.normalizeVisible(raw);
    const loose   = NameNormalizer.looseKey(visible);
    if(!loose) return {type:'skip', official:'', loose, visible, raw};

    // 完全/別名一致（同じテーブル扱い）
    const hit = this.mapLooseToOfficial.get(loose);
    if(hit){ return {type:'exact', official:hit, loose, visible, raw}; }

    // あいまい一致：official の loose と距離を取る
    let best = {dist: 1e9, official:''};
    for(const o of this.officials){
      const d = NameNormalizer.levenshtein(loose, o.loose);
      if(d<best.dist){ best={dist:d, official:o.official}; }
      // 早期打ち切り（速度向上）
      if(best.dist===0) break;
    }
    const baseLen = Math.max(loose.length, 1);
    const ratio = best.dist / baseLen;

    // しきい値（経験則）：0.15（= 15% までの編集差は同一とみなす）
    if(ratio <= 0.15){
      return {type:'fuzzy', official:best.official, loose, visible, raw, score: (1-ratio)};
    }

    return {type:'unmatched', official:'', loose, visible, raw};
  }
}

// ===== CSV utilities =====
function parseCsv(text){
  // シンプルCSVパーサ（ダブルクォート対応）
  const rows=[]; let row=[], cur='', inQ=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(inQ){
      if(ch==='"'){
        if(text[i+1]==='"'){ cur+='"'; i++; }
        else inQ=false;
      }else cur+=ch;
    }else{
      if(ch==='"'){ inQ=true; }
      else if(ch===','){ row.push(cur); cur=''; }
      else if(ch==='\n'){
        row.push(cur); rows.push(row); row=[]; cur='';
      }else if(ch==='\r'){ /* skip */ }
      else cur+=ch;
    }
  }
  if(cur.length>0 || row.length>0){ row.push(cur); rows.push(row); }
  return rows;
}

function toCsv(rows){
  return rows.map(r=>r.map(c=>{
    const s = c==null ? '' : String(c);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
  }).join(',')).join('\n');
}
