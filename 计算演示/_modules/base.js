/* ================================================================
 * ① 进制转换（仅保留 6 种转换：二 ↔ 十 · 二 ↔ 八 · 二 ↔ 十六）
 * 子页面：bin2dec 二转十（按权展开）/ dec2bin 十转二（除基取余·逆序）
 *         bin2oct 二转八（三位分组）/ oct2bin 八转二（逐位展开）
 *         bin2hex 二转十六（四位分组）/ hex2bin 十六转二（逐位展开）
 * 复用全局辅助：B, DIGITS, toBase, fromBase, randDigits, randInt, $, $$, esc
 * ================================================================ */

/* ---------- 本模块大屏幕样式覆盖（作用域限定在 .base-stage，不影响其他模块） ---------- */
(function injectBaseCss(){
  if(document.getElementById('base-module-css')) return;
  const css = `
  .base-stage{font-family:var(--font)}
  .base-stage .stage-hd{font-size:clamp(12px,1.6vw,15px);margin-bottom:clamp(10px,1.8vw,16px)}
  .base-stage .step-title{font-size:clamp(13px,1.8vw,16px);padding:5px 14px;margin-bottom:clamp(12px,2.2vw,18px)}
  /* 竖式除法：计算过程字号放大、行间距加大 */
  .base-stage .div-vertical{font-family:var(--mono);font-size:clamp(16px,2.2vw,22px);line-height:2.1;padding:6px 10px}
  .base-stage .div-vertical .dv-row{margin:clamp(8px,1.6vw,12px) 0;white-space:normal}
  .base-stage .div-vertical .dv-row.empty{opacity:.32;color:#94A3B8}
  .base-stage .div-vertical .dv-row.done{opacity:.85}
  .base-stage .div-vertical .dv-row.done .dv-num,.base-stage .div-vertical .dv-row.done .dv-q{color:#15803D}
  .base-stage .div-vertical .dv-row.done .dv-rem{color:#16A34A}
  .base-stage .div-vertical .dv-num,.base-stage .div-vertical .dv-q{font-size:clamp(18px,2.8vw,28px);min-width:2.1em}
  .base-stage .div-vertical .dv-rem{font-size:clamp(18px,2.8vw,28px)}
  .base-stage .div-vertical .dv-rem::before{font-size:.6em}
  .base-stage .dv-rank{margin-left:1em;font-size:clamp(12px,1.6vw,15px);font-weight:700;color:#15803D;background:#E8FBF0;border:1px solid #BFE9CF;border-radius:8px;padding:2px 10px;white-space:nowrap}
  .base-stage .dv-rank.cur{color:#B45309;background:#FFF4DB;border-color:#F59E0B}
  /* 按权展开表：单元格加大内边距 */
  .base-stage .exp-table{font-size:clamp(12px,1.6vw,16px);margin:clamp(10px,1.8vw,14px) 0}
  .base-stage .exp-table th,.base-stage .exp-table td{padding:clamp(11px,1.8vw,15px) clamp(12px,1.8vw,18px);font-size:clamp(13px,1.8vw,17px)}
  .base-stage .exp-inline{font-size:clamp(16px,2.2vw,22px);padding:clamp(8px,1.4vw,12px) clamp(12px,1.8vw,16px);margin:clamp(12px,2.2vw,18px) 0;line-height:1.8}
  .base-stage .exp-inline .res{color:var(--accent);font-weight:700;margin-left:4px}
  /* 分组框 */
  .base-stage .group-box{font-size:clamp(18px,2.8vw,28px);gap:clamp(8px,1.4vw,12px);margin:clamp(10px,1.8vw,14px) 0;flex-wrap:wrap;justify-content:center}
  .base-stage .group-box .grp{padding:clamp(7px,1.3vw,11px) clamp(13px,2vw,18px);font-size:clamp(18px,2.8vw,28px)}
  .base-stage .group-box .grp .grp-label{top:-20px;font-size:.6em}
  /* 结果框 */
  .base-stage .write-result{font-size:clamp(18px,2.8vw,28px);padding:clamp(11px,1.8vw,16px) clamp(14px,2.2vw,22px);margin-top:clamp(14px,2.6vw,20px);letter-spacing:1px}
  .base-stage .annot{font-size:clamp(13px,1.8vw,17px);line-height:1.7;padding:clamp(8px,1.4vw,12px) clamp(12px,1.8vw,16px)}
  /* 逆序读取箭头条 */
  .base-stage .rev-arrow{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;font-family:var(--mono);font-size:clamp(12px,1.6vw,15px);padding:0 clamp(6px,1.2vw,10px)}
  .base-stage .rev-arrow .ra-top{color:#94A3B8;font-weight:700}
  .base-stage .rev-arrow .ra-mid{color:#22C55E;font-size:clamp(22px,3.2vw,30px);font-weight:700;line-height:1.1}
  .base-stage .rev-arrow .ra-bottom{color:#22C55E;font-weight:700}
  .base-stage .two-col{align-items:stretch}
  .base-stage .two-col>div{min-width:min(100%,220px)}
  `;
  document.head.insertAdjacentHTML('beforeend', `<style id="base-module-css">${css}</style>`);
})();

/* ---------- 输入约束配置（key 对应子页面 id） ---------- */
const BASE_INPUT_SPEC = {
  bin2dec: {kind:'bin',  base:2,  label:'二进制',   chars:'0、1',      max:8,  ph:'如 1011',    re:/^[01]+$/},
  dec2bin: {kind:'dec',  base:10, label:'十进制',   chars:'0~9',      max:3,  ph:'如 25',      re:/^\d+$/},
  bin2oct: {kind:'bin',  base:2,  label:'二进制',   chars:'0、1',      max:12, ph:'如 10110',   re:/^[01]+$/},
  oct2bin: {kind:'oct',  base:8,  label:'八进制',   chars:'0~7',      max:4,  ph:'如 37',      re:/^[0-7]+$/i},
  bin2hex: {kind:'bin',  base:2,  label:'二进制',   chars:'0、1',      max:12, ph:'如 101101',  re:/^[01]+$/},
  hex2bin: {kind:'hex',  base:16,label:'十六进制', chars:'0~9、A~F',max:4,  ph:'如 A3',      re:/^[0-9A-Fa-f]+$/}
};
const BASE_RAND_OPTS = {
  bin2dec: [['2','2 位'],['3','3 位'],['4','4 位']],
  dec2bin: [['small','较小数（5~31）'],['medium','中等数（32~99）']],
  bin2oct: [['6','6 位'],['9','9 位']],
  oct2bin: [['2','2 位'],['3','3 位']],
  bin2hex: [['8','8 位'],['12','12 位']],
  hex2bin: [['2','2 位'],['3','3 位']]
};

MOD.base = {
  name:'进制转换', sub:'二 ↔ 十 · 二 ↔ 八 · 二 ↔ 十六 · 手写逐步演示',
  badge:'数制 · 互转', tag:'进制', cat:'bit',
  color:'#4F6BFF', soft:'#EEF3FF', type:'按权展开 · 除基取余逆序 · 三位/四位分组 · 逐位展开',
  desc:'以二进制为桥梁，演示二转十（按权展开）、十转二（除基取余逆序读取）、二转八/十六（分组法）及其反向逐位展开，每一步都在大屏上手写推导。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 17l6-6-6-6M12 19h8"/></svg>',
  setDesc:'',
  stageDesc:'橙色=当前计算项，绿色=已得出结果，灰色=尚未操作；十转二最后展示余数逆序（倒排）读取。',
  initDesc:'点击「自动演示」或「下一步」，观看进制转换的手写推导过程。',
  currentSub:'bin2dec',
  subTabs:[
    {id:'bin2dec', label:'二转十',   method:'按权展开'},
    {id:'dec2bin', label:'十转二',   method:'除基取余·逆序'},
    {id:'bin2oct', label:'二转八',   method:'三位分组'},
    {id:'oct2bin', label:'八转二',   method:'逐位展开'},
    {id:'bin2hex', label:'二转十六', method:'四位分组'},
    {id:'hex2bin', label:'十六转二', method:'逐位展开'}
  ],
  _customInput:'', _invalidMsg:'', _digits:'', _dec:0,

  theory(){
    return `<div class="rule"><b>数码</b>：进制中使用的符号。二进制用 0、1；八进制用 0~7；十六进制用 0~9、A~F；十进制用 0~9。</div>
    <div class="rule"><b>基数</b>：进制中数码的个数。二进制基数=2，八进制=8，十六进制=16，十进制=10。</div>
    <div class="rule"><b>位权</b>：每一位的权重 = 基数<sup>位序</sup>。位序从右往左由 0 开始，如二进制 1011 中，最左一位的位序是 3、位权是 2³=8。</div>
    <div class="rule"><b>按权展开</b>（二进制 → 十进制）：每一位数码 × 对应位权得到贡献，再把所有贡献相加。</div>
    <div class="rule"><b>除基取余</b>（十进制 → 二进制）：不断除以 2，记录商和余数，直到商为 0；余数要<b>从下往上逆序（倒排）</b>读取才是结果。</div>
    <div class="rule"><b>分组转换</b>：二 ↔ 八 按 <b>3 位一组</b>，二 ↔ 十六 按 <b>4 位一组</b>；反向转换时把每一位展开成对应位数的二进制即可。</div>`;
  },

  /* ---------- 参数设置区 ---------- */
  settingsHTML(){
    const spec = BASE_INPUT_SPEC[this.currentSub];
    const cur = this._customInput || '';
    let bad = '', hint = `仅允许 ${spec.chars}`;
    if(cur){
      if(!spec.re.test(cur)){ bad=' bad'; hint=`「${cur}」含非法数码，仅允许 ${spec.chars}`; }
      else if(cur.length > spec.max){ bad=' bad'; hint=`位数过长：最多 ${spec.max} 位，当前 ${cur.length} 位`; }
    }
    const opts = BASE_RAND_OPTS[this.currentSub].map(([v,t],i)=>`<option value="${v}" ${i===0?'selected':''}>${t}</option>`).join('');
    return `<div class="inline-setting"><div class="setting">
        <label>输入数值（留空则随机出题）</label>
        <input type="text" id="b-custom" value="${esc(cur)}" placeholder="${spec.ph}" maxlength="${spec.max}" class="${bad}" style="width:clamp(100px,15vw,140px);font-family:var(--mono)">
        <div class="chk ${bad?'bad':''}">${hint}</div>
      </div></div>
      <div class="inline-setting"><div class="setting"><label>或随机</label>
        <select id="b-len">${opts}</select></div></div>`;
  },
  afterSettings(){
    const el = $('#b-custom');
    if(el) el.onchange = ()=>{ this._customInput = el.value.trim(); Engine.reload(); };
    const sel = $('#b-len');
    if(sel) sel.onchange = ()=>{ this._customInput = ''; Engine.reload(); };
  },

  /* ---------- 内部辅助 ---------- */
  _spec(){ return BASE_INPUT_SPEC[this.currentSub]; },
  /* 读取并校验手动输入；非法时抛错进入引擎错误态；返回 true=手动输入，false=随机 */
  _readCustom(){
    const spec = this._spec();
    const el = $('#b-custom');
    const raw = (el ? el.value.trim() : (this._customInput || ''));
    if(raw){
      if(!spec.re.test(raw)){
        this._invalidMsg = `「${raw}」不是合法${spec.label}数，仅允许数码 ${spec.chars}`;
        this._customInput = raw; throw new Error(this._invalidMsg);
      }
      if(raw.length > spec.max){
        this._invalidMsg = `位数过长：${spec.label}数最多 ${spec.max} 位（当前 ${raw.length} 位）`;
        this._customInput = raw; throw new Error(this._invalidMsg);
      }
      if(spec.kind === 'dec'){
        const v = Number(raw);
        if(!/^\d+$/.test(raw) || v < 1){
          this._invalidMsg = '请输入正整数十进制数（1 ~ 255）';
          this._customInput = raw; throw new Error(this._invalidMsg);
        }
        if(v > 255){
          this._invalidMsg = '为保证竖式除法在大屏幕上清晰展示，请输入 1 ~ 255 之间的十进制数';
          this._customInput = raw; throw new Error(this._invalidMsg);
        }
        this._dec = v;
      }else{
        this._digits = raw.toUpperCase();
      }
      this._customInput = raw;
      this._invalidMsg = '';
      return true;
    }
    this._invalidMsg = '';
    return false;
  },
  /* 随机出题（无手动输入时调用） */
  _randomInput(){
    const spec = this._spec();
    const sizeEl = $('#b-len');
    const size = sizeEl ? sizeEl.value : null;
    switch(spec.kind){
      case 'bin': {
        const len = +size || 3;
        let s = randDigits(2, len);
        if(s[0] === '0') s = '1' + s.slice(1);
        this._digits = s; break;
      }
      case 'dec': {
        const range = size === 'medium' ? [32, 99] : [5, 31];
        this._dec = randInt(range[0], range[1]); break;
      }
      case 'oct': {
        const len = +size || 2;
        let s = randDigits(8, len);
        if(s[0] === '0') s = DIGITS[randInt(1,7)] + s.slice(1);
        this._digits = s; break;
      }
      case 'hex': {
        const len = +size || 2;
        let s = randDigits(16, len);
        if(s[0] === '0') s = DIGITS[randInt(1,15)] + s.slice(1);
        this._digits = s; break;
      }
    }
  },
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },
  _stage(body){ return `<div class="base-stage">${body}</div>`; },
  _head(label, state){ return `<div class="stage-hd"><span>${label}</span><span class="state">${state}</span></div>`; },
  /* 写出原数的通用画布 */
  drawNumber(num, label, state){
    const srcName = label.split('→')[0].trim();
    return this._stage(`${this._head(label, state)}
      <div style="text-align:center;padding:clamp(18px,3vw,30px) 0">
        <span style="font-family:var(--mono);font-size:clamp(26px,4.5vw,38px);font-weight:700;color:var(--text)">
          ${this.typeChars(num)}<sub style="font-size:.45em;color:var(--text-3)">${srcName}</sub>
        </span><span class="caret pen"></span>
      </div>`);
  },

  /* ========== 二转十：按权展开法 ========== */
  genBin2Dec(){
    if(!this._readCustom()) this._randomInput();
    const digits = this._digits;
    const dec = fromBase(digits, 2);
    const label = '二进制 → 十进制';
    const n = digits.length;
    const powers  = [...digits].map((_,i)=>n-1-i);
    const weights = powers.map(p=>Math.pow(2,p));
    const contribs= [...digits].map((d,i)=>DIGITS.indexOf(d)*weights[i]);
    const cum = []; let s = 0; contribs.forEach(c=>{ s+=c; cum.push(s); });
    const steps = [];

    steps.push({desc:`写出要转换的二进制数：<b class="c-op">${digits}</b>₂`,
      draw:()=>{ $('#stage').innerHTML = this.drawNumber(digits, label, '准备'); }});

    steps.push({desc:`用<b>按权展开法</b>：从左到右每一位的数码 × 位权（2<sup>位序</sup>）得到贡献，再把所有贡献相加。`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'按权展开')}
          <div class="step-title"><span class="st-dot"></span>位权展开表</div>
          <table class="exp-table">
            <tr><th>位序</th>${powers.map(p=>`<th>${p}</th>`).join('')}</tr>
            <tr><th>数码</th>${[...digits].map((d,i)=>`<td class="write-in" style="animation-delay:${i*0.08}s">${d}</td>`).join('')}</tr>
            <tr><th>位权</th>${weights.map(()=>`<td class="empty">?</td>`).join('')}</tr>
            <tr><th>贡献</th>${contribs.map(()=>`<td class="empty">?</td>`).join('')}</tr>
          </table>`);
      }});

    for(let i=0;i<n;i++){
      const idx = i;
      steps.push({desc:`第 ${i+1} 位：数码 <b class="c-op">${digits[idx]}</b>，位序 <b>${powers[idx]}</b>，位权 2<sup>${powers[idx]}</sup> = <b>${weights[idx]}</b>，贡献 = ${digits[idx]} × ${weights[idx]} = <b class="c-op">${contribs[idx]}</b>，累计 ${cum[idx]}`,
        draw:()=>{
          $('#stage').innerHTML = this._stage(`
            ${this._head(label,`第 ${idx+1}/${n} 位`)}
            <div class="step-title"><span class="st-dot"></span>位权展开表</div>
            <table class="exp-table">
              <tr><th>位序</th>${powers.map(p=>`<th>${p}</th>`).join('')}</tr>
              <tr><th>数码</th>${[...digits].map((d,j)=>`<td class="${j<=idx?'done':''}">${d}</td>`).join('')}</tr>
              <tr><th>位权</th>${weights.map((w,j)=>`<td class="${j<idx?'done':j===idx?'cur':'empty'}">${j<=idx?w:'?'}</td>`).join('')}</tr>
              <tr><th>贡献</th>${contribs.map((c,j)=>`<td class="${j<idx?'done':j===idx?'cur':'empty'}">${j<=idx?c:'?'}</td>`).join('')}</tr>
            </table>
            <div class="exp-inline" style="display:flex;justify-content:center;flex-wrap:wrap;align-items:center;gap:8px">
              <span class="term done">${digits[idx]} × 2<sup>${powers[idx]}</sup></span>
              <span class="op">=</span>
              <span class="term cur">${contribs[idx]}</span>
              <span class="op">→</span><span class="term">累计 <b class="res">${cum[idx]}</b></span>
            </div>`);
        }});
    }

    steps.push({desc:`把所有位贡献相加：${contribs.join(' + ')} = <b class="c-op">${dec}</b>`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'完成')}
          <div class="step-title"><span class="st-dot"></span>位权展开表</div>
          <table class="exp-table">
            <tr><th>位序</th>${powers.map(p=>`<th>${p}</th>`).join('')}</tr>
            <tr><th>数码</th>${[...digits].map(d=>`<td class="done">${d}</td>`).join('')}</tr>
            <tr><th>位权</th>${weights.map(w=>`<td class="done">${w}</td>`).join('')}</tr>
            <tr><th>贡献</th>${contribs.map(c=>`<td class="done">${c}</td>`).join('')}</tr>
          </table>
          <div class="exp-inline" style="display:flex;justify-content:center;flex-wrap:wrap;align-items:center;gap:8px">
            ${contribs.map((c,i)=>`<span class="term done">${c}</span>${i<contribs.length-1?'<span class="op">+</span>':''}`).join('')}
            <span class="op">=</span><span class="term cur">${dec}</span>
          </div>
          <div class="write-result">${digits}<sub>二进制</sub> = ${dec}<sub>十进制</sub></div>`);
      }});
    return steps;
  },

  /* ========== 十转二：除基取余 + 倒排（逆序）读取 ========== */
  _rankName(k, R){ return k===0?'最高位' : k===R-1?'最低位' : `第 ${R-k} 高位`; },
  genDec2Bin(){
    if(!this._readCustom()) this._randomInput();
    const dec = this._dec;
    const allRows = []; let temp = dec;
    while(temp > 0){ const q = Math.floor(temp/2), r = temp%2; allRows.push({n:temp, q, r}); temp = q; }
    if(allRows.length === 0) allRows.push({n:0, q:0, r:0});
    const R = allRows.length;
    const rems = allRows.map(x=>x.r);
    const ans = rems.slice().reverse().join('');
    const label = '十进制 → 二进制';
    const steps = [];

    /* 1. 写出原数 */
    steps.push({desc:`写出要转换的十进制数：<b class="c-op">${dec}</b>₁₀`,
      draw:()=>{ $('#stage').innerHTML = this.drawNumber(dec, label, '准备'); }});

    /* 2. 方法说明 */
    steps.push({desc:`用<b>除基取余法</b>：不断除以 <b>2</b>，记录每次的商和余数，直到商为 <b>0</b>；最后把余数<b>从下往上逆序（倒排）</b>读取，才是二进制结果。`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'除基取余')}
          <div class="step-title"><span class="st-dot"></span>竖式除法（基数 = 2）</div>
          <div class="annot" style="display:block">每一行格式：<b>被除数 ÷ 2 = 商 … 余数</b>；先从上到下收集余数，最后再逆序读取。</div>`);
      }});

    /* 3. 逐次除法：所有竖式行同屏展示，已完成绿色、当前橙色、未操作灰色 */
    for(let i=0;i<R;i++){
      const idx = i;
      steps.push({desc:`<b>${allRows[idx].n}</b> ÷ 2 = <b>${allRows[idx].q}</b>，余数 <b class="c-op">${allRows[idx].r}</b>（第 ${idx+1} 个被收集的余数）`,
        draw:()=>{
          const rowsHtml = allRows.map((rd,j)=>{
            const state = j<idx?'done' : j===idx?'cur' : 'empty';
            return `<div class="dv-row ${state} ${j<=idx?'write-in':''}" style="${j<=idx?`animation-delay:${j*0.05}s`:''}">
              <span class="dv-num">${j<=idx?rd.n:'?'}</span><span class="dv-op">÷</span><span class="dv-num">2</span><span class="dv-eq">=</span><span class="dv-q">${j<=idx?rd.q:'?'}</span><span class="dv-rem">${j<=idx?rd.r:'?'}</span>
            </div>`;
          }).join('');
          const rack = rems.slice(0,idx+1).map((r,j)=>`<span class="chip o write-in" style="animation-delay:${j*0.06}s">${r}</span>`).join('');
          $('#stage').innerHTML = this._stage(`
            ${this._head(label,`第 ${idx+1}/${R} 次除法`)}
            <div class="step-title"><span class="st-dot"></span>竖式除法（基数 = 2）</div>
            <div class="two-col">
              <div><div class="div-vertical">${rowsHtml}</div></div>
              <div>
                <div class="annot" style="margin-bottom:10px"><b>余数收集区</b>（先收集 → 后收集，顺序暂存）</div>
                <div class="group-box" style="flex-wrap:wrap;justify-content:flex-start;min-height:48px">${rack||'<span style="color:#CBD5E1;font-family:var(--mono);font-size:clamp(14px,2vw,18px)">（暂无）</span>'}</div>
              </div>
            </div>`);
        }});
    }

    /* 4. 倒排阶段：从最下面一行（最后收集的余数）开始，自下而上依次标注位序 */
    for(let k=0;k<R;k++){
      const rowIdx = R-1-k;
      steps.push({desc:`逆序读取：第 <b>${k+1}</b> 个读出的余数是竖式最下面的 <b>${allRows[rowIdx].r}</b>，对应 <b class="c-op">${this._rankName(k,R)}</b>。`,
        draw:()=>{
          const rowsHtml = allRows.map((rd,j)=>{
            const isMarked = j >= rowIdx;
            const isCur = j === rowIdx;
            let badge = '';
            if(isMarked){
              badge = `<span class="dv-rank ${isCur?'cur':''}">${this._rankName(R-1-j, R)}</span>`;
            }
            return `<div class="dv-row ${isCur?'cur':'done'}">
              <span class="dv-num">${rd.n}</span><span class="dv-op">÷</span><span class="dv-num">2</span><span class="dv-eq">=</span><span class="dv-q">${rd.q}</span><span class="dv-rem">${rd.r}</span>${badge}
            </div>`;
          }).join('');
          const readRack = [];
          for(let t=0;t<=k;t++) readRack.push(rems[R-1-t]);
          const rackHtml = readRack.map((r,t)=>`<span class="grp ${t===k?'cur':'done'} write-in" style="animation-delay:${t*0.08}s">${r}<span class="grp-label">${this._rankName(t,R)}</span></span>`).join('');
          $('#stage').innerHTML = this._stage(`
            ${this._head(label,'逆序读取')}
            <div class="step-title"><span class="st-dot"></span>余数从下往上读（最高位 → 最低位）</div>
            <div class="two-col">
              <div><div class="div-vertical">${rowsHtml}</div></div>
              <div class="rev-arrow">
                <div class="ra-top">最低位<br>（第一个余数）</div>
                <div class="ra-mid">↑</div>
                <div class="ra-bottom">最高位<br>（最后一个余数）</div>
              </div>
              <div>
                <div class="annot" style="margin-bottom:10px"><b>逆序读取区</b>（最高位 → 最低位）</div>
                <div class="group-box" style="flex-wrap:wrap;justify-content:flex-start;min-height:48px">${rackHtml}</div>
              </div>
            </div>`);
        }});
    }

    /* 5. 完成：倒排标注 + 结果框 */
    steps.push({desc:`余数逆序排列：${rems.slice().reverse().join('')}，即 <b class="c-op">${ans}</b>₂。`,
      draw:()=>{
        const rowsHtml = allRows.map((rd,j)=>{
          return `<div class="dv-row done">
            <span class="dv-num">${rd.n}</span><span class="dv-op">÷</span><span class="dv-num">2</span><span class="dv-eq">=</span><span class="dv-q">${rd.q}</span><span class="dv-rem">${rd.r}</span>
            <span class="dv-rank">${this._rankName(R-1-j, R)}</span>
          </div>`;
        }).join('');
        const rackHtml = rems.slice().reverse().map((r,t)=>`<span class="grp done">${r}<span class="grp-label">${this._rankName(t,R)}</span></span>`).join('');
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'完成')}
          <div class="step-title"><span class="st-dot"></span>竖式除法 · 余数逆序读取</div>
          <div class="two-col">
            <div><div class="div-vertical">${rowsHtml}</div></div>
            <div class="rev-arrow">
              <div class="ra-top">最低位<br>（第一个余数）</div>
              <div class="ra-mid">↑</div>
              <div class="ra-bottom">最高位<br>（最后一个余数）</div>
            </div>
            <div>
              <div class="annot" style="margin-bottom:10px"><b>二进制结果</b>（最高位 → 最低位）</div>
              <div class="group-box" style="flex-wrap:wrap;justify-content:flex-start;min-height:48px">${rackHtml}</div>
            </div>
          </div>
          <div class="write-result">${dec}<sub>十进制</sub> = ${ans}<sub>二进制</sub></div>`);
      }});
    return steps;
  },

  /* ========== 二 → 八/十六：分组法 ========== */
  genBinGroup(groupSize, dstB){
    if(!this._readCustom()) this._randomInput();
    const digits = this._digits;
    const padded = digits.padStart(Math.ceil(digits.length/groupSize)*groupSize, '0');
    const groups = [];
    for(let i=0;i<padded.length;i+=groupSize) groups.push(padded.slice(i,i+groupSize));
    const groupVals  = groups.map(g=>fromBase(g,2));
    const groupDigits= groupVals.map(v=>DIGITS[v]);
    const ans = toBase(fromBase(digits,2), dstB);
    const label = `二进制 → ${B[dstB]}`;
    const groupName = groupSize===3?'三位一组':'四位一组';
    const steps = [];

    steps.push({desc:`写出要转换的二进制数：<b class="c-op">${digits}</b>₂`,
      draw:()=>{ $('#stage').innerHTML = this.drawNumber(digits, label, '准备'); }});

    steps.push({desc:`从<b>右往左</b>每 <b>${groupSize}</b> 位分一组，左侧不足补 <b>0</b>：<b class="c-op">${padded}</b>，共 ${groups.length} 组。`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,groupName)}
          <div class="annot" style="margin-bottom:12px">分组方式：${groupSize} 位一组，左侧补零对齐（原数 ${digits}）。</div>
          <div style="text-align:center">
            <div class="group-box">${groups.map((g,i)=>`<span class="grp write-in" style="animation-delay:${i*0.12}s">${g}</span>`).join('')}</div>
          </div>`);
      }});

    for(let i=0;i<groups.length;i++){
      const idx = i;
      steps.push({desc:`第 ${i+1} 组 <b class="c-op">${groups[idx]}</b> = ${groupVals[idx]} → ${B[dstB]}数码 <b class="c-op">${groupDigits[idx]}</b>`,
        draw:()=>{
          $('#stage').innerHTML = this._stage(`
            ${this._head(label,`第 ${idx+1}/${groups.length} 组`)}
            <div style="text-align:center">
              <div class="group-box">${groups.map((g,j)=>`<span class="grp ${j<idx?'done':j===idx?'cur':''}">${g}</span>`).join('')}</div>
              <div class="annot" style="margin:clamp(12px,2.2vw,18px) 0">↓ 逐组换算 ↓</div>
              <div class="group-box">${groups.map((g,j)=>`<span class="grp ${j<idx?'done':j===idx?'cur':''}" style="${j>=idx?'opacity:.35':''}">${j<=idx?groupDigits[j]:'?'}</span>`).join('')}</div>
              <div class="exp-inline" style="display:inline-flex;align-items:center;gap:8px;margin-top:clamp(12px,2.2vw,18px)">
                <span class="term ${idx===0?'cur':'done'}">${groups[idx]}</span><span class="op">=</span>
                <span class="term cur">${groupVals[idx]}</span><span class="op">→</span>
                <span class="term cur">${groupDigits[idx]}</span>
              </div>
            </div>`);
        }});
    }

    steps.push({desc:`各组结果按顺序拼接：${groupDigits.join('')} → <b class="c-op">${ans}</b>`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'完成')}
          <div style="text-align:center">
            <div class="group-box">${groups.map(g=>`<span class="grp done">${g}</span>`).join('')}</div>
            <div class="annot" style="margin:clamp(12px,2.2vw,18px) 0;color:#22C55E">↓ 拼接 ↓</div>
            <div class="group-box">${groupDigits.map(d=>`<span class="grp done" style="border-color:#22C55E;background:#E8FBF0">${d}</span>`).join('')}</div>
          </div>
          <div class="write-result">${digits}<sub>二进制</sub> = ${ans}<sub>${B[dstB]}</sub></div>`);
      }});
    return steps;
  },

  /* ========== 八/十六 → 二：逐位展开法 ========== */
  genToBin(srcB){
    if(!this._readCustom()) this._randomInput();
    const digits = this._digits;
    const groupSize = srcB===8?3:4;
    const bitGroups = [...digits].map(d=>DIGITS.indexOf(d).toString(2).padStart(groupSize,'0'));
    const combined = bitGroups.join('');
    const ans = toBase(fromBase(digits, srcB), 2);
    const label = `${B[srcB]} → 二进制`;
    const steps = [];

    steps.push({desc:`写出要转换的${B[srcB]}数：<b class="c-op">${digits}</b>`,
      draw:()=>{ $('#stage').innerHTML = this.drawNumber(digits, label, '准备'); }});

    steps.push({desc:`每位${B[srcB]}数码展开为 <b>${groupSize}</b> 位二进制（高位补 0 对齐）。`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'逐位展开')}
          <div style="text-align:center">
            <div class="group-box">${[...digits].map((d,i)=>`<span class="grp write-in" style="animation-delay:${i*0.1}s">${d}</span>`).join('')}</div>
            <div class="annot" style="margin-top:clamp(12px,2.2vw,18px)">每位展开为 ${groupSize} 位二进制</div>
          </div>`);
      }});

    for(let i=0;i<digits.length;i++){
      const idx = i; const v = DIGITS.indexOf(digits[idx]);
      steps.push({desc:`第 ${i+1} 位 <b class="c-op">${digits[idx]}</b> = ${v} → 二进制 <b class="c-op">${bitGroups[idx]}</b>`,
        draw:()=>{
          $('#stage').innerHTML = this._stage(`
            ${this._head(label,`第 ${idx+1}/${digits.length} 位`)}
            <div style="text-align:center">
              <div class="group-box">${[...digits].map((d,j)=>`<span class="grp ${j<idx?'done':j===idx?'cur':''}">${d}</span>`).join('')}</div>
              <div class="annot" style="margin:clamp(12px,2.2vw,18px) 0">↓ 逐位展开 ↓</div>
              <div class="group-box">${bitGroups.map((bg,j)=>`<span class="grp ${j<idx?'done':j===idx?'cur':''}" style="${j>=idx?'opacity:.35':''};min-width:${groupSize*1.4}em">${j<=idx?bg:'?'}</span>`).join('')}</div>
              <div class="exp-inline" style="display:inline-flex;align-items:center;gap:8px;margin-top:clamp(12px,2.2vw,18px)">
                <span class="term cur">${digits[idx]}</span><span class="op">=</span>
                <span class="term cur">${v}</span><span class="op">→</span>
                <span class="term cur">${bitGroups[idx]}</span>
              </div>
            </div>`);
        }});
    }

    steps.push({desc:`拼接后为 <b>${combined}</b>${combined!==ans?`，去掉前导零得 <b class="c-op">${ans}</b>`:`，即结果 <b class="c-op">${ans}</b>`}`,
      draw:()=>{
        $('#stage').innerHTML = this._stage(`
          ${this._head(label,'完成')}
          <div style="text-align:center">
            <div class="group-box">${[...digits].map(d=>`<span class="grp done">${d}</span>`).join('')}</div>
            <div class="annot" style="margin:clamp(12px,2.2vw,18px) 0;color:#22C55E">↓ 拼接 ↓</div>
            <div style="font-family:var(--mono);font-size:clamp(18px,2.8vw,28px);font-weight:700;letter-spacing:3px;color:#334155">${combined}</div>
            ${combined!==ans?`<div class="annot" style="margin-top:10px">去掉前导零：<b>${ans}</b></div>`:''}
          </div>
          <div class="write-result">${digits}<sub>${B[srcB]}</sub> = ${ans}<sub>二进制</sub></div>`);
      }});
    return steps;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const map = {
      bin2dec:  ['二进制 → 十进制','按权展开法'],
      dec2bin:  ['十进制 → 二进制','除基取余法（余数逆序读取）'],
      bin2oct:  ['二进制 → 八进制','三位分组法'],
      oct2bin:  ['八进制 → 二进制','逐位展开法'],
      bin2hex:  ['二进制 → 十六进制','四位分组法'],
      hex2bin:  ['十六进制 → 二进制','逐位展开法']
    };
    const [label, method] = map[this.currentSub] || ['进制转换',''];
    $('#stage').innerHTML = this._stage(`
      ${this._head(label,'准备就绪')}
      <div style="text-align:center;padding:clamp(20px,4vw,36px) 0;color:#94A3B8">
        <div style="font-size:clamp(15px,2.2vw,18px);margin-bottom:10px">使用 <b style="color:#4F6BFF">${method}</b></div>
        <div style="font-size:clamp(13px,1.8vw,15px)">点击「自动演示」或「下一步」开始手写推导</div>
      </div>`);
  },

  generate(){
    switch(this.currentSub){
      case 'bin2dec': this.steps = this.genBin2Dec(); break;
      case 'dec2bin': this.steps = this.genDec2Bin(); break;
      case 'bin2oct': this.steps = this.genBinGroup(3, 8); break;
      case 'oct2bin': this.steps = this.genToBin(8); break;
      case 'bin2hex': this.steps = this.genBinGroup(4, 16); break;
      case 'hex2bin': this.steps = this.genToBin(16); break;
      default: this.steps = this.genBin2Dec();
    }
  }
};
