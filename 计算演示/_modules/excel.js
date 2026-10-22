/* ================================================================
 * Excel 公式
 * 子页面：ref_rel 相对引用 / ref_abs 绝对引用 / ref_mix 混合引用 / func 统计函数与数据类型
 * 引用复制路径：起点 B2 → 向右 C2、D2 → 向下 B3、B4
 * 函数数据区域：A1:D4（数字 10/20/30、文本 abc、逻辑 TRUE、空单元格）
 * ================================================================ */

/* ---- 大屏幕专用样式注入（仅作用于本模块 .excel-stage 容器，不影响其他模块） ---- */
if(!document.getElementById('excel-xl-style')){
  document.head.insertAdjacentHTML('beforeend', `
  <style id="excel-xl-style">
    .excel-stage .xl-grid{display:grid;gap:5px;font-family:var(--mono);font-size:clamp(14px,2vw,18px);margin-bottom:clamp(10px,1.8vw,14px)}
    .excel-stage .xl-grid .xl-cell{border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:clamp(10px,1.6vw,16px) clamp(8px,1.2vw,12px);font-weight:700;min-height:clamp(46px,6vw,60px)}
    .excel-stage .xl-grid .xl-cell.hd{background:var(--surface-3);color:var(--accent)}
    .excel-stage .xl-grid .xl-cell.ignored{background:#F8FAFC;color:#94A3B8;text-decoration:line-through;opacity:.72;border-style:dashed}
    .excel-stage .xl-grid .xl-cell.empty-cell{color:#CBD5E1;font-weight:400}
    .excel-stage .xl-formula{font-family:var(--mono);background:var(--code-bg);color:#8FD3FF;border-radius:10px;padding:clamp(10px,1.6vw,14px) clamp(14px,2vw,18px);font-size:clamp(15px,2vw,19px);margin-bottom:clamp(10px,1.8vw,14px);overflow-x:auto;line-height:1.8}
    .excel-stage .ref-move{display:inline-flex;align-items:center;gap:4px;background:var(--orange-soft);color:#B45309;border:1.5px solid var(--orange);border-radius:8px;padding:2px 10px;font-weight:800}
    .excel-stage .ref-lock{display:inline-flex;align-items:center;gap:4px;background:var(--green-soft);color:#15803D;border:1.5px solid var(--green);border-radius:8px;padding:2px 10px;font-weight:800}
    .excel-stage .ref-still{display:inline-flex;align-items:center;gap:4px;background:#F1F5F9;color:var(--text-3);border:1.5px dashed #CBD5E1;border-radius:8px;padding:2px 10px;font-weight:700}
    .excel-stage .type-tag{font-family:var(--font);font-size:clamp(9px,1.1vw,11px);font-weight:700;padding:1px 10px;border-radius:999px;letter-spacing:.3px}
    .excel-stage .type-tag.num{background:var(--green-soft);color:#15803D}
    .excel-stage .type-tag.text{background:#F1F5F9;color:var(--text-3)}
    .excel-stage .type-tag.bool{background:#F3E8FF;color:#7C3AED}
    .excel-stage .type-tag.empty{background:#F8FAFC;color:#CBD5E1}
    .excel-stage .ref-tbl{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:clamp(12px,1.6vw,15px);margin-top:clamp(8px,1.5vw,12px)}
    .excel-stage .ref-tbl th{background:var(--surface-3);color:var(--accent);font-weight:700;padding:clamp(8px,1.2vw,12px) clamp(6px,1vw,10px);border:1px solid var(--border)}
    .excel-stage .ref-tbl td{padding:clamp(8px,1.2vw,12px) clamp(6px,1vw,10px);border:1px solid var(--border);text-align:center;background:#fff}
    .excel-stage .ref-tbl td.note{font-family:var(--font);font-size:clamp(11px,1.4vw,13px);color:var(--text-2);text-align:left}
    .excel-stage .ana-card{background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius);padding:clamp(10px,1.6vw,16px);margin-top:clamp(8px,1.5vw,12px)}
    .excel-stage .ana-row{display:flex;align-items:center;gap:clamp(8px,1.4vw,12px);flex-wrap:wrap;font-size:clamp(13px,1.8vw,17px);padding:clamp(6px,1vw,10px) 0}
    .excel-stage .ana-label{font-family:var(--font);font-weight:700;color:var(--text-3);min-width:64px}
    .excel-stage .ana-note{font-family:var(--font);font-weight:700}
    .excel-stage .judge-row{display:flex;align-items:center;gap:clamp(8px,1.4vw,12px);flex-wrap:wrap;font-family:var(--mono);font-size:clamp(12px,1.6vw,15px);padding:clamp(6px,1vw,10px) 0;border-bottom:1px dashed var(--border)}
    .excel-stage .judge-row:last-child{border-bottom:none}
    .excel-stage .judge-addr{color:var(--accent);font-weight:800;min-width:64px}
    .excel-stage .judge-val{min-width:80px;text-align:center}
    .excel-stage .chip-ok{background:var(--green);color:#fff;border-radius:7px;padding:2px 12px;font-weight:700;font-size:clamp(11px,1.3vw,13px)}
    .excel-stage .chip-no{background:#E2E8F0;color:#64748B;border-radius:7px;padding:2px 12px;font-weight:700;font-size:clamp(11px,1.3vw,13px);text-decoration:line-through}
    .excel-stage .legend{display:flex;gap:clamp(8px,1.5vw,14px);flex-wrap:wrap;margin:clamp(8px,1.5vw,12px) 0}
    .excel-stage .legend .lg{display:inline-flex;align-items:center;gap:6px;font-family:var(--font);font-size:clamp(11px,1.4vw,13px);color:var(--text-2);font-weight:600}
    .excel-stage .legend .swatch{width:14px;height:14px;border-radius:4px;display:inline-block}
  </style>`);
}

/* ---- 引用类型配置 ---- */
const REF_CONFIG = {
  rel: {
    ref:'A1', title:'相对引用', subtitle:'列标、行号都随目标位置偏移',
    intro:'<b>相对引用</b>（如 A1）：列标、行号都没有 $，复制时一起随目标位置偏移——向右复制列号增大，向下复制行号增大。',
    analysisDesc:'先拆解 B2 中公式 =A1 的两个部分，判断谁会动、谁不动。',
    ana:[
      {label:'列标', lock:false, val:'A', relText:'相对 → 向右复制时列标变', lockText:'锁定 → 不动'},
      {label:'行号', lock:false, val:'1', relText:'相对 → 向下复制时行号变', lockText:'不动'}
    ],
    stepDesc:{
      C2:'向右复制 1 格到 <b class="c-op">C2</b>：列偏移 <b>+1</b>（A→B，橙色变化），同行复制行号 1 不变，公式变为 <b class="c-op">=B1</b>。',
      D2:'再向右复制到 <b class="c-op">D2</b>：列偏移 <b>+2</b>（A→C，橙色变化），行号 1 不变，公式变为 <b class="c-op">=C1</b>。',
      B3:'向下复制 1 格到 <b class="c-op">B3</b>：行偏移 <b>+1</b>（1→2，橙色变化），同列复制列标 A 不变，公式变为 <b class="c-op">=A2</b>。',
      B4:'向下复制到 <b class="c-op">B4</b>：行偏移 <b>+2</b>（1→3，橙色变化），列标 A 不变，公式变为 <b class="c-op">=A3</b>。'
    },
    note:{
      C2:'列 A→B（橙色变化）；同行复制，行号 1 不变',
      D2:'列 A→C（橙色变化）；行号 1 不变',
      B3:'行 1→2（橙色变化）；同列复制，列标 A 不变',
      B4:'行 1→3（橙色变化）；列标 A 不变'
    },
    finalDesc:'相对引用复制完成：向右复制只改列标、向下复制只改行号，公式整体随目标位置平移。验证：C2=B1、D2=C1、B3=A2、B4=A3。'
  },
  abs: {
    ref:'$A$1', title:'绝对引用', subtitle:'列标、行号都被 $ 锁定',
    intro:'<b>绝对引用</b>（如 $A$1）：列标前和行号前都有 $，同时锁定列和行，复制到任何位置公式都不变。',
    analysisDesc:'拆解 =$A$1：列标、行号两个部分都被 $ 绿色锁定。',
    ana:[
      {label:'列标', lock:true, val:'A', relText:'', lockText:'$ 锁定 → 复制后列标不动'},
      {label:'行号', lock:true, val:'1', relText:'', lockText:'$ 锁定 → 复制后行号不动'}
    ],
    stepDesc:{
      C2:'向右复制到 <b class="c-op">C2</b>：列、行都被 $ 锁定，<b>两个部分都不动</b>（绿色锁定），公式仍为 <b class="c-op">=$A$1</b>。',
      D2:'再向右复制到 <b class="c-op">D2</b>：列、行依旧锁定不动，公式仍为 <b class="c-op">=$A$1</b>。',
      B3:'向下复制到 <b class="c-op">B3</b>：列、行都不动，公式仍为 <b class="c-op">=$A$1</b>。',
      B4:'向下复制到 <b class="c-op">B4</b>：列、行都不动，公式仍为 <b class="c-op">=$A$1</b>。'
    },
    note:{
      C2:'列、行均被 $ 锁定，公式不变',
      D2:'列、行均被 $ 锁定，公式不变',
      B3:'列、行均被 $ 锁定，公式不变',
      B4:'列、行均被 $ 锁定，公式不变'
    },
    finalDesc:'绝对引用复制完成：无论向右还是向下复制，公式始终 =$A$1，绿色锁定的部分保持不动。'
  },
  mixcol: {
    ref:'$A1', title:'混合引用（列绝对 · 行相对）', subtitle:'列锁死，行偏移',
    intro:'<b>混合引用 $A1</b>：列标前有 $ → 列被锁定；行号前没有 $ → 行相对偏移。',
    analysisDesc:'拆解 =$A1：列标 $A 绿色锁定，行号 1 橙色随动。',
    ana:[
      {label:'列标', lock:true, val:'A', relText:'', lockText:'$ 锁定 → 复制后列标始终是 A'},
      {label:'行号', lock:false, val:'1', relText:'相对 → 向下复制时行号变'}
    ],
    stepDesc:{
      C2:'向右复制到 <b class="c-op">C2</b>：列 $A 锁死不动（绿色）；在同一行内复制，行号 1 未偏移，公式 =<b class="c-op">=$A1</b>。',
      D2:'向右复制到 <b class="c-op">D2</b>：列 $A 仍锁定不动，行号 1 未偏移，公式 =<b class="c-op">=$A1</b>。',
      B3:'向下复制到 <b class="c-op">B3</b>：列 $A 不动（绿色锁定）；行偏移 +1：1→2（橙色变化），公式 =<b class="c-op">=$A2</b>。',
      B4:'向下复制到 <b class="c-op">B4</b>：列 $A 不动；行偏移 +2：1→3（橙色变化），公式 =<b class="c-op">=$A3</b>。'
    },
    note:{
      C2:'列 $A 锁定不动；同行复制，行号 1 未偏移',
      D2:'列 $A 锁定不动；行号 1 未偏移',
      B3:'列 $A 锁定不动；行 1→2（橙色变化）',
      B4:'列 $A 锁定不动；行 1→3（橙色变化）'
    },
    finalDesc:'$A1 复制完成：向右复制列标、行号都不变；向下复制列标仍为 $A，行号随之下移。'
  },
  mixrow: {
    ref:'A$1', title:'混合引用（列相对 · 行绝对）', subtitle:'列偏移，行锁死',
    intro:'<b>混合引用 A$1</b>：行号前有 $ → 行被锁定；列标前没有 $ → 列相对偏移。',
    analysisDesc:'拆解 =A$1：列标 A 橙色随动，行号 $1 绿色锁定。',
    ana:[
      {label:'列标', lock:false, val:'A', relText:'相对 → 向右复制时列标变'},
      {label:'行号', lock:true, val:'1', relText:'', lockText:'$ 锁定 → 复制后行号始终是 1'}
    ],
    stepDesc:{
      C2:'向右复制到 <b class="c-op">C2</b>：行 $1 锁死不动（绿色）；列偏移 +1：A→B（橙色变化），公式 =<b class="c-op">=B$1</b>。',
      D2:'向右复制到 <b class="c-op">D2</b>：行 $1 不动；列偏移 +2：A→C（橙色变化），公式 =<b class="c-op">=C$1</b>。',
      B3:'向下复制到 <b class="c-op">B3</b>：行 $1 锁死不动（绿色）；在同一列内复制，列标 A 未偏移，公式 =<b class="c-op">=A$1</b>。',
      B4:'向下复制到 <b class="c-op">B4</b>：行 $1 不动；列标 A 未偏移，公式 =<b class="c-op">=A$1</b>。'
    },
    note:{
      C2:'行 $1 锁定不动；列 A→B（橙色变化）',
      D2:'行 $1 锁定不动；列 A→C（橙色变化）',
      B3:'行 $1 锁定不动；同列复制，列标 A 未偏移',
      B4:'行 $1 锁定不动；列标 A 未偏移'
    },
    finalDesc:'A$1 复制完成：向下复制行号、列标都不变；向右复制行号仍为 $1，列标随之右移。'
  }
};

/* ---- 函数演示数据（A1:D4） ---- */
const FUNC_CELLS = [
  {addr:'A1', disp:'10', type:'num',  num:10},
  {addr:'B1', disp:'abc', type:'text'},
  {addr:'C1', disp:'TRUE', type:'bool'},
  {addr:'D1', disp:'20', type:'num',  num:20},
  {addr:'A2', disp:'', type:'empty'},
  {addr:'B2', disp:'', type:'empty'},
  {addr:'C2', disp:'30', type:'num',  num:30},
  {addr:'D2', disp:'', type:'empty'}
];
const TYPE_META = {
  num:  {label:'数字'},
  text: {label:'文本'},
  bool: {label:'逻辑值'},
  empty:{label:'空单元格'}
};
const FUNCS = [
  {name:'SUM', title:'求和', accept:['num'],
   verdict:'只把数字单元格相加；文本、逻辑值、空单元格全部忽略。',
   process:(c)=>`${c.map(x=>x.num).join(' + ')} = ${c.reduce((a,b)=>a+b.num,0)}`,
   result:(c)=>c.reduce((a,b)=>a+b.num,0)},
  {name:'AVERAGE', title:'平均值', accept:['num'],
   verdict:'先对数字求和，再除以数字个数；文本、逻辑值、空单元格既不计入和，也不计入分母。',
   process:(c)=>{const s=c.reduce((a,b)=>a+b.num,0);return `(${c.map(x=>x.num).join(' + ')}) ÷ ${c.length} = ${s} ÷ ${c.length} = ${s/c.length}`;},
   result:(c)=>c.reduce((a,b)=>a+b.num,0)/c.length},
  {name:'COUNT', title:'计数（数字个数）', accept:['num'],
   verdict:'只数数字单元格的个数；文本、逻辑值、空单元格不计入。',
   process:(c)=>`数字单元格：${c.map(x=>x.addr).join('、')} → 共 ${c.length} 个`,
   result:(c)=>c.length},
  {name:'COUNTA', title:'计数（非空个数）', accept:['num','text','bool'],
   verdict:'数所有非空单元格：数字、文本、逻辑值都算；只有空单元格被忽略。',
   process:(c)=>`非空单元格：${c.map(x=>x.addr).join('、')} → 共 ${c.length} 个`,
   result:(c)=>c.length},
  {name:'MAX', title:'最大值', accept:['num'],
   verdict:'在数字单元格中找最大；文本、逻辑值、空单元格不参与比较。',
   process:(c)=>`max(${c.map(x=>x.num).join(', ')}) = ${Math.max(...c.map(x=>x.num))}`,
   result:(c)=>Math.max(...c.map(x=>x.num))},
  {name:'MIN', title:'最小值', accept:['num'],
   verdict:'在数字单元格中找最小；文本、逻辑值、空单元格不参与比较。',
   process:(c)=>`min(${c.map(x=>x.num).join(', ')}) = ${Math.min(...c.map(x=>x.num))}`,
   result:(c)=>Math.min(...c.map(x=>x.num))}
];

MOD.excel = {
  name:'Excel公式', sub:'相对 · 绝对 · 混合引用复制 · 统计函数与数据类型',
  badge:'表格 · 引用', tag:'表格', cat:'net',
  color:'#22C55E', soft:'#E8FBF0', type:'相对/绝对/混合引用复制 · SUM/AVERAGE/COUNT/COUNTA/MAX/MIN',
  desc:'演示相对、绝对、混合引用在公式复制时的行列变化；在含数字、文本、逻辑值、空单元格的区域上逐步计算 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg>',
  setDesc:'引用演示固定起点 B2，先向右复制到 C2、D2，再向下复制到 B3、B4；函数演示固定区域 A1:D4。',
  stageDesc:'橙色=相对/变化部分，绿色=锁定/计入，灰色删除线=忽略；观察复制时引用偏移规律与各函数的计入规则。',
  initDesc:'点击「自动演示」，观看公式引用变化或函数逐步判断计入/忽略。',
  currentSub:'ref_rel',
  subTabs:[
    {id:'ref_rel', label:'相对引用', method:'行列都偏移'},
    {id:'ref_abs', label:'绝对引用', method:'行列都锁定'},
    {id:'ref_mix', label:'混合引用', method:'锁列或锁行'},
    {id:'func',    label:'统计函数与数据类型', method:'六函数逐个判断'}
  ],
  theory(){
    return `<div class="rule"><b>相对引用</b>（如 A1）：复制时行号和列号都会随目标位置偏移，向右复制列 +1，向下复制行 +1。</div>
    <div class="rule"><b>绝对引用</b>（如 $A$1）：复制时行号和列号都<b>不变</b>，$ 同时锁定行和列。</div>
    <div class="rule"><b>混合引用 $A1</b>：列绝对、行相对——列锁死不变，行随目标位置偏移。</div>
    <div class="rule"><b>混合引用 A$1</b>：列相对、行绝对——行锁死不变，列随目标位置偏移。</div>
    <div class="rule"><b>移动</b>公式（剪切粘贴）：无论哪种引用，引用地址都不变；只有<b>复制</b>时相对部分才会偏移。</div>
    <div class="rule"><b>SUM / AVERAGE / COUNT / MAX / MIN</b>：只识别<b>数字</b>单元格；区域引用中的文本、逻辑值（TRUE/FALSE）、空单元格一律忽略。</div>
    <div class="rule"><b>COUNTA</b>：统计所有<b>非空</b>单元格，数字、文本、逻辑值都计入，只忽略空单元格；注意与 COUNT（只数数字）区分。</div>
    <div class="rule"><b>AVERAGE</b>：总和 ÷ 数字个数，分母同样只统计数字单元格，文本、逻辑值、空单元格不计入分母。</div>`;
  },

  /* ---------- 参数区 ---------- */
  settingsHTML(){
    if(this.currentSub === 'func'){
      return `<div class="inline-setting"><div class="setting"><label>演示区域</label>
        <div class="annot" style="padding:4px 8px">固定区域 <b>A1:D4</b>：数字 10/20/30、文本 abc、逻辑 TRUE、空单元格；依次演示六个函数</div></div></div>`;
    }
    return `<div class="inline-setting"><div class="setting"><label>复制路径</label>
      <div class="annot" style="padding:4px 8px">起点 <b>B2</b>：先向右复制到 <b>C2、D2</b>，再向下复制到 <b>B3、B4</b></div></div></div>`;
  },
  afterSettings(){},

  /* ---------- 工具方法 ---------- */
  colLetter(n){ return String.fromCharCode(64 + n); },
  parseRef(ref){
    const m = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(ref);
    return { colLock:m[1]==='$', col:m[2], rowLock:m[3]==='$', row:m[4] };
  },
  offsetRef(ref, dc, dr){
    const p = this.parseRef(ref);
    const colN = p.col.charCodeAt(0) - 64;
    const newCol = p.colLock ? p.col : this.colLetter(colN + dc);
    const newRow = p.rowLock ? p.row : String(Number(p.row) + dr);
    return (p.colLock ? '$' : '') + newCol + (p.rowLock ? '$' : '') + newRow;
  },
  lockSVG(){
    return `<svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="10" width="16" height="10" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></svg>`;
  },
  fmtNum(v){ return Number.isInteger(v) ? String(v) : v.toFixed(2); },

  /* 源公式：按引用类型上色（锁定=绿色锁，相对=橙色） */
  renderRefSource(ref){
    const p = this.parseRef(ref);
    const col = p.colLock
      ? `<span class="ref-lock">$${p.col}${this.lockSVG()}</span>`
      : `<span class="ref-move">${p.col}</span>`;
    const row = p.rowLock
      ? `<span class="ref-lock">$${p.row}${this.lockSVG()}</span>`
      : `<span class="ref-move">${p.row}</span>`;
    return col + row;
  },
  /* 目标公式：变化=橙色，锁定=绿色锁，本方向未偏移=灰色虚线 */
  renderRefTarget(ref, dc, dr){
    const p = this.parseRef(ref);
    const newCol = p.colLock ? p.col : this.colLetter(p.col.charCodeAt(0)-64 + dc);
    const newRow = p.rowLock ? p.row : String(Number(p.row) + dr);
    let col, row;
    if(p.colLock)      col = `<span class="ref-lock">$${newCol}${this.lockSVG()}</span>`;
    else if(dc === 0) col = `<span class="ref-still">${newCol}</span>`;
    else               col = `<span class="ref-move">${newCol}</span>`;
    if(p.rowLock)      row = `<span class="ref-lock">$${newRow}${this.lockSVG()}</span>`;
    else if(dr === 0) row = `<span class="ref-still">${newRow}</span>`;
    else               row = `<span class="ref-move">${newRow}</span>`;
    return col + row;
  },

  /* ---------- 引用复制舞台 ---------- */
  drawRefStage(o){
    const { cfg, srcRef, map, visible, phase } = o;
    const cols = ['A','B','C','D'];
    let html = `<div class="excel-stage">
      <div class="stage-hd"><span>${cfg.title} · 公式复制</span><span class="state">起点 B2 · 向右 C2/D2 · 向下 B3/B4</span></div>
      <div class="step-title"><span class="st-dot"></span>${cfg.subtitle}</div>
      <div class="xl-grid" style="grid-template-columns:38px repeat(4,1fr)">
        <div class="xl-cell hd"></div>${cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('')}`;
    for(let r=1; r<=4; r++){
      html += `<div class="xl-cell hd">${r}</div>`;
      for(let c=1; c<=4; c++){
        const cell = this.colLetter(c) + r;
        let cls = '', content = '<span style="color:#E2E8F0">·</span>';
        if(cell === 'B2'){
          cls = 'sel'; content = '=' + this.renderRefSource(srcRef);
        }else if(map[cell]){
          const t = map[cell];
          if(t.state === 'pending'){ cls = ''; content = '<span style="color:#CBD5E1">?</span>'; }
          else { cls = t.state === 'done' ? 'done' : 'cur'; content = '=' + this.renderRefTarget(srcRef, t.dc, t.dr); }
        }
        html += `<div class="xl-cell ${cls}">${content}</div>`;
      }
    }
    html += `</div>`;

    const curTarget = Object.values(map).find(t=>t.state === 'cur');
    html += `<div class="xl-formula">起点 B2 =${this.renderRefSource(srcRef)}`;
    if(curTarget) html += ` &nbsp;→&nbsp; ${curTarget.addr} =${this.renderRefTarget(srcRef, curTarget.dc, curTarget.dr)}`;
    html += `</div>`;

    if(phase === 'source') html += this.analysisCard(cfg);
    html += this.refSummaryTable(o, map, visible);
    html += `</div>`;
    $('#stage').innerHTML = html;
  },

  analysisCard(cfg){
    const rows = cfg.ana.map(a=>{
      const token = a.lock
        ? `<span class="ref-lock">$${a.val}${this.lockSVG()}</span>`
        : `<span class="ref-move">${a.val}</span>`;
      const note = a.lock
        ? `<span class="ana-note" style="color:#15803D">${a.lockText}</span>`
        : `<span class="ana-note" style="color:#B45309">${a.relText}</span>`;
      return `<div class="ana-row"><span class="ana-label">${a.label}</span>${token}${note}</div>`;
    }).join('');
    return `<div class="ana-card">
      <div style="font-size:clamp(12px,1.5vw,14px);font-weight:700;color:var(--text-3);margin-bottom:6px">引用类型拆解</div>${rows}
    </div>`;
  },

  refSummaryTable(o, map, visible){
    let rows = `<tr><td>B2</td><td>原公式（起点）</td><td>=${this.renderRefSource(o.srcRef)}</td><td class="note">复制的起点单元格</td></tr>`;
    Object.values(map).slice(0, visible).forEach(t=>{
      rows += `<tr><td>${t.addr}</td><td>${t.dir === 'right' ? '向右复制' : '向下复制'}</td><td>=${this.renderRefTarget(o.srcRef, t.dc, t.dr)}</td><td class="note">${o.cfg.note[t.addr]}</td></tr>`;
    });
    return `<table class="ref-tbl">
      <thead><tr><th>单元格</th><th>复制方向</th><th>复制后公式</th><th>变化说明</th></tr></thead>
      <tbody>${rows}</tbody></table>`;
  },

  /* ---------- 单个引用类型：源公式 → 向右 C2/D2 → 向下 B3/B4 → 汇总 ---------- */
  genRefCase(type){
    const cfg = REF_CONFIG[type];
    const srcRef = cfg.ref;
    const targets = [
      {addr:'C2', dc:1, dr:0, dir:'right'},
      {addr:'D2', dc:2, dr:0, dir:'right'},
      {addr:'B3', dc:0, dr:1, dir:'down'},
      {addr:'B4', dc:0, dr:2, dir:'down'}
    ].map(t=>({...t, formula:this.offsetRef(srcRef, t.dc, t.dr)}));

    const snap = (activeIdx)=>{
      const map = {};
      targets.forEach((t,i)=>{
        let state = 'pending';
        if(activeIdx >= targets.length) state = 'done';
        else if(activeIdx < 0) state = 'pending';
        else if(i < activeIdx) state = 'done';
        else if(i === activeIdx) state = 'cur';
        map[t.addr] = {...t, state};
      });
      return map;
    };

    const steps = [];
    /* 第一步：原公式 + 引用类型分析 */
    steps.push({
      desc: cfg.intro + ' ' + cfg.analysisDesc,
      draw: ()=>{ this.drawRefStage({ cfg, srcRef, map:snap(-1), visible:0, phase:'source' }); }
    });
    /* 第二、三步：向右逐个复制；第四、五步：向下逐个复制 */
    targets.forEach((t,i)=>{
      steps.push({
        desc: cfg.stepDesc[t.addr],
        draw: ()=>{ this.drawRefStage({ cfg, srcRef, map:snap(i), visible:i+1, phase:t.dir }); }
      });
    });
    /* 第六步：汇总完成 */
    steps.push({
      desc: cfg.finalDesc,
      draw: ()=>{ this.drawRefStage({ cfg, srcRef, map:snap(targets.length), visible:targets.length, phase:'done' }); }
    });
    return steps;
  },

  /* ========== ① 相对引用 ========== */
  genRefRel(){ return this.genRefCase('rel'); },
  /* ========== ② 绝对引用 ========== */
  genRefAbs(){ return this.genRefCase('abs'); },
  /* ========== ③ 混合引用：$A1 与 A$1 两组对比 ========== */
  genRefMix(){ return this.genRefCase('mixcol').concat(this.genRefCase('mixrow')); },

  /* ---------- 函数演示网格 ---------- */
  classify(f, cell){
    if(cell.type === 'empty') return {counted:false, reason:'空单元格：无内容，忽略'};
    if(f.accept.includes(cell.type)){
      if(cell.type === 'num')  return {counted:true, reason:'数字：参与本函数计算'};
      if(cell.type === 'text') return {counted:true, reason:'文本：非空，计入 COUNTA'};
      if(cell.type === 'bool') return {counted:true, reason:'逻辑值：非空，计入 COUNTA'};
    }
    if(cell.type === 'text') return {counted:false, reason:'文本：数值函数不参与，忽略'};
    if(cell.type === 'bool') return {counted:false, reason:'逻辑值：区域引用中被忽略'};
    return {counted:false, reason:'忽略'};
  },
  /* A1:D4 中全部空单元格地址 */
  allEmptyAddrs(judged){
    const filled = judged.filter(x=>x.type !== 'empty').map(x=>x.addr);
    const empty = [];
    ['A','B','C','D'].forEach(col=>{ for(let r=1; r<=4; r++){ const a = col + r; if(!filled.includes(a)) empty.push(a); } });
    return empty;
  },
  funcGrid(judged){
    const cols = ['A','B','C','D'];
    const byAddr = {};
    (judged || FUNC_CELLS).forEach(c=>{ byAddr[c.addr] = c; });
    let html = `<div class="xl-grid" style="grid-template-columns:38px repeat(4,1fr)">
      <div class="xl-cell hd"></div>${cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('')}`;
    for(let r=1; r<=4; r++){
      html += `<div class="xl-cell hd">${r}</div>`;
      for(let c=1; c<=4; c++){
        const addr = this.colLetter(c) + r;
        const cell = byAddr[addr];
        let cls = '', content = '';
        if(!cell || cell.type === 'empty'){
          cls = 'empty-cell'; content = '<span style="font-size:.65em">空</span>';
        }else{
          if(judged) cls = cell.verdict.counted ? 'done' : 'ignored';
          content = `<div>${cell.disp}</div><span class="type-tag ${cell.type}">${TYPE_META[cell.type].label}</span>`;
        }
        html += `<div class="xl-cell ${cls}">${content}</div>`;
      }
    }
    html += `</div>`;
    return html;
  },

  funcStepDesc(f, judged){
    const counted = judged.filter(x=>x.verdict.counted).map(x=>x.addr);
    const ignored = judged.filter(x=>!x.verdict.counted && x.type !== 'empty').map(x=>x.addr);
    return `函数 <b class="c-op">${f.name}(A1:D4)</b>：${f.verdict} 计入单元格 ${counted.join('、')}${ignored.length ? `；忽略 ${ignored.join('、')} 与所有空单元格` : ''}。`;
  },

  /* ========== ④ 统计函数与数据类型 ========== */
  genFunc(){
    const steps = [];

    /* 0. 区域总览 + 类型图例 */
    steps.push({
      desc:`先认识数据区域 <b class="c-op">A1:D4</b> 中的四类数据：<b>数字</b> 10、20、30；<b>文本</b> "abc"；<b>逻辑值</b> TRUE；以及若干<b>空单元格</b>。不同函数对这些类型的计入规则不同。`,
      draw:()=>{
        let html = `<div class="excel-stage">
          <div class="stage-hd"><span>统计函数与数据类型</span><span class="state">先认识 A1:D4 中的数据类型</span></div>
          <div class="step-title"><span class="st-dot"></span>四类数据：数字 · 文本 · 逻辑值 · 空单元格</div>
          <div class="legend">
            <span class="lg"><span class="swatch" style="background:var(--green-soft);border:1px solid var(--green)"></span>数字：SUM/AVERAGE/COUNT/MAX/MIN 计入</span>
            <span class="lg"><span class="swatch" style="background:#F1F5F9;border:1px solid #CBD5E1"></span>文本：数值函数忽略，COUNTA 计入</span>
            <span class="lg"><span class="swatch" style="background:#F3E8FF;border:1px solid #8B5CF6"></span>逻辑值：区域引用中被数值函数忽略，COUNTA 计入</span>
            <span class="lg"><span class="swatch" style="background:#F8FAFC;border:1px dashed #CBD5E1"></span>空单元格：所有函数都忽略</span>
          </div>
          ${this.funcGrid(null)}
          <div class="annot" style="text-align:center;margin-top:8px">接下来依次演示 SUM、AVERAGE、COUNT、COUNTA、MAX、MIN 六个函数如何识别这些类型。</div>
        </div>`;
        $('#stage').innerHTML = html;
      }
    });

    const results = {};
    FUNCS.forEach(f=>{
      const judged = FUNC_CELLS.map(c=>({...c, verdict:this.classify(f, c)}));
      const counted = judged.filter(x=>x.verdict.counted);
      const res = f.result(counted);
      results[f.name] = { res, counted };
      steps.push({
        desc: this.funcStepDesc(f, judged),
        draw: ()=>this.drawFuncStep(f, judged, res)
      });
    });

    /* 汇总对比 */
    steps.push({
      desc:`六个函数在同一区域 <b>A1:D4</b> 上的结果对比完成。注意 <b>COUNT</b> 只数数字（3 个），<b>COUNTA</b> 数所有非空（5 个）；SUM/AVERAGE/MAX/MIN 只在数字中计算。`,
      draw:()=>this.drawFuncSummary(results)
    });
    return steps;
  },

  drawFuncStep(f, judged, res){
    const counted = judged.filter(x=>x.verdict.counted);
    const emptyAddrs = this.allEmptyAddrs(judged);
    const judgeRows = judged.filter(x=>x.type !== 'empty').map(x=>`
      <div class="judge-row">
        <span class="judge-addr">${x.addr}</span>
        <span class="judge-val">${x.disp}</span>
        <span class="type-tag ${x.type}">${TYPE_META[x.type].label}</span>
        ${x.verdict.counted ? '<span class="chip-ok">计入</span>' : '<span class="chip-no">忽略</span>'}
        <span style="font-family:var(--font);font-size:clamp(11px,1.4vw,13px);color:var(--text-3)">${x.verdict.reason}</span>
      </div>`).join('');
    let html = `<div class="excel-stage">
      <div class="stage-hd"><span>${f.name} ${f.title}</span><span class="state">数据区域 A1:D4</span></div>
      <div class="step-title"><span class="st-dot"></span>函数名 → 选中区域 → 逐个判断 → 计算过程 → 结果</div>
      ${this.funcGrid(judged)}
      <div class="xl-formula">=${f.name}(A1:D4) = <b style="color:#86EFAC" class="write-in">${this.fmtNum(res)}</b></div>
      <div class="calc-process">
        <div class="cp-label">逐个判断（绿色 = 计入，灰色删除线 = 忽略）</div>
        ${judgeRows}
        <div class="judge-row">
          <span class="judge-addr">空单元格</span>
          <span class="judge-val">∅</span>
          <span class="type-tag empty">空单元格</span>
          <span class="chip-no">忽略</span>
          <span style="font-family:var(--font);font-size:clamp(11px,1.4vw,13px);color:var(--text-3)">${emptyAddrs.join('、')}：无内容，所有函数都忽略</span>
        </div>
        <div class="cp-sum-line" style="margin-top:10px">${f.process(counted)}</div>
      </div>
      <div class="write-result">${f.name}(A1:D4) = ${this.fmtNum(res)}</div>
    </div>`;
    $('#stage').innerHTML = html;
  },

  drawFuncSummary(results){
    const rows = FUNCS.map(f=>{
      const r = results[f.name];
      const tags = r.counted.map(x=>`<span class="type-tag ${x.type}" style="margin:1px 2px">${x.addr}</span>`).join(' ');
      return `<tr><td><b>${f.name}</b></td><td style="font-family:var(--font)">${f.title}</td><td>${tags}</td><td><b class="c-op">${this.fmtNum(r.res)}</b></td></tr>`;
    }).join('');
    let html = `<div class="excel-stage">
      <div class="stage-hd"><span>六函数结果对比</span><span class="state">A1:D4 同一区域</span></div>
      <div class="step-title"><span class="st-dot"></span>计入规则与结果汇总</div>
      <table class="ref-tbl">
        <thead><tr><th>函数</th><th>作用</th><th>计入的单元格</th><th>结果</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="write-result">COUNT 只数数字（3 个）；COUNTA 数所有非空（5 个）；SUM/AVERAGE/MAX/MIN 只在数字中计算。</div>
    </div>`;
    $('#stage').innerHTML = html;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const sub = this.currentSub;
    const info = this.subTabs.find(t=>t.id === sub) || {};
    const labelMap = { ref_rel:'相对引用复制', ref_abs:'绝对引用复制', ref_mix:'混合引用复制', func:'统计函数与数据类型' };
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub]||'Excel 公式'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:#22C55E">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始推导</div>
      </div>`;
  },

  generate(){
    switch(this.currentSub){
      case 'ref_rel': this.steps = this.genRefRel(); break;
      case 'ref_abs': this.steps = this.genRefAbs(); break;
      case 'ref_mix': this.steps = this.genRefMix(); break;
      case 'func':    this.steps = this.genFunc(); break;
      default: this.steps = this.genRefRel();
    }
  }
};
