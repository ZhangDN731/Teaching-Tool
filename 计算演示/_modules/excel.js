
;;;

/* ================================================================
 * ⑤ Excel 公式（相对 / 绝对 / 混合引用的复制变化）
 * ================================================================ */

/* ================================================================
 * Excel 公式
 * 子页面：ref_rel 相对引用 / ref_abs 绝对引用 / ref_mix 混合引用 / func 统计函数与数据类型
 * 引用复制：用户自主选择方向（上 / 下 / 左 / 右）与步数（1~3 格），起点随方向自动避让边界
 * 函数数据区域：A1:D4（数字 10/20/30、文本 abc、逻辑 TRUE、空单元格）
 * ================================================================ */

/* ---- 大屏幕专用样式注入（仅作用于本模块；方向按钮位于参数区，单独给出全局规则） ---- */
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
    /* 方向选择按钮（位于参数区 #settingsRow，不在 .excel-stage 内，故单独给规则） */
    .xl-dir-row{display:inline-flex;gap:6px;align-items:center}
    .xl-dir-btn{width:38px;height:34px;border:1.5px solid var(--border);background:#fff;border-radius:9px;font-size:17px;font-weight:800;color:var(--text-2);cursor:pointer;display:inline-flex;align-items:center;justify-content:center;line-height:1;transition:all .15s ease;font-family:var(--mono)}
    .xl-dir-btn:hover{border-color:var(--accent);color:var(--accent)}
    .xl-dir-btn.active{background:var(--accent);border-color:var(--accent);color:#fff;box-shadow:0 2px 8px var(--accent-soft)}
  </style>`);
}

/* ---- 方向布局：起点（anchor）、源引用列/行、第 k 格目标地址与偏移量；向左/向上时起点自动避让边界 ---- */
const DIR_LAYOUT = {
  right:{ label:'右', anchorCol:2, anchorRow:2, refCol:1, refRow:1,
          tCol:k=>2+k, tRow:()=>2, dc:k=>k, dr:()=>0 },
  down: { label:'下', anchorCol:2, anchorRow:2, refCol:1, refRow:1,
          tCol:()=>2, tRow:k=>2+k, dc:()=>0, dr:k=>k },
  left: { label:'左', anchorCol:4, anchorRow:2, refCol:5, refRow:1,
          tCol:k=>4-k, tRow:()=>2, dc:k=>-k, dr:()=>0 },
  up:   { label:'上', anchorCol:2, anchorRow:4, refCol:1, refRow:5,
          tCol:()=>2, tRow:k=>4-k, dc:()=>0, dr:k=>-k }
};

/* ---- 引用类型配置（ana[0]=列标是否锁定，ana[1]=行号是否锁定） ---- */
const REF_CONFIG = {
  rel: {
    title:'相对引用', subtitle:'列标、行号都随目标位置偏移',
    intro:'<b>相对引用</b>（如 A1）：列标、行号都没有 $，复制时一起随目标位置偏移——横向复制列号变化，纵向复制行号变化。',
    ana:[
      {label:'列标', lock:false},
      {label:'行号', lock:false}
    ]
  },
  abs: {
    title:'绝对引用', subtitle:'列标、行号都被 $ 锁定',
    intro:'<b>绝对引用</b>（如 $A$1）：列标前和行号前都有 $，同时锁定列和行，复制到任何位置公式都不变。',
    ana:[
      {label:'列标', lock:true},
      {label:'行号', lock:true}
    ]
  },
  mixcol: {
    title:'混合引用（列绝对 · 行相对）', subtitle:'列锁死，行偏移',
    intro:'<b>混合引用 $A1</b>：列标前有 $ → 列被锁定；行号前没有 $ → 行相对偏移。',
    ana:[
      {label:'列标', lock:true},
      {label:'行号', lock:false}
    ]
  },
  mixrow: {
    title:'混合引用（列相对 · 行绝对）', subtitle:'列偏移，行锁死',
    intro:'<b>混合引用 A$1</b>：行号前有 $ → 行被锁定；列标前没有 $ → 列相对偏移。',
    ana:[
      {label:'列标', lock:false},
      {label:'行号', lock:true}
    ]
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
  desc:'演示相对、绝对、混合引用在公式复制时的行列变化（方向、步数可自选）；在含数字、文本、逻辑值、空单元格的区域上逐步计算 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg>',
  setDesc:'在上方选择复制方向（上 / 下 / 左 / 右）与步数（1~3 格）；函数演示固定区域 A1:D4。',
  stageDesc:'橙色=相对/变化部分，绿色=锁定/计入，灰色虚线=未偏移；观察所选方向上引用的偏移规律与各函数的计入规则。',
  initDesc:'点击「自动演示」，观看公式引用变化或函数逐步判断计入/忽略。',
  currentSub:'ref_rel',
  /* 用户所选复制方向与步数（默认向右 2 格） */
  _dir:'right', _steps:2,
  subTabs:[
    {id:'ref_rel', label:'相对引用', method:'行列都偏移'},
    {id:'ref_abs', label:'绝对引用', method:'行列都锁定'},
    {id:'ref_mix', label:'混合引用', method:'锁列或锁行'},
    {id:'func',    label:'统计函数与数据类型', method:'六函数逐个判断'}
  ],
  theory(){
    return `<div class="rule"><b>相对引用</b>（如 A1）：复制时行号和列号都会随目标位置偏移，向右复制列 +1，向下复制行 +1，向左、向上同理反向偏移。</div>
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
    const { dir, N, L } = this._refLayout();
    const anchor = this.colLetter(L.anchorCol) + L.anchorRow;
    const targetAddrs = [];
    for(let k=1; k<=N; k++) targetAddrs.push(this.colLetter(L.tCol(k)) + L.tRow(k));
    const dirs = [['up','↑','向上'],['down','↓','向下'],['left','←','向左'],['right','→','向右']];
    const dirBtns = dirs.map(([d,arrow,label])=>
      `<button type="button" class="xl-dir-btn ${dir===d?'active':''}" data-dir="${d}" title="${label}">${arrow}</button>`).join('');
    const opts = [1,2,3].map(n=>`<option value="${n}" ${N===n?'selected':''}>${n} 格</option>`).join('');
    return `<div class="inline-setting"><div class="setting"><label>复制方向</label>
        <div class="xl-dir-row">${dirBtns}</div></div></div>
      <div class="inline-setting"><div class="setting"><label>复制步数</label>
        <select id="excel-steps">${opts}</select></div></div>
      <div class="inline-setting"><div class="setting"><label>当前路径</label>
        <div class="annot" style="padding:4px 8px">起点 <b>${anchor}</b>，向${L.label}复制 ${N} 格：${anchor} → ${targetAddrs.join(' → ')}</div></div></div>`;
  },
  afterSettings(){
    if(this.currentSub === 'func') return;
    document.querySelectorAll('#settingsRow .xl-dir-btn').forEach(b=>{
      b.onclick = ()=>{ this._dir = b.dataset.dir; Engine.reload(); };
    });
    const sel = $('#excel-steps');
    if(sel) sel.onchange = ()=>{ this._steps = +sel.value; Engine.reload(); };
  },

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

  /* 读取用户选择的方向与步数，归一化到合法值 */
  _refLayout(){
    const dir = ['right','down','left','up'].includes(this._dir) ? this._dir : 'right';
    const N = [1,2,3].includes(Number(this._steps)) ? Number(this._steps) : 2;
    return { dir, N, L: DIR_LAYOUT[dir] };
  },
  buildRef(colLock, col, rowLock, row){
    return (colLock ? '$' : '') + col + (rowLock ? '$' : '') + row;
  },

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

  /* 单个目标单元格的变化说明（desc=步骤文字，note=汇总表说明） */
  describeTarget(srcRef, dc, dr, addr, k, L){
    const p = this.parseRef(srcRef);
    const newCol = p.colLock ? p.col : this.colLetter(p.col.charCodeAt(0)-64 + dc);
    const newRow = p.rowLock ? p.row : String(Number(p.row) + dr);
    const newRef = (p.colLock ? '$' : '') + newCol + (p.rowLock ? '$' : '') + newRow;
    let colPart, rowPart, noteCol, noteRow;
    if(p.colLock){
      colPart = `列 $${p.col} 被 $ 锁定，保持不变（绿色）`;
      noteCol = `列 $${p.col} 锁定不动`;
    }else if(dc === 0){
      colPart = `同一列内复制，列标 ${p.col} 未偏移（灰色）`;
      noteCol = `同列复制，列标 ${p.col} 未偏移`;
    }else{
      colPart = `列偏移 ${dc>0?'+':''}${dc}：${p.col}→${newCol}（橙色变化）`;
      noteCol = `列 ${p.col}→${newCol}（橙色变化）`;
    }
    if(p.rowLock){
      rowPart = `行 $${p.row} 被 $ 锁定，保持不变（绿色）`;
      noteRow = `行 $${p.row} 锁定不动`;
    }else if(dr === 0){
      rowPart = `同一行内复制，行号 ${p.row} 未偏移（灰色）`;
      noteRow = `同行复制，行号 ${p.row} 未偏移`;
    }else{
      rowPart = `行偏移 ${dr>0?'+':''}${dr}：${p.row}→${newRow}（橙色变化）`;
      noteRow = `行 ${p.row}→${newRow}（橙色变化）`;
    }
    const desc = `向${L.label}复制 ${k} 格到 <b class="c-op">${addr}</b>：${colPart}；${rowPart}，公式变为 <b class="c-op">=${newRef}</b>。`;
    return { desc, note: `${noteCol}；${noteRow}`, newRef };
  },

  /* ---------- 引用复制舞台 ---------- */
  drawRefStage(o){
    const { cfg, srcRef, anchor, L, N, map, visible, phase, targets } = o;
    // 根据复制方向动态裁剪网格范围（只显示涉及的行列，避免大片空行空列）
    const colMin = Math.min(L.refCol,L.anchorCol,L.tCol(N));
    const colMax = Math.max(L.refCol,L.anchorCol,L.tCol(N));
    const rowMin = Math.min(L.refRow,L.anchorRow,L.tRow(N));
    const rowMax = Math.max(L.refRow,L.anchorRow,L.tRow(N));
    const cols=[];
    for(let cc=colMin;cc<=colMax;cc++) cols.push(this.colLetter(cc));
    const targetList = targets.map(t=>t.addr).join('、');
    let html = `<div class="excel-stage">
      <div class="stage-hd"><span>${cfg.title} · 公式复制</span><span class="state">起点 ${anchor} · 向${L.label}复制 ${N} 格：${targetList}</span></div>
      <div class="step-title"><span class="st-dot"></span>${cfg.subtitle}</div>
      <div class="xl-grid" style="grid-template-columns:38px repeat(${cols.length},1fr)">
        <div class="xl-cell hd"></div>${cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('')}`;
    for(let r=rowMin; r<=rowMax; r++){
      html += `<div class="xl-cell hd">${r}</div>`;
      for(let c=colMin; c<=colMax; c++){
        const cell = this.colLetter(c) + r;
        let cls = '', content = '<span style="color:#E2E8F0">·</span>';
        if(cell === anchor){
          cls = 'sel'; content = '=' + this.renderRefSource(srcRef);
        }else if(map[cell]){
          const t = map[cell];
          if(t.state === 'pending'){ cls = ''; content = '<span style="color:#CBD5E1">?</span>'; }
          else { cls = t.state === 'done' ? 'done' : 'cur'; content = '=' + this.renderRefTarget(srcRef, t.dc, t.dr); }
        }else{
          cls = 'empty-cell';
        }
        html += `<div class="xl-cell ${cls}">${content}</div>`;
      }
    }
    html += `</div>`;

    const curTarget = Object.values(map).find(t=>t.state === 'cur');
    html += `<div class="xl-formula">起点 ${anchor} =${this.renderRefSource(srcRef)}`;
    if(curTarget) html += ` &nbsp;→&nbsp; ${curTarget.addr} =${this.renderRefTarget(srcRef, curTarget.dc, curTarget.dr)}`;
    html += `</div>`;

    if(phase === 'source') html += this.analysisCard(cfg, srcRef, L);
    html += this.refSummaryTable(o, map, visible);
    html += `</div>`;
    $('#stage').innerHTML = html;
  },

  analysisCard(cfg, srcRef, L){
    const p = this.parseRef(srcRef);
    const horizontal = L.dc(1) !== 0;
    const rows = cfg.ana.map((a,i)=>{
      const isCol = i === 0;
      const val = isCol ? p.col : p.row;
      const token = a.lock
        ? `<span class="ref-lock">$${val}${this.lockSVG()}</span>`
        : `<span class="ref-move">${val}</span>`;
      let note;
      if(a.lock){
        note = `<span class="ana-note" style="color:#15803D">${isCol ? '$ 锁定 → 复制后列标不动' : '$ 锁定 → 复制后行号不动'}</span>`;
      }else if(isCol){
        note = horizontal
          ? `<span class="ana-note" style="color:#B45309">相对 → 向${L.label}复制时列标变</span>`
          : `<span class="ana-note" style="color:var(--text-3)">同列复制，列标不偏移</span>`;
      }else{
        note = horizontal
          ? `<span class="ana-note" style="color:var(--text-3)">同行复制，行号不偏移</span>`
          : `<span class="ana-note" style="color:#B45309">相对 → 向${L.label}复制时行号变</span>`;
      }
      return `<div class="ana-row"><span class="ana-label">${a.label}</span>${token}${note}</div>`;
    }).join('');
    return `<div class="ana-card">
      <div style="font-size:clamp(12px,1.5vw,14px);font-weight:700;color:var(--text-3);margin-bottom:6px">引用类型拆解</div>${rows}
    </div>`;
  },

  refSummaryTable(o, map, visible){
    const { srcRef, anchor, L } = o;
    let rows = `<tr><td>${anchor}</td><td>原公式（起点）</td><td>=${this.renderRefSource(srcRef)}</td><td class="note">复制的起点单元格</td></tr>`;
    Object.values(map).slice(0, visible).forEach(t=>{
      rows += `<tr><td>${t.addr}</td><td>向${L.label}复制 ${t.k} 格</td><td>=${this.renderRefTarget(srcRef, t.dc, t.dr)}</td><td class="note">${t.note}</td></tr>`;
    });
    return `<table class="ref-tbl">
      <colgroup><col style="width:14%"><col style="width:22%"><col style="width:22%"><col style="width:42%"></colgroup>
      <thead><tr><th>单元格</th><th>复制方向</th><th>复制后公式</th><th>变化说明</th></tr></thead>
      <tbody>${rows}</tbody></table>`;
  },

  analysisDesc(anchor, srcRef, L){
    return `先拆解 <b class="c-op">${anchor}</b> 中公式 =${srcRef} 的列标、行号两个部分，判断谁会在本次向${L.label}复制时移动、谁保持不动。`;
  },

  finalDesc(cfg, anchor, srcRef, targets, L){
    const p = this.parseRef(srcRef);
    const parts = [];
    if(p.colLock) parts.push('列标被 $ 锁定，始终不变');
    else if(L.dc(1) === 0) parts.push('列标同列复制，不偏移');
    else parts.push('列标随目标位置横向偏移');
    if(p.rowLock) parts.push('行号被 $ 锁定，始终不变');
    else if(L.dr(1) === 0) parts.push('行号同行复制，不偏移');
    else parts.push('行号随目标位置纵向偏移');
    const verify = targets.map(t=>`${t.addr}=${t.formula}`).join('、');
    return `${cfg.title}复制完成：${parts.join('；')}。验证：${verify}。`;
  },

  /* ---------- 单个引用类型：源公式 → 所选方向逐格复制 → 汇总 ---------- */
  genRefCase(type){
    const cfg = REF_CONFIG[type];
    const { N, L } = this._refLayout();
    const colLock = cfg.ana[0].lock, rowLock = cfg.ana[1].lock;
    const anchor = this.colLetter(L.anchorCol) + L.anchorRow;
    const srcRef = this.buildRef(colLock, this.colLetter(L.refCol), rowLock, String(L.refRow));
    const targets = [];
    for(let k=1; k<=N; k++){
      const addr = this.colLetter(L.tCol(k)) + L.tRow(k);
      const dc = L.dc(k), dr = L.dr(k);
      const meta = this.describeTarget(srcRef, dc, dr, addr, k, L);
      targets.push({ addr, dc, dr, k, formula:this.offsetRef(srcRef, dc, dr), desc:meta.desc, note:meta.note });
    }

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
      desc: cfg.intro + ' ' + this.analysisDesc(anchor, srcRef, L),
      draw: ()=>{ this.drawRefStage({ cfg, srcRef, anchor, L, N, targets, map:snap(-1), visible:0, phase:'source' }); }
    });
    /* 逐个方向复制 */
    targets.forEach((t,i)=>{
      steps.push({
        desc: t.desc,
        draw: ()=>{ this.drawRefStage({ cfg, srcRef, anchor, L, N, targets, map:snap(i), visible:i+1, phase:'copy' }); }
      });
    });
    /* 汇总完成 */
    steps.push({
      desc: this.finalDesc(cfg, anchor, srcRef, targets, L),
      draw: ()=>{ this.drawRefStage({ cfg, srcRef, anchor, L, N, targets, map:snap(targets.length), visible:targets.length, phase:'done' }); }
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
    const srcCells = judged || FUNC_CELLS;
    srcCells.forEach(c=>{ byAddr[c.addr] = c; });
    // 仅显示有数据的行（A1:D4 中数据占第 1 行和 C2），避免空行留白
    let maxRow = 2;
    srcCells.forEach(c=>{ if(c.type!=='empty'){ const rr=parseInt(c.addr.slice(1),10); if(rr>maxRow) maxRow=rr; } });
    let html = `<div class="xl-grid" style="grid-template-columns:38px repeat(4,1fr)">
      <div class="xl-cell hd"></div>${cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('')}`;
    for(let r=1; r<=maxRow; r++){
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
        <span class="judge-reason">${x.verdict.reason}</span>
      </div>`).join('');
    let html = `<div class="excel-stage">
      <div class="stage-hd"><span>${f.name} ${f.title}</span><span class="state">数据区域 A1:D4</span></div>
      <div class="func-split">
        <div class="func-left">
          ${this.funcGrid(judged)}
          <div class="xl-formula">=${f.name}(A1:D4) = <b style="color:#86EFAC">${this.fmtNum(res)}</b></div>
          <div class="write-result">${f.name}(A1:D4) = ${this.fmtNum(res)}</div>
        </div>
        <div class="func-right">
          <div class="cp-label">逐个判断（绿色 = 计入，灰色删除线 = 忽略）</div>
          ${judgeRows}
          <div class="judge-row">
            <span class="judge-addr">空单元格</span>
            <span class="judge-val">∅</span>
            <span class="type-tag empty">空单元格</span>
            <span class="chip-no">忽略</span>
            <span class="judge-reason">${emptyAddrs.join('、')}：无内容，所有函数都忽略</span>
          </div>
          <div class="cp-sum-line">${f.process(counted)}</div>
        </div>
      </div>
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
    let hint = '点击「自动演示」或「下一步」开始推导';
    if(sub !== 'func'){
      const { N, L } = this._refLayout();
      const anchor = this.colLetter(L.anchorCol) + L.anchorRow;
      hint = `在上方选择复制方向与步数（当前：向${L.label} ${N} 格，起点 ${anchor}），点击「自动演示」开始推导`;
    }
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub]||'Excel 公式'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:#22C55E">${info.method||''}</b></div>
        <div style="font-size:clamp(13px,1.6vw,15px)">${hint}</div>
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
