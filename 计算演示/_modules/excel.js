/* ================================================================
 * Excel 公式
 * 子页面：ref_rel 相对引用 / ref_abs 绝对引用 / ref_mix 混合引用 / func 统计函数
 * 固定复制路径：B2 → D4（行偏移 +2，列偏移 +2）
 * ================================================================ */
MOD.excel = {
  name:'Excel公式', sub:'相对 · 绝对 · 混合引用复制 · 统计函数',
  badge:'表格 · 引用', tag:'表格', cat:'net',
  color:'#22C55E', soft:'#E8FBF0', type:'相对/绝对/混合引用复制 · SUM/AVERAGE/MAX/MIN',
  desc:'演示相对、绝对、混合引用在公式复制时的行列变化，并逐步计算 SUM/AVERAGE/MAX/MIN 统计函数。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg>',
  setDesc:'引用类型固定由 B2 复制到 D4；统计函数可随机或手动选择。',
  stageDesc:'橙色=源单元格/当前计算，绿色=结果，观察复制时引用地址的偏移规律。',
  initDesc:'点击「自动演示」，观看公式引用变化或函数逐步计算。',
  currentSub:'ref_rel',
  subTabs:[
    {id:'ref_rel', label:'相对引用', method:'行列都偏移'},
    {id:'ref_abs', label:'绝对引用', method:'行列都锁定'},
    {id:'ref_mix', label:'混合引用', method:'锁列或锁行'},
    {id:'func',    label:'统计函数', method:'逐步求值'}
  ],
  theory(){
    return `<div class="rule"><b>相对引用</b>（如 A1）：复制时行号和列号都会随目标位置偏移，向右复制列 +1，向下复制行 +1。</div>
    <div class="rule"><b>绝对引用</b>（如 $A$1）：复制时行号和列号都<b>不变</b>，$ 同时锁定行和列。</div>
    <div class="rule"><b>混合引用 $A1</b>：列绝对、行相对——列锁死不变，行随目标位置偏移。</div>
    <div class="rule"><b>混合引用 A$1</b>：列相对、行绝对——行锁死不变，列随目标位置偏移。</div>
    <div class="rule"><b>移动</b>公式（剪切粘贴）：无论哪种引用，引用地址都不变；只有<b>复制</b>时相对部分才会偏移。</div>
    <div class="rule"><b>常用函数</b>：SUM 求和、AVERAGE 平均值、MAX 最大值、MIN 最小值。</div>`;
  },
  /* ---------- 参数区 ---------- */
  settingsHTML(){
    if(this.currentSub === 'func'){
      const old = $('#xl-func');
      const cur = old ? old.value : (this._func || 'auto');
      return `<div class="inline-setting"><div class="setting"><label>统计函数</label>
        <select id="xl-func">
          <option value="auto" ${cur==='auto'?'selected':''}>自动随机</option>
          <option value="SUM" ${cur==='SUM'?'selected':''}>SUM 求和</option>
          <option value="AVERAGE" ${cur==='AVERAGE'?'selected':''}>AVERAGE 平均值</option>
          <option value="MAX" ${cur==='MAX'?'selected':''}>MAX 最大值</option>
          <option value="MIN" ${cur==='MIN'?'selected':''}>MIN 最小值</option>
        </select></div></div>`;
    }
    return `<div class="inline-setting"><div class="setting"><label>复制路径</label>
      <div class="annot" style="padding:4px 8px">固定由 <b>B2</b> 复制到 <b>D4</b>（行 +2、列 +2）</div></div></div>`;
  },
  afterSettings(){
    const el = $('#xl-func');
    if(el) el.onchange = ()=>{ this._func = el.value; Engine.reload(); };
  },
  colLetter(n){ return String.fromCharCode(64 + n); },
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.05}s">${esc(ch)}</span>`).join('');
  },

  /* ---------- 引用复制：5×5 网格 ---------- */
  drawRefGrid(srcCell, dstCell, srcFormula, dstFormula, highlight, done, note){
    const cols = ['A','B','C','D','E'];
    let html = `<div class="stage-hd"><span>引用复制演示</span><span class="state">${srcCell} → ${dstCell}</span></div>
      <div class="step-title"><span class="st-dot"></span>公式复制 · 观察引用地址变化</div>`;
    html += `<div class="xl-grid" style="grid-template-columns:34px repeat(5,1fr)">`;
    html += `<div class="xl-cell hd"></div>` + cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('');
    for(let r=1; r<=5; r++){
      html += `<div class="xl-cell hd">${r}</div>`;
      for(let c=1; c<=5; c++){
        const cell = this.colLetter(c) + r;
        let cls = '', content = '';
        if(cell === srcCell){ cls = 'sel'; content = '=' + this.typeChars(srcFormula); }
        else if(cell === dstCell){ cls = done ? 'done' : 'cur'; content = dstFormula ? '=' + this.typeChars(dstFormula) : '?'; }
        html += `<div class="xl-cell ${cls}">${content}</div>`;
      }
    }
    html += `</div>`;
    html += `<div class="xl-formula">源 ${srcCell}: =${this.typeChars(srcFormula)} ${dstFormula ? `→ 目标 ${dstCell}: =<span class="write-in">${this.typeChars(dstFormula)}</span>` : ''}</div>`;
    if(note) html += `<div class="annot" style="text-align:center;margin-top:2px">${note}</div>`;
    $('#stage').innerHTML = html;
  },

  /* ---------- 单个引用类型的复制演示 ---------- */
  genRefCase(type, intro){
    const srcR = 2, srcC = 2;   // B2
    const dstR = 4, dstC = 4;   // D4
    const offR = dstR - srcR;   // +2 行
    const offC = dstC - srcC;   // +2 列
    const srcCell = this.colLetter(srcC) + srcR;
    const dstCell = this.colLetter(dstC) + dstR;
    let ref, dstRef, rule;
    if(type === 'rel'){
      ref = 'A1';
      dstRef = this.colLetter(1 + offC) + (1 + offR);
      rule = `列 A + ${offC} = <b class="c-op">${this.colLetter(1 + offC)}</b>，行 1 + ${offR} = <b class="c-op">${1 + offR}</b>`;
    }else if(type === 'abs'){
      ref = '$A$1';
      dstRef = '$A$1';
      rule = `$ 同时锁定列 A、锁定行 1，复制后<b>都不变</b>，仍为 <b class="c-op">$A$1</b>`;
    }else if(type === 'mixcol'){
      ref = '$A1';
      dstRef = '$A' + (1 + offR);
      rule = `列 A 前有 $ → <b>不变</b>；行 1 无 $ → 1 + ${offR} = <b class="c-op">${1 + offR}</b>`;
    }else{
      ref = 'A$1';
      dstRef = this.colLetter(1 + offC) + '$1';
      rule = `列 A 无 $ → A + ${offC} = <b class="c-op">${this.colLetter(1 + offC)}</b>；行 1 前有 $ → <b>不变</b>`;
    }
    const steps = [];
    // 1. 源单元格公式
    steps.push({ desc:`源单元格 <b class="c-op">${srcCell}</b> 的公式 =<b class="c-op">${ref}</b>，将其<b>复制</b>到 ${dstCell}（行偏移 +${offR}，列偏移 +${offC}）。`,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, '', srcCell); }});
    // 2. 规则说明
    steps.push({ desc:intro,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, '', dstCell); }});
    // 3. 推导目标引用
    steps.push({ desc:rule,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, dstRef, dstCell); }});
    // 4. 复制完成
    steps.push({ desc:`复制完成：${srcCell} =${ref} → ${dstCell} =<b class="c-op">${dstRef}</b>`,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, dstRef, dstCell, true); }});
    return steps;
  },

  /* ========== ① ref_rel：相对引用 ========== */
  genRefRel(){
    return this.genRefCase('rel', '相对引用：列号、行号都随目标位置一起偏移，偏移量 = 目标 − 源。');
  },
  /* ========== ② ref_abs：绝对引用 ========== */
  genRefAbs(){
    return this.genRefCase('abs', '绝对引用：$ 同时锁定列和行，复制后列号、行号都保持不变。');
  },
  /* ========== ③ ref_mix：混合引用（两种类型同页对比） ========== */
  genRefMix(){
    const mixCol = this.genRefCase('mixcol', '第一种 混合引用 $A1：<b>列绝对</b>（列锁死）、行相对（行偏移）。');
    const mixRow = this.genRefCase('mixrow', '第二种 混合引用 A$1：列相对（列偏移）、<b>行绝对</b>（行锁死）。');
    return mixCol.concat(mixRow);
  },

  /* ---------- 统计函数：数据网格 ---------- */
  drawFuncGrid(data, curIdx, funcName, result, done){
    let html = `<div class="stage-hd"><span>${funcName} 函数</span><span class="state">A1:A${data.length}</span></div>
      <div class="step-title"><span class="st-dot"></span>${funcName} 逐步求值 · 数据区域 A1:A${data.length}</div>`;
    html += `<div class="xl-grid" style="grid-template-columns:34px repeat(${data.length},1fr)">`;
    html += `<div class="xl-cell hd"></div>` + data.map((_,i)=>`<div class="xl-cell hd">${String.fromCharCode(65+i)}</div>`).join('');
    html += `<div class="xl-cell hd">1</div>` + data.map((v,i)=>`<div class="xl-cell ${i<curIdx?'done':i===curIdx?'cur':''} ${i===curIdx?'write-in':''}">${v}</div>`).join('');
    html += `</div>`;
    html += `<div class="xl-formula">=${funcName}(A1:A${data.length})${result ? ` = <b style="color:#86EFAC" class="write-in">${result}</b>` : ''}</div>`;
    if(done) html += `<div class="write-result">${funcName}(A1:A${data.length}) = ${result}</div>`;
    $('#stage').innerHTML = html;
  },

  /* ========== ④ func：统计函数逐步求值 ========== */
  genFunc(){
    const el = $('#xl-func');
    let f = el ? el.value : (this._func || 'auto');
    if(!f || f === 'auto') f = pick(['SUM','AVERAGE','MAX','MIN']);
    this._func = f;
    const datasets = {
      SUM:[3,7,2,8,5],
      AVERAGE:[4,8,6,9,3],
      MAX:[5,12,3,9,7],
      MIN:[8,3,10,5,2]
    };
    const data = datasets[f];
    const range = 'A1:A' + data.length;
    const steps = [];
    // 1. 函数与数据区域
    steps.push({ desc:`函数 =<b class="c-op">${f}(${range})</b>，数据区域 ${range} = [${data.join(', ')}]。`,
      draw:()=>{ this.drawFuncGrid(data, -1, f, '', false); }});
    if(f === 'SUM'){
      let acc = 0;
      data.forEach((v,i)=>{
        acc += v;
        const snap = acc;
        steps.push({ desc:`第 ${i+1} 个值 <b class="c-op">${v}</b>：累计 ${snap - v} + ${v} = <b class="c-op">${snap}</b>`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false); }});
      });
      steps.push({ desc:`SUM 求和结果 = <b class="c-op">${acc}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(acc), true); }});
    }else if(f === 'AVERAGE'){
      const sum = data.reduce((a,b)=>a+b,0);
      const avg = sum / data.length;
      steps.push({ desc:`先求和：${data.join(' + ')} = <b class="c-op">${sum}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, sum + '（和）', false); }});
      steps.push({ desc:`再除以个数 ${data.length}：${sum} ÷ ${data.length} = <b class="c-op">${avg.toFixed(2)}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, avg.toFixed(2), true); }});
    }else if(f === 'MAX'){
      let mx = data[0];
      steps.push({ desc:`初始最大值 = 第一个值 <b class="c-op">${data[0]}</b>`,
        draw:()=>{ this.drawFuncGrid(data, 0, f, String(mx), false); }});
      for(let i=1; i<data.length; i++){
        const oldMx = mx;
        const bigger = data[i] > oldMx;
        if(bigger) mx = data[i];
        const snap = mx;
        steps.push({ desc:`比较 ${data[i]} 与当前最大 ${oldMx}：${bigger ? `<b class="c-op">${data[i]} 更大，更新最大值为 ${data[i]}</b>` : `${data[i]} 不更大，保持 ${oldMx}`}`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false); }});
      }
      steps.push({ desc:`MAX 最大值 = <b class="c-op">${mx}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(mx), true); }});
    }else{
      let mn = data[0];
      steps.push({ desc:`初始最小值 = 第一个值 <b class="c-op">${data[0]}</b>`,
        draw:()=>{ this.drawFuncGrid(data, 0, f, String(mn), false); }});
      for(let i=1; i<data.length; i++){
        const oldMn = mn;
        const smaller = data[i] < oldMn;
        if(smaller) mn = data[i];
        const snap = mn;
        steps.push({ desc:`比较 ${data[i]} 与当前最小 ${oldMn}：${smaller ? `<b class="c-op">${data[i]} 更小，更新最小值为 ${data[i]}</b>` : `${data[i]} 不更小，保持 ${oldMn}`}`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false); }});
      }
      steps.push({ desc:`MIN 最小值 = <b class="c-op">${mn}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(mn), true); }});
    }
    return steps;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const sub = this.currentSub;
    const info = this.subTabs.find(t=>t.id===sub) || {};
    const labelMap = { ref_rel:'相对引用复制', ref_abs:'绝对引用复制', ref_mix:'混合引用复制', func:'统计函数逐步求值' };
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
