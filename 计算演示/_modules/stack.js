/* ================================================================
 * ⑧ 栈 Stack（push 入栈 · pop 出栈 · LIFO 后进先出）
 * 子页面：push 连续入栈 / pop 先压后弹 / mixed 混合操作
 * ================================================================ */
MOD.stack = {
  name:'栈 Stack', sub:'入栈 push · 出栈 pop · 后进先出 LIFO',
  badge:'结构 · LIFO', tag:'栈', cat:'ds',
  color:'#8B5CF6', soft:'#F3EEFF', type:'栈 · 入栈 · 出栈 · 栈顶 · 后进先出',
  desc:'可视化演示栈的 push（入栈）与 pop（出栈）操作：栈只能在栈顶一端进出，连续入栈观察栈增长，先压后弹观察后进先出（LIFO）。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="14" width="14" height="6" rx="1.5"/><rect x="5" y="8" width="14" height="6" rx="1.5"/><rect x="5" y="2" width="14" height="6" rx="1.5"/><path d="M12 20v2M12 22l-2-2M12 22l2-2"/></svg>',
  setDesc:'',
  stageDesc:'栈底在下、栈顶在上；橙色=当前操作元素，绿色=已在栈中。',
  initDesc:'点击「自动演示」或「下一步」，观看栈的入栈与出栈过程。',
  currentSub:'push',
  subTabs:[
    {id:'push',  label:'入栈', method:'连续 push · 栈增长'},
    {id:'pop',   label:'出栈', method:'先压后弹 · LIFO'},
    {id:'mixed', label:'混合', method:'push / pop 交替'}
  ],
  theory(){
    return `<div class="rule"><b>栈（Stack）</b>：一种线性数据结构，只能在一端（栈顶）进行插入和删除，另一端为栈底。</div>
    <div class="rule"><b>后进先出 LIFO</b>：Last In First Out，最后入栈的元素最先出栈。</div>
    <div class="rule"><b>push（入栈）</b>：在栈顶添加元素，栈大小 +1。</div>
    <div class="rule"><b>pop（出栈）</b>：从栈顶移除元素并返回，栈大小 -1；栈空时执行 pop 会发生下溢错误。</div>
    <div class="rule">应用：函数调用栈、表达式求值、括号匹配、浏览器后退按钮。</div>`;
  },
  settingsHTML(){
    const cur = $('#st-n') ? $('#st-n').value : '5';
    return `<div class="inline-setting"><div class="setting"><label>操作步数</label>
      <select id="st-n">
        <option value="4" ${cur==='4'?'selected':''}>4 步</option>
        <option value="5" ${cur==='5'?'selected':''}>5 步</option>
        <option value="6" ${cur==='6'?'selected':''}>6 步</option>
      </select></div></div>`;
  },
  afterSettings(){ $('#st-n').onchange=()=>Engine.reload(); },

  /* 逐字出现动画辅助 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },

  /* 统一画布渲染：o = {title, state, stack, mode, curVal, popped, log, expr, result, resultLabel} */
  renderStack(o){
    const subName={push:'入栈演示',pop:'出栈演示',mixed:'混合操作'}[this.currentSub]||'栈 Stack';
    const size=o.stack.length;

    /* 栈内方块：DOM 顺序 栈底 → 栈顶（column-reverse 渲染为底在下、顶在上） */
    const chips=o.stack.map((v,i)=>{
      const isTop=i===size-1;
      const isBottom=i===0;
      const isCur=o.mode==='push' && isTop && v===o.curVal;
      const cls=isCur?'o':'g';
      const anim=isCur?' write-in':'';
      const topTag=isTop?`<span style="position:absolute;right:-34px;top:50%;transform:translateY(-50%);font-size:9.5px;color:#8B5CF6;font-weight:700;white-space:nowrap">栈顶</span>`:'';
      const botTag=isBottom&&size>1?`<span style="position:absolute;left:-34px;top:50%;transform:translateY(-50%);font-size:9.5px;color:#94A3B8;font-weight:700;white-space:nowrap">栈底</span>`:'';
      return `<div class="chip ${cls}${anim}" style="position:relative;width:clamp(38px,5.5vw,46px);height:clamp(20px,2.8vw,24px);font-size:12px;${isCur?'animation-delay:.1s':''}">${esc(v)}${topTag}${botTag}</div>`;
    }).join('');
    const bodyHtml=size===0
      ? '<div style="color:#CBD5E1;font-size:12px;font-family:var(--mono)">（空栈）</div>'
      : chips;

    /* 出栈“弹出”幽灵方块（栈顶高亮后向上离开） */
    const ghost=o.mode==='pop'&&o.popped!==null?`
      <div style="display:flex;justify-content:center;align-items:center;gap:8px;min-height:30px;margin-bottom:2px">
        <span style="color:#F59E0B;font-size:11px;font-family:var(--mono);font-weight:700">弹出 ↑</span>
        <div class="chip o write-in" style="position:relative;width:clamp(38px,5.5vw,46px);height:clamp(20px,2.8vw,24px);font-size:12px">${esc(o.popped)}</div>
      </div>`:'';

    /* 操作表达式（逐字写入 + 笔形光标） */
    let exprHtml='';
    if(o.expr){
      if(o.expr.kind==='push'){
        exprHtml=`<span class="term cur">${this.typeChars(`push(${o.expr.v})`)}</span><span class="op">=</span><span class="term">栈顶 ← ${o.expr.v}</span><span class="op">·</span><span class="term done">大小 ${o.expr.prev} → ${o.expr.size}</span><span class="caret pen"></span>`;
      }else{
        exprHtml=`<span class="term cur">${this.typeChars('pop()')}</span><span class="op">=</span><span class="term cur">返回 ${o.expr.v}</span><span class="op">·</span><span class="term done">大小 ${o.expr.before} → ${o.expr.size}</span><span class="caret pen"></span>`;
      }
    }

    /* 完成态分组框 + 结果框 */
    let resultHtml='';
    if(o.result){
      const r=o.result;
      const grps=r.map((v,i)=>{
        const lbl=r.length===1?'栈顶':i===0?'栈底':i===r.length-1?'栈顶':'';
        return `<span class="grp done"><span class="grp-label">${lbl}</span>${esc(v)}</span>`;
      }).join(r.length>1?'<span class="arrow">→</span>':'');
      resultHtml=`<div class="group-box" style="justify-content:center;margin:4px 0">${grps}</div>
        <div class="write-result" style="margin-top:3px;padding:4px 10px">${o.resultLabel||'栈底 → 栈顶'}：${r.length?r.join(' , '):'（空栈）'} · 共 ${r.length} 个元素</div>`;
    }

    $('#stage').innerHTML=`<div class="stage-hd"><span>栈 Stack · ${subName}</span><span class="state">${o.state||`大小 ${size}`}</span></div>
      <div class="step-title"><span class="st-dot"></span>${o.title||''}</div>
      ${ghost}
      <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:stretch">
        <div style="flex:1 1 140px;min-width:130px">
          <div class="zone-label" style="margin-bottom:4px">栈（栈顶在上 · 栈底在下）</div>
          <div style="display:flex;flex-direction:column-reverse;gap:3px;align-items:center;background:#fff;border:1px solid var(--border);border-radius:10px;padding:8px 10px;min-height:118px;justify-content:center">${bodyHtml}</div>
        </div>
        <div style="flex:1 1 140px;min-width:130px;display:flex;flex-direction:column;gap:5px">
          <div class="zone-label" style="margin-bottom:0">操作日志</div>
          <div class="ds-log" style="margin-top:0;line-height:1.6">${o.log||''}</div>
          ${exprHtml?`<div class="exp-inline" style="margin-top:2px">${exprHtml}</div>`:''}
        </div>
      </div>
      ${resultHtml}`;
  },

  /* ========== 子页面 1：连续入栈 push，观察栈增长 ========== */
  genPush(){
    const n=+$('#st-n').value||5;
    const pool=[1,2,3,4,5,6,7,8,9];
    for(let i=pool.length-1;i>0;i--){ const j=randInt(0,i); const t=pool[i]; pool[i]=pool[j]; pool[j]=t; }
    const vals=pool.slice(0,n);
    const steps=[];
    steps.push({desc:`初始状态：<b class="c-op">空栈</b>，栈顶与栈底重合，元素只能从<b>栈顶</b>压入（push）。`,
      draw:()=>{ this.renderStack({title:'入栈 push', state:'空栈 · 准备就绪', stack:[], mode:'ready', curVal:null, popped:null, log:'（尚未开始操作）', expr:null}); }});
    const stack=[];
    vals.forEach((v,k)=>{
      stack.push(v);
      const snap=[...stack];
      const prev=snap.length-1;
      steps.push({desc:`第 ${k+1} 步：<b class="c-op">push(${v})</b>，将 ${v} 压入栈顶；栈大小 <b>${prev}</b> → <b>${snap.length}</b>（+1），栈逐步向上增长。`,
        draw:()=>{ this.renderStack({title:'入栈 push', state:`大小 ${snap.length}`, stack:snap, mode:'push', curVal:v, popped:null,
          log:`<b>push(${v})</b>：栈顶新增 ${v}<br>栈大小 ${prev} → ${snap.length}（+1）`,
          expr:{kind:'push', v:v, prev:prev, size:snap.length}}); }});
    });
    const finalSnap=[...stack];
    steps.push({desc:`连续压入 ${n} 个元素后，栈中共 <b class="c-op">${finalSnap.length}</b> 个元素；若此时出栈，弹出顺序为 <b>${[...finalSnap].reverse().join('、')}</b>，与压入顺序相反（LIFO）。`,
      draw:()=>{ this.renderStack({title:'入栈完成', state:`共 ${finalSnap.length} 个`, stack:finalSnap, mode:'done', curVal:null, popped:null,
        log:`连续入栈完毕，栈底 → 栈顶：${finalSnap.join('、')}`, expr:null, result:finalSnap, resultLabel:'栈底 → 栈顶'}); }});
    return steps;
  },

  /* ========== 子页面 2：先压入再弹出，观察 LIFO ========== */
  genPop(){
    const n=+$('#st-n').value||5;
    const pushCount=3;
    const popCount=Math.max(1,n-pushCount);
    const pool=[1,2,3,4,5,6,7,8,9];
    for(let i=pool.length-1;i>0;i--){ const j=randInt(0,i); const t=pool[i]; pool[i]=pool[j]; pool[j]=t; }
    const pushVals=pool.slice(0,pushCount);
    const steps=[];
    steps.push({desc:`初始状态：<b class="c-op">空栈</b>。先连续压入几个元素，再逐个弹出，观察出栈顺序是否与压入顺序相反。`,
      draw:()=>{ this.renderStack({title:'出栈 pop', state:'空栈 · 准备就绪', stack:[], mode:'ready', curVal:null, popped:null, log:'（尚未开始操作）', expr:null}); }});
    const stack=[];
    for(let k=0;k<pushCount;k++){
      const v=pushVals[k];
      stack.push(v);
      const snap=[...stack];
      const prev=snap.length-1;
      steps.push({desc:`准备阶段 · 第 ${k+1} 步：<b class="c-op">push(${v})</b>，将 ${v} 压入栈顶；栈大小 <b>${prev}</b> → <b>${snap.length}</b>（+1）。`,
        draw:()=>{ this.renderStack({title:'先压入 · push', state:`大小 ${snap.length}`, stack:snap, mode:'push', curVal:v, popped:null,
          log:`<b>push(${v})</b>：栈顶新增 ${v}<br>栈大小 ${prev} → ${snap.length}（+1）`,
          expr:{kind:'push', v:v, prev:prev, size:snap.length}}); }});
    }
    const poppedSeq=[];
    for(let k=0;k<popCount;k++){
      const v=stack.pop();
      poppedSeq.push(v);
      const snap=[...stack];
      const before=snap.length+1;
      steps.push({desc:`第 ${pushCount+k+1} 步：<b class="c-op">pop()</b>，弹出栈顶元素 <b class="c-op">${v}</b>——后进先出，最后压入的最先弹出；栈大小 <b>${before}</b> → <b>${snap.length}</b>（-1）。`,
        draw:()=>{ this.renderStack({title:'出栈 pop', state:`大小 ${snap.length}`, stack:snap, mode:'pop', curVal:null, popped:v,
          log:`<b>pop()</b>：返回栈顶元素 ${v}<br>栈大小 ${before} → ${snap.length}（-1）`,
          expr:{kind:'pop', v:v, before:before, size:snap.length}}); }});
    }
    const finalSnap=[...stack];
    steps.push({desc:`出栈顺序：<b>${poppedSeq.join('、')}</b>，与压入顺序（${pushVals.join('、')}）正好相反，体现 <b class="c-op">LIFO 后进先出</b>；栈中剩余 <b>${finalSnap.length}</b> 个元素。`,
      draw:()=>{ this.renderStack({title:'出栈完成', state:`共 ${finalSnap.length} 个`, stack:finalSnap, mode:'done', curVal:null, popped:null,
        log:`出栈演示完毕，弹出序列：${poppedSeq.join('、')}`, expr:null, result:finalSnap, resultLabel:'剩余栈 · 栈底 → 栈顶'}); }});
    return steps;
  },

  /* ========== 子页面 3：push / pop 交替混合操作 ========== */
  genMixed(){
    const n=+$('#st-n').value||5;
    const ops=[]; let size=0;
    for(let i=0;i<n;i++){
      if(size===0){ ops.push({op:'push',val:randInt(1,9)}); size++; }
      else if(size>=5){ ops.push({op:'pop'}); size--; }
      else if(size===1 && Math.random()<0.7){ ops.push({op:'push',val:randInt(1,9)}); size++; }
      else if(Math.random()<0.5){ ops.push({op:'push',val:randInt(1,9)}); size++; }
      else { ops.push({op:'pop'}); size--; }
    }
    /* 回放得到每一步快照，保证画面与描述一致 */
    const stack=[]; const replay=[];
    ops.forEach(o=>{
      if(o.op==='push'){ stack.push(o.val); replay.push({op:'push',val:o.val,snap:[...stack]}); }
      else { const v=stack.pop(); replay.push({op:'pop',val:v,snap:[...stack]}); }
    });
    const steps=[];
    steps.push({desc:`初始状态：<b class="c-op">空栈</b>。下面交替执行 push 与 pop，观察栈顶元素的动态变化。`,
      draw:()=>{ this.renderStack({title:'混合操作', state:'空栈 · 准备就绪', stack:[], mode:'ready', curVal:null, popped:null, log:'（尚未开始操作）', expr:null}); }});
    replay.forEach((r,k)=>{
      if(r.op==='push'){
        const prev=r.snap.length-1;
        steps.push({desc:`第 ${k+1} 步：<b class="c-op">push(${r.val})</b>，将 ${r.val} 压入栈顶；栈大小 <b>${prev}</b> → <b>${r.snap.length}</b>（+1）。`,
          draw:()=>{ this.renderStack({title:'入栈 push', state:`大小 ${r.snap.length}`, stack:r.snap, mode:'push', curVal:r.val, popped:null,
            log:`<b>push(${r.val})</b>：栈顶新增 ${r.val}<br>栈大小 ${prev} → ${r.snap.length}（+1）`,
            expr:{kind:'push', v:r.val, prev:prev, size:r.snap.length}}); }});
      }else{
        const before=r.snap.length+1;
        steps.push({desc:`第 ${k+1} 步：<b class="c-op">pop()</b>，弹出栈顶元素 <b class="c-op">${r.val}</b>——后进先出，最后压入的最先弹出；栈大小 <b>${before}</b> → <b>${r.snap.length}</b>（-1）。`,
          draw:()=>{ this.renderStack({title:'出栈 pop', state:`大小 ${r.snap.length}`, stack:r.snap, mode:'pop', curVal:null, popped:r.val,
            log:`<b>pop()</b>：返回栈顶元素 ${r.val}<br>栈大小 ${before} → ${r.snap.length}（-1）`,
            expr:{kind:'pop', v:r.val, before:before, size:r.snap.length}}); }});
      }
    });
    const finalSnap=replay.length?replay[replay.length-1].snap:[];
    const poppedAll=replay.filter(r=>r.op==='pop').map(r=>r.val);
    steps.push({desc:`混合操作结束：栈中剩余 <b class="c-op">${finalSnap.length}</b> 个元素${finalSnap.length?`（${finalSnap.join('、')}）`:'（空栈）'}；本次弹出序列：${poppedAll.length?`<b>${poppedAll.join('、')}</b>`:'无'}。`,
      draw:()=>{ this.renderStack({title:'混合操作完成', state:`共 ${finalSnap.length} 个`, stack:finalSnap, mode:'done', curVal:null, popped:null,
        log:`混合操作执行完毕，剩余 ${finalSnap.length} 个元素`, expr:null, result:finalSnap, resultLabel:'剩余栈 · 栈底 → 栈顶'}); }});
    return steps;
  },

  /* 切换子页面时清空画布，显示空栈与准备状态 */
  initStage(){
    const map={
      push:  {label:'入栈演示 · 连续 push，观察栈增长', title:'入栈 push'},
      pop:   {label:'出栈演示 · 先压入再弹出，观察 LIFO', title:'出栈 pop'},
      mixed: {label:'混合操作 · push / pop 交替执行', title:'混合操作'}
    };
    const m=map[this.currentSub]||map.push;
    $('#stage').innerHTML=`<div class="stage-hd"><span>栈 Stack</span><span class="state">空栈 · 准备就绪</span></div>
      <div class="step-title"><span class="st-dot"></span>${m.title}</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:stretch">
        <div style="flex:1 1 140px;min-width:130px">
          <div class="zone-label" style="margin-bottom:4px">栈（栈顶在上 · 栈底在下）</div>
          <div style="display:flex;flex-direction:column-reverse;gap:3px;align-items:center;justify-content:center;background:#fff;border:1px solid var(--border);border-radius:10px;padding:8px 10px;min-height:118px">
            <div style="color:#CBD5E1;font-size:12px;font-family:var(--mono)">（空栈）</div>
          </div>
        </div>
        <div style="flex:1 1 140px;min-width:130px;display:flex;flex-direction:column;gap:5px">
          <div class="zone-label" style="margin-bottom:0">操作日志</div>
          <div class="ds-log" style="margin-top:0;line-height:1.6">（尚未开始操作）</div>
          <div style="margin-top:4px;color:#94A3B8;font-size:clamp(10.5px,1.3vw,12px);line-height:1.6">当前模式：<b style="color:#8B5CF6">${m.label}</b>，点击「下一步」开始演示</div>
        </div>
      </div>`;
  },

  generate(){
    switch(this.currentSub){
      case 'push':  this.steps=this.genPush(); break;
      case 'pop':   this.steps=this.genPop(); break;
      case 'mixed': this.steps=this.genMixed(); break;
      default:      this.steps=this.genPush();
    }
  }
};
