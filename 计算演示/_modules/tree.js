/* ================================================================
 * 树 Tree（前序 · 中序 · 后序 · 层序遍历）
 * 按 base 样板重构：subTabs + settingsHTML/afterSettings + initStage
 *                 + generate/genXxx + 手写动画逐步演示
 * ================================================================ */
MOD.tree = {
  name:'树 Tree',
  sub:'前序 · 中序 · 后序 · 层序遍历',
  badge:'结构 · 遍历', tag:'树', cat:'ds',
  color:'#EC4899', soft:'#FDF0F7',
  type:'二叉树 · 前中后层序遍历',
  desc:'可视化演示二叉树的四种遍历方式：前序（根左右）、中序（左根右）、后序（左右根）、层序（逐层从左到右，借助队列）。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="4" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="12" r="2.5"/><circle cx="3" cy="20" r="2"/><circle cx="9" cy="20" r="2"/><circle cx="15" cy="20" r="2"/><circle cx="21" cy="20" r="2"/><path d="M12 6.5v3M10.5 11l-3 .5M13.5 11l3 .5M6 14.5v3M5 18l-1.5 1.5M7 18l1.5 1.5M18 14.5v3M17 18l-1.5 1.5M19 18l1.5 1.5"/></svg>',
  setDesc:'选择遍历方式，固定一棵 7 节点二叉树，逐步观看访问顺序。',
  stageDesc:'橙色=当前访问节点，绿色=已访问，蓝色=队列中待访问，灰色=未访问。',
  initDesc:'点击「自动演示」或「下一步」，观看二叉树的手写遍历过程。',
  currentSub:'pre',
  subTabs:[
    {id:'pre',  label:'前序', method:'根 → 左 → 右'},
    {id:'in',   label:'中序', method:'左 → 根 → 右'},
    {id:'post', label:'后序', method:'左 → 右 → 根'},
    {id:'level',label:'层序', method:'队列 · 逐层'}
  ],
  theory(){
    return `<div class="rule"><b>二叉树</b>：每个节点最多有两个子节点（左子节点、右子节点）。</div>
    <div class="rule"><b>前序遍历</b>：根 → 左 → 右。先访问根节点，再前序遍历左子树，最后前序遍历右子树。</div>
    <div class="rule"><b>中序遍历</b>：左 → 根 → 右。先中序遍历左子树，再访问根节点，最后中序遍历右子树。</div>
    <div class="rule"><b>后序遍历</b>：左 → 右 → 根。先后序遍历左子树，再后序遍历右子树，最后访问根节点。</div>
    <div class="rule"><b>层序遍历</b>：从上到下、从左到右逐层访问，需要借助<b>队列</b>（先进先出）实现。</div>`;
  },
  settingsHTML(){
    return `<div class="inline-setting"><div class="setting"><label>树结构</label>
      <span style="font-family:var(--mono);font-size:clamp(11px,1.4vw,12.5px);color:var(--text-2);padding:4px 9px;border:1px solid var(--border);border-radius:8px;background:#fff">7 节点满二叉树（固定）</span></div></div>`;
  },
  afterSettings(){ /* 无动态参数；切换子页面由 tab 触发 Engine.reload() */ },

  /* ---------- 固定二叉树数据 ---------- */
  treeNodes:[
    {id:1,x:50,y:14,label:'1',left:2,right:3},
    {id:2,x:25,y:42,label:'2',left:4,right:5},
    {id:3,x:75,y:42,label:'3',left:6,right:7},
    {id:4,x:12,y:70,label:'4',left:null,right:null},
    {id:5,x:38,y:70,label:'5',left:null,right:null},
    {id:6,x:62,y:70,label:'6',left:null,right:null},
    {id:7,x:88,y:70,label:'7',left:null,right:null}
  ],
  edges:[[1,2],[1,3],[2,4],[2,5],[3,6],[3,7]],

  node(id){ return this.treeNodes.find(n=>n.id===id); },
  parentOf(id){ const e=this.edges.find(([a,b])=>b===id); return e?e[0]:null; },
  subMeta(){
    return {
      pre:  {name:'前序遍历', rule:'根 → 左 → 右', kind:'pre'},
      in:   {name:'中序遍历', rule:'左 → 根 → 右', kind:'in'},
      post: {name:'后序遍历', rule:'左 → 右 → 根', kind:'post'},
      level:{name:'层序遍历', rule:'队列 · 逐层从左到右', kind:'level'}
    }[this.currentSub];
  },
  /* 逐字出现动画 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.05}s">${esc(ch)}</span>`).join('');
  },
  /* 节点角色描述（根 / 左子 / 右子 / 叶子） */
  roleText(id, side){
    if(id===1) return '根节点';
    const pid=this.parentOf(id), p=this.node(pid), n=this.node(id);
    const leaf=(n.left==null && n.right==null);
    return `${p.label} 的${side==='L'?'左':'右'}子节点${leaf?'（叶子）':''}`;
  },
  /* ---------- 递归遍历：由代码实际计算访问顺序 ---------- */
  genRecursiveEvents(kind){
    const visited=[], events=[];
    const walk=(id, parentId, side)=>{
      if(id==null) return;
      const n=this.node(id);
      if(kind==='pre'){ visited.push(id); events.push({curId:id,parentId,side,visited:[...visited]}); }
      walk(n.left, id, 'L');
      if(kind==='in'){ visited.push(id); events.push({curId:id,parentId,side,visited:[...visited]}); }
      walk(n.right, id, 'R');
      if(kind==='post'){ visited.push(id); events.push({curId:id,parentId,side,visited:[...visited]}); }
    };
    walk(1, null, null);
    return events;
  },
  /* 层序遍历：队列模拟（先进先出） */
  genLevelEvents(){
    const visited=[], queue=[1], events=[];
    while(queue.length){
      const cur=queue.shift();
      visited.push(cur);
      const n=this.node(cur), newly=[];
      if(n.left!=null){ queue.push(n.left); newly.push(n.left); }
      if(n.right!=null){ queue.push(n.right); newly.push(n.right); }
      events.push({curId:cur,parentId:this.parentOf(cur),side:null,visited:[...visited],queue:[...queue],newly});
    }
    return events;
  },
  /* ---------- 画布渲染 ---------- */
  renderStage(opts){
    const meta=this.subMeta();
    const visited=opts.visited||[], curId=opts.curId, done=!!opts.done;
    const queue=opts.queue||[], vis=new Set(visited);
    let html=`<div class="stage-hd" style="margin-bottom:1px;line-height:1.2;font-size:clamp(10.5px,1.3vw,12px)"><span>二叉树 · ${meta.name}</span><span class="state">${vis.size}/7</span></div>`;
    /* 步骤标题条 */
    html+=`<div class="step-title" style="margin-bottom:1px;padding:1px 8px;line-height:1.3"><span class="st-dot"></span>${done?'遍历完成':'二叉树结构 · 节点访问状态'}</div>`;
    /* SVG 二叉树 */
    html+=`<div class="svg-box" style="margin-bottom:1px;display:flex;justify-content:center">
      <svg viewBox="0 0 100 75" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:320px;height:110px">`;
    this.edges.forEach(([a,b])=>{
      const na=this.node(a), nb=this.node(b);
      const both=vis.has(a)&&vis.has(b);
      html+=`<line x1="${na.x}" y1="${na.y+4.2}" x2="${nb.x}" y2="${nb.y-4.2}"
        stroke="${both?'#22C55E':'#E2E8F0'}" stroke-width="0.7"/>`;
    });
    this.treeNodes.forEach(n=>{
      let state='un';
      if(curId===n.id) state='cur';
      else if(vis.has(n.id)) state='done';
      else if(queue.includes(n.id)) state='queued';
      const st={
        un:    {f:'#F1F5F9', s:'#CBD5E1', t:'#64748B'},
        cur:   {f:'#F59E0B', s:'#D97706', t:'#fff'},
        done:  {f:'#22C55E', s:'#16A34A', t:'#fff'},
        queued:{f:'#E0F2FE', s:'#38BDF8', t:'#0369A1'}
      }[state];
      if(state==='cur')
        html+=`<circle cx="${n.x}" cy="${n.y}" r="6.6" fill="#F59E0B" opacity="0.18" style="animation:pop .3s ease-out"/>`;
      html+=`<circle cx="${n.x}" cy="${n.y}" r="4.9" fill="${st.f}" stroke="${st.s}" stroke-width="0.8"/>`;
      html+=`<text x="${n.x}" y="${n.y+1.6}" text-anchor="middle" font-size="3.9" font-weight="700" fill="${st.t}">${esc(n.label)}</text>`;
    });
    html+=`</svg></div>`;
    /* 横写式操作行（exp-inline） */
    if(!done){
      html+=`<div class="exp-inline" style="margin:1px 0;padding:1px 8px;line-height:1.3">`;
      if(curId==null){
        /* 起始状态：逐字写出遍历规则 + 光标 */
        html+=`<span class="op">规则</span><span class="term done">${this.typeChars(meta.rule)}</span><span class="caret pen"></span>`;
      }else{
        if(opts.opLabel){
          html+=`<span class="op">${meta.name}操作</span><span class="term cur">${opts.opLabel}</span>`;
          if(opts.extra) html+=`<span class="op">→</span><span class="term done">${opts.extra}</span>`;
        }else{
          const node=this.node(curId);
          html+=`<span class="op">当前操作</span><span class="term cur">访问 ${esc(node.label)}</span>`;
          if(opts.role) html+=`<span class="op">· ${opts.role}</span>`;
        }
      }
      html+=`</div>`;
    }
    /* 访问序列（group-box 逐格写出） */
    html+=`<div class="zone-label" style="margin:0 0 1px;font-size:clamp(10px,1.2vw,11.5px);line-height:1.2">访问序列</div><div class="group-box" style="margin:1px 0;gap:3px;font-size:clamp(10.5px,1.3vw,12px)">`;
    if(visited.length===0){
      html+=`<span style="color:#CBD5E1;font-size:clamp(10.5px,1.3vw,12px)">（尚未访问）</span>`;
    }else{
      visited.forEach((id,i)=>{
        const n=this.node(id);
        const cls=done?'done':(id===curId?'cur':'done');
        html+=`<span class="grp ${cls} write-in" style="padding:1px 5px;border-width:1.5px;font-size:clamp(10.5px,1.3vw,12px);animation-delay:${i*0.08}s">${esc(n.label)}</span>`;
      });
    }
    html+=`</div>`;
    /* 队列状态（仅层序遍历显示） */
    if(meta.kind==='level'){
      html+=`<div class="zone-label" style="margin:1px 0 1px;font-size:clamp(10px,1.2vw,11.5px);line-height:1.2">队列（队首 → 队尾）</div>
        <div class="queue-area" style="min-height:22px;padding:5px 8px 3px;gap:4px;overflow:visible">`;
      if(queue.length===0) html+=`<span style="color:#94A3B8;font-size:11.5px">（空队列）</span>`;
      else queue.forEach((id,i)=>{
        const n=this.node(id);
        html+=`<span class="chip cy write-in" style="padding:1px 7px;font-size:clamp(10.5px,1.3vw,12px);animation-delay:${i*0.08}s">${esc(n.label)}</span>`;
      });
      html+=`</div>`;
    }
    /* 图例 */
    html+=`<div class="legend" style="margin-top:1px;line-height:1.4;font-size:clamp(9.5px,1.1vw,11px)">
      <span style="color:#94A3B8">● 未访问</span>　<span style="color:#F59E0B">● 当前</span>
      <span style="color:#22C55E">● 已访问</span>${meta.kind==='level'?'　<span style="color:#38BDF8">● 排队中</span>':''}
    </div>`;
    if(done){
      const seq=visited.map(v=>this.node(v).label).join(' · ');
      html+=`<div class="write-result" style="margin-top:2px;padding:2px 10px;line-height:1.3;font-size:clamp(11px,1.4vw,13px)">${meta.name}结果：[${seq}]</div>`;
    }
    $('#stage').innerHTML=html;
  },
  /* ---------- 步骤生成：递归三序 ---------- */
  genRecursiveSteps(kind){
    const meta=this.subMeta();
    const events=this.genRecursiveEvents(kind);
    const steps=[];
    steps.push({
      desc:`本次演示<b>${meta.name}</b>（<b>${meta.rule}</b>）：按规则递归遍历左、右子树，点击「下一步」逐个访问节点。`,
      draw:()=>{ this.renderStage({visited:[],curId:null,queue:[],done:false}); }
    });
    events.forEach(ev=>{
      const snap=ev;
      const labels=snap.visited.map(v=>this.node(v).label);
      const node=this.node(snap.curId);
      const role=this.roleText(snap.curId, snap.side);
      steps.push({
        desc:`<b>${meta.name}</b>（${meta.rule}）：访问${role} <b class="c-op">${esc(node.label)}</b>，当前访问序列：[${labels.join(', ')}]`,
        draw:()=>{ this.renderStage({visited:snap.visited,curId:snap.curId,queue:[],done:false,role}); }
      });
    });
    const finalIds=[...events[events.length-1].visited];
    steps.push({
      desc:`${meta.name}完成，最终访问序列：<b class="c-op">[${finalIds.map(v=>this.node(v).label).join(', ')}]</b>`,
      draw:()=>{ this.renderStage({visited:finalIds,curId:null,queue:[],done:true}); }
    });
    return steps;
  },
  /* ---------- 步骤生成：层序 ---------- */
  genLevelSteps(){
    const meta=this.subMeta();
    const events=this.genLevelEvents();
    const steps=[];
    steps.push({
      desc:`本次演示<b>${meta.name}</b>：借助<b>队列</b>从上到下、从左到右逐层访问，先访问根节点 <b class="c-op">1</b>。`,
      draw:()=>{ this.renderStage({visited:[],curId:null,queue:[],done:false}); }
    });
    events.forEach((ev,i)=>{
      const snap=ev;
      const node=this.node(snap.curId);
      const labels=snap.visited.join(', ');
      const newlyLabels=snap.newly.map(id=>this.node(id).label).join('、');
      const queueLabels=snap.queue.map(id=>this.node(id).label).join(' · ');
      steps.push({
        desc:`<b>${meta.name}</b>：取出队首${i===0?'（起点）':''} <b class="c-op">${esc(node.label)}</b>${snap.newly.length?`，将其未访问子节点 ${newlyLabels} 依次入队`:''}，队列：[${queueLabels||'空'}]，访问序列：[${labels}]`,
        draw:()=>{ this.renderStage({visited:snap.visited,curId:snap.curId,queue:snap.queue,done:false,opLabel:`出队 ${node.label}`,extra:newlyLabels?`入队 ${newlyLabels}`:null}); }
      });
    });
    const finalIds=[...events[events.length-1].visited];
    steps.push({
      desc:`层序遍历完成，最终访问序列：<b class="c-op">[${finalIds.map(v=>this.node(v).label).join(', ')}]</b>`,
      draw:()=>{ this.renderStage({visited:finalIds,curId:null,queue:[],done:true}); }
    });
    return steps;
  },
  /* 初始化演示画布（切换子页面时清空旧内容，显示完整灰色树） */
  initStage(){
    const meta=this.subMeta();
    this.renderStage({visited:[],curId:null,queue:[],done:false});
    $('#stage').innerHTML+=`<div class="annot" style="text-align:center;margin-top:1px;line-height:1.3">
      使用 <b>${meta.rule}</b>；点击「自动演示」或「下一步」开始遍历</div>`;
  },
  generate(){
    const sub=this.currentSub;
    if(sub==='level') this.steps=this.genLevelSteps();
    else this.steps=this.genRecursiveSteps(sub);
  }
};
