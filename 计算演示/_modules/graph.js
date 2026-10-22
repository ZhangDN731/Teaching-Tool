/* ================================================================
 * 图 Graph（广度优先 BFS · 深度优先 DFS）
 * 按 base 样板重构：subTabs + settingsHTML/afterSettings + initStage
 *                 + generate/genXxx + 手写动画逐步演示
 * ================================================================ */
MOD.graph = {
  name:'图 Graph',
  sub:'广度优先 BFS · 深度优先 DFS',
  badge:'结构 · 遍历', tag:'图', cat:'ds',
  color:'#6366F1', soft:'#EEF0FF',
  type:'无向图 · BFS · DFS',
  desc:'可视化演示无向图的广度优先搜索（BFS，借助队列逐层扩展）和深度优先搜索（DFS，借助栈一路深入再回溯）。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="6" r="2.5"/><circle cx="18" cy="6" r="2.5"/><circle cx="6" cy="18" r="2.5"/><circle cx="18" cy="18" r="2.5"/><circle cx="12" cy="12" r="2.5"/><path d="M8 6h8M6 8v8M18 8v8M8 18h8M7.5 7.5l3 3M16.5 7.5l-3 3M7.5 16.5l3-3M16.5 16.5l-3-3"/></svg>',
  setDesc:'选择遍历方式（BFS/DFS），从节点 A 开始遍历固定无向图。',
  stageDesc:'橙色=当前节点，绿色=已访问，蓝色=队列/栈中待访问，灰色=未访问。',
  initDesc:'点击「自动演示」或「下一步」，观看图的遍历过程。',
  currentSub:'bfs',
  subTabs:[
    {id:'bfs', label:'广度优先', method:'队列 FIFO'},
    {id:'dfs', label:'深度优先', method:'栈 LIFO'}
  ],
  theory(){
    return `<div class="rule"><b>图</b>：由顶点（节点）和边组成的数据结构，分为有向图和无向图。</div>
    <div class="rule"><b>广度优先搜索 BFS</b>：从起点开始，先访问所有相邻节点，再逐层向外扩展。借助<b>队列</b>实现，先进先出（FIFO）。</div>
    <div class="rule"><b>深度优先搜索 DFS</b>：从起点开始，沿一条路径一直深入到底，再回溯访问其他分支。借助<b>栈</b>或递归实现，后进先出（LIFO）。</div>
    <div class="rule">两种遍历都需要<b>标记已访问</b>的节点，避免重复访问和死循环。</div>
    <div class="rule">应用：BFS 求最短路径、层序遍历；DFS 求连通分量、拓扑排序、迷宫。</div>`;
  },
  settingsHTML(){
    return `<div class="inline-setting"><div class="setting"><label>图结构</label>
      <span style="font-family:var(--mono);font-size:clamp(11px,1.4vw,12.5px);color:var(--text-2);padding:4px 9px;border:1px solid var(--border);border-radius:8px;background:#fff">6 节点无向图（固定）</span></div></div>`;
  },
  afterSettings(){ /* 无动态参数；切换子页面由 tab 触发 Engine.reload() */ },

  /* ---------- 固定无向图数据 ---------- */
  nodes:[
    {id:'A',x:50,y:16},
    {id:'B',x:20,y:46},
    {id:'C',x:80,y:46},
    {id:'D',x:15,y:76},
    {id:'E',x:50,y:76},
    {id:'F',x:85,y:76}
  ],
  adj:{A:['B','C'],B:['A','D','E'],C:['A','F'],D:['B','E'],E:['B','D','F'],F:['C','E']},

  node(id){ return this.nodes.find(n=>n.id===id); },
  subMeta(){
    return {
      bfs:{name:'广度优先 BFS', rule:'队列 · 逐层扩展', ds:'队列', dsLabel:'队列（队首 → 队尾）', opVerb:'出队'},
      dfs:{name:'深度优先 DFS', rule:'栈 · 一路深入再回溯', ds:'栈', dsLabel:'栈（栈底 → 栈顶）', opVerb:'弹栈'}
    }[this.currentSub];
  },
  /* 逐字出现动画 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.05}s">${esc(ch)}</span>`).join('');
  },
  /* ---------- BFS：队列模拟，代码实际计算顺序 ---------- */
  genBfsEvents(){
    const visited=[], seen=new Set(['A']), queue=['A'], events=[];
    while(queue.length){
      const cur=queue.shift();
      visited.push(cur);
      const newly=[];
      this.adj[cur].forEach(nb=>{
        if(!seen.has(nb)){ seen.add(nb); queue.push(nb); newly.push(nb); }
      });
      events.push({curId:cur,visited:[...visited],waiting:[...queue],newly});
    }
    return events;
  },
  /* ---------- DFS：栈模拟（压栈即标记，逆序压栈保证邻居顺序一致） ---------- */
  genDfsEvents(){
    const visited=[], seen=new Set(['A']), stack=['A'], events=[];
    while(stack.length){
      const cur=stack.pop();
      visited.push(cur);
      const newly=[];
      [...this.adj[cur]].reverse().forEach(nb=>{
        if(!seen.has(nb)){ seen.add(nb); stack.push(nb); newly.push(nb); }
      });
      events.push({curId:cur,visited:[...visited],waiting:[...stack],newly});
    }
    return events;
  },
  /* ---------- 画布渲染 ---------- */
  renderGraph(opts){
    const meta=this.subMeta();
    const visited=opts.visited||[], curId=opts.curId, done=!!opts.done;
    const waiting=opts.waiting||[], vis=new Set(visited);
    let html=`<div class="stage-hd" style="margin-bottom:1px;line-height:1.2;font-size:clamp(10.5px,1.3vw,12px)"><span>无向图 · ${meta.name}</span><span class="state">${vis.size}/6</span></div>`;
    /* 步骤标题条 */
    html+=`<div class="step-title" style="margin-bottom:1px;padding:1px 8px;line-height:1.3"><span class="st-dot"></span>${done?'遍历完成':'无向图 · 节点访问状态'}</div>`;
    /* SVG 无向图 */
    html+=`<div class="svg-box" style="margin-bottom:1px;display:flex;justify-content:center">
      <svg viewBox="0 0 100 82" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:320px;height:110px">`;
    const drawn=new Set();
    Object.entries(this.adj).forEach(([a,nbs])=>{
      nbs.forEach(b=>{
        const key=[a,b].sort().join('-');
        if(drawn.has(key)) return;
        drawn.add(key);
        const na=this.node(a), nb=this.node(b);
        const both=vis.has(a)&&vis.has(b);
        const touched=(a===curId||b===curId)&&(waiting.includes(a)||waiting.includes(b));
        const stroke=both?'#22C55E':touched?'#7DD3FC':'#E2E8F0';
        html+=`<line x1="${na.x}" y1="${na.y}" x2="${nb.x}" y2="${nb.y}" stroke="${stroke}" stroke-width="0.7"/>`;
      });
    });
    this.nodes.forEach(n=>{
      let state='un';
      if(curId===n.id) state='cur';
      else if(vis.has(n.id)) state='done';
      else if(waiting.includes(n.id)) state='waiting';
      const st={
        un:     {f:'#F1F5F9', s:'#CBD5E1', t:'#64748B'},
        cur:    {f:'#F59E0B', s:'#D97706', t:'#fff'},
        done:   {f:'#22C55E', s:'#16A34A', t:'#fff'},
        waiting:{f:'#E0F2FE', s:'#38BDF8', t:'#0369A1'}
      }[state];
      if(state==='cur')
        html+=`<circle cx="${n.x}" cy="${n.y}" r="7" fill="#F59E0B" opacity="0.18" style="animation:pop .3s ease-out"/>`;
      html+=`<circle cx="${n.x}" cy="${n.y}" r="5" fill="${st.f}" stroke="${st.s}" stroke-width="0.8"/>`;
      html+=`<text x="${n.x}" y="${n.y+1.8}" text-anchor="middle" font-size="4" font-weight="700" fill="${st.t}">${esc(n.id)}</text>`;
    });
    html+=`</svg></div>`;
    /* 横写式操作行（exp-inline） */
    if(!done){
      html+=`<div class="exp-inline" style="margin:1px 0;padding:1px 8px;line-height:1.3">`;
      if(curId==null){
        html+=`<span class="op">规则</span><span class="term done">${this.typeChars(meta.rule)}</span><span class="caret pen"></span>`;
      }else{
        html+=`<span class="op">${meta.ds}操作</span><span class="term cur">${meta.opVerb} ${esc(curId)}</span>`;
        if(opts.extra) html+=`<span class="op">→</span><span class="term done">${opts.extra}</span>`;
      }
      html+=`</div>`;
    }
    /* 访问序列（group-box 逐格写出） */
    html+=`<div class="zone-label" style="margin:0 0 1px;font-size:clamp(10px,1.2vw,11.5px);line-height:1.2">访问序列</div><div class="group-box" style="margin:1px 0;gap:3px;font-size:clamp(10.5px,1.3vw,12px)">`;
    if(visited.length===0){
      html+=`<span style="color:#CBD5E1;font-size:clamp(10.5px,1.3vw,12px)">（尚未访问）</span>`;
    }else{
      visited.forEach((id,i)=>{
        const cls=done?'done':(id===curId?'cur':'done');
        html+=`<span class="grp ${cls} write-in" style="padding:1px 5px;border-width:1.5px;font-size:clamp(10.5px,1.3vw,12px);animation-delay:${i*0.08}s">${esc(id)}</span>`;
      });
    }
    html+=`</div>`;
    /* 队列 / 栈状态 */
    html+=`<div class="zone-label" style="margin:1px 0 1px;font-size:clamp(10px,1.2vw,11.5px);line-height:1.2">${meta.dsLabel}</div>`;
    if(meta.ds==='队列'){
      html+=`<div class="queue-area" style="min-height:22px;padding:5px 8px 3px;gap:4px;overflow:visible">`;
      if(waiting.length===0) html+=`<span style="color:#94A3B8;font-size:11.5px">（空队列）</span>`;
      else waiting.forEach((id,i)=>{
        const isHead=i===0, isTail=i===waiting.length-1;
        const tag=isHead?'<span style="font-size:8.5px;color:#14B8A6;margin-left:2px;font-weight:700">头</span>'
          :isTail?'<span style="font-size:8.5px;color:#14B8A6;margin-left:2px;font-weight:700">尾</span>':'';
        html+=`<span class="chip cy write-in" style="padding:1px 7px;font-size:clamp(10.5px,1.3vw,12px);animation-delay:${i*0.08}s">${esc(id)}${tag}</span>`;
      });
      html+=`</div>`;
    }else{
      html+=`<div class="queue-area" style="min-height:22px;padding:5px 8px 3px;gap:4px;overflow:visible">`;
      if(waiting.length===0) html+=`<span style="color:#94A3B8;font-size:11.5px">（空栈）</span>`;
      else waiting.forEach((id,i)=>{
        const isTop=i===waiting.length-1;
        const tag=isTop?'<span style="font-size:8.5px;color:#fff;margin-left:2px;font-weight:700">顶</span>':'';
        html+=`<span class="chip cy write-in" style="padding:1px 7px;font-size:clamp(10.5px,1.3vw,12px);animation-delay:${i*0.08}s;${isTop?'background:#8B5CF6;border:1.5px solid #8B5CF6;':''}">${esc(id)}${tag}</span>`;
      });
      html+=`</div>`;
    }
    /* 图例 */
    html+=`<div class="legend" style="margin-top:1px;line-height:1.4;font-size:clamp(9.5px,1.1vw,11px)">
      <span style="color:#94A3B8">● 未访问</span>　<span style="color:#F59E0B">● 当前</span>
      <span style="color:#22C55E">● 已访问</span>　<span style="color:#38BDF8">● ${meta.ds}中待访问</span>
    </div>`;
    if(done){
      const seq=visited.join(' · ');
      html+=`<div class="write-result" style="margin-top:2px;padding:2px 10px;line-height:1.3;font-size:clamp(11px,1.4vw,13px)">${meta.name}结果：[${seq}]</div>`;
    }
    $('#stage').innerHTML=html;
  },
  /* ---------- 步骤生成：BFS ---------- */
  genBfsSteps(){
    const meta=this.subMeta();
    const events=this.genBfsEvents();
    const steps=[];
    steps.push({
      desc:`本次演示<b>${meta.name}</b>：从起点 <b class="c-op">A</b> 出发，先访问所有相邻节点，再借助<b>队列</b>逐层向外扩展。`,
      draw:()=>{ this.renderGraph({visited:[],curId:null,waiting:[],done:false}); }
    });
    events.forEach((ev,i)=>{
      const snap=ev;
      const newlyLabels=snap.newly.join('、');
      const waitingLabels=snap.waiting.join(' · ');
      steps.push({
        desc:`<b>${meta.name}</b>：取出队首${i===0?'（起点）':''} <b class="c-op">${esc(snap.curId)}</b>${snap.newly.length?`，未访问邻居 ${newlyLabels} 依次入队`:''}，队列：[${waitingLabels||'空'}]，访问序列：[${snap.visited.join(', ')}]`,
        draw:()=>{ this.renderGraph({visited:snap.visited,curId:snap.curId,waiting:snap.waiting,done:false,extra:newlyLabels?`入队 ${newlyLabels}`:''}); }
      });
    });
    const finalIds=[...events[events.length-1].visited];
    steps.push({
      desc:`广度优先搜索完成，最终访问序列：<b class="c-op">[${finalIds.join(', ')}]</b>`,
      draw:()=>{ this.renderGraph({visited:finalIds,curId:null,waiting:[],done:true}); }
    });
    return steps;
  },
  /* ---------- 步骤生成：DFS ---------- */
  genDfsSteps(){
    const meta=this.subMeta();
    const events=this.genDfsEvents();
    const steps=[];
    steps.push({
      desc:`本次演示<b>${meta.name}</b>：从起点 <b class="c-op">A</b> 出发，借助<b>栈</b>沿一条路径一直深入到底，再回溯访问其他分支。`,
      draw:()=>{ this.renderGraph({visited:[],curId:null,waiting:[],done:false}); }
    });
    events.forEach((ev,i)=>{
      const snap=ev;
      const newlyLabels=snap.newly.join('、');
      const waitingLabels=snap.waiting.join(' · ');
      steps.push({
        desc:`<b>${meta.name}</b>：弹出栈顶${i===0?'（起点）':''} <b class="c-op">${esc(snap.curId)}</b>${snap.newly.length?`，邻居 ${newlyLabels} 逆序压栈`:''}，栈（栈底 → 栈顶）：[${waitingLabels||'空'}]，访问序列：[${snap.visited.join(', ')}]`,
        draw:()=>{ this.renderGraph({visited:snap.visited,curId:snap.curId,waiting:snap.waiting,done:false,extra:newlyLabels?`压栈 ${newlyLabels}`:''}); }
      });
    });
    const finalIds=[...events[events.length-1].visited];
    steps.push({
      desc:`深度优先搜索完成，最终访问序列：<b class="c-op">[${finalIds.join(', ')}]</b>`,
      draw:()=>{ this.renderGraph({visited:finalIds,curId:null,waiting:[],done:true}); }
    });
    return steps;
  },
  /* 初始化演示画布（切换子页面时清空旧内容，显示完整灰色图） */
  initStage(){
    const meta=this.subMeta();
    this.renderGraph({visited:[],curId:null,waiting:[],done:false});
    $('#stage').innerHTML+=`<div class="annot" style="text-align:center;margin-top:1px;line-height:1.3">
      使用 <b>${meta.rule}</b>；点击「自动演示」或「下一步」开始遍历</div>`;
  },
  generate(){
    if(this.currentSub==='dfs') this.steps=this.genDfsSteps();
    else this.steps=this.genBfsSteps();
  }
};
