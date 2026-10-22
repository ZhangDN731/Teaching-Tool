/* ================================================================
 * 图 Graph（可视化自定义构建 · BFS · DFS · Dijkstra 最短路径）
 * 构建阶段：输入名称添加顶点，点击两个顶点再添加边（可选权重），支持拖拽
 * 遍历阶段：点击「开始遍历」后由 Engine 步骤系统逐步演示 BFS / DFS / 最短路径
 * ================================================================ */
(function () {
  if (document.getElementById('graph-custom-css')) return;
  document.head.insertAdjacentHTML('beforeend', `<style id="graph-custom-css">
    #stage .g-node{cursor:grab;transition:filter .15s ease}
    #stage .g-node:active{cursor:grabbing}
    #stage .g-node:hover{filter:brightness(.95)}
    .order-chip.chip-done{background:var(--green)}
    .order-chip.chip-cur{background:var(--orange)}
    #stage svg text{font-family:var(--font)}
  </style>`);
})();

MOD.graph = {
  name: '图 Graph',
  sub: '广度优先 BFS · 深度优先 DFS · 最短路径',
  badge: '结构 · 遍历', tag: '图', cat: 'ds',
  color: '#6366F1', soft: '#EEF0FF',
  type: '无向图 · 可视化构建 · BFS · DFS · Dijkstra',
  desc: '在画布上添加顶点、点击两两相连组成无向图（可设权重），演示广度优先搜索（BFS，队列逐层扩展）、深度优先搜索（DFS，栈一路深入），并用 Dijkstra 求带权最短路径。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="6" r="2.5"/><circle cx="18" cy="6" r="2.5"/><circle cx="6" cy="18" r="2.5"/><circle cx="18" cy="18" r="2.5"/><circle cx="12" cy="12" r="2.5"/><path d="M8 6h8M6 8v8M18 8v8M8 18h8M7.5 7.5l3 3M16.5 7.5l-3 3M7.5 16.5l3-3M16.5 16.5l-3-3"/></svg>',
  setDesc: '在上方添加顶点、连接边（可填权重），选择起点；切换顶部标签选择 BFS / DFS / 最短路径，再点击「开始遍历」。',
  stageDesc: '构建模式：蓝色虚线圈=选中端点一，紫色虚线圈=选中端点二。遍历模式：橙色=当前顶点，绿色=已访问，蓝色=队列/栈中待访问，灰色=未访问，橙色粗边=当前经过的边。',
  initDesc: '先构建图（或点击「示例图」），选择起点，点击「开始遍历」后用「下一步」观看搜索过程。',
  currentSub: 'bfs',
  subTabs: [
    { id: 'bfs',  label: '广度优先', method: '队列 FIFO' },
    { id: 'dfs',  label: '深度优先', method: '栈 LIFO' },
    { id: 'path', label: '最短路径', method: 'Dijkstra · 加权' }
  ],

  theory() {
    return `<div class="rule"><b>图</b>：由<b>顶点</b>（节点）和<b>边</b>组成的数据结构，分为有向图和无向图；边上可以标注<b>权重</b>（如距离、时间、费用）。</div>
    <div class="rule"><b>可视化构建</b>：输入顶点名称点击「添加顶点」；在画布上依次点击两个顶点，填好可选权重后点击「连接选中两顶点」。顶点可直接拖拽调整位置。</div>
    <div class="rule"><b>广度优先搜索 BFS</b>：从起点开始，先访问所有相邻节点，再逐层向外扩展。借助<b>队列</b>实现，先进先出（FIFO）；无权图中 BFS 得到的就是最短路径。</div>
    <div class="rule"><b>深度优先搜索 DFS</b>：从起点开始，沿一条路径一直深入到底，再回溯访问其他分支。借助<b>栈</b>或递归实现，后进先出（LIFO）。</div>
    <div class="rule"><b>Dijkstra 最短路径</b>：每次取出当前距离最小的未确定顶点，用它松弛邻边（经过它更近就更新距离），直到终点距离确定。</div>
    <div class="rule">两种遍历都需要<b>标记已访问</b>的节点，避免重复访问和死循环。应用：BFS 求最短路径、层序遍历；DFS 求连通分量、拓扑排序、迷宫；Dijkstra 求加权网络最短路径。</div>`;
  },

  settingsHTML() {
    return `
      <div class="settings-row">
        <div class="inline-setting"><div class="setting"><label>顶点名称</label>
          <input type="text" id="gf-v" placeholder="如 E" style="width:clamp(80px,12vw,110px)"></div></div>
        <div class="inline-setting"><button class="btn primary" id="gf-addv">添加顶点</button></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><div class="setting"><label>边权重（可空）</label>
          <input type="number" id="gf-w" placeholder="如 3" style="width:clamp(80px,12vw,110px)"></div></div>
        <div class="inline-setting"><button class="btn" id="gf-adde">连接选中两顶点</button></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><div class="setting"><label>顶点 1</label>
          <input type="text" id="gf-a" placeholder="如 A" style="width:clamp(70px,10vw,100px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>顶点 2</label>
          <input type="text" id="gf-b" placeholder="如 D" style="width:clamp(70px,10vw,100px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>权重</label>
          <input type="number" id="gf-w2" placeholder="可空" style="width:clamp(64px,9vw,96px)"></div></div>
        <div class="inline-setting"><button class="btn" id="gf-pair">按名称加边</button></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><div class="setting"><label>遍历起点</label><select id="gf-start" style="min-width:80px"></select></div></div>
        <div class="inline-setting"><div class="setting"><label>终点（最短路径）</label><select id="gf-end" style="min-width:80px"></select></div></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><button class="btn" id="gf-preset">示例图</button></div>
        <div class="inline-setting"><button class="btn ghost" id="gf-clear">清空画布</button></div>
        <div class="inline-setting"><button class="btn primary" id="gf-startbtn">开始遍历</button></div>
      </div>
      <div class="chk" id="gf-msg"></div>`;
  },

  /* ================= 自定义构建状态 ================= */
  ensureState() {
    if (!this._g) {
      this._g = { nodes: [], edges: [], sel: [], start: null, end: null };
      this.preset();
    }
  },
  node(id) { return this._g.nodes.find(n => n.id === id); },
  nodeByLabel(lbl) { return this._g.nodes.find(n => n.id === lbl); },
  edgeBetween(a, b) {
    if (!a || !b) return null;
    return this._g.edges.find(e => (e.a === a && e.b === b) || (e.a === b && e.b === a)) || null;
  },
  adjMap() {
    const m = {};
    this._g.nodes.forEach(n => m[n.id] = []);
    this._g.edges.forEach(e => { m[e.a].push(e.b); m[e.b].push(e.a); });
    return m;
  },
  autoPos() {
    const n = this._g.nodes.length;
    const ang = (-90 + n * 137.5) * Math.PI / 180;
    const x = 400 + 280 * Math.cos(ang);
    const y = 260 + 175 * Math.sin(ang);
    return { x: Math.max(60, Math.min(740, x)), y: Math.max(60, Math.min(460, y)) };
  },
  preset() {
    const g = this._g;
    g.nodes = [
      { id: 'A', x: 400, y: 80 },
      { id: 'B', x: 190, y: 270 },
      { id: 'C', x: 610, y: 270 },
      { id: 'D', x: 400, y: 450 }
    ];
    g.edges = [
      { a: 'A', b: 'B', w: 1, key: 'A-B' },
      { a: 'A', b: 'C', w: 2, key: 'A-C' },
      { a: 'B', b: 'D', w: 3, key: 'B-D' },
      { a: 'C', b: 'D', w: 1, key: 'C-D' }
    ];
    g.sel = []; g.start = 'A'; g.end = 'D';
  },
  clearGraph() {
    this._g = { nodes: [], edges: [], sel: [], start: null, end: null };
  },
  populateSelects() {
    const g = this._g;
    const fill = (selId, cur) => {
      const el = $(selId); if (!el) return;
      if (g.nodes.length === 0) { el.innerHTML = `<option value="">（无顶点）</option>`; return; }
      el.innerHTML = g.nodes.map(n =>
        `<option value="${esc(n.id)}" ${n.id === cur ? 'selected' : ''}>${esc(n.id)}</option>`).join('');
    };
    fill('#gf-start', g.start);
    fill('#gf-end', g.end);
  },
  afterEdit() {
    this.steps = [];
    Engine.renderAll();
    this.renderBuildStage();
    this.populateSelects();
    const eb = $('#explainBox'); if (eb) eb.innerHTML = this.initDesc;
  },

  subMeta() {
    return {
      bfs:  { name: '广度优先 BFS', rule: '队列 · 逐层扩展', ds: '队列', dsLabel: '队列（队首 → 队尾）', opVerb: '出队', kind: 'bfs' },
      dfs:  { name: '深度优先 DFS', rule: '栈 · 一路深入再回溯', ds: '栈', dsLabel: '栈（栈底 → 栈顶）', opVerb: '弹栈', kind: 'dfs' },
      path: { name: '最短路径 Dijkstra', rule: '权重 · 逐步松弛', ds: null, dsLabel: '', opVerb: '确定距离', kind: 'path' }
    }[this.currentSub];
  },
  typeChars(text) {
    return [...String(text)].map((ch, i) => `<span class="char-in" style="animation-delay:${i * 0.05}s">${esc(ch)}</span>`).join('');
  },

  /* ================= 搜索事件（由当前结构实际计算） ================= */
  genBfsEvents() {
    const g = this._g, start = g.start || g.nodes[0].id;
    const visited = [], seen = new Set([start]), queue = [{ id: start, from: null }], events = [];
    const adj = this.adjMap();
    while (queue.length) {
      const { id: cur, from } = queue.shift();
      visited.push(cur);
      const newly = [];
      adj[cur].forEach(nb => {
        if (!seen.has(nb)) { seen.add(nb); queue.push({ id: nb, from: cur }); newly.push(nb); }
      });
      events.push({ curId: cur, visited: [...visited], waiting: queue.map(q => q.id), newly, edge: from ? [from, cur] : null });
    }
    return events;
  },
  genDfsEvents() {
    const g = this._g, start = g.start || g.nodes[0].id;
    const visited = [], seen = new Set(), stack = [{ id: start, from: null }], events = [];
    const adj = this.adjMap();
    while (stack.length) {
      const { id: cur, from } = stack.pop();
      if (seen.has(cur)) continue;
      seen.add(cur);
      visited.push(cur);
      const newly = [];
      [...adj[cur]].reverse().forEach(nb => {
        if (!seen.has(nb)) { stack.push({ id: nb, from: cur }); newly.push(nb); }
      });
      events.push({ curId: cur, visited: [...visited], waiting: stack.map(q => q.id), newly, edge: from ? [from, cur] : null });
    }
    return events;
  },
  genDijkstraEvents() {
    const g = this._g, start = g.start, end = g.end;
    const adj = this.adjMap();
    const dist = {}, prev = {}, visited = [], done = new Set();
    g.nodes.forEach(n => dist[n.id] = Infinity);
    dist[start] = 0;
    const events = [{ type: 'init', visited: [], curId: null, waiting: [], dist: Object.assign({}, dist), edge: null, relaxed: [] }];
    while (true) {
      let u = null;
      g.nodes.forEach(n => {
        if (!done.has(n.id) && dist[n.id] < Infinity && (u == null || dist[n.id] < dist[u])) u = n.id;
      });
      if (u == null) break;
      done.add(u); visited.push(u);
      const relaxed = [];
      adj[u].forEach(v => {
        if (done.has(v)) return;
        const e = this.edgeBetween(u, v);
        const w = e && e.w != null && e.w !== '' ? Number(e.w) : 1;
        if (dist[u] + w < dist[v]) { dist[v] = dist[u] + w; prev[v] = u; relaxed.push({ v, d: dist[v] }); }
      });
      events.push({ type: 'visit', curId: u, visited: [...visited], waiting: [], dist: Object.assign({}, dist), edge: prev[u] ? [prev[u], u] : null, relaxed });
      if (u === end) break;
    }
    const path = [];
    if (dist[end] !== Infinity) {
      let cur = end;
      while (cur != null) { path.unshift(cur); cur = prev[cur]; }
    }
    events.push({ type: 'done', visited: [...visited], curId: null, waiting: [], dist: Object.assign({}, dist), edge: null, relaxed: [], path, total: dist[end] });
    return events;
  },

  /* ================= 构建模式画布 ================= */
  clientToSvg(ev) {
    const svg = $('#stage svg'); if (!svg) return { x: 0, y: 0 };
    const pt = svg.createSVGPoint(); pt.x = ev.clientX; pt.y = ev.clientY;
    const ctm = svg.getScreenCTM(); if (!ctm) return { x: 0, y: 0 };
    const p = pt.matrixTransform(ctm.inverse());
    return { x: p.x, y: p.y };
  },
  refreshBuildCoords() {
    const g = this._g;
    g.nodes.forEach(n => {
      const c = $(`#stage circle[data-node="${n.id}"]`);
      const tx = $(`#stage text[data-node="${n.id}"]`);
      if (c) { c.setAttribute('cx', n.x); c.setAttribute('cy', n.y); }
      if (tx) { tx.setAttribute('x', n.x); tx.setAttribute('y', n.y + 7); }
    });
    g.edges.forEach(e => {
      const a = this.node(e.a), b = this.node(e.b);
      const line = $(`#stage line[data-edge="${e.key}"]`);
      if (line) { line.setAttribute('x1', a.x); line.setAttribute('y1', a.y); line.setAttribute('x2', b.x); line.setAttribute('y2', b.y); }
      const gEl = $(`#stage g[data-w="${e.key}"]`);
      if (gEl) gEl.setAttribute('transform', `translate(${(a.x + b.x) / 2},${(a.y + b.y) / 2})`);
    });
  },
  renderBuildStage() {
    const g = this._g;
    let html = `<div class="stage-hd" style="margin-bottom:6px;line-height:1.2;font-size:clamp(12px,1.6vw,14px)">
      <span>自定义无向图 · 构建中</span><span class="state">${g.nodes.length} 个顶点 · ${g.edges.length} 条边</span></div>`;
    html += `<div class="step-title" style="margin-bottom:8px;padding:4px 10px;line-height:1.3">
      <span class="st-dot"></span>依次点击两个顶点选中端点，再点击「连接选中两顶点」；可拖拽顶点调整位置</div>`;
    html += `<div class="svg-box" style="margin-bottom:8px;display:flex;justify-content:center;background:var(--surface-2)">
      <svg viewBox="0 0 800 520" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:640px">`;

    g.edges.forEach(e => {
      const a = this.node(e.a), b = this.node(e.b);
      html += `<line data-edge="${e.key}" x1="${a.x}" y1="${a.y}" x2="${b.x}" y2="${b.y}" stroke="#CBD5E1" stroke-width="3" stroke-linecap="round"/>`;
      if (e.w != null && e.w !== '') {
        const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
        html += `<g data-w="${e.key}" transform="translate(${mx},${my})">
          <rect x="-20" y="-15" width="40" height="26" rx="10" fill="#fff" stroke="#CBD5E1" stroke-width="1.5"/>
          <text text-anchor="middle" y="4" font-size="clamp(12px,1.5vw,14px)" font-weight="700" fill="#334155">${esc(e.w)}</text></g>`;
      }
    });

    if (g.nodes.length === 0) {
      html += `<text x="400" y="260" text-anchor="middle" fill="#94A3B8" font-size="20" font-weight="600">画布为空：在上方输入名称，点击「添加顶点」开始构建</text>`;
    }

    g.nodes.forEach(n => {
      const selIdx = g.sel.indexOf(n.id);
      if (selIdx === 0) html += `<circle cx="${n.x}" cy="${n.y}" r="36" fill="none" stroke="#4F6BFF" stroke-width="3" stroke-dasharray="7 6" opacity=".65"/>`;
      if (selIdx === 1) html += `<circle cx="${n.x}" cy="${n.y}" r="36" fill="none" stroke="#8B5CF6" stroke-width="3" stroke-dasharray="7 6" opacity=".65"/>`;
      if (g.start === n.id) html += `<text x="${n.x}" y="${n.y - 40}" text-anchor="middle" font-size="13" font-weight="700" fill="#4F6BFF">起点</text>`;
      html += `<circle class="g-node" data-node="${n.id}" cx="${n.x}" cy="${n.y}" r="26"
        fill="#fff" stroke="${selIdx >= 0 ? '#4F6BFF' : '#CBD5E1'}" stroke-width="3" style="cursor:grab"/>`;
      html += `<text data-node="${n.id}" x="${n.x}" y="${n.y + 7}" text-anchor="middle" font-size="clamp(13px,1.8vw,17px)" font-weight="700" fill="#334155" style="pointer-events:none">${esc(n.id)}</text>`;
    });
    html += `</svg></div>`;

    let info = '尚未选中端点；依次点击两个顶点即可连边。';
    if (g.sel.length === 1) info = `已选端点一：<b style="color:#4F6BFF">${esc(g.sel[0])}</b>，请再点击第二个顶点。`;
    else if (g.sel.length === 2) info = `已选端点：<b style="color:#4F6BFF">${esc(g.sel[0])}</b> — <b style="color:#8B5CF6">${esc(g.sel[1])}</b>，点击「连接选中两顶点」。`;
    html += `<div class="annot" style="text-align:center;margin-bottom:6px;line-height:1.6">${info}</div>`;
    html += `<div class="zone-label">访问序列</div>
      <div class="order-rack"><span style="color:#CBD5E1;font-size:clamp(12px,1.6vw,14px)">构建完成后点击「开始遍历」，访问序列将显示在这里</span></div>`;

    $('#stage').innerHTML = html;
    this.bindBuildInteractions();
  },

  bindBuildInteractions() {
    const g = this._g;
    $$('#stage .g-node').forEach(c => {
      const id = c.dataset.node;
      let downX = 0, downY = 0, moved = false;
      c.addEventListener('pointerdown', ev => {
        downX = ev.clientX; downY = ev.clientY; moved = false;
        c.setPointerCapture(ev.pointerId);
        const n = this.node(id);
        const onMove = mv => {
          if (Math.hypot(mv.clientX - downX, mv.clientY - downY) > 5) moved = true;
          if (!moved) return;
          const p = this.clientToSvg(mv);
          n.x = Math.max(50, Math.min(750, p.x));
          n.y = Math.max(50, Math.min(470, p.y));
          this.refreshBuildCoords();
        };
        const onUp = () => {
          c.removeEventListener('pointermove', onMove);
          c.removeEventListener('pointerup', onUp);
          if (!moved) {
            if (g.sel.includes(id)) g.sel = g.sel.filter(x => x !== id);
            else { g.sel.push(id); if (g.sel.length > 2) g.sel.shift(); }
            this.afterEdit();
          }
        };
        c.addEventListener('pointermove', onMove);
        c.addEventListener('pointerup', onUp);
      });
    });
  },

  /* ================= 遍历模式画布 ================= */
  renderWalk(opts) {
    const meta = this.subMeta();
    const visited = opts.visited || [], curId = opts.curId, done = !!opts.done;
    const waiting = opts.waiting || [], vis = new Set(visited);
    const g = this._g;
    const pathSet = opts.pathKeys || new Set();

    let html = `<div class="stage-hd" style="margin-bottom:6px;line-height:1.2;font-size:clamp(12px,1.6vw,14px)">
      <span>无向图 · ${meta.name}</span><span class="state">${vis.size}/${g.nodes.length}</span></div>`;
    html += `<div class="step-title" style="margin-bottom:8px;padding:4px 10px;line-height:1.3">
      <span class="st-dot"></span>${done ? '遍历完成' : '顶点访问状态'}</div>`;
    html += `<div class="svg-box" style="margin-bottom:8px;display:flex;justify-content:center;background:var(--surface-2)">
      <svg viewBox="0 0 800 520" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:640px">`;

    g.edges.forEach(e => {
      const a = this.node(e.a), b = this.node(e.b);
      const both = vis.has(e.a) && vis.has(e.b);
      const isCur = opts.edge && ((opts.edge[0] === e.a && opts.edge[1] === e.b) || (opts.edge[0] === e.b && opts.edge[1] === e.a));
      const inPath = pathSet.has(e.key);
      let stroke = '#E2E8F0', width = 3;
      if (inPath) { stroke = '#22C55E'; width = 6; }
      else if (isCur) { stroke = '#F59E0B'; width = 5; }
      else if (both) { stroke = '#22C55E'; }
      html += `<line x1="${a.x}" y1="${a.y}" x2="${b.x}" y2="${b.y}" stroke="${stroke}" stroke-width="${width}" stroke-linecap="round"/>`;
      if (e.w != null && e.w !== '') {
        const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
        html += `<g transform="translate(${mx},${my})">
          <rect x="-20" y="-15" width="40" height="26" rx="10" fill="#fff" stroke="${inPath ? '#22C55E' : '#CBD5E1'}" stroke-width="1.5"/>
          <text text-anchor="middle" y="4" font-size="clamp(12px,1.5vw,14px)" font-weight="700" fill="${inPath ? '#16A34A' : '#334155'}">${esc(e.w)}</text></g>`;
      }
    });

    g.nodes.forEach(n => {
      let state = 'un';
      if (curId === n.id) state = 'cur';
      else if (vis.has(n.id)) state = 'done';
      else if (waiting.includes(n.id)) state = 'waiting';
      const st = {
        un:      { f: '#F1F5F9', s: '#CBD5E1', t: '#64748B' },
        cur:     { f: '#F59E0B', s: '#D97706', t: '#fff' },
        done:    { f: '#22C55E', s: '#16A34A', t: '#fff' },
        waiting: { f: '#E0F2FE', s: '#38BDF8', t: '#0369A1' }
      }[state];
      if (state === 'cur')
        html += `<circle cx="${n.x}" cy="${n.y}" r="36" fill="#F59E0B" opacity="0.15" style="animation:pop .3s ease-out"/>`;
      if (g.start === n.id) html += `<text x="${n.x}" y="${n.y - 40}" text-anchor="middle" font-size="13" font-weight="700" fill="#4F6BFF">起点</text>`;
      if (meta.kind === 'path' && g.end === n.id && g.end !== g.start) html += `<text x="${n.x}" y="${n.y - 40}" text-anchor="end" font-size="13" font-weight="700" fill="#8B5CF6" dx="18">终点</text>`;
      html += `<circle cx="${n.x}" cy="${n.y}" r="26" fill="${st.f}" stroke="${st.s}" stroke-width="3"/>`;
      html += `<text x="${n.x}" y="${n.y + 7}" text-anchor="middle" font-size="clamp(13px,1.8vw,17px)" font-weight="700" fill="${st.t}">${esc(n.id)}</text>`;
    });
    html += `</svg></div>`;

    if (!done && meta.kind !== 'path') {
      html += `<div class="exp-inline" style="margin:4px 0 8px;padding:4px 10px;line-height:1.6">`;
      if (curId == null) {
        html += `<span class="op">规则</span><span class="term done">${this.typeChars(meta.rule)}</span><span class="caret pen"></span>`;
      } else {
        html += `<span class="op">${meta.ds}操作</span><span class="term cur">${meta.opVerb} ${esc(curId)}</span>`;
        if (opts.extra) html += `<span class="op">→</span><span class="term done">${opts.extra}</span>`;
      }
      html += `</div>`;
    }

    html += `<div class="zone-label">访问序列</div><div class="order-rack" style="margin-bottom:8px">`;
    if (visited.length === 0) {
      html += `<span style="color:#CBD5E1;font-size:clamp(12px,1.6vw,14px)">（尚未访问）</span>`;
    } else {
      visited.forEach((id, i) => {
        const cls = done ? 'chip-done' : (id === curId ? 'chip-cur' : 'chip-done');
        html += `<span class="order-chip ${cls} write-in" style="animation-delay:${i * 0.08}s">${esc(id)}</span>`;
      });
    }
    html += `</div>`;

    if (meta.kind === 'bfs' || meta.kind === 'dfs') {
      html += `<div class="zone-label">${meta.dsLabel}</div>
        <div class="queue-area" style="min-height:34px;padding:8px 10px 6px;gap:6px;overflow:visible;margin-bottom:8px">`;
      if (waiting.length === 0) html += `<span style="color:#94A3B8;font-size:clamp(12px,1.5vw,14px)">（空${meta.ds}）</span>`;
      else waiting.forEach((id, i) => {
        const mark = meta.kind === 'bfs'
          ? (i === 0 ? '头' : (i === waiting.length - 1 ? '尾' : ''))
          : (i === waiting.length - 1 ? '顶' : '');
        const tag = mark ? `<span style="font-size:10px;color:#fff;margin-left:3px;font-weight:700">${mark}</span>` : '';
        const topStyle = meta.kind === 'dfs' && i === waiting.length - 1 ? 'background:#8B5CF6;border:1.5px solid #8B5CF6;' : '';
        html += `<span class="chip cy write-in" style="padding:4px 10px;font-size:clamp(12px,1.6vw,14px);${topStyle}animation-delay:${i * 0.08}s">${esc(id)}${tag}</span>`;
      });
      html += `</div>`;
    }

    if (meta.kind === 'path' && opts.dist) {
      html += `<div class="zone-label">距离表（当前最短估计距离）</div>
        <div class="tbl-wrap" style="margin-bottom:8px;overflow-x:auto"><table class="tbl" style="table-layout:fixed">
        <thead><tr><th style="width:34%">顶点</th><th style="width:33%">距离</th><th style="width:33%">状态</th></tr></thead><tbody>`;
      g.nodes.forEach(n => {
        const d = opts.dist[n.id];
        const stText = vis.has(n.id) ? '已确定' : (d === Infinity ? '待访问' : '候选');
        html += `<tr><td>${esc(n.id)}</td><td>${d === Infinity ? '∞' : d}</td><td>${stText}</td></tr>`;
      });
      html += `</tbody></table></div>`;
    }

    html += `<div class="legend" style="margin-top:6px;line-height:1.8;font-size:clamp(11px,1.4vw,13px)">
      <span style="color:#94A3B8">● 未访问</span>　<span style="color:#F59E0B">● 当前</span>
      <span style="color:#22C55E">● 已访问</span>${meta.kind !== 'path' ? `　<span style="color:#38BDF8">● ${meta.ds}中待访问</span>` : '　<span style="color:#22C55E">● 最短路径</span>'}
    </div>`;

    if (done) {
      if (meta.kind === 'path') {
        if (opts.path && opts.path.length) {
          html += `<div class="write-result" style="margin-top:8px;padding:6px 14px;line-height:1.6;font-size:clamp(13px,1.8vw,16px)">
            最短路径 ${esc(g.start)} → ${esc(g.end)}：<b class="c-op">${opts.path.join(' → ')}</b>，总权重 = <b class="c-op">${opts.total}</b></div>`;
        } else {
          html += `<div class="write-result" style="margin-top:8px;padding:6px 14px;line-height:1.6;font-size:clamp(13px,1.8vw,16px)">从 ${esc(g.start)} 无法到达 ${esc(g.end)}</div>`;
        }
      } else {
        html += `<div class="write-result" style="margin-top:8px;padding:6px 14px;line-height:1.6;font-size:clamp(13px,1.8vw,16px)">
          ${meta.name}结果：[${visited.join(' · ')}]</div>`;
      }
    }
    $('#stage').innerHTML = html;
  },

  /* ================= 步骤生成 ================= */
  genBfsSteps() {
    const meta = this.subMeta();
    const events = this.genBfsEvents();
    const steps = [];
    steps.push({
      desc: `本次演示<b>${meta.name}</b>：从起点 <b class="c-op">${esc(this._g.start)}</b> 出发，先访问所有相邻节点，再借助<b>队列</b>逐层向外扩展。`,
      draw: () => { this.renderWalk({ visited: [], curId: null, waiting: [], done: false }); }
    });
    events.forEach((ev, i) => {
      const snap = ev;
      const newlyLabels = snap.newly.join('、');
      const waitingLabels = snap.waiting.join(' · ');
      steps.push({
        desc: `<b>${meta.name}</b>：取出队首${i === 0 ? '（起点）' : ''} <b class="c-op">${esc(snap.curId)}</b>${snap.newly.length ? `，未访问邻居 ${newlyLabels} 依次入队` : ''}，队列：[${waitingLabels || '空'}]，访问序列：[${snap.visited.join(', ')}]`,
        draw: () => { this.renderWalk({ visited: snap.visited, curId: snap.curId, waiting: snap.waiting, edge: snap.edge, done: false, extra: newlyLabels ? `入队 ${newlyLabels}` : '' }); }
      });
    });
    const finalIds = [...events[events.length - 1].visited];
    steps.push({
      desc: `广度优先搜索完成，最终访问序列：<b class="c-op">[${finalIds.join(', ')}]</b>`,
      draw: () => { this.renderWalk({ visited: finalIds, curId: null, waiting: [], done: true }); }
    });
    return steps;
  },
  genDfsSteps() {
    const meta = this.subMeta();
    const events = this.genDfsEvents();
    const steps = [];
    steps.push({
      desc: `本次演示<b>${meta.name}</b>：从起点 <b class="c-op">${esc(this._g.start)}</b> 出发，借助<b>栈</b>沿一条路径一直深入到底，再回溯访问其他分支。`,
      draw: () => { this.renderWalk({ visited: [], curId: null, waiting: [], done: false }); }
    });
    events.forEach((ev, i) => {
      const snap = ev;
      const newlyLabels = snap.newly.join('、');
      const waitingLabels = snap.waiting.join(' · ');
      steps.push({
        desc: `<b>${meta.name}</b>：弹出栈顶${i === 0 ? '（起点）' : ''} <b class="c-op">${esc(snap.curId)}</b>${snap.newly.length ? `，邻居 ${newlyLabels} 压栈` : ''}，栈（栈底 → 栈顶）：[${waitingLabels || '空'}]，访问序列：[${snap.visited.join(', ')}]`,
        draw: () => { this.renderWalk({ visited: snap.visited, curId: snap.curId, waiting: snap.waiting, edge: snap.edge, done: false, extra: newlyLabels ? `压栈 ${newlyLabels}` : '' }); }
      });
    });
    const finalIds = [...events[events.length - 1].visited];
    steps.push({
      desc: `深度优先搜索完成，最终访问序列：<b class="c-op">[${finalIds.join(', ')}]</b>`,
      draw: () => { this.renderWalk({ visited: finalIds, curId: null, waiting: [], done: true }); }
    });
    return steps;
  },
  genPathSteps() {
    const g = this._g;
    const events = this.genDijkstraEvents();
    const steps = [];
    steps.push({
      desc: `本次演示<b>Dijkstra 最短路径</b>：从起点 <b class="c-op">${esc(g.start)}</b> 到终点 <b class="c-op">${esc(g.end)}</b>，起点距离记 0，其余顶点距离记 ∞；反复取出距离最小的未确定顶点并松弛邻边。`,
      draw: () => { this.renderWalk({ visited: [], curId: null, waiting: [], dist: events[0].dist, done: false }); }
    });
    events.forEach((ev, i) => {
      if (ev.type === 'init') return;
      if (ev.type === 'visit') {
        const relaxTxt = ev.relaxed.length
          ? `，经 <b class="c-op">${esc(ev.curId)}</b> 松弛邻边：${ev.relaxed.map(r => `${esc(r.v)} 更新为 ${r.d}`).join('、')}`
          : '，邻边无需更新';
        steps.push({
          desc: `<b>Dijkstra</b>：取出距离最小的未确定顶点 <b class="c-op">${esc(ev.curId)}</b>（距离 ${ev.dist[ev.curId]}）${relaxTxt}，已确定顺序：[${ev.visited.join(', ')}]`,
          draw: () => { this.renderWalk({ visited: ev.visited, curId: ev.curId, waiting: [], dist: ev.dist, edge: ev.edge, done: false }); }
        });
      } else if (ev.type === 'done') {
        const pathSet = new Set();
        for (let k = 0; k < ev.path.length - 1; k++) {
          const e = this.edgeBetween(ev.path[k], ev.path[k + 1]);
          if (e) pathSet.add(e.key);
        }
        steps.push({
          desc: ev.path.length
            ? `最短路径确定：<b class="c-op">${ev.path.join(' → ')}</b>，总权重 = <b class="c-op">${ev.total}</b>。`
            : `从 ${esc(g.start)} 无法到达 ${esc(g.end)}（距离为 ∞）。`,
          draw: () => { this.renderWalk({ visited: ev.visited, curId: null, waiting: [], dist: ev.dist, pathKeys: pathSet, path: ev.path, total: ev.total, done: true }); }
        });
      }
    });
    return steps;
  },

  initStage() {
    if (!this._g || this._g.nodes.length === 0) { this.renderBuildStage(); return; }
    this.renderWalk({ visited: [], curId: null, waiting: [], dist: null, done: false });
    $('#stage').innerHTML += `<div class="annot" style="text-align:center;margin-top:8px;line-height:1.6">
      使用 <b>${this.subMeta().rule}</b>；点击「自动演示」或「下一步」开始搜索</div>`;
  },

  generate() {
    this.ensureState();
    const g = this._g;
    if (g.nodes.length === 0 || !g.start) { this.steps = []; return; }
    if (this.currentSub === 'path' && !g.end) { this.steps = []; return; }
    if (this.currentSub === 'dfs') this.steps = this.genDfsSteps();
    else if (this.currentSub === 'path') this.steps = this.genPathSteps();
    else this.steps = this.genBfsSteps();
  },

  afterSettings() {
    this.ensureState();
    const g = this._g;
    const msg = (txt, ok) => {
      const m = $('#gf-msg');
      if (m) { m.textContent = txt; m.className = 'chk ' + (ok ? 'good' : 'bad'); }
    };
    const val = id => { const el = $(id); return el ? el.value.trim() : ''; };
    const clear = id => { const el = $(id); if (el) el.value = ''; };

    this.populateSelects();

    const startSel = $('#gf-start');
    if (startSel) startSel.onchange = () => { g.start = startSel.value || null; this.renderBuildStage(); };
    const endSel = $('#gf-end');
    if (endSel) endSel.onchange = () => { g.end = endSel.value || null; this.renderBuildStage(); };

    const addv = $('#gf-addv');
    if (addv) addv.onclick = () => {
      const label = val('#gf-v');
      if (!label) { msg('请先输入顶点名称', false); return; }
      if (this.nodeByLabel(label)) { msg(`顶点「${label}」已存在，名称不能重复`, false); return; }
      const pos = this.autoPos();
      g.nodes.push({ id: label, x: pos.x, y: pos.y });
      if (!g.start) g.start = label;
      if (!g.end) g.end = label;
      this.afterEdit();
      clear('#gf-v');
      msg(`顶点「${label}」已添加，可拖拽调整位置`, true);
    };

    const parseW = w => (w === '' || w == null) ? null : Number(w);
    const addEdge = (a, b, wRaw) => {
      if (!a || !b) { msg('请先选择或填写两个顶点', false); return false; }
      if (a === b) { msg('两个顶点不能是同一个', false); return false; }
      if (!this.node(a) || !this.node(b)) { msg(`顶点不存在：${!this.node(a) ? a : b}`, false); return false; }
      if (this.edgeBetween(a, b)) { msg(`边 ${a} — ${b} 已经存在`, false); return false; }
      if (wRaw !== '' && wRaw != null && isNaN(Number(wRaw))) { msg('权重必须是数字（可留空表示无权边）', false); return false; }
      const w = parseW(wRaw);
      const key = [a, b].sort().join('-');
      g.edges.push({ a, b, w, key });
      g.sel = [];
      return true;
    };

    const adde = $('#gf-adde');
    if (adde) adde.onclick = () => {
      if (g.sel.length !== 2) { msg('请先在画布上依次点击两个顶点（共两个端点）', false); return; }
      const a = g.sel[0], b = g.sel[1];
      const wRaw = val('#gf-w');
      const ok = addEdge(a, b, wRaw);
      if (ok) { this.afterEdit(); clear('#gf-w'); msg(`已添加边 ${a} — ${b}` + (wRaw ? `（权重 ${wRaw}）` : ''), true); }
    };

    const pair = $('#gf-pair');
    if (pair) pair.onclick = () => {
      const a = val('#gf-a'), b = val('#gf-b');
      const ok = addEdge(a, b, val('#gf-w2'));
      if (ok) { this.afterEdit(); clear('#gf-a'); clear('#gf-b'); clear('#gf-w2'); msg(`已添加边 ${a} — ${b}`, true); }
    };

    const preset = $('#gf-preset');
    if (preset) preset.onclick = () => {
      this.clearGraph(); this.ensureState(); this.preset();
      this.afterEdit();
      msg('已载入示例图：A-B(1)、A-C(2)、B-D(3)、C-D(1)', true);
    };
    const clearBtn = $('#gf-clear');
    if (clearBtn) clearBtn.onclick = () => {
      this.clearGraph(); this.afterEdit();
      msg('画布已清空，可重新构建', true);
    };
    const startBtn = $('#gf-startbtn');
    if (startBtn) startBtn.onclick = () => {
      if (g.nodes.length === 0) { msg('请先添加顶点，或点击「示例图」', false); return; }
      if (!g.start) { msg('请先选择遍历起点', false); return; }
      if (this.currentSub === 'path' && !g.end) { msg('最短路径模式请先选择终点', false); return; }
      msg('开始搜索：点击「下一步」逐个访问顶点', true);
      Engine.regenerate();
    };

    this.renderBuildStage();
  }
};
