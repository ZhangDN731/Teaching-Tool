/* ================================================================
 * 树 Tree（可视化自定义构建 · 前序 · 中序 · 后序 · 层序遍历）
 * 构建阶段：输入名称 → 添加根节点 / 选中节点 → 添加子节点，纯 SVG 可视化
 * 遍历阶段：点击「开始遍历」后由 Engine 步骤系统逐步演示四种遍历
 * ================================================================ */
(function () {
  if (document.getElementById('tree-custom-css')) return;
  document.head.insertAdjacentHTML('beforeend', `<style id="tree-custom-css">
    #stage .t-node{cursor:pointer;transition:filter .15s ease}
    #stage .t-node:hover{filter:brightness(.95)}
    .order-chip.chip-done{background:var(--green)}
    .order-chip.chip-cur{background:var(--orange)}
    #stage svg text{font-family:var(--font)}
  </style>`);
})();

MOD.tree = {
  name: '树 Tree',
  sub: '前序 · 中序 · 后序 · 层序遍历',
  badge: '结构 · 遍历', tag: '树', cat: 'ds',
  color: '#EC4899', soft: '#FDF0F7',
  type: '二叉树 · 可视化构建 · 前中后层序遍历',
  desc: '在画布上可视化搭建二叉树（添加根节点、选中节点添加子节点），再演示前序（根左右）、中序（左根右）、后序（左右根）、层序（逐层，借助队列）四种遍历方式。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="4" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="12" r="2.5"/><circle cx="3" cy="20" r="2"/><circle cx="9" cy="20" r="2"/><circle cx="15" cy="20" r="2"/><circle cx="21" cy="20" r="2"/><path d="M12 6.5v3M10.5 11l-3 .5M13.5 11l3 .5M6 14.5v3M5 18l-1.5 1.5M7 18l1.5 1.5M18 14.5v3M17 18l-1.5 1.5M19 18l1.5 1.5"/></svg>',
  setDesc: '在上方输入节点名称，点击按钮构建二叉树；切换顶部标签选择遍历方式，再点击「开始遍历」。',
  stageDesc: '构建模式：蓝色虚线圈 = 当前选中节点。遍历模式：橙色=当前访问，绿色=已访问，蓝色=队列中待访问，灰色=未访问。',
  initDesc: '先构建二叉树（或点击「示例二叉树」），点击「开始遍历」后，用「下一步」逐个观看节点访问顺序。',
  currentSub: 'pre',
  subTabs: [
    { id: 'pre',   label: '前序', method: '根 → 左 → 右' },
    { id: 'in',    label: '中序', method: '左 → 根 → 右' },
    { id: 'post',  label: '后序', method: '左 → 右 → 根' },
    { id: 'level', label: '层序', method: '队列 · 逐层' }
  ],

  theory() {
    return `<div class="rule"><b>二叉树</b>：每个节点最多有两个子节点（左子节点、右子节点）。根节点没有父节点，没有子节点的节点称为叶子。</div>
    <div class="rule"><b>可视化构建</b>：先添加根节点；在画布上点击选中一个父节点，输入子节点名称后点击「添加子节点」，子节点自动挂到空位（先左后右）。</div>
    <div class="rule"><b>前序遍历</b>：根 → 左 → 右。先访问根节点，再前序遍历左子树，最后前序遍历右子树。</div>
    <div class="rule"><b>中序遍历</b>：左 → 根 → 右。先中序遍历左子树，再访问根节点，最后中序遍历右子树。</div>
    <div class="rule"><b>后序遍历</b>：左 → 右 → 根。先后序遍历左子树，再后序遍历右子树，最后访问根节点。</div>
    <div class="rule"><b>层序遍历</b>：从上到下、从左到右逐层访问，需要借助<b>队列</b>（先进先出）实现。</div>`;
  },

  settingsHTML() {
    return `
      <div class="settings-row">
        <div class="inline-setting"><div class="setting"><label>节点名称</label>
          <input type="text" id="tr-label" placeholder="如 A、B、1" style="width:clamp(90px,14vw,130px)"></div></div>
        <div class="inline-setting"><button class="btn primary" id="tr-root">添加根节点</button></div>
        <div class="inline-setting"><button class="btn" id="tr-child">添加子节点</button></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><div class="setting"><label>父节点名</label>
          <input type="text" id="tr-p" placeholder="如 A" style="width:clamp(70px,10vw,100px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>子节点名</label>
          <input type="text" id="tr-c" placeholder="如 D" style="width:clamp(70px,10vw,100px)"></div></div>
        <div class="inline-setting"><button class="btn" id="tr-pair">按名称添加父子</button></div>
      </div>
      <div class="settings-row" style="margin-top:6px">
        <div class="inline-setting"><button class="btn" id="tr-preset">示例二叉树</button></div>
        <div class="inline-setting"><button class="btn ghost" id="tr-clear">清空画布</button></div>
        <div class="inline-setting"><button class="btn primary" id="tr-start">开始遍历</button></div>
      </div>
      <div class="chk" id="tr-msg"></div>`;
  },

  /* ================= 自定义构建状态 ================= */
  ensureState() {
    if (!this._t) {
      this._t = { nodes: [], rootId: null, selected: null, nextId: 1 };
      this.preset();
    }
  },
  findNode(id) { return this._t.nodes.find(n => n.id === id); },
  nodeByLabel(lbl) { return this._t.nodes.find(n => n.label === lbl); },
  newNode(label) {
    const n = { id: this._t.nextId++, label, left: null, right: null, parent: null, x: 0, y: 0, depth: 0 };
    this._t.nodes.push(n);
    return n;
  },
  layout() {
    const t = this._t;
    if (!t.rootId) { t.nodes.forEach(n => { n.x = 400; n.y = 260; }); return; }
    const assign = (id, depth, minX, maxX) => {
      const n = this.findNode(id);
      if (!n) return;
      n.depth = depth;
      n.x = (minX + maxX) / 2;
      n.y = 70 + depth * 145;
      if (n.left != null)  assign(n.left,  depth + 1, minX, n.x);
      if (n.right != null) assign(n.right, depth + 1, n.x, maxX);
    };
    assign(t.rootId, 0, 70, 730);
  },
  preset() {
    const t = this._t;
    const mk = label => { const n = this.newNode(label); return n.id; };
    const A = mk('A'), B = mk('B'), C = mk('C'), D = mk('D'), E = mk('E');
    t.rootId = A;
    const link = (p, c, side) => {
      const pn = this.findNode(p), cn = this.findNode(c);
      cn.parent = p;
      if (side === 'L') pn.left = c; else pn.right = c;
    };
    link(A, B, 'L'); link(A, C, 'R'); link(B, D, 'L'); link(B, E, 'R');
    t.selected = null;
    this.layout();
  },
  clearTree() {
    this._t = { nodes: [], rootId: null, selected: null, nextId: 1 };
  },
  afterEdit() {
    this.layout();
    this.steps = [];
    Engine.renderAll();
    this.renderBuildStage();
    const eb = $('#explainBox'); if (eb) eb.innerHTML = this.initDesc;
  },

  subMeta() {
    return {
      pre:   { name: '前序遍历', rule: '根 → 左 → 右', kind: 'pre' },
      in:    { name: '中序遍历', rule: '左 → 根 → 右', kind: 'in' },
      post:  { name: '后序遍历', rule: '左 → 右 → 根', kind: 'post' },
      level: { name: '层序遍历', rule: '队列 · 逐层从左到右', kind: 'level' }
    }[this.currentSub];
  },
  typeChars(text) {
    return [...String(text)].map((ch, i) => `<span class="char-in" style="animation-delay:${i * 0.05}s">${esc(ch)}</span>`).join('');
  },
  roleText(id, side) {
    const n = this.findNode(id);
    if (id === this._t.rootId) return '根节点';
    const p = this.findNode(n.parent);
    const leaf = (n.left == null && n.right == null);
    return `${p.label} 的${side === 'L' ? '左' : '右'}子节点${leaf ? '（叶子）' : ''}`;
  },

  /* ================= 遍历事件（由当前结构实际计算） ================= */
  genRecursiveEvents(kind) {
    const visited = [], events = [];
    const walk = (id, parentId, side) => {
      if (id == null) return;
      const n = this.findNode(id);
      if (kind === 'pre')  { visited.push(id); events.push({ curId: id, parentId, side, visited: [...visited] }); }
      walk(n.left, id, 'L');
      if (kind === 'in')   { visited.push(id); events.push({ curId: id, parentId, side, visited: [...visited] }); }
      walk(n.right, id, 'R');
      if (kind === 'post') { visited.push(id); events.push({ curId: id, parentId, side, visited: [...visited] }); }
    };
    walk(this._t.rootId, null, null);
    return events;
  },
  genLevelEvents() {
    const visited = [], queue = [this._t.rootId], events = [];
    while (queue.length) {
      const cur = queue.shift();
      visited.push(cur);
      const n = this.findNode(cur), newly = [];
      if (n.left != null)  { queue.push(n.left);  newly.push(n.left); }
      if (n.right != null) { queue.push(n.right); newly.push(n.right); }
      events.push({ curId: cur, parentId: n.parent, side: null, visited: [...visited], queue: [...queue], newly });
    }
    return events;
  },

  /* ================= 构建模式画布 ================= */
  renderBuildStage() {
    const t = this._t;
    let html = `<div class="stage-hd" style="margin-bottom:6px;line-height:1.2;font-size:clamp(12px,1.6vw,14px)">
      <span>自定义二叉树 · 构建中</span><span class="state">${t.nodes.length} 个节点</span></div>`;
    html += `<div class="step-title" style="margin-bottom:8px;padding:4px 10px;line-height:1.3"><span class="st-dot"></span>点击圆形节点可选中，再输入名称添加子节点</div>`;
    html += `<div class="svg-box" style="margin-bottom:8px;display:flex;justify-content:center;background:var(--surface-2)">
      <svg viewBox="0 0 800 520" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:640px">`;

    t.nodes.forEach(n => {
      [n.left, n.right].forEach(c => {
        if (c == null) return;
        const m = this.findNode(c);
        html += `<line x1="${n.x}" y1="${n.y}" x2="${m.x}" y2="${m.y}" stroke="#CBD5E1" stroke-width="3" stroke-linecap="round"/>`;
      });
    });

    if (t.nodes.length === 0) {
      html += `<text x="400" y="260" text-anchor="middle" fill="#94A3B8" font-size="20" font-weight="600">画布为空：在上方输入名称，点击「添加根节点」开始构建</text>`;
    }

    t.nodes.forEach(n => {
      const sel = t.selected === n.id;
      if (sel) html += `<circle cx="${n.x}" cy="${n.y}" r="36" fill="none" stroke="#4F6BFF" stroke-width="3" stroke-dasharray="7 6" opacity=".65"/>`;
      html += `<circle class="t-node" data-id="${n.id}" cx="${n.x}" cy="${n.y}" r="26"
        fill="${sel ? '#EEF3FF' : '#fff'}" stroke="${sel ? '#4F6BFF' : '#CBD5E1'}" stroke-width="3" style="cursor:pointer"/>`;
      html += `<text x="${n.x}" y="${n.y + 7}" text-anchor="middle" font-size="clamp(13px,1.8vw,17px)" font-weight="700"
        fill="${sel ? '#3D56E0' : '#334155'}" style="pointer-events:none">${esc(n.label)}</text>`;
    });
    html += `</svg></div>`;

    const selN = t.selected ? this.findNode(t.selected) : null;
    let info = '尚未选中节点';
    if (selN) {
      const slots = [];
      slots.push(selN.left == null ? '左子空位'  : `左子 = ${this.findNode(selN.left).label}`);
      slots.push(selN.right == null ? '右子空位' : `右子 = ${this.findNode(selN.right).label}`);
      info = `已选中 <b style="color:var(--accent)">${esc(selN.label)}</b>（${slots.join('，')}）`;
    }
    html += `<div class="annot" style="text-align:center;margin-bottom:6px;line-height:1.6">${info}</div>`;
    html += `<div class="zone-label">访问序列</div>
      <div class="order-rack"><span style="color:#CBD5E1;font-size:clamp(12px,1.6vw,14px)">构建完成后点击「开始遍历」，访问序列将显示在这里</span></div>`;

    $('#stage').innerHTML = html;
    $$('#stage .t-node').forEach(c => {
      c.onclick = () => {
        const id = parseInt(c.dataset.id, 10);
        t.selected = (t.selected === id) ? null : id;
        this.afterEdit();
      };
    });
  },

  /* ================= 遍历模式画布 ================= */
  renderWalk(opts) {
    const meta = this.subMeta();
    const visited = opts.visited || [], curId = opts.curId, done = !!opts.done;
    const queue = opts.queue || [], vis = new Set(visited);
    const t = this._t;

    let html = `<div class="stage-hd" style="margin-bottom:6px;line-height:1.2;font-size:clamp(12px,1.6vw,14px)">
      <span>二叉树 · ${meta.name}</span><span class="state">${vis.size}/${t.nodes.length}</span></div>`;
    html += `<div class="step-title" style="margin-bottom:8px;padding:4px 10px;line-height:1.3">
      <span class="st-dot"></span>${done ? '遍历完成' : '节点访问状态'}</div>`;
    html += `<div class="svg-box" style="margin-bottom:8px;display:flex;justify-content:center;background:var(--surface-2)">
      <svg viewBox="0 0 800 520" preserveAspectRatio="xMidYMid meet" style="width:100%;max-width:640px">`;

    t.nodes.forEach(n => {
      [n.left, n.right].forEach(c => {
        if (c == null) return;
        const m = this.findNode(c);
        const both = vis.has(n.id) && vis.has(c);
        html += `<line x1="${n.x}" y1="${n.y}" x2="${m.x}" y2="${m.y}"
          stroke="${both ? '#22C55E' : '#E2E8F0'}" stroke-width="3" stroke-linecap="round"/>`;
      });
    });

    t.nodes.forEach(n => {
      let state = 'un';
      if (curId === n.id) state = 'cur';
      else if (vis.has(n.id)) state = 'done';
      else if (queue.includes(n.id)) state = 'queued';
      const st = {
        un:     { f: '#F1F5F9', s: '#CBD5E1', t: '#64748B' },
        cur:    { f: '#F59E0B', s: '#D97706', t: '#fff' },
        done:   { f: '#22C55E', s: '#16A34A', t: '#fff' },
        queued: { f: '#E0F2FE', s: '#38BDF8', t: '#0369A1' }
      }[state];
      if (state === 'cur')
        html += `<circle cx="${n.x}" cy="${n.y}" r="36" fill="#F59E0B" opacity="0.15" style="animation:pop .3s ease-out"/>`;
      html += `<circle cx="${n.x}" cy="${n.y}" r="26" fill="${st.f}" stroke="${st.s}" stroke-width="3"/>`;
      html += `<text x="${n.x}" y="${n.y + 7}" text-anchor="middle" font-size="clamp(13px,1.8vw,17px)" font-weight="700" fill="${st.t}">${esc(n.label)}</text>`;
    });
    html += `</svg></div>`;

    if (!done) {
      html += `<div class="exp-inline" style="margin:4px 0 8px;padding:4px 10px;line-height:1.6">`;
      if (curId == null) {
        html += `<span class="op">规则</span><span class="term done">${this.typeChars(meta.rule)}</span><span class="caret pen"></span>`;
      } else {
        const node = this.findNode(curId);
        html += `<span class="op">当前操作</span><span class="term cur">访问 ${esc(node.label)}</span>`;
        if (opts.role) html += `<span class="op">· ${opts.role}</span>`;
        if (opts.extra) html += `<span class="op">→</span><span class="term done">${opts.extra}</span>`;
      }
      html += `</div>`;
    }

    html += `<div class="zone-label">访问序列</div><div class="order-rack" style="margin-bottom:8px">`;
    if (visited.length === 0) {
      html += `<span style="color:#CBD5E1;font-size:clamp(12px,1.6vw,14px)">（尚未访问）</span>`;
    } else {
      visited.forEach((id, i) => {
        const n = this.findNode(id);
        const cls = done ? 'chip-done' : (id === curId ? 'chip-cur' : 'chip-done');
        html += `<span class="order-chip ${cls} write-in" style="animation-delay:${i * 0.08}s">${esc(n.label)}</span>`;
      });
    }
    html += `</div>`;

    if (meta.kind === 'level') {
      html += `<div class="zone-label">队列（队首 → 队尾）</div>
        <div class="queue-area" style="min-height:34px;padding:8px 10px 6px;gap:6px;overflow:visible;margin-bottom:8px">`;
      if (queue.length === 0) html += `<span style="color:#94A3B8;font-size:clamp(12px,1.5vw,14px)">（空队列）</span>`;
      else queue.forEach((id, i) => {
        const n = this.findNode(id);
        html += `<span class="chip cy write-in" style="padding:4px 10px;font-size:clamp(12px,1.6vw,14px);animation-delay:${i * 0.08}s">${esc(n.label)}</span>`;
      });
      html += `</div>`;
    }

    html += `<div class="legend" style="margin-top:6px;line-height:1.8;font-size:clamp(11px,1.4vw,13px)">
      <span style="color:#94A3B8">● 未访问</span>　<span style="color:#F59E0B">● 当前</span>
      <span style="color:#22C55E">● 已访问</span>${meta.kind === 'level' ? '　<span style="color:#38BDF8">● 排队中</span>' : ''}
    </div>`;

    if (done) {
      const seq = visited.map(v => this.findNode(v).label).join(' · ');
      html += `<div class="write-result" style="margin-top:8px;padding:6px 14px;line-height:1.6;font-size:clamp(13px,1.8vw,16px)">${meta.name}结果：[${seq}]</div>`;
    }
    $('#stage').innerHTML = html;
  },

  /* ================= 步骤生成 ================= */
  genRecursiveSteps(kind) {
    const meta = this.subMeta();
    const events = this.genRecursiveEvents(kind);
    const steps = [];
    steps.push({
      desc: `本次演示<b>${meta.name}</b>（<b>${meta.rule}</b>）：按规则递归遍历左、右子树，点击「下一步」逐个访问节点。`,
      draw: () => { this.renderWalk({ visited: [], curId: null, queue: [], done: false }); }
    });
    events.forEach(snap => {
      const labels = snap.visited.map(v => this.findNode(v).label);
      const node = this.findNode(snap.curId);
      const role = this.roleText(snap.curId, snap.side);
      steps.push({
        desc: `<b>${meta.name}</b>（${meta.rule}）：访问${role} <b class="c-op">${esc(node.label)}</b>，当前访问序列：[${labels.join(', ')}]`,
        draw: () => { this.renderWalk({ visited: snap.visited, curId: snap.curId, queue: [], done: false, role }); }
      });
    });
    const finalIds = [...events[events.length - 1].visited];
    steps.push({
      desc: `${meta.name}完成，最终访问序列：<b class="c-op">[${finalIds.map(v => this.findNode(v).label).join(', ')}]</b>`,
      draw: () => { this.renderWalk({ visited: finalIds, curId: null, queue: [], done: true }); }
    });
    return steps;
  },
  genLevelSteps() {
    const meta = this.subMeta();
    const events = this.genLevelEvents();
    const steps = [];
    steps.push({
      desc: `本次演示<b>${meta.name}</b>：借助<b>队列</b>从上到下、从左到右逐层访问，先访问根节点 <b class="c-op">${esc(this.findNode(this._t.rootId).label)}</b>。`,
      draw: () => { this.renderWalk({ visited: [], curId: null, queue: [], done: false }); }
    });
    events.forEach((ev, i) => {
      const snap = ev;
      const node = this.findNode(snap.curId);
      const labels = snap.visited.map(v => this.findNode(v).label).join(', ');
      const newlyLabels = snap.newly.map(id => this.findNode(id).label).join('、');
      const queueLabels = snap.queue.map(id => this.findNode(id).label).join(' · ');
      steps.push({
        desc: `<b>${meta.name}</b>：取出队首${i === 0 ? '（起点）' : ''} <b class="c-op">${esc(node.label)}</b>${snap.newly.length ? `，将其未访问子节点 ${newlyLabels} 依次入队` : ''}，队列：[${queueLabels || '空'}]，访问序列：[${labels}]`,
        draw: () => { this.renderWalk({ visited: snap.visited, curId: snap.curId, queue: snap.queue, done: false, extra: newlyLabels ? `入队 ${newlyLabels}` : null }); }
      });
    });
    const finalIds = [...events[events.length - 1].visited];
    steps.push({
      desc: `层序遍历完成，最终访问序列：<b class="c-op">[${finalIds.map(v => this.findNode(v).label).join(', ')}]</b>`,
      draw: () => { this.renderWalk({ visited: finalIds, curId: null, queue: [], done: true }); }
    });
    return steps;
  },

  initStage() {
    if (!this._t || !this._t.rootId) { this.renderBuildStage(); return; }
    this.renderWalk({ visited: [], curId: null, queue: [], done: false });
    $('#stage').innerHTML += `<div class="annot" style="text-align:center;margin-top:8px;line-height:1.6">
      使用 <b>${this.subMeta().rule}</b>；点击「自动演示」或「下一步」开始遍历</div>`;
  },

  generate() {
    this.ensureState();
    const t = this._t;
    if (!t.rootId || t.nodes.length === 0) { this.steps = []; return; }
    if (this.currentSub === 'level') this.steps = this.genLevelSteps();
    else this.steps = this.genRecursiveSteps(this.currentSub);
  },

  afterSettings() {
    this.ensureState();
    const msg = (txt, ok) => {
      const m = $('#tr-msg');
      if (m) { m.textContent = txt; m.className = 'chk ' + (ok ? 'good' : 'bad'); }
    };
    const val = id => { const el = $(id); return el ? el.value.trim() : ''; };
    const clear = id => { const el = $(id); if (el) el.value = ''; };
    const t = this._t;

    const rootBtn = $('#tr-root');
    if (rootBtn) rootBtn.onclick = () => {
      const label = val('#tr-label');
      if (!label) { msg('请先输入节点名称', false); return; }
      if (t.rootId) { msg('已存在根节点：请先清空画布，或选中节点后「添加子节点」', false); return; }
      if (this.nodeByLabel(label)) { msg(`节点「${label}」已存在，名称不能重复`, false); return; }
      const n = this.newNode(label);
      t.rootId = n.id; t.selected = n.id;
      this.afterEdit();
      msg(`根节点「${label}」已创建`, true);
    };

    const childBtn = $('#tr-child');
    if (childBtn) childBtn.onclick = () => {
      const label = val('#tr-label');
      if (t.selected == null) { msg('请先在画布上点击选中一个父节点', false); return; }
      const p = this.findNode(t.selected);
      if (!p) { t.selected = null; msg('请重新选中父节点', false); return; }
      if (!label) { msg('请先输入子节点名称', false); return; }
      if (this.nodeByLabel(label)) { msg(`节点「${label}」已存在，名称不能重复`, false); return; }
      let slot = null;
      if (p.left == null) slot = 'L'; else if (p.right == null) slot = 'R';
      if (!slot) { msg(`节点「${p.label}」左右子均已满，无法再添加`, false); return; }
      const c = this.newNode(label);
      c.parent = p.id;
      if (slot === 'L') p.left = c.id; else p.right = c.id;
      t.selected = p.id;
      this.afterEdit();
      clear('#tr-label');
      msg(`已添加：${p.label} → ${c.label}（${slot === 'L' ? '左子' : '右子'}）`, true);
    };

    const pairBtn = $('#tr-pair');
    if (pairBtn) pairBtn.onclick = () => {
      const pl = val('#tr-p'), cl = val('#tr-c');
      if (!pl || !cl) { msg('请填写父节点和子节点名称', false); return; }
      const p = this.nodeByLabel(pl);
      if (!p) { msg(`未找到父节点「${pl}」`, false); return; }
      if (this.nodeByLabel(cl)) { msg(`节点「${cl}」已存在，名称不能重复`, false); return; }
      let slot = null;
      if (p.left == null) slot = 'L'; else if (p.right == null) slot = 'R';
      if (!slot) { msg(`节点「${pl}」左右子均已满`, false); return; }
      const c = this.newNode(cl);
      c.parent = p.id;
      if (slot === 'L') p.left = c.id; else p.right = c.id;
      t.selected = c.id;
      this.afterEdit();
      msg(`已添加：${p.label} → ${c.label}`, true);
    };

    const presetBtn = $('#tr-preset');
    if (presetBtn) presetBtn.onclick = () => {
      this.clearTree(); this.ensureState(); this.preset();
      this.afterEdit();
      msg('已载入示例二叉树：A(B(D,E),C)', true);
    };
    const clearBtn = $('#tr-clear');
    if (clearBtn) clearBtn.onclick = () => {
      this.clearTree(); this.afterEdit();
      msg('画布已清空，可重新构建', true);
    };
    const startBtn = $('#tr-start');
    if (startBtn) startBtn.onclick = () => {
      if (!t.rootId) { msg('请先添加根节点，或点击「示例二叉树」', false); return; }
      msg('开始遍历：点击「下一步」逐个访问节点', true);
      Engine.regenerate();
    };

    this.renderBuildStage();
  }
};
