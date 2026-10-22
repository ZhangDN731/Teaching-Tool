
;;;

/* ================================================================
 * ④ IP 地址与子网掩码（分类 + 全套子网计算）
 * ================================================================ */
function ipToArr(ip){return ip.split('.').map(Number);}
function arrToIp(a){return a.join('.');}
function ipAnd(a,b){return a.map((v,i)=>v&b[i]);}
function ipOr(a,b){return a.map((v,i)=>v|b[i]);}
function ipNot(a){return a.map(v=>255-v);}
function ipAdd(a,n){let v=a[0]*16777216+a[1]*65536+a[2]*256+a[3]+n;return [(v>>>24)&255,(v>>>16)&255,(v>>>8)&255,v&255];}
function prefixToMask(p){const m=[];for(let i=0;i<4;i++){const s=Math.max(0,Math.min(8,p-i*8));m.push(s===8?255:s===0?0:256-Math.pow(2,8-s));}return m;}
function ipClass(a){const f=a[0];if(f>=1&&f<=126)return['A','255.0.0.0','/8'];if(f===127)return['A(回环)','255.0.0.0','/8'];if(f>=128&&f<=191)return['B','255.255.0.0','/16'];if(f>=192&&f<=223)return['C','255.255.255.0','/24'];if(f>=224&&f<=239)return['D(组播)','—','—'];return['E(保留)','—','—'];}

/* ================================================================
 * IP 地址与子网掩码
 * 子页面：classify 地址分类 / mask 子网掩码 / network 网络地址 / hostrange 主机范围
 * 复用全局辅助函数：ipToArr, arrToIp, ipAnd, ipOr, ipNot, ipAdd, prefixToMask, ipClass
 * ================================================================ */
/* ================================================================
 * IP 地址与子网掩码
 * 子页面：classify 地址分类 / mask 子网掩码 / network 网络地址 / hostrange 主机范围
 * 复用全局辅助函数：ipToArr, arrToIp, ipAnd, ipOr, ipNot, ipAdd, prefixToMask, ipClass
 * ================================================================ */
/* ================================================================
 * IP 地址与子网掩码
 * 子页面：classify 地址分类 / mask 子网掩码 / network 网络地址 / hostrange 主机范围
 * 复用全局辅助函数：ipToArr, arrToIp, ipAnd, ipOr, ipNot, ipAdd, prefixToMask, ipClass
 * ================================================================ */
/* ================================================================
 * IP 地址与子网掩码
 * 子页面：classify 地址分类 / mask 子网掩码 / network 网络地址 / hostrange 主机范围
 * 复用全局辅助函数（保留在 index.html 中）：ipToArr, arrToIp, ipAnd, ipOr, ipNot, ipAdd, prefixToMask, ipClass
 * 网络地址 / 主机范围统一使用 32 位二进制按 8 位一组加点分隔展示：
 * 网络位蓝色、主机位绿色，AND 结果 / 变化位橙色，两行布局对齐一致。
 * ================================================================ */
if(!document.getElementById('mod-ip-css')){
  document.head.insertAdjacentHTML('beforeend', `<style id="mod-ip-css">
    .bin32-row{display:flex;align-items:center;gap:clamp(8px,1.4vw,12px);margin:clamp(6px,1vw,10px) 0}
    .bin32-label{width:clamp(72px,10vw,96px);flex:none;font-size:clamp(11px,1.4vw,13px);font-weight:700;color:var(--text-3);text-align:right;font-family:var(--font)}
    .bin32-groups{flex:1;display:flex;align-items:center;gap:clamp(6px,1.2vw,10px);min-width:0}
    .grp32{display:grid;grid-template-columns:repeat(8,1fr);gap:clamp(2px,.5vw,4px);flex:1;min-width:0}
    .bin-dot{font-family:var(--mono);font-weight:800;font-size:clamp(16px,2.2vw,22px);color:var(--text-3);flex:none}
    .grp32 .bd32{background:#fff;border:1px solid var(--border);border-radius:6px;text-align:center;font-family:var(--mono);font-size:clamp(14px,2vw,18px);font-weight:700;padding:clamp(5px,1vw,9px) 0;color:var(--text-2)}
    .grp32 .bd32.nb{color:#4F6BFF;background:#EEF3FF;border-color:#C7D2FE}
    .grp32 .bd32.hb{color:#15803D;background:#E8FBF0;border-color:#BFE9CF}
    .grp32 .bd32.hb-zero{color:#CBD5E1;background:#F8FAFC;border-color:#E2E8F0}
    .grp32 .bd32.hb-hot{color:#B45309;background:#FFF4DB;border-color:var(--orange)}
    .grp32 .bd32.rs{color:#B45309;background:#FFF4DB;border-color:var(--orange)}
    .bin32-caption{font-family:var(--mono);font-size:clamp(14px,2vw,18px);font-weight:700;text-align:center;margin:clamp(4px,1vw,8px) 0;word-break:break-all;letter-spacing:1px;color:var(--text-2)}
    .bin32-caption .nb-t{color:#4F6BFF}
    .bin32-caption .hb-t{color:#15803D}
    .bin32-caption .hb-zero-t{color:#CBD5E1}
    .bin32-caption .hb-hot-t{color:#B45309}
    .bin32-caption .bin-dot-t{color:var(--text-3);margin:0 1px}
    .bin32-legend{display:flex;gap:clamp(10px,2vw,18px);justify-content:center;flex-wrap:wrap;font-size:clamp(11px,1.4vw,12.5px);color:var(--text-3);margin:6px 0}
    .bin32-legend .lg{display:inline-flex;align-items:center;gap:5px}
    .bin32-legend .sw{width:12px;height:12px;border-radius:4px;display:inline-block}
  </style>`);
}

MOD.ip = {
  name:'IP与子网掩码', sub:'IPv4分类 · 子网掩码 · 网络地址 · 主机范围',
  badge:'网络 · 子网', tag:'网络', cat:'net',
  color:'#06B6D4', soft:'#E6F9FE', type:'IPv4分类 · 前缀掩码 · 按位与 · 广播与主机范围',
  desc:'IPv4 地址分类、子网掩码二进制推导、网络地址按位与计算、广播地址与可用主机范围及主机数，逐步手写演示；网络位蓝色、主机位绿色，32 位二进制按 8 位一组加点分隔。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15 15 0 0 1 0 20M12 2a15 15 0 0 0 0 20"/></svg>',
  setDesc:'输入 IP 地址并选择子网前缀；各子页面独立演示对应计算。',
  stageDesc:'蓝色=网络位，绿色=主机位，橙色=AND 结果或变化位；32 位二进制按 8 位一组加点分隔，与点分十进制格式一致。',
  initDesc:'点击「自动演示」或「下一步」，观看 IP 分类与子网计算的手写推导。',
  currentSub:'classify',
  subTabs:[
    {id:'classify', label:'地址分类', method:'按第一段判定'},
    {id:'mask',     label:'子网掩码', method:'前缀转二进制'},
    {id:'network',  label:'网络地址', method:'按位与'},
    {id:'hostrange',label:'主机范围', method:'广播与主机数'}
  ],
  theory(){
    return `<div class="rule"><b>IP 地址组成</b>：32 位二进制，分 4 段（每段 8 位），用点分十进制表示，如 192.168.1.10；二进制同样按 8 位一组加点书写，如 <span class="mono">11000000.10101000.00000001.00000000</span>。</div>
    <div class="rule"><b>IP 分类</b>：A 类（1~126，默认掩码 255.0.0.0）、B 类（128~191，255.255.0.0）、C 类（192~223，255.255.255.0）、D 类组播（224~239）、E 类保留（240~255）。</div>
    <div class="rule"><b>子网掩码</b>：连续的 1 后跟连续的 0，1 对应网络位，0 对应主机位，用 /前缀长度 表示，如 /24。</div>
    <div class="rule"><b>网络地址</b> = IP 地址 <b>按位与</b> 子网掩码（主机位全 0）。</div>
    <div class="rule"><b>广播地址</b> = 网络地址的主机位全 1（网络地址 | 反掩码）。</div>
    <div class="rule"><b>可用主机范围</b>：网络地址 +1 ~ 广播地址 −1；可用主机数 = 2<sup>主机位数</sup> − 2（减去网络地址和广播地址）。</div>`;
  },

  /* ---------- 参数区 ---------- */
  settingsHTML(){
    const sub = this.currentSub;
    let html = '';
    if(sub !== 'mask'){
      const oldIp = $('#ip-addr');
      const cur = oldIp ? oldIp.value : (this._ip || '');
      html += `<div class="inline-setting"><div class="setting"><label>IP 地址</label>
        <input type="text" id="ip-addr" value="${esc(cur)}" placeholder="如 192.168.1.10" style="width:clamp(130px,20vw,170px)"></div></div>`;
    }
    if(sub !== 'classify'){
      const oldPf = $('#ip-prefix');
      const cur = oldPf ? oldPf.value : String(this._prefix || 24);
      const opts = [8,16,24,25,26,27,28,29,30];
      html += `<div class="inline-setting"><div class="setting"><label>子网前缀</label>
        <select id="ip-prefix">${opts.map(p=>`<option value="${p}" ${String(p)===cur?'selected':''}>/${p}</option>`).join('')}</select></div></div>`;
    }
    return html;
  },
  afterSettings(){
    const ipEl = $('#ip-addr');
    if(ipEl && !ipEl.value.trim()){
      this._ip = this._ip || this._randIp();
      ipEl.value = this._ip;
    }
    if(ipEl) ipEl.onchange = ()=>{ this._ip = ipEl.value.trim(); Engine.reload(); };
    const pfEl = $('#ip-prefix');
    if(pfEl) pfEl.onchange = ()=>{ this._prefix = +pfEl.value; Engine.reload(); };
  },

  /* ---------- 内部辅助 ---------- */
  segBin(v){ return v.toString(2).padStart(8,'0'); },
  /* 位权展开：返回位权数组、二进制位、求和结果、被激活的位权 */
  weightExpand(binStr){
    const weights = [128,64,32,16,8,4,2,1];
    const bits = binStr.split('');
    let sum = 0;
    const active = [];
    bits.forEach((b,i)=>{ if(b==='1'){ sum += weights[i]; active.push(weights[i]); } });
    return { weights, bits, sum, active };
  },
  /* 生成位权展开计算卡片 HTML */
  weightCard(binStr){
    const { weights, bits, sum, active } = this.weightExpand(binStr);
    const weightRow = weights.map(w=>`<div class="bd weight">${w}</div>`).join('');
    const bitRow = bits.map((b,i)=>`<div class="bd ${b==='1'?'one':'zero'} write-in" style="animation-delay:${i*0.06}s">${b}</div>`).join('');
    const valRow = bits.map((b,i)=>`<div class="bd ${b==='1'?'cur':'empty'}" style="animation-delay:${0.5+i*0.05}s">${b==='1'?weights[i]:'0'}</div>`).join('');
    const sumExpr = active.length ? active.join(' + ') : '0';
    return `<div class="calc-process">
      <div class="cp-label">位权展开求和（从左到右位权：128 → 1）</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">位权</span>
        <div class="bin-grid" style="flex:1">${weightRow}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">二进制</span>
        <div class="bin-grid" style="flex:1">${bitRow}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">取值</span>
        <div class="bin-grid" style="flex:1">${valRow}</div>
      </div>
      <div class="cp-sum-line">${sumExpr} <span class="eq">=</span> <span class="sum">${sum}</span></div>
    </div>`;
  },

  /* 32 位二进制行：4 个 8 位分组、组间加点；网络位蓝、主机位绿，布局全模块统一 */
  binRow32(arr, prefix, opts){
    opts = opts || {};
    const role = opts.role || 'ip';              // ip | mask | net | broad | result
    const label = opts.label || '';
    const delayBase = opts.delayBase || 0;
    const highlightHost = !!opts.highlightHost;
    const groups = [];
    for(let oct=0; oct<4; oct++){
      const bits = this.segBin(arr[oct]).split('');
      groups.push(bits.map((b,k)=>{
        const pos = oct*8+k;
        const isNet = pos < prefix;
        let cls;
        if(role==='mask'){ cls = isNet ? 'nb' : 'hb-zero'; }
        else if(role==='result'){ cls = 'rs'; }
        else if(role==='broad' && highlightHost && !isNet){ cls = 'hb-hot'; }
        else { cls = isNet ? 'nb' : 'hb'; }
        return `<div class="bd32 ${cls} write-in" style="animation-delay:${delayBase+pos*0.018}s">${b}</div>`;
      }).join(''));
    }
    const groupsHtml = groups.map((g,i)=>`<div class="grp32">${g}</div>${i<3?'<div class="bin-dot">.</div>':''}`).join('');
    return `<div class="bin32-row">
      <span class="bin32-label">${label}</span>
      <div class="bin32-groups">${groupsHtml}</div>
    </div>`;
  },
  /* 点分格式二进制字符串（逐位着色：网络位蓝、主机位绿/橙，8 位一组加点） */
  binDottedColored(arr, prefix, role){
    const segs = arr.map((v, oct)=>{
      const bits = this.segBin(v).split('');
      return bits.map((b,k)=>{
        const pos = oct*8+k;
        const isNet = pos < prefix;
        let cls;
        if(role==='mask'){ cls = isNet ? 'nb-t' : 'hb-zero-t'; }
        else if(role==='broad'){ cls = isNet ? 'nb-t' : 'hb-hot-t'; }
        else { cls = isNet ? 'nb-t' : 'hb-t'; }
        return `<span class="${cls}">${b}</span>`;
      }).join('');
    });
    return segs.join('<span class="bin-dot-t">.</span>');
  },
  /* 图例：网络位 / 主机位 / 结果色 */
  legendHTML(){
    return `<div class="bin32-legend">
      <span class="lg"><span class="sw" style="background:#EEF3FF;border:1px solid #C7D2FE"></span>网络位（蓝）</span>
      <span class="lg"><span class="sw" style="background:#E8FBF0;border:1px solid #BFE9CF"></span>主机位（绿）</span>
      <span class="lg"><span class="sw" style="background:#FFF4DB;border:1px solid var(--orange)"></span>AND 结果 / 变化位（橙）</span>
    </div>`;
  },

  _randIp(){
    return [randInt(1,254), randInt(0,255), randInt(0,255), randInt(1,254)].join('.');
  },
  _parseIp(){
    const el = $('#ip-addr');
    let s = el ? el.value.trim() : '';
    if(!s){ s = this._ip || this._randIp(); if(el) el.value = s; }
    const parts = s.split('.');
    if(parts.length !== 4) throw new Error('IP 必须为 4 段点分十进制，如 192.168.1.10');
    parts.forEach(p=>{
      if(!/^\d{1,3}$/.test(p)) throw new Error(`段「${p}」不是数字`);
      if(+p < 0 || +p > 255) throw new Error(`段「${p}」超出 0~255`);
    });
    const arr = ipToArr(s);
    this._ip = arrToIp(arr);
    return { str:this._ip, arr };
  },
  _readPrefix(){
    const el = $('#ip-prefix');
    let v = el ? el.value : '24';
    if(v === 'auto' || v === '' || v == null) v = 24;
    this._prefix = +v;
    return this._prefix;
  },
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },
  classRows(){
    return [
      {key:'A', range:'1 ~ 126', mask:'255.0.0.0', prefix:'/8',  note:''},
      {key:'B', range:'128 ~ 191', mask:'255.255.0.0', prefix:'/16', note:''},
      {key:'C', range:'192 ~ 223', mask:'255.255.255.0', prefix:'/24', note:''},
      {key:'D', range:'224 ~ 239', mask:'—', prefix:'—', note:'组播地址，无默认主机掩码'},
      {key:'E', range:'240 ~ 255', mask:'—', prefix:'—', note:'保留地址，无默认主机掩码'}
    ];
  },
  classifyOf(arr){
    const cls = ipClass(arr);
    const key = cls[0][0];
    const rows = this.classRows();
    const row = rows.find(r=>r.key===key) || rows[0];
    const note = cls[0].includes('(') ? '127.x 为回环地址（仍属 A 类保留段）' : row.note;
    return Object.assign({}, row, { mask:cls[1], prefix:cls[2], note });
  },
  ipRow(arr, curSeg, label, doneAll){
    return `<div class="ip-row">${arr.map((v,i)=>`<div class="seg ${i===curSeg?'cur':doneAll?'done':''} ${i===curSeg?'write-in':''}"><div class="seg-hd">${label||('第 '+(i+1)+' 段')}</div><div class="val">${v}</div></div>`).join('')}</div>`;
  },

  /* ========== ① classify：IPv4 地址分类 ========== */
  genClassify(){
    const { str, arr } = this._parseIp();
    const f = arr[0];
    const cat = this.classifyOf(arr);
    const rows = this.classRows();
    const label = 'IPv4 地址分类';
    const steps = [];
    // 1. 逐段写出 IP 地址
    steps.push({ desc:`写出待分类的 IPv4 地址：<b class="c-op">${str}</b>`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(18px,3vw,28px) 0">
            <span style="font-family:var(--mono);font-size:clamp(22px,4vw,32px);font-weight:700;color:var(--text)">${this.typeChars(str)}</span>
            <span class="caret pen"></span>
          </div>`;
      }});
    // 2. 高亮第一段
    steps.push({ desc:`IPv4 只依据<b>第一段</b>判断类别，当前第一段 = <b class="c-op">${f}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">观察第一段</span></div>
          <div class="ip-row">${arr.map((v,i)=>`<div class="seg ${i===0?'cur':''} write-in" style="animation-delay:${i*0.08}s"><div class="seg-hd">第 ${i+1} 段</div><div class="val">${v}</div></div>`).join('')}</div>
          <div class="annot" style="text-align:center">分类只看<b>第一段</b>的取值范围，后三段不影响类别</div>`;
      }});
    // 3. 对照表判定类别
    steps.push({ desc:`对照分类范围：第一段 ${f} 落在 <b>${cat.range}</b>，属于 <b class="c-op">${cat.key} 类</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">对照分类表</span></div>
          <div class="step-title"><span class="st-dot"></span>IPv4 分类对照表（看第一段）</div>
          <table class="exp-table">
            <tr><th>类别</th><th>第一段范围</th><th>默认掩码</th><th>前缀</th></tr>
            ${rows.map((r,i)=>`<tr class="write-in" style="animation-delay:${i*0.07}s">
              <td class="${r.key===cat.key?'cur':'done'}" style="font-weight:700">${r.key} 类</td>
              <td class="${r.key===cat.key?'cur':''}">${r.range}</td>
              <td class="${r.key===cat.key?'cur':''}">${r.mask}</td>
              <td class="${r.key===cat.key?'cur':''}">${r.prefix}</td>
            </tr>`).join('')}
          </table>`;
      }});
    // 4. 写出默认掩码结果
    steps.push({ desc:`${cat.key} 类默认子网掩码 = <b class="c-op">${cat.mask}</b>${cat.prefix!=='—'?`（${cat.prefix}）`:''}${cat.note?`，${cat.note}`:''}。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          ${this.ipRow(arr,-1,'',true)}
          <div class="write-result">${str} → ${cat.key} 类 · 默认掩码 ${cat.mask}${cat.prefix!=='—'?` (${cat.prefix})`:''}</div>`;
      }});
    return steps;
  },

  /* ========== ② mask：子网掩码计算 ========== */
  genMask(){
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const hostBits = 32 - prefix;
    const bin = mask.map(v=>this.segBin(v));
    const label = '子网掩码计算';
    const steps = [];
    // 1. 写出前缀
    steps.push({ desc:`子网前缀 <b class="c-op">/${prefix}</b>：前 ${prefix} 位为网络位（1），后 ${hostBits} 位为主机位（0）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(16px,3vw,24px) 0">
            <span style="font-family:var(--mono);font-size:clamp(22px,4vw,30px);font-weight:700;color:var(--text)">/${this.typeChars(String(prefix))}</span>
            <span class="caret pen"></span>
          </div>
          <div class="flow-line" style="justify-content:center">
            <span class="part">网络位 <b class="res">${prefix}</b> 位</span><span class="part">主机位 <b class="res">${hostBits}</b> 位</span>
          </div>`;
      }});
    // 2. 32 位二进制逐位填写（前 N 位 1，后 M 位 0）
    steps.push({ desc:`写出掩码二进制：前 <b>${prefix}</b> 位填 1（网络位），后 <b>${hostBits}</b> 位填 0（主机位）。`,
      draw:()=>{
        let bits = '';
        for(let oct=0; oct<4; oct++){
          for(let b=0; b<8; b++){
            const pos = oct*8+b;
            const isNet = pos < prefix;
            bits += `<div class="bd ${isNet?'mask':''} write-in" style="animation-delay:${pos*0.018}s">${isNet?'1':'0'}</div>`;
          }
        }
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">32 位二进制</span></div>
          <div class="step-title"><span class="st-dot"></span>二进制掩码（蓝=网络位 1，灰=主机位 0）</div>
          <div class="bin-grid">${bits}</div>`;
      }});
    // 3-6. 逐段二进制 → 十进制（位权展开求和）
    for(let i=0; i<4; i++){
      const oct = i;
      const { active, sum } = this.weightExpand(bin[oct]);
      const sumExpr = active.length ? active.join(' + ') : '0';
      steps.push({ desc:`第 ${oct+1} 段：二进制 <b class="c-op">${bin[oct]}</b>，按位权展开 ${sumExpr} = <b class="c-op">${mask[oct]}</b>。`,
        draw:()=>{
          const gridBits = mask.map((v,j)=>this.segBin(v).split('').map((b,k)=>
            `<div class="bd ${j<oct?'net':j===oct?'cur':'empty'} ${j===oct?'write-in':''}" style="${j===oct?`animation-delay:${k*0.05}s`:''}">${j<=oct?b:'?'}</div>`).join('')).join('');
          $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">第 ${oct+1}/4 段</span></div>
            <div class="step-title"><span class="st-dot"></span>二进制 → 点分十进制（位权展开）</div>
            <div class="bin-grid">${gridBits}</div>
            ${this.weightCard(bin[oct])}
            <div class="ip-row">${mask.map((v,j)=>`<div class="seg ${j<oct?'done':j===oct?'cur write-in':''}"><div class="seg-hd">第 ${j+1} 段</div><div class="val">${j<=oct?v:'?'}</div></div>`).join('')}</div>`;
        }});
    }
    // 7. 拼接点分十进制结果
    steps.push({ desc:`四段拼接得子网掩码：<b class="c-op">${arrToIp(mask)}</b>（/${prefix}），网络位 ${prefix} 位、主机位 ${hostBits} 位。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>子网掩码</div>
          ${this.ipRow(mask,-1,'掩码',true)}
          <div class="write-result">/${prefix} → 掩码 ${arrToIp(mask)} · 网络位 ${prefix} 位 / 主机位 ${hostBits} 位</div>`;
      }});
    return steps;
  },

  /* ========== ③ network：网络地址计算 ========== */
  /* 步骤：IP 二进制 → 掩码二进制 → 逐位 AND → 网络地址二进制 → 转十进制
   * 展示与主机范围页完全一致：8 位分组加点、网络位蓝 / 主机位绿、相同布局对齐 */
  genNetwork(){
    const { str, arr:ip } = this._parseIp();
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const net = ipAnd(ip, mask);
    const hostBits = 32 - prefix;
    const label = '网络地址计算';
    const steps = [];
    // 1. 写出 IP 与前缀
    steps.push({ desc:`IP 地址 <b class="c-op">${str}</b>，前缀 /${prefix}，对应掩码 ${arrToIp(mask)}；网络地址 = IP 按位与 子网掩码。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(14px,2.5vw,22px) 0">
            <span style="font-family:var(--mono);font-size:clamp(20px,3.5vw,28px);font-weight:700;color:var(--text)">${this.typeChars(str)}<sub style="font-size:.45em;color:var(--text-3)"> /${prefix}</sub></span>
            <span class="caret pen"></span>
          </div>
          ${this.legendHTML()}`;
      }});
    // 2. IP 地址二进制（8 位一组加点）
    steps.push({ desc:`把 IP 地址 ${str} 写成 32 位二进制，按 8 位一组用点分隔；前 ${prefix} 位为网络位（蓝），后 ${hostBits} 位为主机位（绿）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">IP 二进制</span></div>
          <div class="step-title"><span class="st-dot"></span>IP 地址 32 位二进制（8 位一组，加点分隔）</div>
          ${this.binRow32(ip, prefix, {label:'IP', role:'ip'})}
          <div class="bin32-caption">${this.binDottedColored(ip, prefix, 'ip')}</div>
          ${this.legendHTML()}`;
      }});
    // 3. 子网掩码二进制（8 位一组加点）
    steps.push({ desc:`子网掩码 ${arrToIp(mask)} 的 32 位二进制：前 ${prefix} 位为 1（网络位，蓝），后 ${hostBits} 位为 0（主机位，灰）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">掩码二进制</span></div>
          <div class="step-title"><span class="st-dot"></span>子网掩码 32 位二进制（8 位一组，加点分隔）</div>
          ${this.binRow32(mask, prefix, {label:'掩码', role:'mask'})}
          <div class="bin32-caption">${this.binDottedColored(mask, prefix, 'mask')}</div>
          ${this.legendHTML()}`;
      }});
    // 4. 按位与运算（逐位展示，三行对齐）
    steps.push({ desc:`逐位按位与（AND）：只有 1 AND 1 = 1，其余均为 0；掩码为 1 的位保留 IP 原值，掩码为 0 的位清 0（主机位全 0）。`,
      draw:()=>{
        const table = `<table class="exp-table" style="margin-top:10px">
          <tr><th>项目</th>${[0,1,2,3].map(k=>`<th>第 ${k+1} 段</th>`).join('')}</tr>
          <tr><th>IP</th>${ip.map(v=>`<td class="done">${v}</td>`).join('')}</tr>
          <tr><th>掩码</th>${mask.map(v=>`<td class="done">${v}</td>`).join('')}</tr>
          <tr><th>AND 结果</th>${net.map(v=>`<td class="cur">${v}</td>`).join('')}</tr>
        </table>`;
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">逐位 AND</span></div>
          <div class="step-title"><span class="st-dot"></span>IP AND 掩码（逐位计算，8 位一组对齐）</div>
          <div class="calc-process" style="margin-top:6px">
            <div class="cp-label">逐位按位与：1 AND 1 = 1，1 AND 0 = 0，0 AND 1 = 0，0 AND 0 = 0</div>
            ${this.binRow32(ip, prefix, {label:'IP', role:'ip'})}
            ${this.binRow32(mask, prefix, {label:'掩码', role:'mask'})}
            ${this.binRow32(net, prefix, {label:'结果', role:'result'})}
          </div>
          ${table}
          ${this.legendHTML()}`;
      }});
    // 5. 网络地址二进制（主机位全 0）
    steps.push({ desc:`按位与结果即网络地址二进制：网络位保持不变，主机位全部清 0（绿色 0）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">网络地址二进制</span></div>
          <div class="step-title"><span class="st-dot"></span>网络地址 32 位二进制（主机位全 0）</div>
          ${this.binRow32(net, prefix, {label:'网络地址', role:'net'})}
          <div class="bin32-caption">${this.binDottedColored(net, prefix, 'net')}</div>
          ${this.legendHTML()}`;
      }});
    // 6. 网络地址二进制 → 十进制
    steps.push({ desc:`网络地址二进制转点分十进制 = <b class="c-op">${arrToIp(net)}</b>（主机位全 0），/${prefix}。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>网络地址（点分十进制）</div>
          ${this.ipRow(net,-1,'网络地址',true)}
          <div class="write-result">网络地址 = ${arrToIp(net)} /${prefix}</div>`;
      }});
    return steps;
  },

  /* ========== ④ hostrange：主机范围与数量 ========== */
  /* 二进制展示与网络地址页完全一致：8 位分组加点、网络位蓝 / 主机位绿、相同布局对齐 */
  genHostrange(){
    const { str, arr:ip } = this._parseIp();
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const net = ipAnd(ip, mask);
    const notMask = ipNot(mask);
    const broad = ipOr(net, notMask);
    const first = ipAdd(net, 1);
    const last = ipAdd(broad, -1);
    const hostBits = 32 - prefix;
    const count = Math.pow(2, hostBits) - 2;
    const label = '主机范围与数量';
    const steps = [];
    // 1. 网络地址（统一二进制展示，与网络地址页一致）
    steps.push({ desc:`先求网络地址：${str} AND ${arrToIp(mask)} = <b class="c-op">${arrToIp(net)}</b>（主机位全 0）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">网络地址</span></div>
          <div class="step-title"><span class="st-dot"></span>网络地址（主机位全 0）</div>
          ${this.binRow32(net, prefix, {label:'网络地址', role:'net'})}
          <div class="bin32-caption">${this.binDottedColored(net, prefix, 'net')}</div>
          ${this.ipRow(net,-1,'网络地址',true)}
          <div class="write-result">网络地址 = ${arrToIp(net)} /${prefix}</div>
          ${this.legendHTML()}`;
      }});
    // 2. 广播地址：主机位全 1（逐位 OR，主机位橙色高亮变化）
    steps.push({ desc:`广播地址 = 网络地址的<b>主机位全 1</b>：网络位保持不变，主机位 0→1，得到 <b class="c-op">${arrToIp(broad)}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">广播地址</span></div>
          <div class="step-title"><span class="st-dot"></span>主机位全 1 → 广播地址（32 位逐位 OR）</div>
          <div class="calc-process">
            <div class="cp-label">网络位（前 ${prefix} 位）保持不变，主机位（后 ${hostBits} 位）全部置 1</div>
            ${this.binRow32(net, prefix, {label:'网络地址', role:'net'})}
            ${this.binRow32(broad, prefix, {label:'广播地址', role:'broad', highlightHost:true})}
          </div>
          <div class="bin32-caption">${this.binDottedColored(broad, prefix, 'broad')}</div>
          <table class="exp-table" style="margin-top:8px">
            <tr><th>项目</th>${[0,1,2,3].map(k=>`<th>段 ${k+1}</th>`).join('')}</tr>
            <tr><th>网络地址</th>${net.map(v=>`<td class="done">${v}</td>`).join('')}</tr>
            <tr><th>广播地址</th>${broad.map(v=>`<td class="cur">${v}</td>`).join('')}</tr>
          </table>
          <div class="exp-inline" style="margin-top:8px"><span class="term done">${arrToIp(net)}</span><span class="op">OR</span><span class="term">${arrToIp(notMask)}</span><span class="op">=</span><span class="term cur">${arrToIp(broad)}</span></div>
          ${this.legendHTML()}`;
      }});
    // 3. 可用主机范围（首台 / 末台主机二进制，统一展示风格）
    steps.push({ desc:`第一个可用主机 = 网络地址 + 1 = <b class="c-op">${arrToIp(first)}</b>；最后一个 = 广播地址 − 1 = <b class="c-op">${arrToIp(last)}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">可用范围</span></div>
          <div class="step-title"><span class="st-dot"></span>排除网络地址与广播地址（首台 / 末台主机）</div>
          ${this.binRow32(first, prefix, {label:'首台主机', role:'net'})}
          <div class="bin32-caption">${this.binDottedColored(first, prefix, 'net')}</div>
          ${this.binRow32(last, prefix, {label:'末台主机', role:'net'})}
          <div class="bin32-caption">${this.binDottedColored(last, prefix, 'net')}</div>
          <div class="exp-inline" style="margin:8px 0 6px"><span class="term done">${arrToIp(net)}</span><span class="op">+1 =</span><span class="term cur">${arrToIp(first)}</span></div>
          <div class="exp-inline" style="margin-bottom:6px"><span class="term done">${arrToIp(broad)}</span><span class="op">−1 =</span><span class="term cur">${arrToIp(last)}</span></div>
          <div class="write-result">可用主机范围：${arrToIp(first)} ~ ${arrToIp(last)}</div>
          ${this.legendHTML()}`;
      }});
    // 4. 可用主机数 = 2^主机位 − 2（幂运算展开）
    steps.push({ desc:`主机位 = ${hostBits} 位，2<sup>${hostBits}</sup> = ${Math.pow(2,hostBits).toLocaleString()}，减去网络地址和广播地址各 1 个，可用主机数 = <b class="c-op">${count.toLocaleString()}</b>。`,
      draw:()=>{
        const total = Math.pow(2, hostBits);
        let expandExpr;
        if(hostBits <= 12){
          expandExpr = Array(hostBits).fill('2').join(' × ');
        } else {
          expandExpr = `2 × 2 × … × 2（${hostBits} 个 2 相乘）`;
        }
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">可用主机数</span></div>
          <div class="step-title"><span class="st-dot"></span>2<sup>主机位</sup> − 2（幂运算展开）</div>
          <div class="calc-process">
            <div class="cp-label">每一位主机位有 0/1 两种可能，${hostBits} 位共有 2<sup>${hostBits}</sup> 种组合</div>
            <div class="cp-sum-line" style="margin-bottom:8px">
              2<sup>${hostBits}</sup> = ${expandExpr} = <span class="sum">${total.toLocaleString()}</span>
            </div>
            <div class="cp-sum-line">
              ${total.toLocaleString()} <span class="eq">−</span> 2（网络地址+广播地址） = <span class="sum">${count.toLocaleString()} 台</span>
            </div>
          </div>
          <table class="exp-table" style="margin-bottom:8px">
            <tr><th>主机位数</th><th>2<sup>主机位</sup>（总地址数）</th><th>减网络/广播</th><th>可用主机数</th></tr>
            <tr><td class="cur">${hostBits}</td><td>${total.toLocaleString()}</td><td>−2</td><td class="done">${count.toLocaleString()}</td></tr>
          </table>`;
      }});
    // 5. 汇总
    steps.push({ desc:`汇总：${str}/${prefix} 的网络地址、广播地址、可用范围与主机数量。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">汇总</span></div>
          <div class="ip-info">
            <div class="ii"><div class="ii-hd">网络地址</div><div class="ii-val">${arrToIp(net)}</div></div>
            <div class="ii"><div class="ii-hd">广播地址</div><div class="ii-val">${arrToIp(broad)}</div></div>
            <div class="ii"><div class="ii-hd">可用主机范围</div><div class="ii-val">${arrToIp(first)} ~ ${arrToIp(last)}</div></div>
            <div class="ii"><div class="ii-hd">可用主机数</div><div class="ii-val">${count.toLocaleString()} 台</div></div>
          </div>
          <div class="write-result">${arrToIp(net)} ~ ${arrToIp(broad)} · 可用 ${count.toLocaleString()} 台（${arrToIp(first)} ~ ${arrToIp(last)}）</div>`;
      }});
    return steps;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const sub = this.currentSub;
    const info = this.subTabs.find(t=>t.id===sub) || {};
    const labelMap = { classify:'IPv4 地址分类', mask:'子网掩码计算', network:'网络地址计算', hostrange:'主机范围与数量' };
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub]||'IP 与子网'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:#06B6D4">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始手写推导</div>
      </div>`;
  },
  generate(){
    switch(this.currentSub){
      case 'classify': this.steps = this.genClassify(); break;
      case 'mask':     this.steps = this.genMask(); break;
      case 'network':  this.steps = this.genNetwork(); break;
      case 'hostrange':this.steps = this.genHostrange(); break;
      default: this.steps = this.genClassify();
    }
  }
};
