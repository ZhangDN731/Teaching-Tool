/* ================================================================
 * 音视频容量（未压缩文件大小计算）
 * 子页面：audio · video
 * ================================================================ */
MOD.media = {
  name:'音视频容量', sub:'音频 / 视频 文件大小（未压缩）计算',
  badge:'存储 · 公式', tag:'容量', cat:'store',
  color:'#38BDF8', soft:'#EAF7FF', type:'音频采样 · 视频帧率 · 字节换算',
  desc:'根据时长、采样率、位深、帧率与分辨率等参数，逐步代入公式计算未压缩音视频文件的容量并换算单位。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M10 8.5v7l6-3.5z"/></svg>',
  setDesc:'',
  stageDesc:'橙色=正在代入的参数，绿色=已得出结果，逐步累乘并换算成方便阅读的容量单位。',
  initDesc:'点击「自动演示」或「下一步」，一步步代入公式算出文件容量。',
  currentSub:'audio',
  subTabs:[
    {id:'audio',label:'音频容量',method:'时长×采样率×位深×声道÷8'},
    {id:'video',label:'视频容量',method:'时长×帧率×宽×高×位深÷8'}
  ],
  theory(){
    return `<div class="rule"><b>音频容量</b> = 时长(秒) × 采样率(Hz) × 位深(bit) × 声道数 ÷ 8（字节）</div>
    <div class="rule"><b>视频容量</b> = 时长(秒) × 帧率(fps) × 宽 × 高 × 位深(bit) ÷ 8（字节）</div>
    <div class="rule"><b>单位换算</b>：1 KB = 1024 B，1 MB = 1024 KB，1 GB = 1024 MB。</div>
    <div class="rule">位深：每个采样点/像素用多少 bit 表示；除以 8 是把 bit 换成字节。</div>
    <div class="rule">常见采样率：电话 8kHz、广播 22.05kHz、CD 44.1kHz、专业 48kHz；视频位深常见 24 bit。</div>`;
  },
  settingsHTML(){
    const sub=this.currentSub;
    if(sub==='audio'){
      const curT=$('#md-a-time')?$('#md-a-time').value:'60';
      const curS=$('#md-a-sr')?$('#md-a-sr').value:'44100';
      const curB=$('#md-a-bits')?$('#md-a-bits').value:'16';
      const curC=$('#md-a-ch')?$('#md-a-ch').value:'2';
      return `<div class="inline-setting"><div class="setting"><label>时长</label>
        <select id="md-a-time">
          <option value="30" ${curT==='30'?'selected':''}>30 秒</option>
          <option value="60" ${curT==='60'?'selected':''}>60 秒</option>
          <option value="120" ${curT==='120'?'selected':''}>120 秒</option>
          <option value="240" ${curT==='240'?'selected':''}>240 秒</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>采样率</label>
        <select id="md-a-sr">
          <option value="22050" ${curS==='22050'?'selected':''}>22.05 kHz</option>
          <option value="44100" ${curS==='44100'?'selected':''}>44.1 kHz</option>
          <option value="48000" ${curS==='48000'?'selected':''}>48 kHz</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>位深</label>
        <select id="md-a-bits">
          <option value="8" ${curB==='8'?'selected':''}>8 bit</option>
          <option value="16" ${curB==='16'?'selected':''}>16 bit</option>
          <option value="24" ${curB==='24'?'selected':''}>24 bit</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>声道</label>
        <select id="md-a-ch">
          <option value="1" ${curC==='1'?'selected':''}>单声道</option>
          <option value="2" ${curC==='2'?'selected':''}>立体声</option>
        </select></div></div>`;
    }
    if(sub==='video'){
      const curT=$('#md-v-time')?$('#md-v-time').value:'10';
      const curF=$('#md-v-fps')?$('#md-v-fps').value:'24';
      const curR=$('#md-v-res')?$('#md-v-res').value:'1280x720';
      return `<div class="inline-setting"><div class="setting"><label>时长</label>
        <select id="md-v-time">
          <option value="10" ${curT==='10'?'selected':''}>10 秒</option>
          <option value="15" ${curT==='15'?'selected':''}>15 秒</option>
          <option value="30" ${curT==='30'?'selected':''}>30 秒</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>帧率</label>
        <select id="md-v-fps">
          <option value="24" ${curF==='24'?'selected':''}>24 fps</option>
          <option value="25" ${curF==='25'?'selected':''}>25 fps</option>
          <option value="30" ${curF==='30'?'selected':''}>30 fps</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>分辨率</label>
        <select id="md-v-res">
          <option value="640x480" ${curR==='640x480'?'selected':''}>640×480</option>
          <option value="1280x720" ${curR==='1280x720'?'selected':''}>1280×720</option>
          <option value="1920x1080" ${curR==='1920x1080'?'selected':''}>1920×1080</option>
        </select></div></div>`;
    }
    return '';
  },
  afterSettings(){
    const bind=id=>{ const el=$('#'+id); if(el) el.onchange=()=>Engine.reload(); };
    ['md-a-time','md-a-sr','md-a-bits','md-a-ch','md-v-time','md-v-fps','md-v-res'].forEach(bind);
  },
  /* 字节 → 易读单位 */
  humanSize(bytes){
    if(bytes>=1073741824) return [bytes/1073741824,'GB'];
    if(bytes>=1048576) return [bytes/1048576,'MB'];
    if(bytes>=1024) return [bytes/1024,'KB'];
    return [bytes,'B'];
  },
  /* 参数 chip 行（done / cur / dim 高亮） */
  paramChips(params, activeIdx){
    return params.map((p,j)=>{
      const cls=j<activeIdx?'g':j===activeIdx?'o':'dim';
      return `<span class="chip ${cls}">${p.label} ${p.value}${p.unit?' '+p.unit:''}</span>`;
    }).join('');
  },
  /* ========== 音频容量 ========== */
  genAudio(){
    const g=id=>{const el=$(id);return el?+el.value:null;};
    const t=g('#md-a-time')??60;
    const sr=g('#md-a-sr')??44100;
    const bits=g('#md-a-bits')??16;
    const ch=g('#md-a-ch')??2;
    const label='音频文件（未压缩）';
    const params=[
      {label:'时长',value:t,unit:'秒'},
      {label:'采样率',value:sr,unit:'Hz'},
      {label:'位深',value:bits,unit:'bit'},
      {label:'声道数',value:ch,unit:''}
    ];
    const steps=[];
    /* 步骤1：写出已知参数 */
    steps.push({desc:`已知参数：时长 ${t} 秒、采样率 ${sr} Hz、位深 ${bits} bit、声道数 ${ch}；音频容量 = 时长 × 采样率 × 位深 × 声道数 ÷ 8（字节）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">整理参数</span></div>
          <div class="step-title"><span class="st-dot"></span>已知参数</div>
          <div class="flow-line" style="justify-content:center;margin-top:8px">
            ${params.map((p,i)=>`<span class="chip write-in" style="animation-delay:${i*0.12}s">${p.label} ${p.value}${p.unit?' '+p.unit:''}</span>`).join('')}
          </div>
          <div class="annot" style="margin-top:10px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            公式：时长 × 采样率 × 位深 × 声道数 ÷ 8 = 字节
          </div>`;
      }});
    /* 步骤2~4：逐步累乘 */
    const ops=[
      {name:'采样率',value:sr,unit:'Hz',before:t,after:t*sr,afterUnit:'采样点'},
      {name:'位深',value:bits,unit:'bit',before:t*sr,after:t*sr*bits,afterUnit:'bit'},
      {name:'声道数',value:ch,unit:'',before:t*sr*bits,after:t*sr*bits*ch,afterUnit:'bit'}
    ];
    ops.forEach((o,i)=>{
      const idx=i;
      steps.push({desc:`乘以${o.name} <b class="c-op">${o.value}${o.unit?' '+o.unit:''}</b>：${o.before.toLocaleString()} × ${o.value} = <b class="c-op">${o.after.toLocaleString()}</b> ${o.afterUnit}。`,
        draw:()=>{
          $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">代入 ${idx+1}/3</span></div>
            <div class="step-title"><span class="st-dot"></span>逐步代入公式</div>
            <div class="flow-line" style="justify-content:center;margin-top:8px">
              ${this.paramChips(params, idx+1)}
            </div>
            <div class="exp-inline" style="display:inline-block;margin-top:10px">
              <span class="term done">${o.before.toLocaleString()}</span><span class="op">×</span>
              <span class="term cur">${o.value}</span><span class="op">=</span>
              <span class="term done write-in" style="animation-delay:.2s">${o.after.toLocaleString()} ${o.afterUnit}</span>
            </div>`;
        }});
    });
    /* 步骤5：÷8 换算字节 */
    const bitsTotal=t*sr*bits*ch;
    const bytes=bitsTotal/8;
    steps.push({desc:`除以 8，把 bit 换算为字节：${bitsTotal.toLocaleString()} ÷ 8 = <b class="c-op">${bytes.toLocaleString()}</b> 字节。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">换算字节</span></div>
          <div class="step-title"><span class="st-dot"></span>bit ÷ 8 = 字节</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term done">${bitsTotal.toLocaleString()} bit</span><span class="op">÷ 8 =</span>
            <span class="term cur write-in" style="animation-delay:.2s">${bytes.toLocaleString()} B</span>
          </div>`;
      }});
    /* 步骤6：单位换算 + 最终结果 */
    const h=this.humanSize(bytes);
    steps.push({desc:`总容量 = <b class="c-op">${bytes.toLocaleString()}</b> 字节 ≈ <b class="c-op">${h[0].toFixed(2)} ${h[1]}</b>（1 KB = 1024 B）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>容量单位换算</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term done">${bytes.toLocaleString()} B</span><span class="op">≈</span>
            <span class="term cur write-in" style="animation-delay:.2s">${h[0].toFixed(2)} ${h[1]}</span>
          </div>
          <div class="write-result">${t}s × ${sr}Hz × ${bits}bit × ${ch} ÷ 8 = ${bytes.toLocaleString()} B ≈ ${h[0].toFixed(2)} ${h[1]}</div>`;
      }});
    return steps;
  },
  /* ========== 视频容量 ========== */
  genVideo(){
    const g=id=>{const el=$(id);return el?+el.value:null;};
    const t=g('#md-v-time')??10;
    const fps=g('#md-v-fps')??24;
    const res=$('#md-v-res')?$('#md-v-res').value:'1280x720';
    const [w,h]=res.split('x').map(Number);
    const bits=24;
    const label='视频文件（未压缩）';
    const params=[
      {label:'时长',value:t,unit:'秒'},
      {label:'帧率',value:fps,unit:'fps'},
      {label:'宽',value:w,unit:'px'},
      {label:'高',value:h,unit:'px'},
      {label:'位深',value:bits,unit:'bit'}
    ];
    const steps=[];
    /* 步骤1：写出已知参数 */
    steps.push({desc:`已知参数：时长 ${t} 秒、帧率 ${fps} fps、分辨率 ${w}×${h}、位深 ${bits} bit；视频容量 = 时长 × 帧率 × 宽 × 高 × 位深 ÷ 8（字节）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">整理参数</span></div>
          <div class="step-title"><span class="st-dot"></span>已知参数</div>
          <div class="flow-line" style="justify-content:center;margin-top:8px">
            ${params.map((p,i)=>`<span class="chip write-in" style="animation-delay:${i*0.12}s">${p.label} ${p.value}${p.unit?' '+p.unit:''}</span>`).join('')}
          </div>
          <div class="annot" style="margin-top:10px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            公式：时长 × 帧率 × 宽 × 高 × 位深 ÷ 8 = 字节
          </div>`;
      }});
    /* 步骤2~5：逐步累乘 */
    const ops=[
      {name:'帧率',value:fps,unit:'fps',before:t,after:t*fps,afterUnit:'帧'},
      {name:'宽',value:w,unit:'px',before:t*fps,after:t*fps*w,afterUnit:'点·帧'},
      {name:'高',value:h,unit:'px',before:t*fps*w,after:t*fps*w*h,afterUnit:'像素'},
      {name:'位深',value:bits,unit:'bit',before:t*fps*w*h,after:t*fps*w*h*bits,afterUnit:'bit'}
    ];
    ops.forEach((o,i)=>{
      const idx=i;
      steps.push({desc:`乘以${o.name} <b class="c-op">${o.value}${o.unit?' '+o.unit:''}</b>：${o.before.toLocaleString()} × ${o.value} = <b class="c-op">${o.after.toLocaleString()}</b> ${o.afterUnit}。`,
        draw:()=>{
          $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">代入 ${idx+1}/4</span></div>
            <div class="step-title"><span class="st-dot"></span>逐步代入公式</div>
            <div class="flow-line" style="justify-content:center;margin-top:8px;overflow-x:auto">
              ${this.paramChips(params, idx+1)}
            </div>
            <div class="exp-inline" style="display:inline-block;margin-top:10px">
              <span class="term done">${o.before.toLocaleString()}</span><span class="op">×</span>
              <span class="term cur">${o.value}</span><span class="op">=</span>
              <span class="term done write-in" style="animation-delay:.2s">${o.after.toLocaleString()} ${o.afterUnit}</span>
            </div>`;
        }});
    });
    /* 步骤6：÷8 换算字节 */
    const bitsTotal=t*fps*w*h*bits;
    const bytes=bitsTotal/8;
    steps.push({desc:`除以 8，把 bit 换算为字节：${bitsTotal.toLocaleString()} ÷ 8 = <b class="c-op">${bytes.toLocaleString()}</b> 字节。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">换算字节</span></div>
          <div class="step-title"><span class="st-dot"></span>bit ÷ 8 = 字节</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term done">${bitsTotal.toLocaleString()} bit</span><span class="op">÷ 8 =</span>
            <span class="term cur write-in" style="animation-delay:.2s">${bytes.toLocaleString()} B</span>
          </div>`;
      }});
    /* 步骤7：单位换算 + 最终结果 */
    const hh=this.humanSize(bytes);
    steps.push({desc:`总容量 = <b class="c-op">${bytes.toLocaleString()}</b> 字节 ≈ <b class="c-op">${hh[0].toFixed(2)} ${hh[1]}</b>（1 KB = 1024 B）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>容量单位换算</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term done">${bytes.toLocaleString()} B</span><span class="op">≈</span>
            <span class="term cur write-in" style="animation-delay:.2s">${hh[0].toFixed(2)} ${hh[1]}</span>
          </div>
          <div class="write-result">${t}s × ${fps}fps × ${w}×${h} × ${bits}bit ÷ 8 = ${bytes.toLocaleString()} B ≈ ${hh[0].toFixed(2)} ${hh[1]}</div>`;
      }});
    return steps;
  },
  /* 初始化演示画布（切换子页面时清空旧内容） */
  initStage(){
    const sub=this.currentSub;
    const info=this.subTabs.find(t=>t.id===sub)||{};
    const labelMap={
      audio:'音频文件容量（未压缩）',
      video:'视频文件容量（未压缩）'
    };
    $('#stage').innerHTML=`<div class="stage-hd"><span>${labelMap[sub]||'音视频容量'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:${this.color}">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始手写推导</div>
      </div>`;
  },
  generate(){
    const sub=this.currentSub;
    switch(sub){
      case 'audio': this.steps=this.genAudio(); break;
      case 'video': this.steps=this.genVideo(); break;
      default: this.steps=this.genAudio();
    }
  }
};
