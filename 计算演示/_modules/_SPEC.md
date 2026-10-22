# 计算演示模块统一修改规范（所有子代理必读）

## 1. 项目架构

- 入口文件：`计算演示/index.html`（约4875行，含CSS+HTML骨架+所有JS模块内联）
- 模块文件：`计算演示/_modules/*.js`（每个文件包含一个 `MOD.xxx = {...}` 赋值，与index.html中内联代码一致）
- **你的工作方式**：读取 index.html 中对应模块的内联代码（或直接读 _modules/xxx.js），修改后写回 `_modules/xxx.js`。不要直接修改 index.html，最终由组织者统一替换回内联。
- 设计规范：`_设计规范V2.md`（蓝紫渐变现代数据看板风格）

## 2. MOD 模块接口契约（必须严格遵守）

每个模块是一个赋值给全局 `MOD` 对象的字面量：

```js
MOD.xxx = {
  // ---- 元数据 ----
  name: '模块名',
  sub: '副标题',
  badge: '徽章文字',
  tag: '标签',
  cat: '分类',       // bit(数制编码) / store(存储容量) / net(网络表格) / code(程序设计) / ds(数据结构)
  color: '#4F6BFF',  // 主题色
  soft: '#EEF3FF',   // 浅色背景
  type: '类型描述',
  desc: '详细描述',
  icon: '<svg ...>...</svg>',

  // ---- 子页面tab（可选） ----
  currentSub: 'subId',
  subTabs: [{id:'subId', label:'显示名', method:'方法说明'}],

  // ---- 方法 ----
  theory() { return 'HTML字符串'; },           // 知识点讲解
  settingsHTML() { return 'HTML字符串'; },      // 参数设置区
  afterSettings() { /* 绑定事件 */ },            // 设置区渲染后绑定
  restoreSettings() { /* 可选：恢复设置值 */ },
  initStage() { /* 可选：初始化画布 */ },
  generate() { this.steps = [...]; },           // 生成演示步骤
  steps: [],                                     // 步骤数组，由generate()填充
  initDesc: '初始提示文字',
  stageDesc: '画布说明',
  canRegenerate: true
};
```

### 步骤对象格式
```js
{
  desc: '本步骤的文字说明（HTML，支持<b class="c-op">高亮</b>）',
  draw: () => { $('#stage').innerHTML = '...'; }  // 渲染画布
}
```

### Engine 引擎行为
- `Engine.load(key)` → 渲染设置区 → 调用 `generate()` → 等待用户操作
- 点击"下一步" → `Engine.i++` → 调用当前步骤的 `draw()` → 显示 `desc` 到 `#explainBox`
- 点击"上一步" → `Engine.i--` → 重新渲染
- `#stage` 是画布容器，`#explainBox` 是说明区
- 全局工具：`$`(querySelector), `$$`(querySelectorAll), `esc`(HTML转义), `randInt(a,b)`, `pick(arr)`

## 3. 可用CSS类（index.html中已定义，直接用）

### 布局/容器
- `.stage-hd` + `.state` — 画布顶部标题条
- `.step-title` + `.st-dot` — 步骤标题胶囊
- `.calc-process` — 计算过程重点卡片（橙色边框，右上角"计算过程"标签）
- `.result-box` — 结果展示框
- `.two-col` — 两列布局
- `.tbl-wrap` > `.tbl` — 数据表格

### 文字/代码
- `.code-block` — Python代码块（深色背景，支持 `.l.cur` 当前行、`.kw`关键字、`.fn`函数、`.str`字符串、`.cm`注释、`.num`数字）
- `.fml` — 公式逐步块
- `.console` — 控制台输出（深绿字）
- `.mono` — 等宽字体行内样式
- `.var-rack` > `.var-pill` — 变量展示架（`.hot`高亮变化）

### 动画
- `.write-in` — 写入动画（淡入上移）
- `.char-in` — 逐字出现
- `.pop` — 弹出动画（chip自动带）
- `.caret` / `.caret.pen` — 闪烁光标

### 芯片/标签
- `.chip`（`.g`绿 `.o`橙 `.r`红 `.cy`青 `.v`紫 `.pk`粉 `.dim`灰）— 圆角标签
- `.byte-tag` — 字节标签

### 进制/计算
- `.exp-table` — 按权展开表（`.done`已算 `.cur`当前 `.empty`未算）
- `.div-vertical` > `.dv-row` — 竖式除法（`.dv-num` `.dv-op` `.dv-eq` `.dv-q` `.dv-rem` `.dv-arrow`）
- `.group-box` > `.grp` — 分组框（`.done` `.cur`）
- `.write-result` — 最终结果行
- `.exp-inline` — 行内表达式（`.term` `.op` `.res`）
- `.flow-line` > `.part` — 流程线（`.cur` `.done`）

### IP/网络
- `.bin-grid` > `.bd` — 8位二进制网格（`.cur` `.mask` `.net` `.zero` `.one` `.weight` `.and-result`）
- `.ip-row` > `.seg` — IP分段行
- `.ip-info` > `.ii` — IP信息卡
- `.cp-bit-grid` > `.cb` — 计算过程中的位网格（`.ip` `.mk` `.rs` `.lb`）

### Excel
- `.xl-grid` > `.xl-cell` — Excel单元格网格（`.hd`表头 `.sel`选中 `.cur`当前 `.done`已完成 `.src`源）
- `.xl-formula` — 公式栏

### 数据结构
- `.stack-area` — 栈区域（纵向，column-reverse）
- `.queue-area` — 队列区域（横向）
- `.deque-area` — 双端队列区域
- `.ds-log` — 操作日志
- `.zone-label` — 区域标签
- `.ptr`（`.head` `.tail`）— 指针标注

### 树/图
- `.svg-box` — SVG容器
- `.order-rack` > `.order-chip` — 遍历序列架
- `.legend` — 图例

### 字形码
- `.dot-matrix` > `i.on` — 点阵字形

### 状态
- `.err` — 错误框
- `.ok-box` — 成功框
- `.annot` — 标注文字

## 4. 设计Token（CSS变量，直接用 var(--xxx)）

```
--accent:#4F6BFF  --accent-hover:#3D56E0  --accent-soft:rgba(79,107,255,.09)
--cyan:#38BDF8  --green:#22C55E  --orange:#F59E0B  --red:#EF4444
--violet:#8B5CF6  --pink:#EC4899
--text:#0F172A  --text-2:#334155  --text-3:#64748B
--surface:#FFFFFF  --surface-2:#F6F9FF  --surface-3:#EEF3FF
--border:rgba(79,107,255,.14)
--radius-lg:18px  --radius:12px  --radius-sm:8px
--mono:ui-monospace,"SF Mono",Menlo,Consolas,"Courier New",monospace
```

## 5. 大屏幕展示要求（核心！）

当前字号太小，必须放大：

- **计算过程区域**：等宽字体至少 `clamp(16px, 2.2vw, 22px)`，关键数字 `clamp(18px, 2.8vw, 28px)`
- **步骤说明（desc）**：至少 `clamp(13px, 1.8vw, 17px)`，行高1.7
- **竖式除法**：`.div-vertical` 字号提升到 `clamp(14px, 2vw, 18px)`，行间距加大
- **表格**：单元格内边距加大，字号 `clamp(12px, 1.6vw, 16px)`
- **步骤卡片间距**：每个步骤块之间 `margin-bottom: clamp(10px, 2vw, 16px)`
- **颜色对比**：当前操作项用橙色(`--orange`)高亮，已完成用绿色(`--green`)，未操作灰色
- 可以在模块内用 `<style>` 标签或内联style覆盖字号，但优先用已有CSS类

## 6. 步骤累加展示模式

对于要求"保留之前步骤，逐步增加"的模块：
- 每一步的 `draw()` 应该渲染**所有已完成步骤 + 当前步骤**，而不是只渲染当前步骤
- 已完成步骤用绿色/低透明度样式，当前步骤用橙色高亮
- 新步骤以 `.write-in` 或 `.pop` 动画出现
- 可以用一个累积HTML变量：`this._accumHTML += '...'`，每步draw时渲染完整累积内容

## 7. 输出约定

- 修改完成后，将完整的 `MOD.xxx = {...};` 代码写入 `计算演示/_modules/xxx.js`（覆盖原文件）
- 文件开头保留注释横幅说明模块名称
- 不要修改 index.html
- 不要修改其他模块的 _modules 文件
- 如果需要新增CSS类，在模块代码顶部用 `document.head.insertAdjacentHTML` 注入 `<style>`，或在返回的HTML中用内联样式

## 8. 验证要求（每个模块必须做）

修改完成后，用具体测试用例验证：
1. 列出你测试的输入参数
2. 逐步走查每一步的 desc 和 draw 输出
3. 验证中间结果和最终结果的数学正确性
4. 验证大屏幕下字号和布局可读
5. 验证"上一步"能正确回退
6. 将验证结果写入回复中

## 9. 禁止事项

- 不要修改 Engine 引擎代码
- 不要修改全局工具函数（$, $$, esc, randInt, pick）
- 不要修改其他模块
- 不要删除 MOD_ORDER 中需要保留的模块
- 不要使用emoji作为图标（用内联SVG）
- 不要引入外部CDN（除jsDelivr/cdnjs外，但尽量不用）
