---
name: "prd-and-prototype-kit"
description: "产品经理向：把需求沉淀为高保真原型图（PNG）与结构化PRD（docx），并可产出导出交付表格（xlsx）。当用户要求'画原型/原型图'、'出PRD/需求文档'、'仿某系统页面风格'、'导出交付表格'时调用。内置HTML建模→puppeteer截图→docx-js出稿的完整流水线与避坑清单。"
---

# PRD & 原型图 生产套件（prd-and-prototype-kit）

面向 B 端/中后台产品经理的需求交付流水线。核心方法：**用 HTML/CSS 高保真复刻目标系统的页面风格 → 用无头 Chrome 截高清 PNG → 用 docx-js 生成结构化 PRD → 用 openpyxl 生成交付用的 Excel 模板**。

## 何时使用本 Skill
- 用户说：画原型图 / 出原型 / 仿照某个系统页面 / 高保真线框图
- 用户说：出一份 PRD / 需求文档 / 产品方案（要 .docx）
- 用户说：把配置字段整理成导出模板 / Sheet 表头（要 .xlsx）

---

## 一、总原则

1. **中间产物**（HTML、生成脚本、临时数据）放临时工作目录；**最终交付**（PNG/docx/xlsx）放用户工作区目录，并用 `computer://` 链接给用户。
2. 原型风格要**贴目标系统**：主色、圆角、字号、表格线、Tab、按钮形态尽量还原；对「新增/改动点」用红框(`outline`)+红色 badge 标注，让评审一眼看到差异。
3. 所有交付**可迭代**：改一处 → 重跑脚本 → 覆盖输出。脚本化是关键，不要手工画图。
4. 每次改动后**回读一次 PNG** 自检（Read 图片），确认无错位/文字溢出再交付。

---

## 二、环境准备（Windows / macOS，一次性）

**Windows**

```powershell
# npm 走国内镜像，规避证书问题
npm config set registry https://registry.npmmirror.com
# 原型需要 puppeteer-core + 系统 Chrome（勿装完整 puppeteer，太重）
npm install puppeteer-core docx --registry=https://registry.npmmirror.com
# 系统 Chrome 路径（Windows 默认，按实际安装位置调整）
# $env:CHROME = "C:\Program Files\Google\Chrome\Application\chrome.exe"
# Excel
pip install openpyxl
```

**macOS**

```bash
# npm 走国内镜像，规避证书问题
npm config set registry https://registry.npmmirror.com
# 原型需要 puppeteer-core + 系统 Chrome（勿装完整 puppeteer，太重）
npm install puppeteer-core docx --registry=https://registry.npmmirror.com
# 系统 Chrome 路径（macOS 默认，按实际安装位置调整）
# export CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
# Excel（pip3 报错时改 python3 -m pip install openpyxl）
pip3 install openpyxl
```

**避坑清单（本套件实战总结）**
- **macOS 差异**：Chrome 路径用 `export CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"`；pip 用 `pip3` / `python3 -m pip`；macOS 上 `launch` 不需要 `--no-sandbox`（留着也无碍）。
- **numpy 破坏 openpyxl**：若报 `numpy has no attribute 'short'/__version__`，说明环境里有残缺 numpy。**最省事**：脚本顶部加 `import sys; sys.modules["numpy"] = None`，让 openpyxl 走无 numpy 分支即可。
- **npm SSL 证书错误**：加 `--strict-ssl=false`（Windows 企业网代理场景偶发，正常环境不需要）。
- **Chrome 退出时 Crashpad/Sandbox 报错**：无害，输出正常时忽略即可（可 `grep -v -i "crashpad\|sandbox\|settings.dat"` 过滤）。
- **LibreOffice 转 PDF 校验易失败**（profile 锁）：非必要不做；docx 用同一套样式代码即可，改动前先目视 PNG。

---

## 三、原型图流水线（只产出 PNG，不产 SVG）

### 3.1 建 HTML（proto.html）
- 一个文件里放多张原型，每张用 `<div class="shot" id="pN">…</div>` 包裹，便于按 id 单独截图。
- 复用一套 CSS 设计令牌（主色 `#3B6EF5`、深蓝标题 `#1F4E79`、表头灰 `#fafbfc`、红标注 `#ff4d4f`）。
- 标注改动：`.redbox{outline:2px solid #ff4d4f}`、`.badge{背景红/白字「新增」}`、`.rednote{红色小字说明}`。
- 常见组件：顶部 Tab、筛选区（filter/fitem/inp）、工具条按钮、数据表格（th/td/ck 复选框）、分页（pager/pg）、弹窗（modal-mask/modal/mhd/mbd/mft）、二次确认对话框（cdlg/cicon）。
- 表单/弹窗内**给出真实示例数据**，评审代入感更强。

### 3.2 截高清 PNG（shot.js）
```js
const puppeteer = require("puppeteer-core"); const path=require("path");
const CHROME="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
const OUT="<用户工作区目录>";
(async()=>{
  const b=await puppeteer.launch({executablePath:CHROME,headless:"new",
    args:["--no-sandbox","--force-device-scale-factor=2"]});
  const p=await b.newPage();
  await p.setViewport({width:1400,height:1200,deviceScaleFactor:2});// 2倍清晰
  await p.goto("file://"+path.join(__dirname,"proto.html"));
  await new Promise(r=>setTimeout(r,400));
  for(const [id,file] of [["p1","原型1.png"],["p2","原型2.png"]]){
    const el=await p.$("#"+id); await el.screenshot({path:path.join(OUT,file)});
  }
  await b.close();
})();
```

### 3.3 自检
用 Read 打开生成的 PNG 目视：文字有无溢出、红框位置对不对、示例数据是否合理。

---

## 四、PRD 文档流水线（docx-js / gen_prd.js）

- 用 `docx` 包，**中文字体**在 run 里设 `{ font: { eastAsia: "微软雅黑" } }`，否则中文可能不美观。
- 封装小工具函数：`h1/h2`（标题）、`p`（正文）、`bullet/num`（列表）、`table/cell`（表格）。
- **推荐 PRD 结构（8 章）**：
  1. 需求背景　2. 目标与非目标　3. 功能需求总览　4. 功能详细说明（分 4.1/4.2…）　5. 字段/规则清单（表格）　6. 影响分析　7. 验收标准　8. 待确认事项
- 交付一版就出一版**干净稿**：不要把 V1/V2 对比留在最终文档里。
- 需求有变更时：先**只做评估分析、不动文档**，等用户拍板每个决策点后再改稿。

---

## 五、Excel 交付模板（openpyxl / gen_xlsx.py）

用于把「一坨 JSON 配置」拆成可读的导出模板。推荐 **4 个 Sheet**：
1. **主表**（一账户/一模板一行，明细汇总成可读文本）
2. **明细表**（一事件一行，所有可能字段取**并集**，不适用填「—」，含关联键；核心交付）
3. **字段说明**（逐列：源字段 key / 枚举或格式 / 适用条件，给研发照此取数）
4. **枚举字典**（code→中文，导出时统一翻译）

要点：多值字段用「、」分隔；条件类拆成「字段/比较符/值」三列；枚举必须有字典，缺则列入「待确认」。样式：表头深蓝底白字、分组行浅色底、冻结首行+关联列。

---

## 六、交付话术

- 给**直接可点的** `computer://` 链接（PNG / docx / xlsx），后置说明尽量简短。
- 结尾主动提一个**合理的下一步**（如：是否把该规则同步进 PRD 某章 / 是否补一张状态图），但不啰嗦。

---

## 七、目录约定（每次开工先套用）
- 中间产物：临时工作目录（proto.html / *.js / gen_*.py / 原始数据）
- 最终交付：用户工作区目录（*.png / *.docx / *.xlsx）
- 输出文件名用**中文语义名 + 序号**（如 `原型4_批量设置上报模板弹窗.png`），改结构时同步清理旧名文件。

---

## 八、参考示例（本仓库 examples/ 目录）

实际产出的 PRD 交付文档样例（PDF），供对照参考：

- [examples/PRD_广告维度数据上报_01.pdf](examples/PRD_广告维度数据上报_01.pdf) —— 广告维度数据上报 PRD（示例 01）
- [examples/PRD_广告维度数据上报_02.pdf](examples/PRD_广告维度数据上报_02.pdf) —— 广告维度数据上报 PRD（示例 02）
