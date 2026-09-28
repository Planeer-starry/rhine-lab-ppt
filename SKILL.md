---
name: rhine-lab-ppt
description: "Generate 16:9 Chinese presentation decks in the RHINE LAB (莱茵生命) house style — warm ivory ground, ink black, a single champagne-gold accent, wide letter-spacing, hairline rules and square nodes. Use this skill whenever the user asks for a PPT, 演示文稿, 幻灯片, 汇报, 竞选 or 答辩 deck in this style, or says 用莱茵生命风格 / 按上次那个风格做. It covers cover, agenda, section divider, statement, photo, data and closing layouts, and renders an editable .pptx via the bundled PowerShell renderer."
whenToUse: "Any request to build or extend a slide deck that should follow the RHINE LAB visual system, including class-monitor campaigns, project reports, competition results, and data readouts."
---

# RHINE LAB PPT 风格技能

生成莱茵生命视觉体系的中文演示文稿：暖象牙纸面、墨黑信息、**香槟金只做一处强调**、宽字距、发丝线与方块节点。

## 何时用

- 用户要「做 PPT / 演示文稿 / 幻灯片 / 汇报 / 竞选 / 答辩」，且希望沿用这套风格
- 用户说「用莱茵生命风格」「按上次那个风格做」
- 用户给出素材（证书、照片、数据）要排成这种克制的档案感版式

**不适用**：需要炫色、渐变、卡片圆角、图标库的场合 —— 这套风格刻意不做这些。

## 三步出片

### 1. 安装（首次）

```powershell
# 用户级（推荐，对当前用户所有会话生效）
Copy-Item -Recurse <本仓库> "$env:APPDATA\dsh-desktop\harness\skills\rhine-lab-ppt"
# 或项目级：<项目根>\.dsh\skills\rhine-lab-ppt
```

也可以直接跑 `scripts/install-skill.ps1`。

### 2. 写内容 JSON

复制 `templates/content.example.json`，按业务替换文字。**这是唯一需要改的文件**，风格全部固化在渲染器里。

所有 `src` 图片路径**必须是绝对路径**，且建议放在纯 ASCII 路径下（含中文或空格的路径会让 Office COM 导入图片失败）。

### 3. 渲染

```powershell
$env:RHINE_CONTENT = 'D:\path\to\my-content.json'
$env:RHINE_PPTX    = 'D:\path\to\my-deck.pptx'
$env:RHINE_PNG     = 'D:\path\to\my-preview'   # 可选，输出逐页预览图
$code = [IO.File]::ReadAllText('scripts\rhine-style.ps1', [Text.Encoding]::UTF8)
Invoke-Expression $code
```

用 `Invoke-Expression` 读取而不是直接执行，是为了规避 Windows PowerShell 5.1 把无 BOM 的 UTF-8 脚本按 ANSI 解码、导致中文字面量乱码的问题。

产出：可编辑 `.pptx` + 全部页面 PNG 预览。

> 若渲染时 PowerPoint 报错在插图上，说明当前进程没有读取图片所在目录的权限。让 Office 能读到素材即可（换到有权限的路径，或以更高权限重跑渲染）。

## 版式选择

| layout | 用途 | 必填字段 |
|---|---|---|
| `cover` | 封面 | `title`, `subtitle` |
| `agenda` | 目录 | `kicker`, `title`, `items[{no,zh,en}]` |
| `process` | 阶段/步骤（3–4 列） | `kicker`, `title`, `stages[{no,title,desc}]`, `readout`, `note` |
| `dark` | 深场章节页 / 承诺页 | `title`, `subtitle`, `label`, `caption`, `card[]` |
| `statement` | 论点 + 证据面板 | `section`, `title`, `leadLabel`, `lead`, `explain`, `panelTitle`, `rows[]` |
| `statement-image` | 左图右论点 | 同上 + `src`, `cropBias` |
| `photos` | 2–3 图作品墙 | `section`, `title`, `kicker`, `photos[{src,caption}]` |
| `statement-photo` | 深场全宽大图 | `kicker`, `title`, `src`, `cropBias`, `caption`, `note` |
| `data` | 柱状数据页 | `section`, `title`, `kicker`, `bars[{value,label,height,accent}]`, `readout`, `note` |
| `closing` | 结尾 | `kicker`, `band`, `mark`, `note` |

**选版式看信息关系，不看页序**：≥6 个内容页时至少用 4 种结构，别让相邻页重复同一构图；密集页之后留一个安静页。

## 内容纪律（风格的一部分）

1. **一页一个论点**：抓人的主张放标题，主展示占主导，结尾一句读解或下一步。
2. **数据必须有口径**：单位、样本、统计窗口写在图旁或来源注。
3. **只高亮一个重点**：一根金柱、一个金序号、一条金强调条。
4. **不虚构**：占位文字保留「待填写」；示例数据显式标注。
5. **不压字号**：中文正文 12–16pt，最小 9pt；塞不下就换版式或拆页。

## 禁止事项

- 不使用渐变填充、外阴影、三维效果、艺术字、圆角矩形
- 不引入第 4 种主色；金色每页至多一次
- 不删除品牌字标（左上三行）与页脚（右下 `POWERED BY RHINE LAB` + 方块）
- 深场页不用浅色页的元素配色
- 原模板里的 `FFFF00` / `92D050` / `00FF00` 是编辑占位提示色，**不是设计色**，禁止使用

## 继续深挖

- 完整色值 / 字体栈 / 网格常量 / 逐版式坐标：`references/design-tokens.md`
- 可复用渲染器：`scripts/rhine-style.ps1`
- 现成母版（想在 PPT 里手动续写）：`templates/rhine-lab-master.pptx` 或 `.potx`
- 成品效果：`docs/preview/*.png`
