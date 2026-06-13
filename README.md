# NJUPT 毕业设计助手 · njupt-thesis skill

一个面向**南京邮电大学本科毕业设计(论文)全流程**的 Claude Agent Skill,覆盖从立题、开题、外文翻译、中期检查,到论文撰写、查重降重、AIGC 降重、答辩准备的各个环节。

*A Claude Agent Skill for the **entire NJUPT (Nanjing University of Posts and Telecommunications) undergraduate thesis workflow** — from topic approval, proposal, foreign-literature translation, and mid-term review, all the way through thesis writing, plagiarism reduction, AIGC (AI-content) reduction, and defense preparation.*

> 本 skill 由一篇已通过答辩的毕设项目沉淀而来。课题本身是通信方向(英文论文 + MATLAB 仿真 → 中文论文),但流程、格式规范与文档工程方法对**各专业通用**——涉及具体学科的地方都标注了"以你的源材料为准"。
>
> *This skill was distilled from a real thesis project that passed its defense. The original topic was in wireless communications (an English paper + MATLAB simulations turned into a Chinese thesis), but the workflow, formatting rules, and document-engineering methods are **discipline-agnostic** — anywhere something is field-specific, it is marked "defer to your own source material."*

## 它能帮你做什么 · What it can do

- **按南邮格式写/改论文**:宋体黑体混排、1.5 倍行距、章节层级、页眉页码、图题表题,自动符合规范
  *Write/edit your thesis to NJUPT formatting: mixed SimSun/SimHei fonts, 1.5× line spacing, heading levels, headers & page numbers, figure/table captions — all compliant by default.*
- **公式 LaTeX → MathType 工作流**:生成的 Word 里公式以 LaTeX 源码呈现,在 MathType 里一键 `Toggle TeX` 转换,不再手敲
  *LaTeX → MathType formula workflow: formulas are emitted as LaTeX source inside the Word file, converted in one click via MathType's `Toggle TeX` — no manual retyping.*
- **GB/T 7714 参考文献**:格式化、核实、补足篇数
  *GB/T 7714 references: formatting, verification, and topping up the count.*
- **过程材料代拟**:任务书、开题报告、中期检查表、指导记录(含双周写法)、外文翻译
  *Drafting process documents: task assignment, proposal, mid-term review form, supervision log (with a biweekly-entry pattern), and foreign-literature translation.*
- **定稿三关**:查重降重、AIGC 检测降 AI 率(按检测报告标色提取段落 + 系统化改写)、逐条回应导师批注
  *The three finalization gates: plagiarism reduction; AIGC reduction (extracting flagged passages by the report's highlight colors, then rewriting systematically); and addressing supervisor comments one by one.*
- **答辩准备**:PPT 框架、逐页讲稿与计时表、预设问答、评委背景调研
  *Defense prep: slide-deck structure, per-slide script with a timing table, anticipated Q&A, and committee-member background research.*

## 核心原则 · Core principles

skill 内置三条贯穿全流程的铁律,这也是它和"直接让 AI 写论文"最大的区别:

*Three iron rules run through the whole workflow — this is the biggest difference from simply "letting an AI write your thesis":*

1. **先读后写,绝不编造** —— 写任何内容前先读源材料(英文原文、代码输出、真实仿真图);没把握的公式/数据用灰色斜体标注"待核对",而不是悄悄给一个看起来合理的版本。
   ***Read before writing, never fabricate*** *— read the source material first (the English original, code output, real simulation figures); anything uncertain is flagged in gray italics as "to be verified," rather than quietly filled in with a plausible-looking guess.*
2. **增量交付,先定风格** —— 先出一章确认风格,再批量推进;每章给出"论文小节 ↔ 原文 ↔ 代码"对应关系。
   ***Deliver incrementally, lock the style first*** *— produce one chapter, confirm the style, then scale up; each chapter comes with a "thesis section ↔ source paper ↔ code" mapping.*
3. **完成必核,渲染检查** —— 每个 Word 文档生成后都校验并转 PDF 逐页核对版式。
   ***Always verify on completion, check the rendering*** *— every generated Word file is validated and converted to PDF for page-by-page layout checking.*

## 目录结构 · Repository structure

```
njupt-thesis-skill/
├── README.md
├── LICENSE
├── .gitignore
└── njupt-thesis/            # skill 本体 · the skill itself
    ├── SKILL.md             # 主文件:触发与全流程导航 · entry point: triggering & workflow routing
    ├── references/
    │   ├── format-spec.md           # 论文格式规格 · thesis formatting spec
    │   ├── materials-checklist.md   # 11 项归档材料的结构与写法 · structure & writing of the 11 archived documents
    │   ├── writing-workflows.md     # 公式/降重/降AI率/批注/TOC修复/仿真图 · formulas / plagiarism & AIGC reduction / comments / TOC repair / figures
    │   └── defense.md               # 答辩 PPT / 讲稿 / Q&A / 评委调研 · defense slides / script / Q&A / committee research
    └── scripts/
        └── njupt_docx_template.js   # 按南邮规格封装的 docx-js 模板模块 · docx-js template module preconfigured to NJUPT spec
```

## 安装 · Installation

### 方式一:Claude Code / 本地 skills 目录 · Option 1: Claude Code / local skills directory

```bash
git clone https://github.com/HenryLiu0405/njupt-thesis-skill.git
# 拷贝 skill 文件夹(注意:是里层的 njupt-thesis,不是整个仓库)
# Copy the skill folder (note: the inner njupt-thesis, not the whole repo)
cp -r njupt-thesis-skill/njupt-thesis ~/.claude/skills/
```

重启 Claude Code,用 `/skills` 确认已加载。之后正常提问即可,例如"帮我把开题报告的研究现状写一下"。

*Restart Claude Code and run `/skills` to confirm it loaded. Then just ask normally, e.g. "help me write the related-work section of my proposal."*

### 方式二:Claude.ai / Claude API 上传 · Option 2: upload to Claude.ai / Claude API

将 `njupt-thesis/` 文件夹按 [Anthropic 官方 Skills 文档](https://docs.claude.com)中"上传自定义 skill"的说明打包上传。

*Package the `njupt-thesis/` folder and upload it following the "upload a custom skill" instructions in the [official Anthropic Skills docs](https://docs.claude.com).*

## 依赖 · Dependencies

skill 中的文档生成脚本依赖:

*The document-generation scripts in this skill require:*

- Node.js + [`docx`](https://www.npmjs.com/package/docx)(`npm install -g docx`)生成 Word · *to generate Word files*
- LibreOffice + Poppler(`pdftoppm`)用于渲染核对 · *for render-checking*
- 学生侧:Word + MathType(公式转换)、MATLAB(仿真图,可选) · *on your side: Word + MathType (formula conversion), MATLAB (figures, optional)*

## 重要声明 · Important notes

- **格式以学院当年下发的官方模板为准**。本 skill 记录的格式规格核对自某一届的定稿论文,用于在没有模板或模板未写清的细节上提供一套已被验收的参考;各项数字红线(查重率、翻译字数等)每年可能调整。
  ***Always defer to your school's official template for the current year.*** *The formatting spec here was verified against one year's finalized thesis; it serves as a known-accepted reference for details when no template exists or the template is unclear. Numeric thresholds (plagiarism rate, translation word count, etc.) may change yearly.*
- 本 skill **不替你做研究、不替你思考**。它把毕设的格式杂活和文档工程自动化,核心研究内容仍须你自己完成,且所有产出都需你核对后对自己负责。
  *This skill **does not do your research or thinking for you.** It automates the formatting chores and document engineering; the core research is still yours to do, and you are responsible for verifying every output.*
- 仓库内不含任何真实学生的个人信息;课题示例仅用于说明流程。
  *The repository contains no real student's personal information; topic examples are used only to illustrate the workflow.*

## License

[MIT](LICENSE)

## 贡献 · Contributing

欢迎其他南邮同学提 Issue 或 PR 补充:不同学院的格式差异、新增的检查环节、踩过的新坑。

*Fellow NJUPT students are welcome to open Issues or PRs to add: formatting differences across schools/colleges, newly added review steps, and new pitfalls you've hit.*
