/**
 * njupt_docx_template.js — 南邮本科毕业论文 docx-js 模板模块
 *
 * 用法(在生成脚本里):
 *   const T = require('/path/to/njupt-thesis/scripts/njupt_docx_template.js');
 *   const doc = T.makeThesisDoc({
 *     year: 2027,                       // 页眉届数
 *     startPage: 1,                     // 本章起始页码(分章生成时按真实页码填)
 *     children: [
 *       T.h1('第二章  系统模型'),
 *       T.h2('2.1 场景描述'),
 *       T.body('考虑一个单站基站同时服务……(正文,自动首行缩进2字符)'),
 *       T.bodyFlush('用户的接收信干噪比为:'),        // 公式引导句,顶格
 *       T.formula(`y = \\frac{1}{\\sqrt{2\\pi}\\sigma} \\exp\\left(-\\frac{(x-\\mu)^2}{2\\sigma^2}\\right)`, '式(2-5)'),
 *       T.note('式(2-5) 请对照英文原稿核对求和下标范围。'),  // 灰色斜体待核注
 *       T.figCaption('图2.1 系统场景示意图'),
 *       T.tableCaption('表2.1 主要符号说明'),
 *       T.refEntry(1, 'LIU F, CUI Y, MASOUROS C, et al. Integrated sensing and communications[J]. IEEE JSAC, 2022, 40(6): 1728-1767.'),
 *     ],
 *   });
 *   T.save(doc, '/mnt/user-data/outputs/chapter2.docx');
 *
 * 注意事项:
 *  - 所有中文内容字符串请用反引号模板字符串,避免内容中的英文双引号炸普通字符串;
 *  - formula() 第一个参数是 LaTeX 源码(学生用 MathType "Toggle TeX" 转换),
 *    源码以 Consolas 等宽+浅灰底呈现,编号在同段右侧制表位,转换后公式自然居中、编号右对齐;
 *  - 行距默认 1.5 倍;若学院模板为固定值 24 磅,改 LINE 常量为 {line:480, lineRule:'exact'};
 *  - 前置部分(摘要/目录)无页眉页码:用 makeFrontMatterDoc() 或自行去掉 headers/footers。
 */
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Header, Footer,
  PageNumber, TabStopType, BorderStyle, PageBreak,
} = require('docx');
const fs = require('fs');

/* ───────── 常量(南邮规格,详见 references/format-spec.md) ───────── */
const ZH = '宋体', HEI = '黑体', EN = 'Times New Roman', MONO = 'Consolas';
const SZ_BODY = 24;   // 小四 12pt(半磅单位)
const SZ_H1 = 32;     // 三号 16pt
const SZ_H2 = 28;     // 四号 14pt
const SZ_CAP = 21;    // 五号 10.5pt(图题/表题/页眉/页码)
const LINE = { line: 360, lineRule: 'auto' };       // 1.5倍行距;院模板若为固定24磅改为 {line:480,lineRule:'exact'}
const INDENT2 = 480;                                 // 首行缩进2字符(12pt下)
// A4 + 近似官方边距(上下2.54cm,左右3.17cm);版心宽 ≈ 11906-2*1797 = 8312 DXA
const PAGE = {
  size: { width: 11906, height: 16838 },
  margin: { top: 1440, bottom: 1440, left: 1797, right: 1797 },
};
const CONTENT_W = PAGE.size.width - PAGE.margin.left - PAGE.margin.right;

/* ───────── run 辅助:中西文字体混排 ───────── */
function run(text, opts = {}) {
  return new TextRun({
    text,
    font: { name: EN, eastAsia: ZH },   // 西文 Times,中文宋体
    size: SZ_BODY,
    ...opts,
  });
}
function heiRun(text, size) {
  return new TextRun({ text, font: { name: HEI, eastAsia: HEI }, size, bold: false });
}

/* ───────── 标题 ───────── */
// 章标题:黑体三号居中,另起一页(pageBreakBefore),挂 Heading1 大纲级别供目录收集
function h1(text, { newPage = true } = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240, ...LINE },
    pageBreakBefore: newPage,
    outlineLevel: 0,
    children: [heiRun(text, SZ_H1)],
  });
}
// 二级标题:黑体四号左对齐顶格
function h2(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 180, after: 120, ...LINE },
    outlineLevel: 1,
    children: [heiRun(text, SZ_H2)],
  });
}
// 三级标题:黑体小四左对齐顶格
function h3(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 120, after: 60, ...LINE },
    outlineLevel: 2,
    children: [heiRun(text, SZ_BODY)],
  });
}

/* ───────── 正文 ───────── */
// 普通正文段:两端对齐 + 首行缩进2字符。
// 入参可为字符串,或 [{t:'文字', sup:true(上标如引用[1]), bold, italic}] 数组
// runs 数组通用构造:支持 {t:'文字'} | {t:'[1]',sup:true} | {t:'B_{p}',mono:true 行内LaTeX}
// 行内变量(γ、B_p 等)必须用 mono 标志给行内 LaTeX 微代码,由学生逐个 MathType 转换,
// 正文中不得出现裸变量名或 Unicode 数学字符(详见 references/format-spec.md 第7节)。
function mkRuns(content) {
  const items = typeof content === 'string' ? [{ t: content }] : content;
  return items.map(r => r.mono
    ? new TextRun({
        text: r.t, font: { name: MONO, eastAsia: MONO }, size: 20,
        shading: { type: 'clear', fill: 'F2F2F2' },
      })
    : new TextRun({
        text: r.t,
        font: { name: EN, eastAsia: ZH },
        size: r.sup ? 14 : SZ_BODY,
        superScript: !!r.sup,
        bold: !!r.bold,
        italics: !!r.italic,
      }));
}
function body(content) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { ...LINE },
    indent: { firstLine: INDENT2 },
    children: mkRuns(content),
  });
}
// 顶格正文段:公式引导句("……为:")或承接公式的"其中,……"
function bodyFlush(content) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { ...LINE },
    children: mkRuns(content),
  });
}

/* ───────── 公式占位段(LaTeX → MathType) ───────── */
// 结构:[居中制表位]LaTeX源码  [右制表位]式(x-y)
// 学生选中源码 → MathType "Toggle TeX" → 公式落在居中位、编号已右对齐。
function formula(latex, label) {
  return new Paragraph({
    tabStops: [
      { type: TabStopType.CENTER, position: Math.round(CONTENT_W / 2) },
      { type: TabStopType.RIGHT, position: CONTENT_W },
    ],
    spacing: { before: 60, after: 60, ...LINE },
    children: [
      new TextRun({ text: '\t' }),
      new TextRun({
        text: latex,
        font: { name: MONO, eastAsia: MONO },
        size: 20,                       // 10pt,与正文区分
        shading: { type: 'clear', fill: 'F2F2F2' },
      }),
      new TextRun({ text: '\t' }),
      new TextRun({ text: label || '', font: { name: EN, eastAsia: ZH }, size: SZ_BODY }),
    ],
  });
}
// 行内公式微代码:在 body() 的 runs 数组里用 {t:'$\\gamma_k$', mono:true} 不直观,
// 故提供独立小函数,返回可拼入 children 的 TextRun:
function inlineLatex(latex) {
  return new TextRun({
    text: latex,
    font: { name: MONO, eastAsia: MONO },
    size: 20,
    shading: { type: 'clear', fill: 'F2F2F2' },
  });
}

/* ───────── 待核对注(灰色斜体,交付前提醒学生,定稿时删除) ───────── */
function note(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { ...LINE },
    children: [new TextRun({
      text: `【待核对】${text}`,
      font: { name: EN, eastAsia: ZH }, size: 20, italics: true, color: '808080',
    })],
  });
}

/* ───────── 图题(图下方)/ 表题(表上方):宋体五号居中 ───────── */
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120, ...LINE },
    children: [new TextRun({ text, font: { name: EN, eastAsia: ZH }, size: SZ_CAP })],
  });
}
const figCaption = caption;     // 语义别名:放在插图段之后
const tableCaption = caption;   // 放在表格之前

/* ───────── 参考文献条目:悬挂缩进 ───────── */
function refEntry(n, text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { ...LINE },
    indent: { left: 480, hanging: 480 },
    children: [new TextRun({
      text: `[${n}] ${text}`,
      font: { name: EN, eastAsia: ZH }, size: SZ_BODY,
    })],
  });
}

/* ───────── 页眉页脚 ───────── */
function njuptHeader(year) {
  return new Header({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
      children: [new TextRun({
        text: `南京邮电大学${year}届本科毕业设计(论文)`,
        font: { name: EN, eastAsia: ZH }, size: SZ_CAP,
      })],
    })],
  });
}
function njuptFooter() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: '- ', font: { name: EN }, size: SZ_CAP }),
        new TextRun({ children: [PageNumber.CURRENT], font: { name: EN }, size: SZ_CAP }),
        new TextRun({ text: ' -', font: { name: EN }, size: SZ_CAP }),
      ],
    })],
  });
}

/* ───────── 文档组装 ───────── */
// 正文文档(带页眉页码)。children 为上面各 helper 的返回值数组。
function makeThesisDoc({ year, startPage = 1, children }) {
  return new Document({
    styles: { default: { document: { run: { font: { name: EN, eastAsia: ZH }, size: SZ_BODY } } } },
    sections: [{
      properties: {
        page: { ...PAGE, pageNumbers: { start: startPage } },
      },
      headers: { default: njuptHeader(year) },
      footers: { default: njuptFooter() },
      children,
    }],
  });
}
// 前置部分文档(摘要/目录等):无页眉无页码
function makeFrontMatterDoc({ children }) {
  return new Document({
    styles: { default: { document: { run: { font: { name: EN, eastAsia: ZH }, size: SZ_BODY } } } },
    sections: [{ properties: { page: PAGE }, children }],
  });
}

function save(doc, path) {
  return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(path, buf); console.log('written:', path); });
}

module.exports = {
  // 组装
  makeThesisDoc, makeFrontMatterDoc, save,
  // 内容 helper
  h1, h2, h3, body, bodyFlush, formula, inlineLatex, note,
  figCaption, tableCaption, caption, refEntry,
  // 低层
  run, heiRun, PAGE, CONTENT_W, LINE,
  PageBreak, Paragraph, TextRun,
};
