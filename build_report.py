# -*- coding: utf-8 -*-
"""Build the assignment report docx from the school template + real ScanCode stats."""
import json
import docx
from docx.shared import Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

BASE = r"C:\Users\caizhengyun\Desktop\开源创新思维作业\作业"
TPL = BASE + r"\平时作业模板(命名要求：学号-姓名-作业1，2，3，4) .docx"
OUT = BASE + r"\班级-学号-姓名-作业1.docx"
S = json.load(open(BASE + r"\scan\report_stats.json", encoding="utf-8"))

SONG = "宋体"
FS = "仿宋_GB2312"
WESTERN = "Times New Roman"


def set_run(r, text=None, cn=SONG, size=10.5, bold=False, western=WESTERN):
    if text is not None:
        r.text = text
    r.font.name = western
    r.font.size = Pt(size)
    r.bold = bold
    r.font.color.rgb = RGBColor(0, 0, 0)
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.append(rf)
    rf.set(qn("w:ascii"), western)
    rf.set(qn("w:hAnsi"), western)
    rf.set(qn("w:eastAsia"), cn)


def make_para(doc, text, cn=SONG, size=10.5, bold=False, align=None, indent=True,
              spacing=1.25, before=0, after=0, western=WESTERN):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = spacing
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent:
        pf.first_line_indent = Pt(size * 2)
    if align:
        p.alignment = align
    if text:
        r = p.add_run()
        set_run(r, text, cn=cn, size=size, bold=bold, western=western)
    return p


class Inserter:
    """Inserts generated blocks (paragraphs/tables) before the anchor element."""

    def __init__(self, doc, anchor_el):
        self.doc = doc
        self.anchor = anchor_el

    def _move(self, el):
        self.anchor.addprevious(el)
        return el

    def para(self, *a, **kw):
        p = make_para(self.doc, *a, **kw)
        self._move(p._p)
        return p

    def code(self, text):
        p = make_para(self.doc, None, indent=False, spacing=1.0, before=1, after=1)
        r = p.add_run()
        set_run(r, text, cn=SONG, size=9, western="Consolas")
        self._move(p._p)
        return p

    def caption(self, text):
        p = make_para(self.doc, text, size=9, bold=True,
                      align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, before=4, after=2)
        self._move(p._p)
        return p

    def table(self, rows, widths=None):
        t = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        t._tbl.tblPr.append(layout)
        for i, row in enumerate(rows):
            for j, cell_text in enumerate(row):
                c = t.cell(i, j)
                c.text = ""
                p = c.paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                r = p.add_run()
                set_run(r, cell_text, cn=SONG, size=9, bold=(i == 0))
        if widths:
            for j, w in enumerate(widths):
                for i in range(len(rows)):
                    t.cell(i, j).width = Emu(int(w * 360000))
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
        self._move(t._tbl)
        return t


doc = docx.Document(TPL)

# ---------------- cover fill ----------------
cover = doc.tables[0]


def center_fill(cell, val):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(), val, cn=FS, size=12)


for row in cover.rows:
    cells = list(row.cells)
    texts = [c.text.strip() for c in cells]
    joined = "".join(texts)
    if "姓" in joined and "名" in joined and "学" in joined:
        # [姓名label, nameval, 学号label(merged), sidval]
        for i, c in enumerate(cells):
            if c.text.strip() == "":
                center_fill(c, STATS := (S["name"] if i == 1 else S["sid"]))
    elif "班" in joined and "级" in joined:
        for i, c in enumerate(cells):
            if c.text.strip() == "":
                center_fill(c, S["klass"] if i == 1 else S["teacher"])
    elif "报告日期" in joined:
        for c in cells:
            if c.text.strip() == "":
                center_fill(c, S["report_date"])
                break

# ---------------- body replacement ----------------
body_el = doc.element.body
children = list(body_el.iterchildren())
head_para = eval_para = None
for el in children:
    if el.tag == qn("w:p"):
        txt = "".join(el.itertext())
        if "上海杉达学院" in txt and "内容标题1" in txt:
            head_para = el
        elif "评价表格" in txt:
            eval_para = el

between = False
for el in children:
    if el is head_para:
        between = True
        continue
    if el is eval_para:
        break
    if between:
        body_el.remove(el)

# clean the first heading paragraph: keep page break + 上海杉达学院, drop placeholder heading text
hp = Paragraph(head_para, doc)
seen_school = False
for r in list(hp.runs):
    if "上海杉达学院" in (r.text or ""):
        seen_school = True
        continue
    if seen_school and (r.text or "").strip():
        r._element.getparent().remove(r._element)

# force the school-name paragraph to start on a fresh page (replace inner br behaviour)
pPr = head_para.find(qn("w:pPr"))
if pPr is None:
    pPr = OxmlElement("w:pPr")
    head_para.insert(0, pPr)
if pPr.find(qn("w:pageBreakBefore")) is None:
    pbb = OxmlElement("w:pageBreakBefore")
    pPr.insert(0, pbb)
# drop the inner page-break run so it does not create an extra blank page
for r in list(hp.runs):
    if r._element.find(qn("w:br")) is not None:
        r._element.getparent().remove(r._element)

ins = Inserter(doc, eval_para)

# ============ 1 项目介绍 ============
ins.para("1、开源项目介绍", cn=FS, size=12, bold=True, indent=False, before=0, after=2)
ins.para("1.1 项目概况", bold=True, indent=False, before=3, after=1)
ins.para("本次作业选取 Apache 软件基金会 Commons 子项目的 Apache Commons Text（Maven 坐标 org.apache.commons:commons-text:1.12.0，2024 年 4 月发布）作为被审计的开源项目。该项目 2014 年自 Commons Lang 的文本处理模块拆分而来，提供字符串插值（StringSubstitutor）、文本相似度（Jaro-Winkler、Levenshtein 等）、字符串转义、最长公共子序列、文本差分等通用算法，是 Java 生态中使用广泛的底层类库。")
ins.para("对官方源码发布包的统计：共 " + S["total_files"] + " 个文件，其中 Java 源文件 " + S["java_files"] + " 个（main " + S["main_java"] + " 个、test " + S["test_java"] + " 个），约 " + S["loc"] + " 行代码；其余为 pom.xml、文档、CI 配置及少量测试资源文件。")
ins.para("1.2 发布物的许可证构成", bold=True, indent=False, before=3, after=1)
ins.para("该项目的合规声明符合 ASF 标准发布要求：根目录 LICENSE.txt 为 Apache License 2.0 全文；NOTICE.txt 声明“Apache Commons Text，Copyright 2014-2024 The Apache Software Foundation”；pom.xml 中的 <license> 元数据与上述声明一致。运行时依赖仅 commons-lang3 3.14.0（同为 Apache-2.0）；另声明 junit、mockito 等 " + S["test_deps"] + " 个 test 作用域依赖，不随发布物分发。")
ins.para("1.3 选取理由", bold=True, indent=False, before=3, after=1)
ins.para("一是治理规范：ASF 发布物必须附带 LICENSE 与 NOTICE 文件，且项目用 Apache Rat 做发布审计门禁，便于将工具结果与项目自身声明逐条对照；二是文件类型多样，可检验工具对非代码文件的判定能力；三是规模适中，" + S["total_files"] + " 个文件可全部人工复核，符合“自动扫描＋人工确认”的合规方法论。")

# ============ 2 工具选型与安装 ============
ins.para("2、许可证合规工具的选型与安装", cn=FS, size=12, bold=True, indent=False, before=4, after=2)
ins.para("2.1 选型过程与依据", bold=True, indent=False, before=3, after=1)
ins.para("按照《33 款开源软件许可合规工具选型手册》的“选型三步法”：本次要解决的问题是“对源码目录做许可证、版权与包依赖识别”，属于源码扫描引擎范畴；被审对象是 Java 源码加测试资源的普通源码树；扫描环境为 Windows 11 个人电脑（无 JDK/Maven，不宜部署服务端）。据此对候选工具逐条对比（表 1），最终选定 ScanCode Toolkit 32.5.0。")
ins.caption("表 1  候选工具选型对比")
ins.table(S["table1"], widths=[2.6, 1.9, 4.9, 5.1])
ins.para("2.2 安装过程", bold=True, indent=False, before=3, after=1)
ins.para("扫描环境为 Windows 11、Python 3.13.12。为避免污染全局环境，先以 python -m venv .venv-scancode 创建独立虚拟环境，再在其中安装 ScanCode Toolkit 32.5.0，并以 scancode --version 验证可用。ScanCode 为纯 Python 实现，无需数据库或外部服务，首次扫描时自动构建许可证规则索引。安装问题及解决见表 2。")
ins.para("2.3 安装过程中遇到的问题及解决", bold=True, indent=False, before=3, after=1)
ins.caption("表 2  安装问题记录与解决方案")
ins.table(S["table2"], widths=[2.6, 7.0, 4.9])

# ============ 3 检测报告 ============
ins.para("3、工具输出的检测报告", cn=FS, size=12, bold=True, indent=False, before=4, after=2)
ins.para("3.1 扫描命令与参数", bold=True, indent=False, before=3, after=1)
ins.code("scancode -clpieu -n 4 --json-pp scancode-result.json --html scancode-report.html --csv scancode-summary.csv commons-text-rel-commons-text-1.12.0")
ins.para("参数含义：-l 许可证、-c 版权、-p 软件包、-i 文件信息、-e/-u 邮箱与 URL，-n 4 为并行进程数。全部文件扫描耗时约 " + S["duration"] + " 秒，同时输出机器可读的 JSON 与人可读的 HTML、CSV 报告。")
ins.para("3.2 总体结果", bold=True, indent=False, before=3, after=1)
ins.caption("表 3  扫描总体统计")
ins.table(S["table3"], widths=[7.0, 7.5])
ins.para("3.3 许可证检出分布", bold=True, indent=False, before=3, after=1)
ins.para("按 SPDX 标识符汇总各许可证命中的文件数如表 4（同一文件命中多个规则时分别计数）。")
ins.caption("表 4  许可证检出分布")
ins.table(S["table4"], widths=[4.0, 2.4, 8.1])
ins.para("3.4 未检出许可证的文件", bold=True, indent=False, before=3, after=1)
ins.para("共 " + S["no_license"] + " 个文件未检出许可证，人工归类为三类：一是项目自身在 pom.xml 的 apache-rat-plugin excludes 中显式豁免的文件（stringEscapeUtilsTestData.txt、download_text.cgi 等）；二是构建与 CI 元数据（.gitignore、.asf.yaml、.github/workflows/*.yml 等）；三是纯二进制资源（.xcf、.png）与代码覆盖率文件。上述文件均非需要许可声明的源代码。")

# ============ 4 详细分析 ============
ins.para("4、检测报告详细分析", cn=FS, size=12, bold=True, indent=False, before=4, after=2)
ins.para("4.1 主许可证判定与声明一致性", bold=True, indent=False, before=3, after=1)
ins.para(S["s41"])
ins.para("4.2 GPL-2.0 命中的性质判定", bold=True, indent=False, before=3, after=1)
ins.para(S["s42"])
ins.para("4.3 未检出文件与项目自身 RAT 排除清单的交叉验证", bold=True, indent=False, before=3, after=1)
ins.para("将未检出许可证的文件清单与 pom.xml 中 apache-rat-plugin 的 excludes 配置逐条比对：全部未检出文件要么落在该排除清单内（如 stringEscapeUtilsTestData.txt、download_text.cgi），要么属于 RAT 默认不要求许可头的二进制/统计文件（logo.xcf、logo.png、profile.jacoco）。这说明 ScanCode 的“无许可证”结果与 ASF 自身的发布审计口径一致，未发现应声明而未声明许可的文件。")
ins.para("4.4 版权信息集中化对扫描结果的影响", bold=True, indent=False, before=3, after=1)
ins.para(S["s44"])
ins.para("4.5 依赖层面的义务分析", bold=True, indent=False, before=2, after=1)
ins.para("运行时依赖 commons-lang3 3.14.0 为 Apache-2.0，随 jar 传递分发时义务仅为随附其 LICENSE 与 NOTICE、不将 Apache 商标用于宣传；8 个 test 作用域依赖不进入发布物，不触发任何许可证义务。这印证了手册“误区 1”：义务取决于“是否修改并对外分发”，而非“是否使用了某许可证的组件”；本项目依赖均为宽松许可证，无 copyleft 传染风险。")
ins.para("4.6 工具误报、漏报与局限性", bold=True, indent=False, before=3, after=1)
ins.para(S["s46"])
ins.caption("表 5  合规风险清单与整改建议")
ins.table(S["table5"], widths=[3.6, 1.4, 4.7, 4.8])
ins.para("4.7 结论", bold=True, indent=False, before=2, after=1)
ins.para("本次扫描表明，Apache Commons Text 1.12.0 许可证合规状态良好：主许可证 Apache-2.0 与项目自带声明完全一致；唯一 GPL-2.0 命中经人工复核确认为测试语料文本巧合，不构成传染。ScanCode 识别准确、报告丰富，适合作为课程级合规审计的主力工具。")

doc.save(OUT)
print("saved:", OUT)
