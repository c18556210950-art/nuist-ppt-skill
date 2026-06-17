---
name: 南京信息工程大学 PPT 生成器
description: "生成南京信息工程大学 (NUIST) 官方风格的演示文稿。使用 python-pptx 直接操作模板，简单高效。适用于学术报告、论文答辩、课程作业、项目汇报。"
---

# 南京信息工程大学 PPT 生成器

## ⚠️ 重要规则

- **skill 文件夹内的文件是只读模板，禁止修改**
- **在工作文件夹（用户当前目录）生成 Python 脚本和 PPTX 输出**
- 模板路径在生成的脚本中硬编码为 skill 文件夹内的绝对路径

## 工作流程

1. 根据用户需求，在工作文件夹生成一个 `generate_ppt.py`
2. 运行 `python generate_ppt.py`
3. 输出在工作文件夹生成 `output.pptx`（或用户指定文件名）

---

## 模板信息

模板位置：`c:\Users\PC\.claude\skills\ppt\南京信息工程大学ppt模版.pptx`（16:9 宽屏）

### 配色

| 用途 | 色值 |
|------|------|
| 主色 | `#00679C` |
| 辅色 | `#2660AD` |
| 强调色 | `#4472C4` |
| 正文 | `#000000` |
| 白色 | `#FFFFFF` |

渐变：所有色块使用 `#00679C` → `#2660AD` 对角线渐变。

### 字体

中文：**微软雅黑** | 西文/数字：**Times New Roman**

### 校训

**明德格物  立己达人** — 封面和结尾页保留。

---

## 11 种版式

| # | 版式名称 | 用途 | find_layout 关键词 |
|---|---------|------|-------------------|
| 1 | 标题幻灯片 | 封面 | `标题幻灯片` |
| 2 | 目录1 | 目录（左侧深色块） | `目录1` |
| 3 | 目录2 | 目录（右侧深色块） | `目录2` |
| 4 | 自我介绍 | 个人/团队介绍 | `自我介绍` |
| 5 | 节标题 | 章节分隔 | `节标题` |
| 6 | 标题和内容 | 标准内容页 | `标题和内容` |
| 7 | 仅标题 | 自由排版 | `仅标题` |
| 8 | 三段式 | 三列并排 | `三段式` |
| 9 | 图文对比 | 左图右文 | `图文对比` |
| 10 | 空白 | 完全自定义 | `空白` |
| 11 | 结尾幻灯片 | 致谢/结束 | `结尾幻灯片` |

---

## 生成脚本模板

在工作文件夹生成的 `generate_ppt.py` 应包含以下结构：

```python
"""
南京信息工程大学 PPT 生成器
运行: python generate_ppt.py
依赖: pip install python-pptx
"""
from pptx import Presentation
import os

# ═══ 配置区 ═══

OUTPUT_FILE = "output.pptx"

COVER = {
    "title": "演示标题",
    "subtitle": "南京信息工程大学",
    "student": "姓名",
    "advisor": "导师",
}

TOC_ITEMS = ["目录项1", "目录项2", "目录项3"]

SECTIONS = [
    {"num": "1", "title": "章节标题", "subtitle": "Section Subtitle"},
]

CONTENT_SLIDES = [
    {"title": "内容标题", "bullets": ["要点一", "要点二", "要点三"]},
]

CLOSING = {
    "thanks": "感谢观看！",
    "subtitle": "Thank you!",
    "student": "姓名",
    "advisor": "导师",
}

# ═══ 模板路径（指向 skill 文件夹，不要改） ═══

TEMPLATE = r"c:\Users\PC\.claude\skills\ppt\南京信息工程大学ppt模版.pptx"

# ═══ 占位符类型 ═══

TITLE, BODY, CENTER_TITLE, SUBTITLE, OBJECT = 1, 2, 3, 4, 7

# ═══ 工具函数 ═══

def find_layout(prs, keyword):
    for ly in prs.slide_layouts:
        if keyword in ly.name:
            return ly
    return None

def fill_ph(slide, ph_type, text, idx=None):
    for ph in slide.placeholders:
        if ph.placeholder_format.type == ph_type:
            if idx is not None and ph.placeholder_format.idx != idx:
                continue
            ph.text = text
            return ph
    return None

def get_body_ph(slide):
    candidates = [ph for ph in slide.placeholders
                  if ph.placeholder_format.type in (BODY, OBJECT)]
    return max(candidates, key=lambda ph: ph.width) if candidates else None

def set_bullets(ph, texts):
    tf = ph.text_frame
    tf.clear()
    for i, text in enumerate(texts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text

def remove_template_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    while len(sldIdLst) > 0:
        rId = sldIdLst[0].get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        prs.part.drop_rel(rId)
        del sldIdLst[0]

# ═══ 幻灯片生成函数 ═══

def add_cover(prs):
    slide = prs.slides.add_slide(find_layout(prs, "标题幻灯片"))
    fill_ph(slide, CENTER_TITLE, COVER["title"])
    fill_ph(slide, SUBTITLE, COVER["subtitle"])
    for ph in slide.placeholders:
        if ph.placeholder_format.type == BODY:
            if ph.placeholder_format.idx == 17:
                ph.text = f">> 答辩学生：{COVER['student']}"
            elif ph.placeholder_format.idx == 18:
                ph.text = f">> 指导教师：{COVER['advisor']}"
    return slide

def add_toc(prs):
    slide = prs.slides.add_slide(find_layout(prs, "目录1"))
    ph = get_body_ph(slide)
    if ph:
        set_bullets(ph, TOC_ITEMS)
    return slide

def add_section(prs, num, title, subtitle):
    slide = prs.slides.add_slide(find_layout(prs, "节标题"))
    fill_ph(slide, TITLE, title)
    for ph in slide.placeholders:
        if ph.placeholder_format.type == BODY:
            if ph.placeholder_format.idx == 13:
                ph.text = num
            elif ph.placeholder_format.idx == 1:
                ph.text = subtitle
    return slide

def add_content(prs, title, bullets):
    slide = prs.slides.add_slide(find_layout(prs, "标题和内容"))
    fill_ph(slide, TITLE, title)
    ph = get_body_ph(slide)
    if ph:
        set_bullets(ph, bullets)
    return slide

def add_closing(prs):
    slide = prs.slides.add_slide(find_layout(prs, "结尾幻灯片"))
    fill_ph(slide, CENTER_TITLE, CLOSING["thanks"])
    fill_ph(slide, SUBTITLE, CLOSING["subtitle"])
    for ph in slide.placeholders:
        if ph.placeholder_format.type == BODY:
            if ph.placeholder_format.idx == 17:
                ph.text = f">> 答辩学生：{CLOSING['student']}"
            elif ph.placeholder_format.idx == 18:
                ph.text = f">> 指导教师：{CLOSING['advisor']}"
    return slide

# ═══ 主流程 ═══

def generate():
    print(f"打开模板: {TEMPLATE}")
    prs = Presentation(TEMPLATE)
    remove_template_slides(prs)

    print("[封面]")
    add_cover(prs)

    print("[目录]")
    add_toc(prs)

    for sec in SECTIONS:
        print(f"[章节] {sec['title']}")
        add_section(prs, sec["num"], sec["title"], sec["subtitle"])

    for cs in CONTENT_SLIDES:
        print(f"[内容] {cs['title']}")
        add_content(prs, cs["title"], cs["bullets"])

    print("[结尾]")
    add_closing(prs)

    prs.save(OUTPUT_FILE)
    print(f"完成: {OUTPUT_FILE} ({len(prs.slides)} slides)")

if __name__ == "__main__":
    generate()
```

---

## 关键点

| 要点 | 说明 |
|------|------|
| 模板路径 | 硬编码为 `c:\Users\PC\.claude\skills\ppt\南京信息工程大学ppt模版.pptx` |
| 占位符 type | TITLE=1, BODY=2, CENTER_TITLE=3, SUBTITLE=4, OBJECT=7 |
| 封面/结尾标题 | 用 CENTER_TITLE (3) |
| 内容页/节标题标题 | 用 TITLE (1) |
| 节标题章节号 | BODY idx=13 |
| 节标题副标题 | BODY idx=1 |
| 封面/结尾学生信息 | BODY idx=17 |
| 封面/结尾导师信息 | BODY idx=18 |
| 项目符号 | 用 `tf.add_paragraph()` 逐条添加，不要设 `ph.text` |

---

## 设计原则

1. **多样化版式** — 至少混用 3-4 种版式
2. **每页有视觉元素** — 图片、图表或色块，避免纯文字
3. **校徽不删** — 模板内嵌校徽保留不动
4. **不要标题下划线** — 用留白代替

---

## 依赖

```bash
pip install python-pptx
```
