import base64
import datetime
import io
import json
import math
import re
import urllib.parse
import uuid

import requests

import pandas as pd
import streamlit as st
from lxml import etree
from openai import OpenAI
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ----------------------------------------------------------------------------
# 基础配置
# ----------------------------------------------------------------------------
st.set_page_config(page_title="医学文献 → PPT 生成器", page_icon="🩺", layout="wide")

FREE_MODELS = [
    "nvidia/nemotron-3-super-120b-a12b:free",
    "qwen/qwen3.8-27b:free",
    "google/gemma-4-31b-it:free",
    "openrouter/free",
]
PAID_MODELS = [
    "anthropic/claude-3.5-sonnet",
    "deepseek/deepseek-chat",
    "openai/gpt-4o",
    "google/gemini-pro-1.5",
]
ALL_MODELS = FREE_MODELS + PAID_MODELS

# 医学风格配色：深蓝 + 青绿，白底
PRIMARY = RGBColor(0x0E, 0x5A, 0x8A)
ACCENT = RGBColor(0x1B, 0x9C, 0x9A)
TEXT = RGBColor(0x2B, 0x2B, 0x2B)
GRAY = RGBColor(0x6B, 0x72, 0x80)
LIGHT = RGBColor(0xEA, 0xF3, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT_LATIN = "Calibri"
FONT_EA = "Microsoft YaHei"

CAP_RE = re.compile(r"^\s*(fig(ure)?\.?\s*\d+|supplementary fig|图\s*\d+|附图)", re.I)
MIN_SIDE, MIN_AREA = 120, 30000

for _k, _v in dict(outline=None, images={}, ppt=None, source="", sel=0, added=set()).items():
    st.session_state.setdefault(_k, _v)


# ----------------------------------------------------------------------------
# 一、文献解析：文本 + 图片 + 表格
# ----------------------------------------------------------------------------
def rows_to_md(rows):
    return "\n".join("| " + " | ".join(r) + " |" for r in rows)


def norm_image(raw: bytes):
    """校验并统一为 PNG；过滤过小/极端长宽比的图（Logo、分隔线等）。"""
    try:
        im = Image.open(io.BytesIO(raw))
        im.load()
    except Exception:
        return None
    w, h = im.size
    if w < MIN_SIDE or h < MIN_SIDE or w * h < MIN_AREA or w / h > 8 or h / w > 8:
        return None
    if im.mode == "P":
        im = im.convert("RGBA")
    elif im.mode not in ("RGB", "RGBA", "L"):
        im = im.convert("RGB")
    im.thumbnail((1800, 1800))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue(), im.size[0], im.size[1]


def find_caption(blocks, rect):
    best, bd = "", 1e9
    for b in blocks:
        if len(b) > 6 and b[6] != 0:  # 非文本块
            continue
        x0, y0, x1, y1, txt = b[:5]
        txt = txt.strip().replace("\n", " ")
        if not CAP_RE.match(txt) or x1 < rect.x0 or x0 > rect.x1:
            continue
        if y0 >= rect.y1 - 20:
            d = y0 - rect.y1
        elif y1 <= rect.y0 + 20:
            d = (rect.y0 - y1) + 60
        else:
            continue
        if d < bd:
            best, bd = txt, d
    return best[:220] if bd < 260 else ""


def parse_pdf(data: bytes):
    import fitz  # PyMuPDF

    doc = fitz.open(stream=data, filetype="pdf")
    texts, imgs, seen = [], [], set()
    for pno, page in enumerate(doc, 1):
        texts.append(f"\n[第{pno}页]\n{page.get_text('text')}")
        try:
            for k, tb in enumerate(page.find_tables().tables):
                rows = [[(c or "").replace("\n", " ").strip() for c in r] for r in tb.extract()]
                if len(rows) >= 2 and len(rows[0]) >= 2:
                    texts.append(f"[原文表格 第{pno}页 #{k + 1}]\n" + rows_to_md(rows))
        except Exception:
            pass
        blocks = None
        for info in page.get_images(full=True):
            xref = info[0]
            if xref in seen:
                continue
            seen.add(xref)
            raw = None
            try:
                pix = fitz.Pixmap(doc, xref)
                if info[1]:
                    try:
                        pix = fitz.Pixmap(pix, fitz.Pixmap(doc, info[1]))  # 合并透明蒙版
                    except Exception:
                        pass
                if pix.n - pix.alpha >= 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                raw = pix.tobytes("png")
            except Exception:
                try:
                    raw = doc.extract_image(xref)["image"]
                except Exception:
                    continue
            n = norm_image(raw)
            if not n:
                continue
            caption = ""
            try:
                rects = page.get_image_rects(xref)
                if rects:
                    if blocks is None:
                        blocks = page.get_text("blocks")
                    caption = find_caption(blocks, rects[0])
            except Exception:
                pass
            imgs.append(dict(id=len(imgs) + 1, page=pno, caption=caption, bytes=n[0], w=n[1], h=n[2]))
    return "\n".join(texts), imgs


def parse_docx(data: bytes):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(io.BytesIO(data))
    out, imgs, pending = [], [], None
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            txt = Paragraph(child, d).text.strip()
            for rid in child.xpath(".//a:blip/@r:embed"):
                try:
                    n = norm_image(d.part.related_parts[rid].blob)
                except Exception:
                    n = None
                if n:
                    imgs.append(dict(id=len(imgs) + 1, page=None, caption="", bytes=n[0], w=n[1], h=n[2]))
                    pending = imgs[-1]
                    out.append(f"[图片{pending['id']}位置]")
            if txt:
                if pending is not None and CAP_RE.match(txt):
                    pending["caption"] = txt[:220]
                pending = None
                out.append(txt)
        elif tag == "tbl":
            rows = [[c.text.strip().replace("\n", " ") for c in r.cells] for r in Table(child, d).rows]
            if rows:
                out.append("[原文表格]\n" + rows_to_md(rows))
    return "\n".join(out), imgs


def parse_text(data: bytes):
    for enc in ("utf-8", "gbk", "latin-1"):
        try:
            return data.decode(enc), []
        except Exception:
            continue
    return "", []


def parse_file(uploaded):
    data = uploaded.getvalue()
    ext = uploaded.name.rsplit(".", 1)[-1].lower()
    if ext == "pdf":
        return parse_pdf(data)
    if ext == "docx":
        return parse_docx(data)
    return parse_text(data)


def strip_references(text: str) -> str:
    ms = list(re.finditer(r"\n\s*(references|bibliography|参考文献)\s*\n", text, re.I))
    if ms and ms[-1].start() > len(text) * 0.5:
        return text[: ms[-1].start()]
    return text


# ----------------------------------------------------------------------------
# 二、调用大模型生成大纲（JSON）
# ----------------------------------------------------------------------------
SYSTEM_PROMPT = """你是一位资深的医学科研汇报助手，负责把医学文献整理成组会/学术汇报用的幻灯片大纲。

【硬性规则】
1. 内容只能来自文献，严禁编造数据；数值、P 值、置信区间、样本量等必须与原文一致。
2. 常见结构（可按文献类型调整）：研究背景与目的 → 研究设计与方法（对象、纳排标准、干预/暴露、终点、统计）→ 结果（基线、主要终点、次要终点、亚组、安全性）→ 讨论 → 结论与临床意义 → 局限性。综述/指南类按主题逻辑组织。
3. 每页 3-6 条要点，每条精炼（中文不超过 40 字，英文不超过 25 词），专业术语保留英文缩写（如 HR、95%CI、RCT）。
4. 图片：只能使用「图片清单」里的图，通过 image_ids 引用，放在对应的结果/方法页；与主题无关的图不要用。不要要求生成任何新图片。
5. 表格：可以重新加工——只保留关键行列、合并或重命名表头、整理成适合幻灯片展示的小表（建议不超过 8 行 × 6 列），数值必须忠于原文。table.rows 的第一行是表头。没有合适的表格时 table 为 null。
6. 一页里图片、表格、要点可以组合，但不要塞得太满；有表格的页要点保持 0-3 条。
7. notes 写简短的讲者备注（可为空字符串）。
8. 只输出一个合法 JSON 对象，不要任何解释，不要 markdown 代码围栏。

【JSON 结构】
{
  "title": "PPT 标题",
  "subtitle": "副标题（期刊/年份/作者等，可为空）",
  "slides": [
    {
      "title": "页面标题",
      "bullets": ["要点1", "要点2", "  二级要点（行首两个空格表示二级）"],
      "image_ids": [1],
      "table": {"rows": [["表头1", "表头2"], ["数据", "数据"]]},
      "notes": ""
    }
  ]
}"""


def extract_json(s: str):
    s = re.sub(r"```(?:json)?", "", s)
    a, b = s.find("{"), s.rfind("}")
    if a < 0 or b < 0:
        raise ValueError("模型没有返回 JSON")
    s = s[a : b + 1]
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        return json.loads(re.sub(r",\s*([}\]])", r"\1", s))


def call_llm(api_key, models, user_prompt, on_fail):
    client = OpenAI(api_key=api_key, base_url="https://openrouter.ai/api/v1", timeout=240)
    for m in models:
        try:
            r = client.chat.completions.create(
                model=m,
                temperature=0.2,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
            )
            return extract_json(r.choices[0].message.content), m
        except Exception as e:  # 换下一个模型
            on_fail(m, e)
    raise RuntimeError("所有候选模型都调用失败，请检查 API Key、额度或换一个模型。")


def to_int(x):
    m = re.search(r"\d+", str(x))
    return int(m.group()) if m else None


def new_slide(title="", bullets=None, image_ids=None, table=None, notes=""):
    return dict(
        uid=uuid.uuid4().hex[:8], title=title, bullets=bullets or [], image_ids=image_ids or [],
        table=table, use_table=table is not None, table_cur=None, tv=0, notes=notes,
    )


def normalize_outline(data, images):
    slides = []
    for s in data.get("slides", []):
        bullets = s.get("bullets") or []
        if isinstance(bullets, str):
            bullets = bullets.split("\n")
        bullets = [str(b).rstrip() for b in bullets if str(b).strip()]
        ids = [to_int(i) for i in (s.get("image_ids") or [])]
        ids = [i for i in ids if i in images]
        t = s.get("table")
        rows = None
        if isinstance(t, dict):
            rows = t.get("rows")
            if rows and t.get("headers"):
                rows = [t["headers"]] + rows
        elif isinstance(t, list):
            rows = t
        if rows:
            rows = [[str(c) for c in r] for r in rows if isinstance(r, (list, tuple))]
            nc = max((len(r) for r in rows), default=0)
            rows = [r + [""] * (nc - len(r)) for r in rows] if nc and len(rows) >= 1 else None
        slides.append(new_slide(str(s.get("title", "")), bullets, ids, rows, str(s.get("notes", "") or "")))
    return dict(uid=uuid.uuid4().hex[:8], title=str(data.get("title", "")), subtitle=str(data.get("subtitle", "")), slides=slides)


# ----------------------------------------------------------------------------
# 三、生成原生、可编辑的 PPTX
# ----------------------------------------------------------------------------
def text_w(s, size):
    """估算文本宽度(pt)：中日韩字符占 1 个字号宽，其余约 0.55。"""
    return sum(size if ord(c) > 0x2E80 else size * 0.55 for c in s)


def text_height(lines, w_in, size):
    avail = max(w_in * 72 - 36, 50)
    total = 0
    for text, lvl in lines:
        n = max(1, math.ceil(text_w(text, size) / max(avail - lvl * 22, 30)))
        total += n * size * 1.25 + size * 0.45
    return total / 72


def fit_font(lines, w_in, h_in):
    for size in (20, 18, 16, 15, 14, 13, 12, 11, 10):
        if text_height(lines, w_in, size) <= h_in - 0.15:
            return size
    return 10


def parse_bullets(bullets):
    out = []
    for raw in bullets:
        lvl = 1 if raw.startswith(("\t", "  ")) else 0
        t = raw.strip().lstrip("-–•·*").strip()
        if t:
            out.append((t, lvl))
    return out


def style_run(run, size, bold=False, color=TEXT):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT_LATIN
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT_EA)


def set_bullet(p, level):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.3 + 0.35 * level))))
    pPr.set("indent", str(-int(Inches(0.28))))
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr")).set("val", "1B9C9A")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•" if level == 0 else "–")


def rect(slide, x, y, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def add_text(slide, x, y, w, h, text, size, color=GRAY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    style_run(r, size, bold, color)
    return tb


def set_placeholder(ph, text, size, x, y, w, h, anchor, color, bold=True):
    ph.left, ph.top, ph.width, ph.height = Inches(x), Inches(y), Inches(w), Inches(h)
    tf = ph.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.text = text
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for r in p.runs:
        style_run(r, size, bold, color)


def add_bullets(slide, x, y, w, h, lines):
    size = fit_font(lines, w, h)
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for k, (text, lvl) in enumerate(lines):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(size * 0.45)
        r = p.add_run()
        r.text = text
        style_run(r, size - 2 * lvl, False, TEXT)
        set_bullet(p, lvl)


NUM_RE = re.compile(r"^[<>≤≥~±+\-–]?\s*\d[\d.,%\s/()±\-–:]*$")


def add_table(slide, rows, x, y, w, max_h):
    nc = max(len(r) for r in rows)
    rows = [r + [""] * (nc - len(r)) for r in rows]
    nr = len(rows)
    weights = [max(4, min(30, max(text_w(r[j], 10) / 10 for r in rows))) for j in range(nc)]
    col_w = [w * wt / sum(weights) for wt in weights]

    def row_heights(font):
        hs = []
        for r in rows:
            lines = max(max(1, math.ceil(text_w(c, font) / max(col_w[j] * 72 - 12, 20))) for j, c in enumerate(r))
            hs.append(lines * font * 1.25 / 72 + 0.14)
        return hs

    fonts = (14, 13, 12, 11, 10, 9) if nc < 6 else (12, 11, 10, 9)
    for font in fonts:
        hs = row_heights(font)
        if sum(hs) <= max_h:
            break

    shape = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(sum(hs)))
    tbl = shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    for j in range(nc):
        tbl.columns[j].width = Inches(col_w[j])
    for i, r in enumerate(rows):
        tbl.rows[i].height = Inches(hs[i])
        for j, val in enumerate(r):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.08)
            c.margin_top = c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = PRIMARY if i == 0 else (LIGHT if i % 2 == 0 else WHITE)
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if (j == 0 and i > 0) else PP_ALIGN.CENTER
            run = p.add_run()
            run.text = val
            style_run(run, font, i == 0, WHITE if i == 0 else TEXT)
    return sum(hs)


def add_images(slide, imgs, x, y, w, h):
    n = len(imgs)
    if n == 1:
        cols = 1
    elif n == 2:
        cols = 2 if w >= h * 1.3 else 1
    elif n <= 4:
        cols = 2
    else:
        cols = 3
    rows = math.ceil(n / cols)
    gap = 0.2
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    cap_size = 11 if n == 1 else 9
    for k, im in enumerate(imgs):
        r, c = divmod(k, cols)
        cx, cy = x + c * (cw + gap), y + r * (ch + gap)
        cap = (im.get("caption") or "").strip()
        cap_h = 0.55 if (cap and ch > 2.2) else 0
        bw, bh = cw, ch - cap_h
        scale = min(bw / im["w"], bh / im["h"])
        pw, ph_ = im["w"] * scale, im["h"] * scale
        px, py = cx + (bw - pw) / 2, cy + (bh - ph_) / 2
        slide.shapes.add_picture(io.BytesIO(im["bytes"]), Inches(px), Inches(py), Inches(pw), Inches(ph_))
        if cap_h:
            text = cap if len(cap) <= 130 else cap[:129] + "…"
            add_text(slide, cx, py + ph_ + 0.04, cw, cap_h - 0.04, text, cap_size, GRAY, align=PP_ALIGN.CENTER)


def final_table(s):
    if not s.get("use_table"):
        return None
    rows = s.get("table_cur") or s.get("table")
    if not rows:
        return None
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    return rows or None


def build_pptx(outline, images, source_name):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    # ---- 封面 ----
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    rect(s0, 0, 0, 0.45, 7.5, PRIMARY)
    rect(s0, 1.2, 4.05, 2.2, 0.06, ACCENT)
    title = outline["title"] or "文献汇报"
    tsize = next((sz for sz in (38, 34, 30, 26, 22) if math.ceil(text_w(title, sz) / (10.8 * 72)) <= 2), 22)
    set_placeholder(s0.shapes.title, title, tsize, 1.2, 1.9, 10.8, 2.0, MSO_ANCHOR.BOTTOM, PRIMARY)
    set_placeholder(s0.placeholders[1], outline["subtitle"] or "", 20, 1.2, 4.3, 10.8, 1.4, MSO_ANCHOR.TOP, GRAY, bold=False)
    if source_name:
        add_text(s0, 1.2, 6.6, 10.8, 0.4, f"文献来源：{source_name}", 11, GRAY)

    # ---- 内容页 ----
    X0, Y0, W, H = 0.6, 1.6, 12.13, 5.2
    for idx, s in enumerate(outline["slides"], start=2):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        t = s["title"] or " "
        tsz = next((sz for sz in (28, 26, 24, 22, 20) if text_w(t, sz) <= 12.1 * 72), 20)
        set_placeholder(slide.shapes.title, t, tsz, 0.6, 0.35, 12.13, 0.9, MSO_ANCHOR.MIDDLE, PRIMARY)
        rect(slide, 0.6, 1.3, 12.13, 0.02, PRIMARY)
        rect(slide, 0.6, 1.285, 1.6, 0.05, ACCENT)

        imgs = [images[i] for i in s["image_ids"] if i in images]
        lines = parse_bullets(s["bullets"])
        table = final_table(s)
        has_left = bool(lines or table)

        if imgs and has_left:
            lw = W * (0.55 if table else 0.45)
            rx, rw = X0 + lw + 0.3, W - lw - 0.3
        else:
            lw, rx, rw = W, X0, W

        if lines and table:
            bh = min(text_height(lines, lw, 16) + 0.2, H * 0.4)
            add_bullets(slide, X0, Y0, lw, bh, lines)
            add_table(slide, table, X0, Y0 + bh + 0.15, lw, H - bh - 0.15)
        elif lines:
            add_bullets(slide, X0, Y0, lw, H, lines)
        elif table:
            add_table(slide, table, X0, Y0, lw, H)
        if imgs:
            add_images(slide, imgs, rx, Y0, rw, H)

        rect(slide, 0.6, 6.98, 12.13, 0.01, RGBColor(0xD1, 0xD5, 0xDB))
        if source_name:
            add_text(slide, 0.6, 7.03, 9.0, 0.3, source_name, 10, GRAY)
        add_text(slide, 11.7, 7.03, 1.03, 0.3, str(idx), 10, GRAY, align=PP_ALIGN.RIGHT)
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# 四、编辑器回调
# ----------------------------------------------------------------------------
def github_put(token, repo, branch, path, content: bytes, message):
    """通过 GitHub Contents API 创建/更新文件，返回文件网页链接。"""
    url = f"https://api.github.com/repos/{repo}/contents/{urllib.parse.quote(path)}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    r = requests.get(url, headers=headers, params={"ref": branch}, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    payload = {"message": message, "content": base64.b64encode(content).decode(), "branch": branch}
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=headers, json=payload, timeout=120)
    if r.status_code not in (200, 201):
        try:
            msg = r.json().get("message", r.text)
        except Exception:
            msg = r.text
        raise RuntimeError(f"GitHub 返回 {r.status_code}：{msg[:200]}")
    return r.json()["content"]["html_url"]


def outline_to_json(outline):
    slides = [
        dict(title=s["title"], bullets=s["bullets"], image_ids=s["image_ids"], table=final_table(s), notes=s["notes"])
        for s in outline["slides"]
    ]
    return json.dumps(dict(title=outline["title"], subtitle=outline["subtitle"], slides=slides), ensure_ascii=False, indent=2)


def commit_tables():
    """把表格编辑结果固化为新的基准，并更换 key，避免 data_editor 状态叠加。"""
    if not st.session_state.outline:
        return
    for s in st.session_state.outline["slides"]:
        if s.get("table_cur") is not None:
            s["table"], s["table_cur"] = s["table_cur"], None
            s["tv"] += 1


def cb_move(i, d):
    commit_tables()
    sl = st.session_state.outline["slides"]
    j = i + d
    if 0 <= j < len(sl):
        sl[i], sl[j] = sl[j], sl[i]
        st.session_state.sel = j


def cb_delete(i):
    commit_tables()
    sl = st.session_state.outline["slides"]
    if sl:
        sl.pop(i)
        st.session_state.sel = max(0, min(i, len(sl) - 1))


def cb_add(i):
    commit_tables()
    sl = st.session_state.outline["slides"]
    sl.insert(i + 1, new_slide("新页面", ["要点 1"]))
    st.session_state.sel = i + 1


def get_thumb(im):
    """生成并缓存小尺寸缩略图，仅用于网页预览，不影响 PPT 中的原图。"""
    if "thumb" not in im:
        t = Image.open(io.BytesIO(im["bytes"]))
        t.thumbnail((320, 320))
        buf = io.BytesIO()
        t.convert("RGB").save(buf, "JPEG", quality=80)
        im["thumb"] = buf.getvalue()
    return im["thumb"]


def img_label(images, i):
    im = images[i]
    cap = (im.get("caption") or "")[:24]
    pg = f"第{im['page']}页 " if im.get("page") else ""
    return f"图{i}（{pg}{cap}）"


def render_slide_editor(s, i, images):
    uid = s["uid"]
    s["title"] = st.text_input("页面标题", s["title"], key=f"t_{uid}")
    txt = st.text_area(
        "要点（每行一条；行首空两格 = 二级要点）",
        "\n".join(s["bullets"]), height=200, key=f"b_{uid}",
    )
    s["bullets"] = [l.rstrip() for l in txt.split("\n") if l.strip()]

    # 图片
    if images:
        st.markdown("**本页使用的原文图片**（勾选缩略图下方的复选框）")
        opts = sorted(images)
        chosen = []
        for row_start in range(0, len(opts), 5):
            cols = st.columns(5)
            for col, iid in zip(cols, opts[row_start : row_start + 5]):
                with col:
                    st.image(get_thumb(images[iid]), use_container_width=True)
                    if st.checkbox(img_label(images, iid), iid in s["image_ids"], key=f"pick_{uid}_{iid}"):
                        chosen.append(iid)
        s["image_ids"] = chosen
        if chosen:
            st.markdown("**已选图片的图注**（显示在幻灯片图片下方，可修改或清空）")
            cols = st.columns(min(3, len(chosen)))
            for k, iid in enumerate(chosen):
                with cols[k % len(cols)]:
                    images[iid]["caption"] = st.text_input(f"图{iid} 图注", images[iid].get("caption", ""), key=f"cap_{iid}")

    # 表格
    s["use_table"] = st.checkbox("本页包含表格", s["use_table"], key=f"ut_{uid}")
    if s["use_table"]:
        if not s["table"]:
            s["table"] = [["指标", "组 A", "组 B"], ["", "", ""], ["", "", ""]]
        base = s["table"]
        nc = len(base[0])
        st.caption("第一行为表头；可直接增删行、修改单元格。")
        new_nc = st.number_input("列数", 1, 10, nc, key=f"nc_{uid}")
        if new_nc != nc:
            cur = s["table_cur"] or s["table"]
            s["table"] = [(list(r) + [""] * new_nc)[:new_nc] for r in cur]
            s["table_cur"] = None
            s["tv"] += 1
            st.rerun()
        df = pd.DataFrame(base, columns=[f"列{j + 1}" for j in range(nc)]).astype(str)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"tb_{uid}_{nc}_{s['tv']}")
        s["table_cur"] = edited.fillna("").astype(str).values.tolist()

    s["notes"] = st.text_area("讲者备注（可选）", s["notes"], height=80, key=f"n_{uid}")


# ----------------------------------------------------------------------------
# 五、界面
# ----------------------------------------------------------------------------
st.title("🩺 医学文献 → 可编辑 PPT")
st.caption("上传文献 → AI 提取大纲 → 在网页中修改 → 生成原生 PowerPoint（无动画、文字/表格均可编辑，仅使用文中图片）")

with st.sidebar:
    st.header("⚙️ 模型设置")
    try:
        secret_key = st.secrets.get("OPENROUTER_API_KEY", "")
    except Exception:
        secret_key = ""
    api_key = st.text_input("OpenRouter API Key", type="password", placeholder="sk-or-v1-...") or secret_key
    model_name = st.selectbox("选择模型", ALL_MODELS, index=0)
    custom_model = st.text_input("自定义模型 ID（可选，优先使用）", placeholder="如 anthropic/claude-sonnet-4.5")
    auto_fallback = st.checkbox("失败时自动尝试其他免费模型", True)
    max_chars = st.slider("送入模型的最大字符数", 10000, 120000, 45000, 5000,
                          help="免费模型上下文较短，长文献请调小或换用大上下文模型。")
    st.markdown("---")
    st.header("☁️ GitHub 保存")
    try:
        gh_token_secret = st.secrets.get("GITHUB_TOKEN", "")
        gh_repo_secret = st.secrets.get("GITHUB_REPO", "")
    except Exception:
        gh_token_secret, gh_repo_secret = "", ""
    gh_token = st.text_input("GitHub Token", type="password", placeholder="github_pat_...",
                             help="需要对目标仓库有 Contents: Read and write 权限。") or gh_token_secret
    gh_repo = st.text_input("仓库（owner/repo）", value=gh_repo_secret, placeholder="your-name/ppt-output")
    gh_branch = st.text_input("分支", value="main")
    gh_folder = st.text_input("保存目录", value="ppt_output")
    gh_with_json = st.checkbox("同时保存大纲 JSON", True)
    st.markdown("---")
    st.markdown("提示：扫描版 PDF 无文字层，需先做 OCR；矢量绘制的图表无法被提取，可在编辑阶段手动上传截图。")

st.subheader("① 上传文献并生成大纲")
uploaded = st.file_uploader("支持 PDF / Word(.docx) / Markdown / TXT", type=["pdf", "docx", "md", "txt"])
extra = st.text_area("补充要求（可选）", placeholder="例如：用于科室组会汇报，重点讲方法学和主要结果，不讲背景细节。")
c1, c2 = st.columns(2)
n_slides = c1.slider("正文页数（不含封面）", 6, 30, 12)
lang = c2.selectbox("PPT 语言", ["与文献一致", "中文", "English"])

if st.button("🚀 分析文献并生成大纲", type="primary"):
    if not api_key:
        st.error("请先在左侧填写 OpenRouter API Key。")
    elif not uploaded:
        st.error("请先上传文献。")
    else:
        try:
            with st.status("正在处理…", expanded=True) as status:
                st.write("解析文献（文本 / 图片 / 表格）…")
                text, imgs = parse_file(uploaded)
                text = strip_references(text).strip()
                if len(text) < 200:
                    raise RuntimeError("几乎没有提取到文字，可能是扫描版 PDF，请先 OCR。")
                images = {im["id"]: im for im in imgs}
                st.write(f"提取到 {len(text)} 字符、{len(images)} 张图片。")
                if len(text) > max_chars:
                    half = max_chars // 2
                    text = text[:half] + "\n\n[……中间部分因长度限制省略……]\n\n" + text[-half:]

                img_desc = "\n".join(
                    f"图片{i}: " + (f"第{m['page']}页; " if m["page"] else "") + f"{m['w']}x{m['h']}; 原文图注: {m['caption'] or '无'}"
                    for i, m in images.items()
                ) or "（无可用图片）"
                lang_req = {"与文献一致": "与文献语言一致", "中文": "全部使用中文（专业缩写保留英文）", "English": "全部使用英文"}[lang]
                user_prompt = (
                    f"请为下面这篇医学文献生成幻灯片大纲。\n- 正文约 {n_slides} 页（不含封面）\n- 语言：{lang_req}\n"
                    f"- 额外要求：{extra.strip() or '无'}\n\n【图片清单】\n{img_desc}\n\n【文献正文】\n{text}"
                )
                first = custom_model.strip() or model_name
                models = [first] + ([m for m in FREE_MODELS if m != first] if auto_fallback else [])
                st.write(f"调用模型 `{first}` 生成大纲…")
                data, used = call_llm(api_key, models, user_prompt, lambda m, e: st.write(f"⚠️ `{m}` 失败：{str(e)[:160]}"))
                outline = normalize_outline(data, images)
                if not outline["slides"]:
                    raise RuntimeError("模型返回的大纲为空，请重试或换模型。")
                st.session_state.update(outline=outline, images=images, ppt=None, sel=0, source=uploaded.name, added=set())
                status.update(label=f"大纲已生成（模型：{used}）", state="complete")
        except Exception as e:
            st.error(f"失败：{e}")

# ---------------- 编辑区 ----------------
outline = st.session_state.outline
if outline:
    images = st.session_state.images
    st.markdown("---")
    st.subheader("② 编辑大纲")

    with st.expander(f"📷 文献中可用的图片（{len(images)} 张）/ 手动补充图片"):
        extra_imgs = st.file_uploader("补充上传图片（如矢量图表的截图）", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="extra_imgs")
        for f in extra_imgs or []:
            sig = (f.name, f.size)
            if sig in st.session_state.added:
                continue
            n = norm_image(f.getvalue())
            if n:
                nid = max(images, default=0) + 1
                images[nid] = dict(id=nid, page=None, caption=f.name.rsplit(".", 1)[0], bytes=n[0], w=n[1], h=n[2])
                st.session_state.added.add(sig)
        if images:
            cols = st.columns(5)
            for k, iid in enumerate(sorted(images)):
                with cols[k % 5]:
                    st.image(images[iid]["bytes"], use_container_width=True)
                    st.caption(img_label(images, iid))
        else:
            st.info("未从文献中提取到图片。")

    with st.expander("封面", expanded=False):
        outline["title"] = st.text_input("PPT 标题", outline["title"], key=f"ot_{outline['uid']}")
        outline["subtitle"] = st.text_input("副标题", outline["subtitle"], key=f"os_{outline['uid']}")

    sl = outline["slides"]
    n = len(sl)
    with st.expander(f"📋 大纲总览（共 {n + 1} 页，含封面）"):
        for k, s in enumerate(sl, start=2):
            extras = (" 🖼" if s["image_ids"] else "") + (" 📊" if final_table(s) else "")
            st.markdown(f"**{k}. {s['title']}**{extras}")
            for b in s["bullets"]:
                st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;{'– ' if b.startswith('  ') else '• '}{b.strip()}")

    if n:
        if st.session_state.sel >= n:
            st.session_state.sel = n - 1
        idx = st.selectbox(
            "选择要编辑的页面", range(n), key="sel", on_change=commit_tables,
            format_func=lambda i: f"第 {i + 2} 页 · {sl[i]['title'] or '(无标题)'}",
        )
        b1, b2, b3, b4, _ = st.columns([1, 1, 1, 1.4, 3])
        b1.button("⬆ 上移", on_click=cb_move, args=(idx, -1), disabled=idx == 0)
        b2.button("⬇ 下移", on_click=cb_move, args=(idx, 1), disabled=idx == n - 1)
        b3.button("🗑 删除", on_click=cb_delete, args=(idx,))
        b4.button("➕ 在其后新增", on_click=cb_add, args=(idx,))
        with st.container(border=True):
            render_slide_editor(sl[idx], idx, images)
    else:
        st.warning("没有页面了。")
        st.button("➕ 新增页面", on_click=cb_add, args=(-1,))

    # ---------------- 生成 ----------------
    st.markdown("---")
    st.subheader("③ 生成 PPT")
    if st.button("📊 生成 PowerPoint", type="primary"):
        try:
            st.session_state.ppt = build_pptx(outline, images, st.session_state.source)
        except Exception as e:
            st.error(f"生成失败：{e}")
    if st.session_state.ppt:
        st.success("生成成功！文字、表格、图片均为原生对象，可在 PowerPoint / WPS 中继续编辑。")
        fname = re.sub(r'[\\/:*?"<>|]', "_", outline["title"] or "presentation")[:40] + ".pptx"
        st.download_button(
            "📥 下载 .pptx", st.session_state.ppt, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        if st.button("☁️ 保存到 GitHub"):
            if not (gh_token and gh_repo.strip()):
                st.error("请先在左侧填写 GitHub Token 和仓库（owner/repo）。")
            else:
                try:
                    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    base = f"{stamp}_{fname[:-5]}"
                    folder = gh_folder.strip().strip("/")
                    prefix = f"{folder}/" if folder else ""
                    repo, branch = gh_repo.strip(), (gh_branch.strip() or "main")
                    with st.spinner("正在上传到 GitHub…"):
                        url = github_put(gh_token, repo, branch, f"{prefix}{base}.pptx",
                                         st.session_state.ppt, f"Add presentation: {base}")
                        if gh_with_json:
                            github_put(gh_token, repo, branch, f"{prefix}{base}.outline.json",
                                       outline_to_json(outline).encode("utf-8"), f"Add outline: {base}")
                    st.success(f"已保存到 GitHub：[{url}]({url})")
                except Exception as e:
                    st.error(f"保存失败：{e}")
