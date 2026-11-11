#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从各卷《README_溯源总表.md》的「总览速查」表生成可公开的逐题对照表。

设计原则：
1. 零外部依赖（xlsx 由标准库 zipfile + XML 直接构造，避免写入本机用户名元数据）。
2. 只使用相对路径，脚本本身可随交付文件夹一同公开。
3. 不生成 DOI。DOI 仅来自：原表记载 / 项目内交叉检索 /（可选）CrossRef 标题核验。
   查不到一律留空并标注「待核」，来源分级写入「DOI来源」列。
4. 输出前做敏感信息黑名单校验，命中即中止。

用法：
    python build_public_table.py            # 离线：仅项目内交叉补全
    python build_public_table.py --online   # 追加 CrossRef 标题核验（单次探测，失败即跳过）
"""

import csv
import datetime
import difflib
import html
import json
import re
import sys
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

OUT = Path(__file__).resolve().parent
ROOT = OUT.parent
VERSION = "第一版"
BUILD_DATE = "2026-11-10"
BASENAME = "高考生物真题×顶刊论文对照表-" + VERSION

HEADERS = [
    "序号", "年份", "卷别", "卷型", "题号",
    "考查情境（概括）", "溯源论文标题", "期刊/年份", "DOI",
    "DOI来源", "关联度", "关联度记法", "确证", "备注",
]

# 输出字段中禁止出现的内容（本地路径、身份信息、内部通道）
BLACKLIST = [
    "file:///", "dyyti", "C:\\Users", "C:/Users", "阳阳", "富平",
    "科研通", "ablesci", "Sci-Hub", "sci-hub", ".workbuddy",
    "dyynas", "spaceId", "folderId", "192.168.",
]

DOI_RE = re.compile(r"10\.\d{4,9}/[A-Za-z0-9._()\-/:;\[\]]+")
MD_LINK_RE = re.compile(r"!?\[([^\]]*)\]\([^)]*\)")


def clean(text):
    """清除 markdown 标记与本地链接，压缩空白。"""
    if text is None:
        return ""
    s = str(text)
    s = MD_LINK_RE.sub(r"\1", s)
    s = re.sub(r"file:///\S+", "", s)
    s = s.replace("`", "").replace("**", "")
    s = re.sub(r"\*([^*\n]+)\*", r"\1", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def split_row(line):
    parts = line.strip().split("|")
    if parts and parts[0].strip() == "":
        parts = parts[1:]
    if parts and parts[-1].strip() == "":
        parts = parts[:-1]
    return [p.strip() for p in parts]


def extract_doi(text):
    m = DOI_RE.search(text or "")
    if not m:
        return ""
    return m.group(0).rstrip(".,;:）)]】》")


def star_count(text):
    return min(5, (text or "").count("★"))


def parse_confidence(raw):
    """归一化置信度。返回 (显示值, 星级数或 None, 记法, 备注)。

    原表混用三种记法：纯星级（★★★★ / ★★★★☆）、简写（★5）、文字等级（高/中高/待核）。
    文字等级不推断为星级，避免制造原表不存在的判定。
    """
    t = clean(raw)
    if not t:
        return "未标注", None, "未标注", ""
    m = re.fullmatch(r"[★☆\s]*★\s*([1-5])[★☆\s]*", t)
    if m:
        n = int(m.group(1))
        return "★" * n + "☆" * (5 - n), n, "星级", ""
    n = star_count(t)
    if n:
        n = min(5, n)
        extra = re.sub(r"[★☆\s]", "", t)
        extra = re.sub(r"^[\（\(/、,，:：]+", "", extra).rstrip("）) ")
        return "★" * n + "☆" * (5 - n), n, "星级", extra.strip()
    for lv in ("中高", "中", "高", "低", "待核"):
        if t.startswith(lv):
            extra = t[len(lv):].lstrip("（(、,，/ ").rstrip("）) ")
            return lv, None, "文字等级", extra.strip()
    return t, None, "文字等级", ""


def norm(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


def parse_volume(dirname):
    name = dirname.replace("_论文溯源", "")
    m = re.match(r"^(\d{4})(.*)$", name)
    year, vol = (m.group(1), m.group(2)) if m else ("", name)
    if "合格考" in vol:
        kind = "合格考"
    elif any(k in vol for k in ("模", "月考", "九调", "调考", "联考")):
        kind = "模考/调研"
    elif "选考" in vol:
        kind = "选考（真题）"
    else:
        kind = "高考真题"
    return year, vol, kind


def find_index(cols, *names):
    for i, c in enumerate(cols):
        if any(n in c for n in names):
            return i
    return -1


def sort_key(rec):
    nums = re.findall(r"\d+", rec["题号"]) or ["9999"]
    return (rec["年份"], rec["卷别"], int(nums[0]))


def collect_records():
    records = []
    for readme in sorted(ROOT.glob("*论文溯源/README_溯源总表.md")):
        lines = readme.read_text(encoding="utf-8", errors="ignore").splitlines()
        header_at = -1
        cols = []
        for i, line in enumerate(lines):
            if not line.strip().startswith("|"):
                continue
            c = split_row(line)
            if c and c[0] == "题号" and find_index(c, "溯源论文") >= 0 and find_index(c, "DOI") >= 0:
                header_at, cols = i, c
                break
        if header_at < 0:
            print("  [跳过] 未找到速查表头:", readme.parent.name)
            continue

        ix = {
            "q": find_index(cols, "题号"),
            "scene": find_index(cols, "试题情境", "试题主题"),
            "title": find_index(cols, "溯源论文"),
            "journal": find_index(cols, "期刊"),
            "doi": find_index(cols, "DOI"),
            "star": find_index(cols, "置信度"),
        }
        year, vol, kind = parse_volume(readme.parent.name)
        j = header_at + 2
        count = 0
        while j < len(lines) and lines[j].strip().startswith("|"):
            c = split_row(lines[j])
            j += 1

            def get(key):
                k = ix[key]
                return c[k] if 0 <= k < len(c) else ""

            q = clean(get("q"))
            title = clean(get("title"))
            scene = clean(get("scene"))
            if not q or q in ("—", "-", "–") or not title or title in ("无", "—", "-"):
                continue
            if "未检出" in scene or "未发现" in title:
                continue
            # 以「待核」开头的是占位/线索行，不是已建立的溯源映射，不进公开表
            if re.match(r"^待核", title):
                continue

            star_raw = get("star")
            star_txt, stars, notation, note = parse_confidence(star_raw)
            doi = extract_doi(get("doi"))
            if stars is not None:
                confirmed = "是" if stars >= 3 else "否"
            elif star_txt == "待核":
                confirmed = "否"
            else:
                confirmed = "未分级"

            records.append({
                "年份": year, "卷别": vol, "卷型": kind, "题号": q,
                "考查情境（概括）": scene, "溯源论文标题": title,
                "期刊/年份": clean(get("journal")),
                "DOI": doi,
                "DOI来源": "原表记载" if doi else "",
                "关联度": star_txt,
                "关联度记法": notation,
                "确证": confirmed,
                "备注": note,
                "_stars": stars,
                "_file": readme.parent.name,
            })
            count += 1
        print(f"  {count:3d} 行  {readme.parent.name}")
    records.sort(key=sort_key)
    return records


def build_doi_index():
    """扫描项目内 md，建立 (上下文文本 -> DOI) 索引，用于交叉补全。"""
    index = []
    seen = set()
    for p in sorted(ROOT.rglob("*.md")):
        s = str(p)
        if ".workbuddy" in s or OUT.name in s:
            continue
        lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
        for i, line in enumerate(lines):
            for d in set(DOI_RE.findall(line)):
                d = d.rstrip(".,;:）)]】》")
                if len(d) < 8:
                    continue
                ctx = " ".join(lines[max(0, i - 2): i + 3])
                key = (d, norm(ctx)[:600])
                if key in seen:
                    continue
                seen.add(key)
                index.append((d, ctx, norm(ctx)))
    return index


def fill_doi_internal(records, index):
    filled = 0
    for r in records:
        if r["DOI"]:
            continue
        title = r["溯源论文标题"]
        nt = norm(title)
        cands = set()
        if len(nt) >= 20:
            probe = nt[:30]
            for d, raw, nctx in index:
                if probe in nctx:
                    cands.add(d)
        else:
            probe = title[:12]
            for d, raw, nctx in index:
                if probe and probe in raw:
                    cands.add(d)
        if len(cands) == 1:
            r["DOI"] = cands.pop()
            r["DOI来源"] = "项目内交叉检索"
            filled += 1
    return filled


def crossref_probe():
    try:
        req = urllib.request.Request(
            "https://api.crossref.org/works?rows=1",
            headers={"User-Agent": "public-table-builder/1.0 (educational research)"},
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            return resp.status == 200
    except Exception as e:
        print("  [网络不可用，跳过 CrossRef]", type(e).__name__)
        return False


def _crossref_query(title):
    if not title or len(norm(title)) < 20:
        return ""
    url = "https://api.crossref.org/works?rows=3&query.title=" + urllib.parse.quote(title[:180])
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": "public-table-builder/1.0 (educational research)"}
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.load(resp)
        target = norm(title)
        best, score = "", 0.0
        for item in data.get("message", {}).get("items", []):
            cand = (item.get("title") or [""])[0]
            ratio = difflib.SequenceMatcher(None, target, norm(cand)).ratio()
            if ratio > score:
                best, score = (item.get("DOI") or ""), ratio
        return best if score >= 0.88 else ""
    except Exception:
        return ""


def crossref_lookup(title):
    """先用完整标题；若原表标题带中文附注，回退到首个括号前的英文主体再试一次。"""
    candidates = [title]
    cut = re.split(r"[（(]", title)[0].strip()
    if len(cut) >= 20 and cut != title:
        candidates.append(cut)
    for t in candidates:
        d = _crossref_query(t)
        if d:
            return d
    return ""


def crossref_fetch_by_doi(doi):
    """按 DOI 反查正式标题，用于闭环核验，返回 (标题, 相似度基准文本)。"""
    url = "https://api.crossref.org/works/" + urllib.parse.quote(doi, safe="")
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": "public-table-builder/1.0 (educational research)"}
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            item = json.load(resp).get("message", {})
        title = (item.get("title") or [""])[0]
        container = (item.get("container-title") or [""])[0]
        return title, container
    except Exception:
        return "", ""


def fill_doi_online(records):
    if not crossref_probe():
        return 0
    filled, tried = 0, 0
    for r in records:
        if r["DOI"] or tried >= 45:
            continue
        if r["确证"] == "否":
            continue
        tried += 1
        d = crossref_lookup(r["溯源论文标题"])
        if not d:
            continue
        official, container = crossref_fetch_by_doi(d)
        nt, no = norm(r["溯源论文标题"]), norm(official)
        if official and no and (no in nt or nt.startswith(no)):
            ratio = 1.0
        else:
            ratio = difflib.SequenceMatcher(None, nt, no).ratio() if official else 0.0
        # 反查失败不剔除（避免网络波动误删），如实记录核验状态
        r["_cr_title"] = official
        r["_cr_journal"] = container
        r["_cr_ratio"] = f"{ratio:.2f}" if official else "未反查"
        r["DOI"], r["DOI来源"] = d, "CrossRef 标题核验"
        filled += 1
    return filled


def validate(records):
    problems = []
    for i, r in enumerate(records, 1):
        for k, v in r.items():
            if k.startswith("_") or not isinstance(v, str):
                continue
            for bad in BLACKLIST:
                if bad.lower() in v.lower():
                    problems.append((i, k, bad, v[:80]))
        if r["DOI"] and not re.fullmatch(r"10\.\d{4,9}/\S+", r["DOI"]):
            problems.append((i, "DOI", "格式非法", r["DOI"]))
    return problems


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def xesc(s):
    return html.escape(str(s), quote=False)


def write_xlsx(path, headers, rows):
    n = len(rows) + 1
    last = col_letter(len(headers))

    def cell(ref, value, style):
        return (f'<c r="{ref}" s="{style}" t="inlineStr">'
                f'<is><t xml:space="preserve">{xesc(value)}</t></is></c>')

    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
           f'<dimension ref="A1:{last}{n}"/>',
           '<sheetViews><sheetView workbookViewId="0">',
           '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
           '</sheetView></sheetViews>',
           '<sheetFormatPr defaultRowHeight="15"/>',
           '<cols>']
    widths = [6, 7, 24, 13, 9, 52, 62, 26, 34, 18, 11, 12, 9, 26]
    for i, w in enumerate(widths[:len(headers)], 1):
        out.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
    out.append('</cols><sheetData>')
    out.append('<row r="1">' + "".join(
        cell(f"{col_letter(i)}1", h, 1) for i, h in enumerate(headers, 1)) + '</row>')
    for ri, row in enumerate(rows, 2):
        out.append(f'<row r="{ri}">' + "".join(
            cell(f"{col_letter(i)}{ri}", v, 0) for i, v in enumerate(row, 1)) + '</row>')
    out.append('</sheetData>')
    out.append(f'<autoFilter ref="A1:{last}{n}"/>')
    out.append('</worksheet>')
    sheet = "".join(out)

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="逐题对照表" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
    ts = BUILD_DATE + "T00:00:00Z"
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>高考生物真题×顶刊论文对照表 {VERSION}</dc:title>
<dc:creator>新高考生物试题与顶刊论文相关性研究</dc:creator>
<cp:lastModifiedBy>新高考生物试题与顶刊论文相关性研究</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">{ts}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">{ts}</dcterms:modified>
</cp:coreProperties>'''
    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>public-table-builder</Application></Properties>'''

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in [
            ("[Content_Types].xml", content_types), ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook), ("xl/_rels/workbook.xml.rels", wb_rels),
            ("xl/worksheets/sheet1.xml", sheet), ("xl/styles.xml", styles),
            ("docProps/core.xml", core), ("docProps/app.xml", app),
        ]:
            z.writestr(name, data.encode("utf-8"))


def write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


def write_readme(path, records, online_used):
    total = len(records)
    confirmed = sum(1 for r in records if r["确证"] == "是")
    ungraded = sum(1 for r in records if r["确证"] == "未分级")
    rejected = sum(1 for r in records if r["确证"] == "否")
    with_doi = sum(1 for r in records if r["DOI"])
    real = sum(1 for r in records if r["卷型"] in ("高考真题", "选考（真题）") and r["确证"] == "是")
    src = {}
    for r in records:
        k = r["DOI来源"] or "待核"
        src[k] = src.get(k, 0) + 1
    src_txt = "\n".join(f"| {k} | {v} |" for k, v in sorted(src.items(), key=lambda x: -x[1]))
    years = sorted({r["年份"] for r in records if r["年份"]})
    vols = len({r["卷别"] for r in records})

    def md(s, limit=90):
        return s.replace("|", "／").replace("\n", " ")[:limit]

    cr_rows = [r for r in records if r["DOI来源"] == "CrossRef 标题核验"]
    cr_txt = "\n".join(
        f"| {md(r['卷别'])} Q{md(r['题号'], 12)} | {r['DOI']} | "
        f"{md(r.get('_cr_title', ''), 70)} | {r.get('_cr_ratio', '未反查')} |"
        for r in cr_rows
    ) or "| — | — | 本次未启用联网核验 | — |"
    todo_rows = [r for r in records if not r["DOI"]]
    todo_txt = "\n".join(
        f"| {md(r['卷别'])} Q{md(r['题号'], 12)} | {md(r['溯源论文标题'], 70)} | "
        f"{md(r['期刊/年份'], 30)} |"
        for r in todo_rows
    ) or "| — | 无待核条目 | — |"

    text = f"""# 高考生物真题 × 顶刊论文对照表 · {VERSION}

> 生成日期：{BUILD_DATE}　|　数据版本：{VERSION}　|　许可：CC BY-NC-SA 4.0（见 `LICENSE.txt`）

## 这是什么

逐题登记新高考生物试题与原始研究论文的对应关系，供教研比对、检索与核查使用。
**本表不含**试题原文、答案、解析、题目图片、论文全文 PDF，也不含任何第三方题库整理稿。
「考查情境（概括）」为对题目素材的简述，不是题干复制。

## 数据概况

| 指标 | 数值 |
|---|---:|
| 记录总数 | {total} |
| 「确证」（星级 ≥ ★★★） | {confirmed} |
| 「未分级」（原表用文字等级，未折算星级） | {ungraded} |
| 「否」（低于三星或明确标待核） | {rejected} |
| 含 DOI 的记录 | {with_doi}（{with_doi * 100 // total if total else 0}%） |
| 覆盖年份 | {"、".join(years)} |
| 覆盖卷别 | {vols} |
| 仅计高考真题/选考卷的确证记录 | {real} |

口径说明：本表按「逐题记录」统计，卷型列区分高考真题、选考、模考/调研、合格考。
各卷原始速查表混用三种置信度记法——纯星级、简写「★5」、文字等级「高/中高」；
文字等级**不折算为星级**，以免制造原表不存在的判定，故单列「未分级」。
不同资料里出现过的 35 / 97 / 100 / 101 等数字，分别来自不同时点与不同卷型口径，
**本表只采用上表口径**，请勿与其他文档的数字直接相加或互换。

## 字段说明

| 字段 | 含义 |
|---|---|
| 卷型 | 高考真题 / 选考（真题）/ 模考·调研 / 合格考。模考类记录仅供教研参考 |
| 题号 | 原卷题号；以「/」分隔的表示同卷多题共用同一论文 |
| 考查情境（概括） | 题目素材的简述，非题干原文 |
| 溯源论文标题 | 原始研究论文标题；部分中文期刊论文无英文标题 |
| 期刊/年份 | 原表记载的期刊与发表年份 |
| DOI | 论文数字对象标识符，可拼 `https://doi.org/<DOI>` 访问 |
| DOI来源 | 见下节，用于判断可信度 |
| 关联度 | 星级（★★★★★ 形式）；若原表用文字等级则原样保留「高 / 中高 / 待核」 |
| 关联度记法 | 「星级」/「文字等级」/「未标注」，用于区分上列数值的来源 |
| 确证 | 「是」= 星级 ≥ ★★★；「否」= 低于三星或明确标待核；「未分级」= 原表为文字等级，不折算 |
| 备注 | 原表置信度列中的附加说明（已清洗） |

## DOI 来源分级

本表**不生成、不推测 DOI**。所有 DOI 均有可追溯来源：

| 来源 | 条数 |
|---|---:|
{src_txt}

- **原表记载**：各卷溯源总表速查表中已写明的 DOI，可信度最高。
- **项目内交叉检索**：原表 DOI 列为空，但同一论文标题在项目其他文档中与某 DOI 共现，且候选唯一。
- **CrossRef 标题核验**：按论文标题查询 CrossRef，仅当标题规范化相似度 ≥ 0.88 时采纳。{"（本次已启用）" if online_used else "（本次未启用）"}
- **待核**：尚无可靠来源，DOI 留空。请自行按论文标题检索，勿据空值推测。

> 注意：DOI 存在不等于已逐篇访问出版商页面核对全文。本表仅保证标识符来源可追溯，
> 不对每篇论文的正文内容作逐条复核。

## 关联度判定标准

| 星级 | 判定 |
|---|---|
| ★★★★★ | 研究对象、关键分子/基因名、实验设计逻辑、结论方向四项全部对应 |
| ★★★★☆ | 四项中三项对应 |
| ★★★☆☆ | 两项对应，或仅机制骨架一致 |
| ★★☆☆☆ 及以下 | 仅主题相似，列为待核，「确证」记为「否」 |

原表中的「高 / 中高」等文字等级无对应星级换算规则，本表保留原文并记为「未分级」，
使用者不应自行按字面把它等同于五星。

## 使用与免责

1. 本表是**个人教研判断**，不是教育考试机构或命题组的官方结论，不代表命题来源已被官方确认。
2. 标注「否」及「待核」的记录不得作为定论引用。
3. 试卷权属归相应考试机构；论文版权归各出版商与作者。本表仅登记事实性书目信息与对照判断，
   不含受版权保护的正文，使用者如需原文请自行通过合法渠道获取。
4. 引用请标注：新高考生物试题与顶刊论文相关性研究《高考生物真题×顶刊论文对照表 {VERSION}》，并遵守 CC BY-NC-SA 4.0（署名、非商业性使用、相同方式共享）。

## 数据缺口：目前还缺哪些试卷

本表是**逐题溯源记录**，不等于已覆盖全部试卷。截至 2026-11-11，缺口如下，欢迎提供可公开访问的权威来源线索；**不接收试题原文、答案、解析、题目图片或论文全文**。

### 1. 整卷未到手（2026 年 5 个卷别）

北京、天津、上海、福建、海南。

2026 年生物按 25 个卷别计（浙江一年两考计 2 卷），目前到手 20 个；上述 5 卷尚未获得可靠来源，因此表中没有它们的记录。相关公开题库暂未收录，会定期复查。

### 2. 只有部分题目（2026 年 7 个卷别）

浙江 1 月选考、浙江 6 月选考、甘肃、新课标、重庆、广西、江苏。

来源站点仅部分收录，每卷 6–13 题不等，**不是完整原卷**；且题号按收录顺序编号，**不等于原卷真实题号**。这些记录只能作为线索，不能当完整卷面引用。

### 3. 已入库卷的内部题号缺口

| 卷别 | 缺失题号 |
|---|---|
| 2025 年新课标卷 | 12–21 |
| 2025 年重庆卷 | 1–15 |

### 4. 其他限制

- 上述「部分收录」的卷别多只有题干，暂无答案与解析。
- 表中 8 条记录尚未锁定唯一 DOI，标为「待核」（见附录 B）；40 条关联度为「未分级」（原始记录用文字等级，未折算星级）。

> 如果你手上有上述卷别的**官方或权威公开来源链接**，欢迎开 Issue 只提供链接与卷别说明。请勿在 Issue 中粘贴试题原文、答案、解析、题目图片或论文全文 PDF。

## 附录 A：联网核验补全的 DOI（逐条可复核）

以下条目原表 DOI 列为空，经 CrossRef 按论文标题检索、再按所得 DOI 反查正式标题得到。
「相似度」为反查标题与原表标题的规范化比对值；低于 0.85 或显示「未反查」的请人工复核后使用。
**本表不因联网核验成功就免除核查责任**。

| 题目 | DOI | 反查标题（CrossRef） | 相似度 |
|---|---|---|---|
{cr_txt}

## 附录 B：仍待核的条目

以下条目缺少可验证的 DOI，多为中文期刊论文、学位论文或专利资料，CrossRef 未收录或标题不足以唯一锁定。
**本表不推测、不编造 DOI**，请按论文标题自行检索。

| 题目 | 论文标题 | 期刊/年份 |
|---|---|---|
{todo_txt}

## 文件清单

| 文件 | 说明 |
|---|---|
| `{BASENAME}.xlsx` | Excel 版，首行冻结、自动筛选 |
| `{BASENAME}.csv` | UTF-8 BOM，与 xlsx 数据完全一致 |
| `build_public_table.py` | 生成脚本（零依赖、相对路径、可核查数据来源） |
| `LICENSE.txt` | 许可声明 |
| `数据说明.md` | 本文件 |
"""
    path.write_text(text, encoding="utf-8")


def main():
    online = "--online" in sys.argv
    print("=== 1. 抽取各卷速查表 ===")
    records = collect_records()
    print("  合计记录:", len(records))

    print("=== 2. 建立项目内 DOI 索引并交叉补全 ===")
    index = build_doi_index()
    print("  索引条目:", len(index))
    n = fill_doi_internal(records, index)
    print("  交叉补全:", n)

    if online:
        print("=== 3. CrossRef 标题核验 ===")
        n = fill_doi_online(records)
        print("  联网补全:", n)

    print("=== 4. 敏感信息校验 ===")
    problems = validate(records)
    if problems:
        for p in problems[:20]:
            print("  [拦截]", p)
        raise SystemExit("存在敏感内容或非法 DOI，已中止，未写出任何交付文件。")
    print("  通过：无本地路径 / 身份信息 / 内部通道 / 非法 DOI")

    rows = []
    for i, r in enumerate(records, 1):
        r["序号"] = str(i)
        rows.append([r.get(h, "") for h in HEADERS])

    xlsx = OUT / (BASENAME + ".xlsx")
    csvp = OUT / (BASENAME + ".csv")
    readme = OUT / "数据说明.md"
    write_xlsx(xlsx, HEADERS, rows)
    write_csv(csvp, HEADERS, rows)
    write_readme(readme, records, online)

    print("=== 5. 输出 ===")
    for p in (xlsx, csvp, readme):
        print(f"  {p.name}  {p.stat().st_size:,} bytes")
    print("  确证(星级≥3):", sum(1 for r in records if r["确证"] == "是"))
    print("  未分级(文字等级):", sum(1 for r in records if r["确证"] == "未分级"))
    print("  否      :", sum(1 for r in records if r["确证"] == "否"))
    print("  含 DOI  :", sum(1 for r in records if r["DOI"]))
    for k in ("原表记载", "项目内交叉检索", "CrossRef 标题核验"):
        print(f"    {k}:", sum(1 for r in records if r["DOI来源"] == k))
    print("  待核    :", sum(1 for r in records if not r["DOI"]))


if __name__ == "__main__":
    main()
