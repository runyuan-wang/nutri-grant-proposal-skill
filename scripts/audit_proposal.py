#!/usr/bin/env python3
"""申报书提交前自检：金额配平 + 字体审计 + 错别字/口径雷区。

用法: python3 audit_proposal.py <申报书.docx>
退出码: 0 = 全部通过/仅提示, 1 = 存在 ❌ 级问题。
"""
import re
import sys
from collections import Counter

from docx import Document
from docx.oxml.ns import qn


def all_paragraph_text(doc):
    """正文 + 所有表格（含嵌套）的纯文本拼接。"""
    parts = []
    def walk_tbl(tbl):
        for tr in tbl.findall(qn("w:tr")):
            for tc in tr.findall(qn("w:tc")):
                for ch in tc:
                    if ch.tag == qn("w:p"):
                        parts.append("".join(t.text or "" for t in ch.iter(qn("w:t"))))
                    elif ch.tag == qn("w:tbl"):
                        walk_tbl(ch)
    for el in doc.element.body:
        if el.tag == qn("w:p"):
            parts.append("".join(t.text or "" for t in el.iter(qn("w:t"))))
        elif el.tag == qn("w:tbl"):
            walk_tbl(el)
    return parts


def unique_cell_rows(doc):
    """每张表按行给出单元格文本（合并单元格按 xml 去重）。"""
    rows = []
    def walk_tbl(tbl):
        for tr in tbl.findall(qn("w:tr")):
            cells = []
            for tc in tr.findall(qn("w:tc")):
                txt = "\n".join(
                    "".join(t.text or "" for t in p.iter(qn("w:t")))
                    for p in tc.findall(qn("w:p"))
                ).strip()
                cells.append(txt)
                for nt in tc.findall(qn("w:tbl")):
                    walk_tbl(nt)
            rows.append(cells)
    for el in doc.element.body:
        if el.tag == qn("w:tbl"):
            walk_tbl(el)
    return rows


def num(s):
    m = re.search(r"(-?\d+(?:\.\d+)?)", s.replace(",", ""))
    return float(m.group(1)) if m else None


def check_amounts(rows):
    issues, found = [], {}
    for cells in rows:
        joined = " ".join(cells)
        for label in ("一、经费支出", "二、经费来源", "经费合计", "申请资金"):
            if label in joined:
                vals = [num(c) for c in cells[1:] if num(c) is not None]
                if vals and label not in found:
                    found[label] = vals[0]
        for c in cells:
            m = re.search(r"申请金额（万元）[:：]\s*(\d+\.\d+)", c)
            if m:
                found["封面申请金额"] = float(m.group(1))
            m = re.search(r"申请金额", c)
            if m and "万元" in joined and "封面申请金额" not in found:
                v = re.search(r"(\d+\.\d+)\s*万元", joined)
                if v:
                    found["信息表申请金额"] = float(v.group(1))
    # 分项（顶级科目 1.~11.）
    item_sum = 0.0
    item_seen = False
    for cells in rows:
        if len(cells) >= 3 and re.match(r"^\d+\.\s*(设备费|材料费|测试化验加工费|燃料动力费|差旅费|会议费|国际合作与交流费|出版|劳务费|专家咨询费|其他支出)", cells[1] if len(cells) > 1 else ""):
            v = num(cells[2])
            if v is not None:
                item_sum += v
                item_seen = True
    vals = list(found.values())
    if vals:
        base = vals[0]
        for k, v in found.items():
            if abs(v - base) > 0.005:
                issues.append(f"❌ 金额不一致：{k}={v}，封面/主口径={base}")
        if item_seen and abs(item_sum - base) > 0.005:
            issues.append(f"❌ 分项加总 {item_sum:.2f} ≠ 经费支出/总额 {base:.2f}")
        if not issues:
            issues.append(f"✅ 金额一致（{base:.2f} 万元，分项加总配平）")
    else:
        issues.append("⚠️ 未找到经费预算表，跳过金额校验")
    return issues


def check_fonts(doc):
    fonts, sizes = Counter(), Counter()
    for r in doc.element.body.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        ea = asc = sz = None
        if rPr is not None:
            f = rPr.find(qn("w:rFonts"))
            if f is not None:
                ea, asc = f.get(qn("w:eastAsia")), f.get(qn("w:ascii"))
            s = rPr.find(qn("w:sz"))
            if s is not None:
                sz = s.get(qn("w:val"))
        fonts[(ea, asc)] += 1
        sizes[sz] += 1
    lines = [f"ℹ️ 字体分布: {dict(fonts)}", f"ℹ️ 字号分布(half-points): {dict(sizes)}"]
    bad = {k: v for k, v in fonts.items() if k[0] not in (None, "宋体", "黑体")}
    if bad:
        lines.append(f"⚠️ 异常中文字体: {bad}")
    else:
        lines.append("✅ 字体收敛（宋体/黑体/样式继承）")
    return lines


def check_typos(paras):
    text = "\n".join(paras)
    issues = []
    for pat, msg in [
        (r"FQQ", "FFQ 误写为 FQQ"),
        (r"多民族t门诊", "误植字母“多民族t门诊”"),
        (r"\d\.\d万元", "金额未写两位小数（如 0.6万元）"),
        (r"视网膜患者", "疑漏字（应为“视网膜病变患者”）"),
        (r"检索检索|核查核查|分析分析|调查调查", "叠词"),
        (r"逐篇|逐条", "暴露流程用词（建议改“核查/核对/复筛”）"),
    ]:
        for m in re.finditer(pat, text):
            ctx = text[max(0, m.start() - 15):m.end() + 15].replace("\n", " ")
            issues.append(f"⚠️ {msg}: …{ctx}…")
    dup = re.findall(r"([一-龥]{2})\1", text)
    for w in set(dup):
        if w not in ("研究研究",):
            issues.append(f"⚠️ 疑似叠词“{w}{w}”")
    # 阶段数口径
    n2, n3 = len(re.findall(r"两阶段", text)), len(re.findall(r"三阶段|第三阶段", text))
    if n2 and n3:
        issues.append(f"❌ 阶段数口径冲突：两阶段×{n2} vs 三阶段×{n3}")
    if not issues:
        issues.append("✅ 未见已知错别字/口径雷区")
    return issues


def main():
    if len(sys.argv) != 2:
        sys.exit("用法: python3 audit_proposal.py <申报书.docx>")
    doc = Document(sys.argv[1])
    rows = unique_cell_rows(doc)
    paras = all_paragraph_text(doc)
    amounts = check_amounts(rows)
    fonts = check_fonts(doc)
    typos = check_typos(paras)
    print("== 金额配平 ==")
    print("\n".join(amounts))
    print("== 字体审计 ==")
    print("\n".join(fonts))
    print("== 错别字与口径 ==")
    print("\n".join(typos))
    report = amounts + fonts + typos
    sys.exit(1 if any(l.startswith("❌") for l in report) else 0)


if __name__ == "__main__":
    main()
