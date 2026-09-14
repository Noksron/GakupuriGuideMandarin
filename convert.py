#!/usr/bin/env python3
"""把攻略 Excel 转成网页用的 data.js。

用法:
    python3 convert.py *.xlsx             # 多个学院的表一起转
    python3 convert.py 青学.xlsx -o out.js

Excel 约定（每个角色一个工作表）:
    B2                角色名
    C3                特殊话题        J3  嫉妒程度
    A 列 "*"          该行有差分
    B~H (第6行起)     日期/时间/类别/场所/选择/备注/获取
    J8:V11            好感度区间 × 话题加成表
    J13 起的小表      可选的备用路线（如海堂的「制服cg回收」）
"""
import json
import os
import re
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

TIME_ALIAS = {"朝": "晨"}          # 统一写法

# 文件名里的学院关键字 -> 显示名与侧栏徽记
SCHOOLS = [
    ("青学",     "青春学园",   "青"),
    ("氷帝",     "冰帝学园",   "冰"),
    ("冰帝",     "冰帝学园",   "冰"),
    ("ルドルフ", "圣鲁道夫学院", "鲁"),
    ("圣鲁道夫", "圣鲁道夫学院", "鲁"),
    ("不动峰",   "不动峰中学", "峰"),
    ("山吹",     "山吹中学",   "吹"),
    ("立海",     "立海大附属", "立"),
]

# 折叠侧栏里显示的缩写，默认取姓氏首字，同形的在这里指定
INITIALS = {
    "不二裕太": "裕太",
}


def school_of(path):
    for key, name, crest in SCHOOLS:
        if key in path:
            return name, crest
    return os.path.splitext(os.path.basename(path))[0], "校"


def norm_date(v):
    """46256 -> ('8/22', False)；'（8/25）' -> ('8/25', True) 非固定日期。"""
    if isinstance(v, (int, float)):
        d = from_excel(v)
        return f"{d.month}/{d.day}", False
    s = str(v).strip().strip("（）()")
    return s, True


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\u3000", " ").strip()
    return "" if s == "-" else s


def norm_note(s):
    """28号私服cg / 28日私服cg -> 28日私服CG"""
    s = re.sub(r"(\d+)号", r"\1日", s)
    return re.sub(r"cg", "CG", s, flags=re.I)


def parse_sheet(ws):
    rows, affinity, extra = [], None, None
    aff_head, extra_head, extra_rows = [], "", []

    for r in ws.iter_rows(min_row=6, max_row=ws.max_row):
        cell = {get_column_letter(c.column): c.value for c in r}

        # ---- 右侧：好感度加成表 ----
        j = cell.get("J")
        if j == "好感度":
            aff_head = [clean(cell.get(col)) for col in "KLMNOPQRSTUV"]
        elif isinstance(j, str) and re.match(r"^\d+\s*[—–~～\-]\s*\d+$", j.strip()):
            affinity = affinity or []
            affinity.append({
                "range": re.sub(r"\s*[—–~～\-]\s*", "–", j.strip()),
                "values": [cell.get(col) for col in "KLMNOPQRSTUV"],
            })
        elif isinstance(j, str) and "回收" in j:
            extra_head = norm_note(j.strip())
        elif isinstance(j, (int, float)) and extra_head:
            extra_rows.append({
                "date": f"8/{int(j)}" if int(j) >= 22 else f"9/{int(j)}",
                "time": TIME_ALIAS.get(clean(cell.get("K")), clean(cell.get("K"))),
                "kind": clean(cell.get("L")),
                "place": clean(cell.get("M")),
                "choice": clean(cell.get("N")) or clean(cell.get("O")),
            })

        # ---- 主表 ----
        if cell.get("B") in (None, ""):
            continue
        date, loose = norm_date(cell["B"])
        time = clean(cell.get("C"))
        note = norm_note(clean(cell.get("G")))
        gain = norm_note(clean(cell.get("H")))
        rows.append({
            "date": date,
            "loose": loose,                         # 非固定日期
            "time": TIME_ALIAS.get(time, time) or "-",
            "kind": clean(cell.get("D")) or "自动",
            "place": clean(cell.get("E")),
            "choice": clean(cell.get("F")),
            "note": note,
            "gain": gain,
            "branch": str(cell.get("A") or "").strip() == "*",
            "cg": "CG" in gain or "CG" in note,
            "game": "小游戏" in note,
        })

    if extra_rows:
        extra = {"title": extra_head, "rows": extra_rows}
    return rows, {"head": aff_head, "rows": affinity or []}, extra


def split_topics(cell):
    """「A +6 / B +7 丨 额外提示」 -> (["A +6", "B +7"], "额外提示")"""
    parts = re.split(r"[丨|｜]", clean(cell), maxsplit=1)
    topics = [t.strip() for t in re.split(r"[/／]", parts[0]) if t.strip()]
    return topics, (parts[1].strip() if len(parts) > 1 else "")


def main():
    args = [a for a in sys.argv[1:] if a != "-o"]
    out = "data.js"
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o") + 1]
        args = [a for a in args if a != out]

    schools = []
    for src in args:
        sname, crest = school_of(src)
        wb = load_workbook(src, data_only=True)
        chars = []
        for name in wb.sheetnames:
            ws = wb[name]
            rows, affinity, extra = parse_sheet(ws)
            topics, tip = split_topics(ws["C3"].value)
            chars.append({
                "name": (ws["B2"].value or name).strip(),
                "ini": INITIALS.get((ws["B2"].value or name).strip(), ""),
                "id": f"{sname}·{name}",
                "topics": topics,
                "tip": tip,
                "jealousy": clean(ws["J3"].value).replace("嫉妒：", ""),
                "steps": rows,
                "affinity": affinity,
                "extra": extra,
            })
        schools.append({"name": sname, "crest": crest, "chars": chars})

    payload = {"schools": schools}
    with open(out, "w", encoding="utf-8") as f:
        f.write("const DATA = " + json.dumps(payload, ensure_ascii=False, indent=1) + ";\n")

    total = sum(len(c["steps"]) for s_ in schools for c in s_["chars"])
    n_chars = sum(len(s_["chars"]) for s_ in schools)
    print(f"✓ {out}：{len(schools)} 个学院，{n_chars} 个角色，{total} 条行动")
    for s_ in schools:
        for c in s_["chars"]:
            a, b = (c["steps"][0]["date"], c["steps"][-1]["date"]) if c["steps"] else ("-", "-")
            print(f"  {s_['name']} {c['name']}：{a} – {b}（{len(c['steps'])} 条）")


if __name__ == "__main__":
    main()
