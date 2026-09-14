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
    ("立海",     "立海大附属中学", "立"),
    ("六角",     "六角中学",   "六"),
    ("四天宝寺", "四天宝寺中学", "四"),
]

# 全学院名单：顺序、分组、学院色都以这里为准。
# 还没整理的角色也会出现在站上（标成「待补充」），Excel 里补了哪个就自动填哪个。
# 注意：Excel 工作表 B2 的角色名要和这里写的一致，否则会被当成新角色单列出来。
ROSTER = [
    ("青春学园", "青", "#595ccc", [
        "越前龙马", "手塚国光", "大石秀一郎", "不二周助", "菊丸英二", "河村隆", "乾贞治", "桃城武", "海堂熏"]),
    ("不动峰中学", "峰", "#646464", [
        "橘桔平", "神尾アキラ", "伊武深司"]),
    ("圣鲁道夫学院", "鲁", "#5e504b", [
        "赤澤吉朗", "観月はじめ", "不二裕太"]),
    ("山吹中学", "吹", "#5c8161", [
        "亚久津仁", "千石清纯", "坛太一"]),
    ("冰帝学园", "冰", "#587d95", [
        "迹部景吾", "忍足侑士", "向日岳人", "凤长太郎", "宍户亮", "日吉若", "芥川慈郎"]),
    ("六角中学", "六", "#804749", [
        "天根光", "佐伯虎次郎", "黑羽春风"]),
    ("立海大附属中学", "立", "#cab134", [
        "切原赤也", "真田弦一郎", "柳莲二", "丸井文太", "仁王雅治", "柳生比吕士", "幸村精市"]),
    ("四天宝寺中学", "四", "#89b16e", [
        "远山金太郎", "千岁千里", "白石藏之介"]),
]


# 表格之外的补充说明，按「学院·角色」挂，会显示在角色名下面
EXTRA_TIPS = {
    "冰帝学园·迹部景吾": ["奖励CG为「学院祭扣杀BINGO」中概率获得"],
    "冰帝学园·忍足侑士": ["奖励CG为「学院祭扣杀BINGO」中概率获得"],
}


def blank(sname, cname, out_dir):
    """还没整理的角色：只有名字和头像，正文那边会显示「待补充」。"""
    return {
        "name": cname,
        "ini": INITIALS.get(cname, ""),
        "img": asset("avatar", cname, out_dir, (".png", ".webp")),
        "portrait": asset("portrait", cname, out_dir),
        "id": f"{sname}·{cname}",
        "topics": [], "tip": "", "notes": EXTRA_TIPS.get(f"{sname}·{cname}", []),
        "jealousy": "",
        "steps": [], "affinity": {"head": [], "rows": []}, "extra": None,
        "todo": True,
    }


def asset(kind, name, out_dir, exts=(".webp", ".png")):
    """assets/<kind>/<角色名>.<后缀> 存在就返回相对路径，否则空串。"""
    for ext in exts:
        rel = f"assets/{kind}/{name}{ext}"
        if os.path.exists(os.path.join(out_dir, rel)):
            return rel
    return ""

# 折叠侧栏里显示的缩写，默认取姓氏首字，同形的在这里指定
INITIALS = {
    "不二裕太": "裕",          # 两个字塞进小方框会换行，统一用单字
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

    out_dir = os.path.dirname(os.path.abspath(out))

    # 先把所有 Excel 读进来，再按名单的顺序拼装
    parsed = {}
    for src in args:
        sname, crest = school_of(src)
        wb = load_workbook(src, data_only=True)
        for name in wb.sheetnames:
            ws = wb[name]
            rows, affinity, extra = parse_sheet(ws)
            topics, tip = split_topics(ws["C3"].value)
            cname = (ws["B2"].value or name).strip()
            parsed[(sname, cname)] = {
                "name": cname,
                "ini": INITIALS.get(cname, ""),
                "img": asset("avatar", cname, out_dir, (".png", ".webp")),  # 侧栏头像
                "portrait": asset("portrait", cname, out_dir),  # 正文立绘
                "id": f"{sname}·{name}",
                "topics": topics,
                "tip": norm_note(tip),
                "notes": EXTRA_TIPS.get(f"{sname}·{cname}", []),
                "jealousy": clean(ws["J3"].value).replace("嫉妒：", ""),
                "steps": rows,
                "affinity": affinity,
                "extra": extra,
                "todo": not rows,
            }

    schools, used = [], set()
    for sname, crest, color, names in ROSTER:
        chars = []
        for cname in names:
            got = parsed.get((sname, cname))
            if got:
                used.add((sname, cname))
            chars.append(got or blank(sname, cname, out_dir))
        schools.append({"name": sname, "crest": crest, "color": color, "chars": chars})

    # 名单上没有的（多半是 Excel 里名字写法不一样），单独挂到对应学院末尾并提醒
    for key, got in parsed.items():
        if key in used:
            continue
        sname, cname = key
        print(f"  ! 名单里没有「{sname} {cname}」，已单独加在该学院末尾")
        for s_ in schools:
            if s_["name"] == sname:
                s_["chars"].append(got); break
        else:
            schools.append({"name": sname, "crest": "校", "color": "", "chars": [got]})

    payload = {"schools": schools}
    with open(out, "w", encoding="utf-8") as f:
        f.write("const DATA = " + json.dumps(payload, ensure_ascii=False, indent=1) + ";\n")

    total = sum(len(c["steps"]) for s_ in schools for c in s_["chars"])
    n_chars = sum(len(s_["chars"]) for s_ in schools)
    n_todo = sum(1 for s_ in schools for c in s_["chars"] if c.get("todo"))
    print(f"✓ {out}：{len(schools)} 个学院，{n_chars} 个角色（{n_todo} 个待补充），{total} 条行动")
    for s_ in schools:
        for c in s_["chars"]:
            if c.get("todo"):
                print(f"  {s_['name']} {c['name']}：待补充")
            else:
                a, b = c["steps"][0]["date"], c["steps"][-1]["date"]
                print(f"  {s_['name']} {c['name']}：{a} – {b}（{len(c['steps'])} 条）")


if __name__ == "__main__":
    main()
