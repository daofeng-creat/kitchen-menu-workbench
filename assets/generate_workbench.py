#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_workbench.py
=====================
把厨师的「历史菜单 .docx」解析成按品类分类的菜品库，并填充进单文件 HTML 工作台模板，
生成一个手机微信可直接打开的「每周配菜工作台」。

用法：
    python3 generate_workbench.py [菜单.docx] [-o 输出.html] [--title 标题] [--author 作者]

若不传参数，默认读取当前目录下的「菜单.docx」，输出「厨房配菜工作台.html」。

菜品位置规则（与厨师历史菜单一致）：
    ①主荤 ②主荤 ③花荤 ④花荤 ⑤纯素菜 ⑥汤 ⑦固定「水果」
配菜硬约束：
    1. 严格按 主荤×2 / 花荤×2 / 素菜×1 / 汤×1 / 水果×1 出菜，位置不错位。
    2. 每周水产（鱼/虾/贝等）次数 2~3 次；工作日不足时取最大可能。
    3. 本周主荤自动避开「上周主荤」（跨周不重复），并回存供下周避开。
    4. 同一周内主荤尽量不重复。
"""

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET


# ---------------------------------------------------------------------------
# 1) 解析 docx
# ---------------------------------------------------------------------------
def read_docx_text(path):
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(xml)
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    paras = []
    for p in root.iter(ns + "p"):
        runs = [t.text or "" for t in p.iter(ns + "t")]
        paras.append("".join(runs))
    return [l.strip() for l in paras if l.strip()]


def parse_menu(lines):
    day_re = re.compile(r"^(星期[一二三四五六日天])$")
    date_re = re.compile(r"^\d{4}年")

    zhuhun, huahun, suchai, tang = [], [], [], []
    i = 0
    while i < len(lines):
        l = lines[i]
        if date_re.match(l):
            i += 1
            continue
        if day_re.match(l):
            dish_line = lines[i + 1] if i + 1 < len(lines) else ""
            if dish_line and not day_re.match(dish_line) and not date_re.match(dish_line):
                items = re.split(r"[，,\s]+", dish_line.rstrip("。"))
                items = [x for x in items if x]
                # 清洗历史错位项
                if "韭菜豆芽西红柿蛋汤" in items:
                    items = [x for x in items if x != "韭菜豆芽西红柿蛋汤"] + ["韭菜豆芽", "西红柿蛋汤"]
                if "紫菜蛋汤蛋汤" in items:
                    items = [x for x in items if x != "紫菜蛋汤蛋汤"] + ["紫菜蛋汤"]
                for k, it in enumerate(items[:7]):
                    if k in (0, 1):
                        if it not in zhuhun:
                            zhuhun.append(it)
                    elif k in (2, 3):
                        if it not in huahun:
                            huahun.append(it)
                    elif k == 4:
                        if it not in suchai:
                            suchai.append(it)
                    elif k == 5:
                        if it not in tang:
                            tang.append(it)
                i += 2
                continue
        i += 1
    return zhuhun, huahun, suchai, tang


# ---------------------------------------------------------------------------
# 2) 水产识别
# ---------------------------------------------------------------------------
FISH_KW = ["鱼", "鲈", "鲶", "虾", "带鱼", "泥鳅", "黄丫头", "黄牙头",
           "鳝", "花甲", "圣子", "墨鱼", "鱿"]
EXCLUDE_FISH = {"鱼香肉丝"}


def is_fish(name):
    if name in EXCLUDE_FISH:
        return False
    return any(k in name for k in FISH_KW)


# ---------------------------------------------------------------------------
# 3) 填充模板
# ---------------------------------------------------------------------------
def build_pools(zhuhun, huahun, suchai, tang):
    zh = [{"n": n, "f": is_fish(n)} for n in zhuhun]
    hh = [{"n": n, "f": is_fish(n)} for n in huahun]
    sc = [x for x in suchai if x != "水果"]
    tg = [x for x in tang if x not in ("水果", "韭菜豆芽")]
    return zh, hh, sc, tg


def generate(docx_path, out_path, title, author):
    lines = read_docx_text(docx_path)
    zhuhun, huahun, suchai, tang = parse_menu(lines)
    zh, hh, sc, tg = build_pools(zhuhun, huahun, suchai, tang)

    zh_names = [x["n"] for x in zh]
    zh_fish = [x["n"] for x in zh if x["f"]]
    hh_names = [x["n"] for x in hh]
    hh_fish = [x["n"] for x in hh if x["f"]]

    here = os.path.dirname(os.path.abspath(__file__))
    tpl = open(os.path.join(here, "template.html"), encoding="utf-8").read()

    tpl = tpl.replace("__ZH__", json.dumps(zh_names, ensure_ascii=False))
    tpl = tpl.replace("__ZH_FISH__", json.dumps(zh_fish, ensure_ascii=False))
    tpl = tpl.replace("__HH__", json.dumps(hh_names, ensure_ascii=False))
    tpl = tpl.replace("__HH_FISH__", json.dumps(hh_fish, ensure_ascii=False))
    tpl = tpl.replace("__SC__", json.dumps(sc, ensure_ascii=False))
    tpl = tpl.replace("__TG__", json.dumps(tg, ensure_ascii=False))

    # 标题与署名
    if title:
        tpl = tpl.replace("<title>抚州市教育体育局厨房每周配菜工作台 by：何康</title>",
                          "<title>%s</title>" % title)
        tpl = re.sub(r"<h1>.*?抚州市教育体育局厨房每周配菜工作台.*?</h1>",
                     "<h1>%s</h1>" % title, tpl, flags=re.S)
    if author:
        tpl = re.sub(r"<p class=\"sub\">by：何康 · .*?</p>",
                     "<p class=\"sub\">by：%s · 勾选本周上班日 → 一键自动配菜 · 手机微信可直接打开</p>" % author,
                     tpl, flags=re.S)

    open(out_path, "w", encoding="utf-8").write(tpl)
    print("✓ 已生成：%s" % out_path)
    print("  主荤 %d（水产 %d）· 花荤 %d（水产 %d）· 素菜 %d · 汤 %d"
          % (len(zh_names), len(zh_fish), len(hh_names), len(hh_fish), len(sc), len(tg)))


def main():
    ap = argparse.ArgumentParser(description="从菜单 docx 生成厨房每周配菜工作台 HTML")
    ap.add_argument("docx", nargs="?", default="菜单.docx", help="历史菜单 .docx 路径")
    ap.add_argument("-o", "--output", default="厨房配菜工作台.html", help="输出 HTML 路径")
    ap.add_argument("--title", default="", help="页面标题（默认保留模板标题）")
    ap.add_argument("--author", default="", help="署名（默认保留模板署名）")
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print("✗ 找不到菜单文件：%s" % args.docx, file=sys.stderr)
        sys.exit(1)

    generate(args.docx, args.output, args.title, args.author)


if __name__ == "__main__":
    main()
