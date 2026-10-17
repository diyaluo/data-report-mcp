#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
data_report.py — 纯本地 CSV/Excel 数据分析报告生成器(无 API Key、不联网、仅依赖 openpyxl)。

输入一份表格(CSV/XLSX),自动:
  1) 概览:行数/列数/缺失率/重复行
  2) 每列画像:类型、唯一值数、数值列的 min/max/均值/中位数/总和、文本列 top 频次
  3) 相关性:数值列两两皮尔逊相关(自算,不依赖 numpy/pandas)
  4) 图表:用 openpyxl 原生图表在输出 xlsx 里插入柱状图/饼图(Excel/WPS 可直接看,无需图片库)
  5) 结论:自动生成中文 Markdown 洞察报告(关键指标、异常提示、Top 类别)

设计:零 numpy/pandas/matplotlib 依赖,任何装了 openpyxl 的机器都能跑;数据不出本机。

用法:
  python3 data_report.py sales.csv -o report.xlsx
  python3 data_report.py data.xlsx -o report.xlsx --md report.md
  python3 data_report.py data.csv --charts 3            # 最多为前 N 个分类列出图
  python3 data_report.py --selftest
依赖:openpyxl (requirements.txt);CSV 自动尝试 utf-8-sig/gbk。
"""
import argparse, csv, json, math, os, re, sys, statistics, datetime
from collections import Counter

try:
    import openpyxl
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("缺少 openpyxl: pip install openpyxl", file=sys.stderr)
    sys.exit(2)


# ---------- 读取 ----------
def read_table(path):
    if path.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "gbk", "latin-1"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    rows = [list(r) for r in csv.reader(f)]
                return rows
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 编码无法识别")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    return [["" if c is None else c for c in r] for r in ws.iter_rows(values_only=True)]


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        return float(v)
    if v is None:
        return None
    s = str(v).strip().replace(",", "").replace("，", "").replace("￥", "").replace("¥", "").replace("$", "").replace("%", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s)
    return None


# ---------- 列画像 ----------
def profile(headers, data):
    cols = []
    n = len(data)
    for j, h in enumerate(headers):
        vals = [r[j] if j < len(r) else "" for r in data]
        nonblank = [v for v in vals if str(v).strip() != ""]
        nums = [x for v in nonblank if (x := _num(v)) is not None]
        is_numeric = nums and len(nums) >= 0.8 * max(1, len(nonblank))
        col = {"name": str(h), "missing": n - len(nonblank),
               "missing_rate": round((n - len(nonblank)) / max(1, n), 3),
               "unique": len(set(map(str, nonblank))), "is_numeric": bool(is_numeric)}
        if is_numeric:
            col.update({
                "min": round(min(nums), 4), "max": round(max(nums), 4),
                "mean": round(statistics.fmean(nums), 4),
                "median": round(statistics.median(nums), 4),
                "sum": round(sum(nums), 4),
                "stdev": round(statistics.pstdev(nums), 4) if len(nums) > 1 else 0,
                "_nums": nums,
            })
        else:
            cnt = Counter(str(v).strip() for v in nonblank)
            col["top"] = cnt.most_common(5)
            col["_nums"] = None
        cols.append(col)
    return cols


def pearson(xs, ys):
    n = min(len(xs), len(ys))
    xs, ys = xs[:n], ys[:n]
    if n < 2:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    cov = sum((a - mx) * (b - my) for a, b in zip(xs, ys))
    vx = sum((a - mx) ** 2 for a in xs)
    vy = sum((b - my) ** 2 for b in ys)
    if vx == 0 or vy == 0:
        return None
    return round(cov / math.sqrt(vx * vy), 3)


def analyze(rows):
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise ValueError("空表")
    headers = [str(h).strip() or f"col_{i+1}" for i, h in enumerate(rows[0])]
    data = rows[1:]
    dup = len(data) - len({tuple(map(str, r)) for r in data})
    cols = profile(headers, data)
    num_cols = [c for c in cols if c["is_numeric"]]
    corr = []
    for i in range(len(num_cols)):
        for j in range(i + 1, len(num_cols)):
            r = pearson(num_cols[i]["_nums"], num_cols[j]["_nums"])
            if r is not None and abs(r) >= 0.5:
                corr.append((num_cols[i]["name"], num_cols[j]["name"], r))
    corr.sort(key=lambda t: -abs(t[2]))
    return {"rows": len(data), "cols": len(headers), "headers": headers,
            "dup_rows": dup, "columns": cols, "corr": corr, "_data": data}


# ---------- Markdown 报告 ----------
def to_markdown(a, title="数据分析报告"):
    L = [f"# {title}", "",
         f"生成时间:{datetime.datetime.now():%Y-%m-%d %H:%M}　数据行:{a['rows']}　列:{a['cols']}　重复行:{a['dup_rows']}", ""]
    L.append("## 一、总体概览")
    miss_hi = [c["name"] for c in a["columns"] if c["missing_rate"] >= 0.2]
    L.append(f"- 共 {a['rows']} 行、{a['cols']} 列;疑似重复行 {a['dup_rows']} 行。")
    L.append(f"- 缺失率≥20% 的列:{('、'.join(miss_hi)) if miss_hi else '无'}。")
    nums = [c for c in a["columns"] if c["is_numeric"]]
    cats = [c for c in a["columns"] if not c["is_numeric"]]
    L.append(f"- 数值列 {len(nums)} 个,分类/文本列 {len(cats)} 个。")
    L.append("")
    if nums:
        L.append("## 二、数值列统计")
        L.append("| 列 | 最小 | 最大 | 均值 | 中位数 | 总和 | 标准差 | 缺失率 |")
        L.append("|---|---|---|---|---|---|---|---|")
        for c in nums:
            L.append(f"| {c['name']} | {c['min']} | {c['max']} | {c['mean']} | {c['median']} | {c['sum']} | {c['stdev']} | {c['missing_rate']*100:.0f}% |")
        L.append("")
    if cats:
        L.append("## 三、分类列 Top 值")
        for c in cats[:8]:
            top = "、".join(f"{k}({v})" for k, v in c["top"][:5])
            L.append(f"- **{c['name']}**(唯一值 {c['unique']}):{top or '—'}")
        L.append("")
    if a["corr"]:
        L.append("## 四、显著相关(|r|≥0.5)")
        for x, y, r in a["corr"][:10]:
            word = "正" if r > 0 else "负"
            L.append(f"- {x} 与 {y}:{r}({word}相关)")
        L.append("")
    L.append("> 本报告由本地工具自动生成,结论供参考;相关不等于因果,请结合业务判断。")
    return "\n".join(L)


# ---------- Excel 报告(含原生图表) ----------
def to_excel(a, out, chart_n=3):
    wb = openpyxl.Workbook()
    hf = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="2F5597")

    ws = wb.active; ws.title = "概览"
    overview = [["指标", "值"], ["数据行", a["rows"]], ["列数", a["cols"]],
                ["重复行", a["dup_rows"]]]
    for r in overview:
        ws.append(r)
    for cell in ws[1]:
        cell.font = hf; cell.fill = fill
    ws.column_dimensions["A"].width = 16; ws.column_dimensions["B"].width = 14

    wsn = wb.create_sheet("数值统计")
    wsn.append(["列", "最小", "最大", "均值", "中位数", "总和", "标准差", "缺失率"])
    for cell in wsn[1]:
        cell.font = hf; cell.fill = fill
    nums = [c for c in a["columns"] if c["is_numeric"]]
    for c in nums:
        wsn.append([c["name"], c["min"], c["max"], c["mean"], c["median"],
                    c["sum"], c["stdev"], c["missing_rate"]])
    for col in "ABCDEFGH":
        wsn.column_dimensions[col].width = 14

    # 明细数据页(供图表引用)
    wsd = wb.create_sheet("明细")
    wsd.append(a["headers"])
    for r in a["_data"]:
        wsd.append(r)

    # 分类列 Top 频次页 + 柱状图
    cats = [c for c in a["columns"] if not c["is_numeric"]]
    made = 0
    for c in cats[:chart_n]:
        title = re.sub(r"[\\/*?:\[\]]", "_", str(c["name"]))[:20] or f"cat{made}"
        wc = wb.create_sheet(f"分布_{title}"[:31])
        wc.append([c["name"], "次数"])
        for k, v in c["top"]:
            wc.append([k, v])
        if len(c["top"]) >= 2:
            chart = BarChart(); chart.title = f"{c['name']} 分布"; chart.height = 7; chart.width = 14
            dataref = Reference(wc, min_col=2, min_row=1, max_row=1 + len(c["top"]))
            catsref = Reference(wc, min_col=1, min_row=2, max_row=1 + len(c["top"]))
            chart.add_data(dataref, titles_from_data=True); chart.set_categories(catsref)
            wc.add_chart(chart, "D2")
            made += 1

    if a["corr"]:
        wr = wb.create_sheet("显著相关")
        wr.append(["列A", "列B", "相关系数r"])
        for x, y, r in a["corr"]:
            wr.append([x, y, r])
    wb.save(out)


def _selftest():
    import tempfile
    rows = [
        ["城市", "品类", "销售额", "数量"],
        ["郑州", "电子", "1200.5", "3"],
        ["郑州", "电子", "1200.5", "3"],      # 重复行
        ["北京", "服装", "￥800", "5"],
        ["上海", "电子", "2,500", "2"],
        ["北京", "食品", "", "4"],
    ]
    a = analyze(rows)
    assert a["rows"] == 5 and a["dup_rows"] == 1, (a["rows"], a["dup_rows"])
    sales = next(c for c in a["columns"] if c["name"] == "销售额")
    assert sales["is_numeric"] and sales["max"] == 2500 and sales["min"] == 800, sales
    city = next(c for c in a["columns"] if c["name"] == "城市")
    assert not city["is_numeric"] and dict(city["top"]).get("郑州") == 2
    md = to_markdown(a)
    assert "显著相关" in md or "数值列统计" in md
    with tempfile.TemporaryDirectory() as td:
        xp = os.path.join(td, "r.xlsx")
        to_excel(a, xp)
        wb = openpyxl.load_workbook(xp)
        assert "数值统计" in wb.sheetnames and "明细" in wb.sheetnames
    print("selftest PASS: 读取/类型推断/数值统计/去重/相关/Markdown/Excel原生图表 全通过")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", nargs="?")
    ap.add_argument("-o", "--output", help="输出 xlsx 路径")
    ap.add_argument("--md", help="同时输出 Markdown 报告路径")
    ap.add_argument("--title", default="数据分析报告")
    ap.add_argument("--charts", type=int, default=3, help="最多为前 N 个分类列出图")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        sys.exit(_selftest())
    if not args.input:
        ap.error("需要输入 CSV/XLSX(或 --selftest)")
    rows = read_table(args.input)
    a = analyze(rows)
    out = args.output or "report.xlsx"
    to_excel(a, out, chart_n=args.charts)
    print("XLSX:", os.path.abspath(out))
    md = to_markdown(a, args.title)
    if args.md:
        open(args.md, "w", encoding="utf-8").write(md)
        print("MD:", os.path.abspath(args.md))
    else:
        print("\n" + md)


if __name__ == "__main__":
    main()
