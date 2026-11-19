"""백마대리점 재고 엑셀 3종 → data.json 변환기
사용법: python3 convert.py <즉시출고.xlsx> <한정재고.xlsx> <선배정.xlsx> <출력폴더>
"""
import sys, json, re, datetime, collections, openpyxl

def model_of(sheet):
    return re.sub(r"\s*\(\d+\)\s*$", "", sheet).strip()

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def strip_prefix(name, model):
    for p in (model, "봉고" + model, "봉고1톤", "봉고1.2톤"):
        if name.startswith(p + " "):
            return name[len(p) + 1:]
    return name

def grouped(rows, key_fields):
    g = collections.OrderedDict()
    for r in rows:
        k = tuple(r[f] for f in key_fields)
        if k in g:
            g[k]["q"] += 1
            if r.get("d"):
                ds = g[k].setdefault("ds", [])
                ds.append(r["d"])
        else:
            nr = dict(r); nr["q"] = 1
            if r.get("d"): nr["ds"] = [r["d"]]
            g[k] = nr
    out = []
    for v in g.values():
        if "ds" in v:
            ds = sorted(v.pop("ds")); v["d"] = ds[0]; v["d2"] = ds[-1]
        out.append(v)
    return out

def read_now(path):
    wb = openpyxl.load_workbook(path, data_only=True); rows = []
    for ws in wb:
        m = model_of(ws.title)
        for r in ws.iter_rows(values_only=True):
            if not r or not isinstance(r[4], (int, float)): continue
            rows.append(dict(m=m, t=strip_prefix(clean(r[0]), m), o=clean(r[1]) or "무옵션",
                             e=clean(r[2]), i=clean(r[3]), p=int(r[4])))
    return grouped(rows, ["m", "t", "o", "e", "i", "p"])

def read_sale(path):
    wb = openpyxl.load_workbook(path, data_only=True); rows = []
    for ws in wb:
        m = model_of(ws.title)
        for r in ws.iter_rows(values_only=True):
            if not r or not isinstance(r[4], (int, float)): continue
            d = r[5].strftime("%Y-%m-%d") if isinstance(r[5], (datetime.date, datetime.datetime)) else ""
            rows.append(dict(m=m, t=strip_prefix(clean(r[0]), m), o=clean(r[1]) or "무옵션",
                             e=clean(r[2]), i=clean(r[3]), p=int(r[4]),
                             s=int(r[6] or 0) if isinstance(r[6], (int, float)) else 0, d=d))
    return grouped(rows, ["m", "t", "o", "e", "i", "p", "s"])

def read_pre(path):
    wb = openpyxl.load_workbook(path, data_only=True); ws = wb.worksheets[0]; out = []
    for n, r in enumerate(ws.iter_rows(values_only=True)):
        if n == 0 or not r or not isinstance(r[7], (int, float)): continue
        name = clean(r[3]); m = name.split(" ")[0]
        out.append(dict(m=m, t=strip_prefix(name, m), o=clean(r[4]) or "무옵션", e=clean(r[5]),
                        i=clean(r[6]), p=int(r[7]), st=clean(r[0]), no=clean(r[1]), c=clean(r[2]), q=1))
    return out

if __name__ == "__main__":
    now, sale, pre, outdir = sys.argv[1:5]
    kst = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
    data = dict(updated=kst.strftime("%Y-%m-%d %H:%M"),
                now=read_now(now), sale=read_sale(sale), pre=read_pre(pre))
    with open(f"{outdir}/data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    for k in ("now", "sale", "pre"):
        print(k, "그룹", len(data[k]), "대수", sum(x["q"] for x in data[k]))
