from __future__ import annotations

import csv
import json
import math
import os
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
from flask import Flask, jsonify, redirect, render_template_string, request

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
PRICING_DIR = Path("/root/.hermes/profiles/mengjie/workspace/pricing")
LIB_CSV = DATA_DIR / "cost_library.csv"
ORDERS_CSV = DATA_DIR / "orders_light.csv"
SALES_CSV = DATA_DIR / "sales_light.csv"
LOG_CSV = DATA_DIR / "library_changes.csv"
SUPPLIERS_CSV = DATA_DIR / "suppliers.csv"

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "").strip()
USE_SUPABASE = bool(SUPABASE_URL and SUPABASE_KEY)
SB_HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}
TABLE_MAP = {
    "cost_library.csv": "cost_library",
    "orders_light.csv": "orders_light",
    "sales_light.csv": "sales_light",
    "library_changes.csv": "library_changes",
    "suppliers.csv": "suppliers",
}

PREFERRED = [
    PRICING_DIR / "报价库_v6_20260507更新.xlsx",
    PRICING_DIR / "报价库_v6_20260507更新.bak_before_20260512_柒天茗鲜实单.xlsx",
    PRICING_DIR / "报价库_v5_含完整米油蛋.xlsx",
]
ALIASES = {"毛毛肉": "圣农琵琶腿XL", "鸡腿肉": "圣农琵琶腿XL"}
LIB_COLS = ["id", "sheet", "品类", "品名", "规格", "单位", "单价", "供应商", "启用", "来源", "更新时间"]
ORDER_COLS = ["记录ID", "日期", "批次", "食材", "输入数量", "输入单位", "计价数量", "计价单位", "单价", "小计", "状态", "报价库品名", "规格", "备注"]
SUPPLIER_COLS = ["id", "类别", "供货商", "联系人", "电话", "地址", "备注", "启用", "更新时间"]
SUPPLIER_CATEGORIES = ["蔬菜", "海鲜", "冻品", "调料", "猪肉"]


def find_pricing_file():
    for p in PREFERRED:
        if p.exists():
            return p
    xs = sorted(PRICING_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return xs[0] if xs else None


def norm(s):
    return re.sub(r"[（）()\s·/\-]", "", str(s or "").strip().lower())


def today_str():
    return date.today().isoformat()


def init_library(force=False):
    if LIB_CSV.exists() and not force:
        return
    path = find_pricing_file()
    rows = []
    if path:
        sheets = pd.read_excel(path, sheet_name=None)
        i = 1
        for sheet, df in sheets.items():
            if "品名" not in df.columns or "单价(元)" not in df.columns:
                continue
            for _, r in df.iterrows():
                name = str(r.get("品名", "")).strip()
                price = pd.to_numeric(r.get("单价(元)"), errors="coerce")
                if not name or name == "nan" or pd.isna(price):
                    continue
                rows.append({
                    "id": i,
                    "sheet": sheet,
                    "品类": "" if pd.isna(r.get("品类", "")) else str(r.get("品类", "")).strip(),
                    "品名": name,
                    "规格": "" if pd.isna(r.get("规格", "")) else str(r.get("规格", "")).strip(),
                    "单位": "" if pd.isna(r.get("单位", "")) else str(r.get("单位", "")).strip(),
                    "单价": float(price),
                    "供应商": "" if pd.isna(r.get("供应商", "")) else str(r.get("供应商", "")).strip(),
                    "启用": "是",
                    "来源": path.name,
                    "更新时间": today_str(),
                })
                i += 1
    pd.DataFrame(rows, columns=LIB_COLS).to_csv(LIB_CSV, index=False, encoding="utf-8-sig")


def read_csv(path, cols):
    if USE_SUPABASE:
        try:
            table = TABLE_MAP.get(Path(path).name)
            if table:
                r = requests.get(f"{SUPABASE_URL}/rest/v1/{table}?select=*", headers=SB_HEADERS, timeout=20)
                if r.status_code < 400:
                    data = r.json()
                    df = pd.DataFrame(data)
                    for c in cols:
                        if c not in df.columns:
                            df[c] = ""
                    return df[cols].fillna("")
        except Exception:
            pass
    if not path.exists():
        return pd.DataFrame(columns=cols)
    return pd.read_csv(path, dtype={"id": str}).fillna("")


def _clean_records(df):
    out = []
    for rec in df.fillna("").to_dict(orient="records"):
        clean = {}
        for k, v in rec.items():
            if hasattr(v, "item"):
                v = v.item()
            if isinstance(v, float) and math.isnan(v):
                v = ""
            clean[k] = v
        out.append(clean)
    return out


def write_csv(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")
    if USE_SUPABASE:
        table = TABLE_MAP.get(Path(path).name)
        if table:
            try:
                requests.delete(f"{SUPABASE_URL}/rest/v1/{table}?id=neq.__never__", headers=SB_HEADERS, timeout=30)
                rows = _clean_records(df)
                if rows:
                    for i in range(0, len(rows), 500):
                        requests.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HEADERS, data=json.dumps(rows[i:i+500], ensure_ascii=False), timeout=30)
            except Exception:
                pass


def classify_category(sheet='', cat='', name=''):
    text = f"{sheet} {cat} {name}"
    if any(k in text for k in ['蔬菜', '菜', '菇', '椒', '笋', '豆腐', '豆芽']):
        return '蔬菜'
    if any(k in text for k in ['海鲜', '鱼', '虾', '蛎', '贝', '蟹', '花蛤']):
        return '海鲜'
    if any(k in text for k in ['冻品', '鸡', '鸭血', '肉类冻品', '丸', '肠', '鱼卷', '糍粑']):
        return '冻品'
    if any(k in text for k in ['调料', '米油蛋调料', '酱', '油', '盐', '味精', '蚝油', '醋', '粉', '料']):
        return '调料'
    if any(k in text for k in ['猪肉', '五花', '龙骨', '猪', '瘦肉', '肉沫', '排骨']):
        return '猪肉'
    return '调料' if '米油' in text else '冻品' if '肉类' in text else '蔬菜' if '蔬菜' in text else ''


def init_suppliers(force=False):
    if SUPPLIERS_CSV.exists() and not force:
        return
    init_library()
    lib = read_csv(LIB_CSV, LIB_COLS)
    rows = []
    seen = set()
    i = 1
    if not lib.empty and '供应商' in lib.columns:
        for _, r in lib.iterrows():
            supplier = str(r.get('供应商', '')).strip()
            if not supplier:
                continue
            category = classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) or '其他'
            key = (category, supplier)
            if key in seen:
                continue
            seen.add(key)
            rows.append({'id': str(i), '类别': category, '供货商': supplier, '联系人': '', '电话': '', '地址': '', '备注': '从成本库供应商字段自动整理', '启用': '是', '更新时间': today_str()})
            i += 1
    write_csv(SUPPLIERS_CSV, pd.DataFrame(rows, columns=SUPPLIER_COLS))


def suppliers_df():
    init_suppliers()
    df = read_csv(SUPPLIERS_CSV, SUPPLIER_COLS)
    for c in SUPPLIER_COLS:
        if c not in df.columns:
            df[c] = ''
    df['id'] = df['id'].astype(str)
    return df[SUPPLIER_COLS]


def library_df():
    init_library()
    df = read_csv(LIB_CSV, LIB_COLS)
    if "id" in df.columns:
        df["id"] = df["id"].astype(str)
    df["单价"] = pd.to_numeric(df.get("单价", 0), errors="coerce")
    return df


def orders_df():
    df = read_csv(ORDERS_CSV, ORDER_COLS)
    if '记录ID' not in df.columns:
        df.insert(0, '记录ID', '')
    # 兼容旧数据：以前没有记录ID，这里自动补上，方便单行删除。
    changed = False
    for i in df.index:
        if not str(df.at[i, '记录ID']).strip():
            df.at[i, '记录ID'] = datetime.now().strftime('%Y%m%d%H%M%S%f') + f'_{i}'
            changed = True
    for c in ORDER_COLS:
        if c not in df.columns:
            df[c] = ''
    df = df[ORDER_COLS]
    if changed:
        write_csv(ORDERS_CSV, df)
    return df


def sales_df():
    return read_csv(SALES_CSV, ["日期", "营业额"])


def log_change(action, detail):
    row = {"时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "动作": action, "详情": detail}
    df = read_csv(LOG_CSV, ["时间", "动作", "详情"])
    write_csv(LOG_CSV, pd.concat([df, pd.DataFrame([row])], ignore_index=True))


def search_price(item_name):
    df = library_df()
    df = df[df["启用"].astype(str).fillna("是") != "否"].copy()
    if df.empty:
        return None, pd.DataFrame()
    key = ALIASES.get(str(item_name).strip(), str(item_name).strip())
    nk = norm(key)
    df["_n"] = df["品名"].map(norm)
    cand = df[df["_n"] == nk]
    if cand.empty:
        cand = df[df.apply(lambda r: nk in norm(str(r.get("品名", "")) + str(r.get("规格", ""))), axis=1)]
    if cand.empty:
        return None, pd.DataFrame()
    cand = cand.sort_values("单价", ascending=False).drop(columns=["_n"], errors="ignore")
    return cand.iloc[0].to_dict(), cand


def parse_pack_per_case(spec):
    text = str(spec or "")
    m = re.search(r"[×xX*]\s*(\d+)\s*包\s*/?\s*件", text) or re.search(r"(\d+)\s*包\s*/?\s*件", text)
    return int(m.group(1)) if m else 10


def calc_line(item, qty, unit, manual_total=""):
    qty = float(qty or 0)
    unit = str(unit or "").strip()
    if str(manual_total or "").strip():
        total = float(manual_total)
        price = total / qty if qty else total
        return qty, unit or "整项", price, total, "✅ 手填整价", "-", "", "按手填整价入账"
    hit, _ = search_price(item)
    if hit is None:
        return qty, unit, math.nan, math.nan, "❌ 库里缺价", "-", "", "需要补单价或整价"
    price_unit = str(hit.get("单位", "")).strip()
    spec = str(hit.get("规格", "")).strip()
    calc_qty = qty
    note = "按成本库取价；多价格取贵"
    if unit == "包" and price_unit == "件":
        per = parse_pack_per_case(spec)
        calc_qty = qty / per
        note = f"{qty:g}包 ÷ {per}包/件 = {calc_qty:g}件"
    elif unit and price_unit and unit != price_unit and price_unit == "斤" and unit in ["包", "个", "袋", "板", "盒"]:
        return qty, unit, math.nan, math.nan, "⚠️ 单位待确认", hit["品名"], spec, f"库里按{price_unit}，你填{unit}；请填重量或整价"
    price = float(hit["单价"])
    return calc_qty, price_unit or unit, price, calc_qty * price, "✅ 已取价", hit["品名"], spec, note


def recalc_order_row(row):
    """按当前行里的食材/数量/单位/单价重新算小计，用于采购页直接改单价后刷新计算过程。"""
    qty = float(pd.to_numeric(row.get("输入数量", 0), errors="coerce") or 0)
    input_unit = str(row.get("输入单位", "")).strip()
    calc_qty = float(pd.to_numeric(row.get("计价数量", qty), errors="coerce") or 0)
    calc_unit = str(row.get("计价单位", input_unit)).strip() or input_unit
    price = float(pd.to_numeric(row.get("单价", 0), errors="coerce") or 0)
    subtotal = calc_qty * price
    row["小计"] = round(subtotal, 2)
    if not str(row.get("状态", "")).strip() or str(row.get("状态", "")).startswith("❌") or str(row.get("状态", "")).startswith("⚠️"):
        row["状态"] = "✅ 手动改价"
    else:
        row["状态"] = "✅ 手动改价"
    note = str(row.get("备注", "")).strip()
    if "采购页手动改价" not in note:
        row["备注"] = (note + "；" if note else "") + "采购页手动改价"
    row["计价单位"] = calc_unit
    row["计价数量"] = calc_qty
    return row


def ensure_library_price_from_order(row, price):
    """采购页缺价时，按这行食材新增/更新成本库；以后同名食材自动取价。"""
    df = library_df()
    item = str(row.get("食材", "")).strip()
    unit = str(row.get("计价单位") or row.get("输入单位") or "").strip()
    category = str(row.get("批次") or "").strip() or classify_category('', '', item) or '其他'
    if category == '海鲜/整价':
        category = '海鲜'
    price = float(price or 0)
    nk = norm(item)
    target = df[(df["品名"].map(norm) == nk) & (df["单位"].astype(str) == unit)] if not df.empty else pd.DataFrame()
    if not target.empty:
        i = target.index[0]
        old = df.loc[i].to_dict()
        df.at[i, "单价"] = price
        df.at[i, "品类"] = df.at[i, "品类"] or category
        df.at[i, "启用"] = "是"
        df.at[i, "来源"] = "采购测算页补价"
        df.at[i, "更新时间"] = today_str()
        write_csv(LIB_CSV, df[LIB_COLS])
        log_change("采购页补价更新成本库", json.dumps({"旧": old, "新": df.loc[i].to_dict()}, ensure_ascii=False))
        return df.loc[i].to_dict(), "更新成本库"
    next_id = str((pd.to_numeric(df["id"], errors="coerce").max() or 0) + 1) if not df.empty else "1"
    lib_row = {c: "" for c in LIB_COLS}
    lib_row.update({
        "id": next_id,
        "sheet": "采购页补价",
        "品类": category,
        "品名": item,
        "规格": "",
        "单位": unit,
        "单价": price,
        "供应商": "",
        "启用": "是",
        "来源": "采购测算页缺价补录",
        "更新时间": today_str(),
    })
    write_csv(LIB_CSV, pd.concat([df, pd.DataFrame([lib_row])], ignore_index=True)[LIB_COLS])
    log_change("采购页补价新增成本库", json.dumps(lib_row, ensure_ascii=False))
    return lib_row, "新增成本库"


def parse_bulk_text(text):
    rows = []
    # 支持“半件/半包/半斤”等口述；也支持瓶、罐等调料常见单位。
    unit_pat = r"斤|件|包|个|袋|板|盒|桶|瓶|罐|条|只"
    for raw in str(text or "").splitlines():
        line = raw.strip()
        if not line or re.search(r"营业额|销售额", line):
            continue
        # 跳过类似“富临店（5月21号）”这种抬头，不当成食材。
        if re.search(r"店", line) and re.search(r"\d+\s*月\s*\d+\s*号?", line):
            continue
        line = line.replace("：", ":").replace("，", ",")
        m_half = re.search(rf"(.+?)\s*半\s*({unit_pat})", line)
        if m_half:
            unit = m_half.group(2)
            if unit in ["条", "只"]:
                unit = "个"
            rows.append({"食材": m_half.group(1).strip(" ,-，:：（("), "数量": 0.5, "单位": unit, "整价": "", "原文": raw})
            continue
        m = re.search(rf"(.+?)\s*([0-9]+(?:\.[0-9]+)?)\s*({unit_pat})", line)
        if not m:
            m2 = re.search(r"(.+?)\s*([0-9]+(?:\.[0-9]+)?)\s*元?$", line)
            if m2:
                rows.append({"食材": m2.group(1).strip(" ,-，:："), "数量": 1.0, "单位": "整项", "整价": float(m2.group(2)), "原文": raw})
            else:
                # 不认识的也保留到计算过程里，方便梦洁姐看见后补数量/单价或删除。
                rows.append({"食材": line, "数量": 1.0, "单位": "待确认", "整价": "", "原文": raw})
            continue
        unit = m.group(3)
        if unit in ["条", "只"]:
            unit = "个"
        rows.append({"食材": m.group(1).strip(" ,-，:：（("), "数量": float(m.group(2)), "单位": unit, "整价": "", "原文": raw})
    return rows


def prev_date_str(biz_date):
    try:
        return (datetime.strptime(str(biz_date), "%Y-%m-%d").date() - timedelta(days=1)).isoformat()
    except Exception:
        return (date.today() - timedelta(days=1)).isoformat()


def day_review_rows(biz_date):
    """回溯前一天：总结成本率、缺价/手动补价、金额较大的品项，给第二天录单时参考。"""
    prev = prev_date_str(biz_date)
    od = orders_df()
    sd = sales_df()
    if od.empty:
        return {"前日日期": prev, "有数据": False, "总结": "前一天暂无采购明细", "建议": [], "高金额": [], "缺价": [], "手动补价": [], "批次": []}
    d = od[od["日期"].astype(str) == prev].copy()
    if d.empty:
        return {"前日日期": prev, "有数据": False, "总结": "前一天暂无采购明细", "建议": [], "高金额": [], "缺价": [], "手动补价": [], "批次": []}
    d["小计数"] = pd.to_numeric(d.get("小计", 0), errors="coerce").fillna(0)
    total = float(d["小计数"].sum())
    sale = 0.0
    if not sd.empty:
        m = sd[sd["日期"].astype(str) == prev]
        if not m.empty:
            sale = float(pd.to_numeric(m.iloc[-1]["营业额"], errors="coerce") or 0)
    rate_text = "未填营业额"
    suggestions = []
    if sale:
        rate = total / sale
        rate_text = f"{rate*100:.2f}%"
        if rate > 0.40:
            suggestions.append(f"前一天已超40%，超出约¥{total - sale*0.4:.2f}，今天同类高金额品项要优先压价或减量。")
        elif rate > 0.38:
            suggestions.append("前一天接近40%红线，今天录单时重点看猪肉/冻品/调料这些大额项。")
        else:
            suggestions.append("前一天成本率安全，今天按同样口径录单即可。")
    else:
        suggestions.append(f"前一天总成本¥{total:.2f}，但没填营业额；补营业额后才能判断是否超40%。")
    high = d.sort_values("小计数", ascending=False).head(8)
    high_rows = [{"批次": r.get("批次",""), "食材": r.get("食材",""), "数量": f"{r.get('输入数量','')}{r.get('输入单位','')}", "单价": r.get("单价",""), "小计": round(float(r.get("小计数") or 0),2)} for _, r in high.iterrows()]
    missing = d[d["状态"].astype(str).str.contains("缺价|单位待确认", na=False)].head(8)
    missing_rows = [{"食材": r.get("食材",""), "状态": r.get("状态",""), "备注": r.get("备注","")} for _, r in missing.iterrows()]
    manual = d[d["状态"].astype(str).str.contains("手动改价|已补价入库|手填整价", na=False)].head(8)
    manual_rows = [{"食材": r.get("食材",""), "单价": r.get("单价",""), "状态": r.get("状态",""), "备注": r.get("备注","")} for _, r in manual.iterrows()]
    batch = d.groupby("批次", dropna=False)["小计数"].sum().reset_index().sort_values("小计数", ascending=False)
    batch_rows = [{"批次": r["批次"], "金额": round(float(r["小计数"]),2)} for _, r in batch.iterrows()]
    if len(missing_rows):
        suggestions.append("前一天还有缺价/单位待确认，今天如果再出现同品项，先补单价并勾选存入成本库。")
    if len(manual_rows):
        suggestions.append("前一天手动补过价的品项已可作为今天参考；若已勾选入库，下次会自动取价。")
    return {"前日日期": prev, "有数据": True, "总结": f"前一天总成本¥{total:.2f}，营业额¥{sale:.2f}，成本率{rate_text}", "建议": suggestions, "高金额": high_rows, "缺价": missing_rows, "手动补价": manual_rows, "批次": batch_rows}


def current_vs_previous_rows(biz_date):
    """把当天已录品项和前一天同品项对比，提示涨价/降价。"""
    prev = prev_date_str(biz_date)
    od = orders_df()
    if od.empty:
        return []
    cur = od[od["日期"].astype(str) == str(biz_date)].copy()
    pre = od[od["日期"].astype(str) == prev].copy()
    if cur.empty or pre.empty:
        return []
    out=[]
    pre_map={}
    for _, r in pre.iterrows():
        key=(norm(r.get("食材","")), str(r.get("计价单位") or r.get("输入单位") or ""))
        try: pre_map[key]=float(r.get("单价") or 0)
        except Exception: pass
    for _, r in cur.iterrows():
        key=(norm(r.get("食材","")), str(r.get("计价单位") or r.get("输入单位") or ""))
        if key in pre_map:
            try: cur_price=float(r.get("单价") or 0)
            except Exception: continue
            old=pre_map[key]
            if old and abs(cur_price-old) >= 0.01:
                out.append({"食材": r.get("食材",""), "单位": key[1], "昨日单价": old, "今日单价": cur_price, "变化": round(cur_price-old,2)})
    return out[:10]


def calc_summary(biz_date):
    od = orders_df()
    today = od[od["日期"].astype(str) == str(biz_date)].copy() if not od.empty else od
    total = pd.to_numeric(today.get("小计", 0), errors="coerce").fillna(0).sum() if not today.empty else 0.0
    sd = sales_df()
    sale = 0.0
    if not sd.empty:
        m = sd[sd["日期"].astype(str) == str(biz_date)]
        if not m.empty:
            sale = float(pd.to_numeric(m.iloc[-1]["营业额"], errors="coerce") or 0)
    if sale:
        rate = total / sale
        red = sale * 0.4
        if rate <= 0.38:
            conclusion = f"✅ OK：当前成本率 {rate*100:.2f}%"
        elif rate <= 0.40:
            conclusion = f"⚠️ 临界：当前成本率 {rate*100:.2f}%"
        else:
            conclusion = f"❌ 已超 {(rate-0.40)*100:.2f}%：需削减 ¥{total-red:.2f}"
        return {"结论": conclusion, "总成本": round(total, 2), "营业额": round(sale, 2), "成本率": f"{rate*100:.2f}%", "红线": round(red, 2), "公式": f"总成本 ¥{total:.2f} ÷ 营业额 ¥{sale:.2f} = 成本率 {rate*100:.2f}%；40%红线 = ¥{sale:.2f} × 40% = ¥{red:.2f}"}
    return {"结论": "营业额待填", "总成本": round(total, 2), "营业额": 0, "成本率": "待算", "红线": round(total / 0.4 if total else 0, 2), "公式": f"当前总成本 ¥{total:.2f}；若要不超40%，营业额至少需要 ¥{(total/0.4 if total else 0):.2f}"}


def build_calc_process_rows(df):
    if df.empty:
        return []
    rows = []
    for r in df.to_dict('records'):
        try:
            calc_qty = float(r.get('计价数量') or 0)
            price = float(r.get('单价') or 0)
            subtotal = float(r.get('小计') or 0)
            formula = f"{calc_qty:g}{r.get('计价单位','')} × ¥{price:g} = ¥{subtotal:.2f}"
        except Exception:
            formula = "待补价/待确认，暂不计入"
        input_qty = f"{r.get('输入数量','')}{r.get('输入单位','')}"
        match = str(r.get('报价库品名','-'))
        if r.get('规格'):
            match += f"（{r.get('规格')}）"
        rows.append({**r, '输入': input_qty, '匹配': match, '计算公式': formula})
    return rows


def batch_summary_rows(df):
    if df.empty:
        return []
    tmp = df.copy()
    tmp['小计数'] = pd.to_numeric(tmp.get('小计', 0), errors='coerce').fillna(0)
    out = tmp.groupby('批次', dropna=False)['小计数'].sum().reset_index()
    return [{'批次': r['批次'], '金额': round(float(r['小计数']), 2)} for _, r in out.iterrows()]


BASE_HTML = """
<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1,maximum-scale=1'>
<title>梦洁姐采购成本实时看板</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--text:#111;--muted:#666;--line:#e8e8ee;--black:#111;--green:#087f3e;--red:#b00020;--orange:#b26b00;--blue:#2454d6}*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,'Microsoft YaHei',sans-serif;margin:0;background:var(--bg);color:var(--text);font-size:16px}.wrap{max-width:1180px;margin:auto;padding:14px 14px 78px}.top{position:sticky;top:0;z-index:9;background:rgba(246,247,249,.96);backdrop-filter:blur(8px);padding:10px 0 8px;border-bottom:1px solid var(--line)}h2{font-size:21px;margin:4px 0 8px}.nav{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}.nav a,.btn{background:var(--black);color:#fff;text-decoration:none;border:0;border-radius:13px;padding:12px 14px;display:inline-flex;align-items:center;justify-content:center;min-height:44px;font-weight:700;font-size:15px}.nav a.light,.btn.light{background:#fff;color:#111;border:1px solid #ddd}.btn.full{width:100%}.card{background:var(--card);border-radius:18px;padding:16px;margin:14px 0;box-shadow:0 1px 8px rgba(0,0,0,.06);border:1px solid #eee}.section-title{display:flex;align-items:center;gap:8px;margin:0 0 10px;font-size:19px}.step{display:inline-flex;width:28px;height:28px;border-radius:50%;align-items:center;justify-content:center;background:#111;color:#fff;font-weight:800}.hint{background:#f0f5ff;border:1px solid #d9e5ff;color:#173b91;border-radius:14px;padding:10px 12px;margin:10px 0}input,select,textarea{width:100%;box-sizing:border-box;padding:13px 12px;border:1px solid #ccd0d8;border-radius:13px;font-size:16px;background:#fff;min-height:46px}textarea{min-height:150px;line-height:1.45}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}.metric{font-size:22px;font-weight:800;line-height:1.25}.ok{color:#067d28}.warn{color:#b26b00}.bad{color:#b00020}.inline-form{display:flex;gap:6px;align-items:center}.price-save-form{min-width:170px}.mini-check{display:flex;gap:5px;align-items:center;margin-top:6px;fon... [truncated]
</style>
</head><body><div class='wrap'>
<div class='top'><h2>⭐Clarity⭐ 梦洁姐采购成本实时看板</h2><div class='nav desktop-only'><a href='/'>①采购测算</a><a href='/library'>②成本库/改价</a><a href='/suppliers'>③供货商/价格</a><a class='light' href='/logs'>修改记录</a></div></div>
{{content|safe}}
</div><div class='bottom-tabs'><a href='/'><b>🧾</b>采购</a><a href='/library'><b>💰</b>成本库</a><a href='/suppliers'><b>🚚</b>供货商</a><a href='/logs'><b>📝</b>记录</a></div></body></html>
"""


def page(content):
    return render_template_string(BASE_HTML, content=content)


@app.route("/")
def index():
    biz_date = request.args.get("date", today_str())
    summary = calc_summary(biz_date)
    od = orders_df()
    today = od[od["日期"].astype(str) == str(biz_date)].tail(200) if not od.empty else od
    if not today.empty:
        _cat_order = {"蔬菜": 1, "猪肉": 2, "冻品": 3, "米油蛋": 4, "调料": 5, "海鲜/整价": 6, "其他": 7}
        today = today.copy()
        today["_批次排序"] = today["批次"].map(_cat_order).fillna(99)
        today = today.sort_values(["_批次排序", "批次", "食材"]).drop(columns=["_批次排序"], errors="ignore")
    rows = build_calc_process_rows(today)
    batch_rows = batch_summary_rows(today)
    review = day_review_rows(biz_date)
    price_changes = current_vs_previous_rows(biz_date)
    content = render_template_string("""
<div class='card'><h3 class='section-title'><span class='step'>1</span>今天先做什么？</h3><div class='quick'><a href='/?date={{biz_date}}'>录采购单<span>粘贴蔬菜/猪肉/冻品单，自动算成本</span></a><a href='/library'>查/改成本价<span>单个 SKU 改价、新增报价</span></a><a href='/suppliers'>供货商价格<span>新增供货商整份报价、批量调价</span></a></div></div>
<div class='card'><h3 class='section-title'><span class='step'>2</span>填写营业额</h3><form method='post' action='/save_sale' class='row'><div><label>日期</label><input name='date' value='{{biz_date}}'></div><div><label>今日营业额</label><input name='sale' type='number' step='0.01' value='{{summary["营业额"]}}'></div><button class='btn'>保存营业额</button></form><br><form method='post' action='/clear_day' onsubmit="return confirm('确定清空这一天的采购明细和营业额吗？输错了才点。');"><input type='hidden' name='date' value='{{biz_date}}'><button class='btn danger'>清空当天全部数据</button> <span class='muted'>输错一整天时用；只错一行可以在下面单独删除。</span></form></div>
<div class='grid'>
<div class='card'><div class='muted'>结论</div><div class='metric'>{{summary['结论']}}</div></div><div class='card'><div class='muted'>总采购成本</div><div class='metric'>¥{{summary['总成本']}}</div></div><div class='card'><div class='muted'>营业额</div><div class='metric'>¥{{summary['营业额']}}</div></div><div class='card'><div class='muted'>成本率 / 40%红线</div><div class='metric'>{{summary['成本率']}} / ¥{{summary['红线']}}</div></div>
</div>
<div class='card'><h3 class='section-title'><span class='step'>3</span>总账计算过程</h3><p><b>{{summary['公式']}}</b></p><div class='scroll'><table><tr><th>批次</th><th>批次小计</th></tr>{% for b in batch_rows %}<tr><td>{{b['批次']}}</td><td>¥{{b['金额']}}</td></tr>{% endfor %}</table></div></div>
<div class='card'><h3 class='section-title'><span class='step'>4</span>昨日回溯 / 自动学习</h3><p><b>{{review['前日日期']}}</b>：{{review['总结']}}</p>{% for x in review['建议'] %}<div class='hint'>{{x}}</div>{% endfor %}{% if review['批次'] %}<h4>昨日批次小计</h4><div class='scroll'><table><tr><th>批次</th><th>金额</th></tr>{% for b in review['批次'] %}<tr><td>{{b['批次']}}</td><td>¥{{b['金额']}}</td></tr>{% endfor %}</table></div>{% endif %}{% if review['高金额'] %}<h4>昨日金额较大的品项</h4><div class='scroll'><table><tr><th>批次</th><th>食材</th><th>数量</th><th>单价</th><th>小计</th></tr>{% for r in review['高金额'] %}<tr><td>{{r['批次']}}</td><td>{{r['食材']}}</td><td>{{r['数量']}}</td><td>¥{{r['单价']}}</td><td>¥{{r['小计']}}</td></tr>{% endfor %}</table></div>{% endif %}{% if price_changes %}<h4>今天 vs 昨天同品项价格变化</h4><div class='scroll'><table><tr><th>食材</th><th>单位</th><th>昨日单价</th><th>今日单价</th><th>变化</th></tr>{% for r in price_changes %}<tr><td>{{r['食材']}}</td><td>{{r['单位']}}</td><td>¥{{r['昨日单价']}}</td><td>¥{{r['今日单价']}}</td><td>¥{{r['变化']}}</td></tr>{% endfor %}</table></div>{% endif %}</div>
<div class='card'><h3 class='section-title'><span class='step'>5</span>整段复制采购单自动计算</h3><form method='post' action='/bulk_add'><input type='hidden' name='date' value='{{biz_date}}'><div class='row'><div><label>批次</label><select name='batch'><option>蔬菜</option><option>猪肉</option><option>冻品</option><option>米油蛋</option><option>调料</option><option>海鲜/整价</option><option>其他</option></select></div></div><p class='muted'>一行一个：圆包菜10斤 / 毛毛肉2包 / 冠军鸭血1件 / 海鲜250</p><textarea name='text' placeholder='把采购内容整段粘贴到这里'></textarea><br><br><button class='btn green'>整批加入并测算</button></form></div>
<div class='card'><h3 class='section-title'><span class='step'>6</span>逐行计算过程 / 可直接改单价</h3><p class='muted'>这里能看到每一项怎么算：你输入的数量 → 匹配成本库 → 单位换算 → 计价数量 × 单价 = 小计。库里缺价时，直接在本行填单价，勾选“存入成本库”，点保存；下次再录这个食材就会自动取到这个价格。</p><div class='scroll'><table><tr><th>操作</th><th>批次</th><th>食材</th><th>你输入</th><th>匹配成本库</th><th>换算说明</th><th>计价数量</th><th>单价可改/可入库</th><th>计算公式</th><th>小计</th><th>状态</th></tr>{% for r in rows %}<tr><td><form method='post' action='/delete_order' onsubmit="return confirm('确定删除这一行吗？');"><input type='hidden' name='date' value='{{biz_date}}'><input type='hidden' name='rid' value='{{r['记录ID']}}'><button class='btn danger small'>删除</button></form></td><td>{{r['批次']}}</td><td>{{r['食材']}}</td><td>{{r['输入']}}</td><td>{{r['匹配']}}</td><td>{{r['备注']}}</td><td>{{r['计价数量']}}{{r['计价单位']}}</td><td><form method='post' action='/order/update_price' class='price-save-form'><input type='hidden' name='date' value='{{biz_date}}'><input type='hidden' name='rid' value='{{r['记录ID']}}'><div class='inline-form'><input class='price-input' name='price' type='number' step='0.01' value='{{r['单价']}}'><button class='btn small green'>保存</button></div><label class='mini-check'><input type='checkbox' name='save_to_library' value='1' {% if '缺价' in r['状态'] %}checked{% endif %}> 存入成本库</label></form></td><td><b>{{r['计算公式']}}</b></td><td>¥{{r['小计']}}</td><td>{{r['状态']}}</td></tr>{% endfor %}</table></div></div>
""", biz_date=biz_date, summary=summary, rows=rows, batch_rows=batch_rows, review=review, price_changes=price_changes)
    return page(content)


@app.post("/save_sale")
def save_sale():
    biz_date = request.form.get("date") or today_str()
    sale = float(request.form.get("sale") or 0)
    sd = sales_df()
    sd = sd[sd["日期"].astype(str) != str(biz_date)] if not sd.empty else sd
    write_csv(SALES_CSV, pd.concat([sd, pd.DataFrame([{"日期": biz_date, "营业额": sale}])], ignore_index=True))
    return redirect(f"/?date={biz_date}")


@app.post("/bulk_add")
def bulk_add():
    biz_date = request.form.get("date") or today_str()
    batch = request.form.get("batch") or "其他"
    text = request.form.get("text") or ""
    od = orders_df()
    new_rows = []
    for r in parse_bulk_text(text):
        if r["单位"] == "待识别" or r["数量"] == "":
            line = (r["数量"], r["单位"] or "待确认", math.nan, math.nan, "⚠️ 单位待确认", "-", "", "未识别数量单位，请补单价或删除")
        else:
            line = calc_line(r["食材"], r["数量"], r["单位"], r.get("整价", ""))
        new_rows.append(dict(zip(ORDER_COLS, [datetime.now().strftime('%Y%m%d%H%M%S%f'), biz_date, batch, r["食材"], r["数量"], r["单位"], *line])))
    if new_rows:
        od = pd.concat([od, pd.DataFrame(new_rows)], ignore_index=True)
        write_csv(ORDERS_CSV, od[ORDER_COLS])
    return redirect(f"/?date={biz_date}")


@app.post("/delete_order")
def delete_order():
    biz_date = request.form.get("date") or today_str()
    rid = str(request.form.get("rid") or "")
    od = orders_df()
    before = len(od)
    if rid and not od.empty:
        deleted = od[od["记录ID"].astype(str) == rid].to_dict("records")
        od = od[od["记录ID"].astype(str) != rid].copy()
        write_csv(ORDERS_CSV, od[ORDER_COLS])
        if before != len(od):
            log_change("删除采购明细", json.dumps({"日期": biz_date, "记录ID": rid, "删除": deleted[:1]}, ensure_ascii=False))
    return redirect(f"/?date={biz_date}")


@app.post("/order/update_price")
def order_update_price():
    biz_date = request.form.get("date") or today_str()
    rid = str(request.form.get("rid") or "")
    new_price = float(request.form.get("price") or 0)
    save_to_library = str(request.form.get("save_to_library") or "") == "1"
    od = orders_df()
    if rid and not od.empty:
        idx = od.index[od["记录ID"].astype(str) == rid]
        if len(idx):
            i = idx[0]
            old = od.loc[i].to_dict()
            od.at[i, "单价"] = new_price
            row = recalc_order_row(od.loc[i].to_dict())
            lib_action = ""
            lib_row = None
            if save_to_library:
                lib_row, lib_action = ensure_library_price_from_order(row, new_price)
                row["报价库品名"] = lib_row.get("品名", row.get("食材", ""))
                row["规格"] = lib_row.get("规格", "")
                row["状态"] = "✅ 已补价入库"
                note = str(row.get("备注", "")).strip()
                if lib_action not in note:
                    row["备注"] = (note + "；" if note else "") + lib_action
            for c in ORDER_COLS:
                od.at[i, c] = row.get(c, od.at[i, c])
            write_csv(ORDERS_CSV, od[ORDER_COLS])
            log_change("采购明细改单价", json.dumps({"日期": biz_date, "记录ID": rid, "食材": old.get("食材"), "旧单价": old.get("单价"), "新单价": new_price, "新小计": row.get("小计"), "是否入库": save_to_library, "入库动作": lib_action}, ensure_ascii=False))
    return redirect(f"/?date={biz_date}")


@app.post("/clear_day")
def clear_day():
    biz_date = request.form.get("date") or today_str()
    od = orders_df()
    sd = sales_df()
    removed_orders = 0
    removed_sales = 0
    if not od.empty:
        removed_orders = int((od["日期"].astype(str) == str(biz_date)).sum())
        od = od[od["日期"].astype(str) != str(biz_date)].copy()
        write_csv(ORDERS_CSV, od[ORDER_COLS])
    if not sd.empty:
        removed_sales = int((sd["日期"].astype(str) == str(biz_date)).sum())
        sd = sd[sd["日期"].astype(str) != str(biz_date)].copy()
        write_csv(SALES_CSV, sd[["日期", "营业额"]] if set(["日期", "营业额"]).issubset(sd.columns) else sd)
    log_change("清空当天数据", json.dumps({"日期": biz_date, "清空采购行数": removed_orders, "清空营业额记录": removed_sales}, ensure_ascii=False))
    return redirect(f"/?date={biz_date}")


@app.route("/library")
def library():
    kw = request.args.get("kw", "").strip()
    cat = request.args.get("cat", "全部")
    df = library_df()
    if cat != "全部":
        if cat == "蔬菜": df = df[df.apply(lambda r: classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) == '蔬菜', axis=1)]
        elif cat == "海鲜": df = df[df.apply(lambda r: classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) == '海鲜', axis=1)]
        elif cat == "冻品": df = df[df.apply(lambda r: classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) == '冻品', axis=1)]
        elif cat == "调料": df = df[df.apply(lambda r: classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) == '调料', axis=1)]
        elif cat == "猪肉": df = df[df.apply(lambda r: classify_category(r.get('sheet',''), r.get('品类',''), r.get('品名','')) == '猪肉', axis=1)]
        elif cat == "肉类冻品": df = df[df["sheet"].str.contains("肉类|冻品", na=False)]
        elif cat == "米油蛋调料": df = df[df["sheet"].str.contains("米油|蛋|调料", na=False)]
        elif cat == "调料/海鲜/其他": df = df[df["sheet"].str.contains("调料|海鲜|其他|新增", na=False) | df["品类"].str.contains("调料|海鲜|其他", na=False)]
    if kw:
        nk = norm(kw)
        df = df[df.apply(lambda r: nk in norm(str(r["品名"]) + str(r["规格"]) + str(r["品类"]) + str(r["供应商"])), axis=1)]
    rows = df.sort_values(["sheet", "品类", "品名"]).head(300).to_dict("records")
    content = render_template_string("""
<div class='card'><h3 class='section-title'><span class='step'>1</span>成本库价格 / 可改价</h3><div class='hint'>手机上先搜索品名，再改价格；表格可以左右滑动。</div><form class='row' method='get'><div><label>搜索</label><input name='kw' value='{{kw}}' placeholder='鸡腿、鸭血、花菜、调料、海鲜'></div><div><label>类别</label><select name='cat'>{% for c in ['全部','蔬菜','海鲜','冻品','调料','猪肉','肉类冻品','米油蛋调料','调料/海鲜/其他'] %}<option {% if c==cat %}selected{% endif %}>{{c}}</option>{% endfor %}</select></div><button class='btn'>查询</button></form></div>
<div class='card'><h3 class='section-title'><span class='step'>2</span>单个新增报价</h3><form method='post' action='/library/add' class='grid'><input name='品名' placeholder='品名，如：蚝油 / 海蛎'><input name='规格' placeholder='规格，如：6kg/桶'><select name='单位'><option>斤</option><option>件</option><option>包</option><option>桶</option><option>瓶</option><option>袋</option><option>盒</option><option>整项</option></select><input name='单价' type='number' step='0.01' placeholder='单价'><select name='品类'><option>蔬菜</option><option>海鲜</option><option>冻品</option><option>调料</option><option>猪肉</option><option>其他</option></select><input name='供应商' placeholder='供应商，可空'><button class='btn green'>新增到成本库</button></form></div>
<div class='card'><div class='scroll'><table><tr><th>ID</th><th>品类</th><th>品名</th><th>规格</th><th>单位</th><th>单价</th><th>供应商</th><th>启用</th><th>操作</th></tr>{% for r in rows %}<tr><form method='post' action='/library/update'><input type='hidden' name='id' value='{{r['id']}}'><td>{{r['id']}}</td><td><input name='品类' value='{{r['品类']}}'></td><td><input name='品名' value='{{r['品名']}}'></td><td><input name='规格' value='{{r['规格']}}'></td><td><input name='单位' value='{{r['单位']}}'></td><td><input name='单价' type='number' step='0.01' value='{{r['单价']}}'></td><td><input name='供应商' value='{{r['供应商']}}'></td><td><select name='启用'><option {% if r['启用']=='是' %}selected{% endif %}>是</option><option {% if r['启用']=='否' %}selected{% endif %}>否</option></select></td><td><button class='btn'>保存</button></td></form></tr>{% endfor %}</table></div></div>
""", rows=rows, kw=kw, cat=cat)
    return page(content)


@app.post("/library/add")
def library_add():
    df = library_df()
    next_id = str((pd.to_numeric(df["id"], errors="coerce").max() or 0) + 1)
    row = {c: "" for c in LIB_COLS}
    row.update({"id": next_id, "sheet": "新增报价", "启用": "是", "来源": "网页新增", "更新时间": today_str()})
    for c in ["品类", "品名", "规格", "单位", "供应商"]:
        row[c] = request.form.get(c, "").strip()
    row["单价"] = float(request.form.get("单价") or 0)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    write_csv(LIB_CSV, df[LIB_COLS])
    log_change("新增报价", json.dumps(row, ensure_ascii=False))
    # 若新增报价时填了供应商，同步进供货商管理页；没填则不强制。
    if row.get("供应商"):
        sdf = suppliers_df()
        exists = ((sdf["类别"].astype(str) == str(row.get("品类"))) & (sdf["供货商"].astype(str) == str(row.get("供应商")))).any() if not sdf.empty else False
        if not exists:
            sid = str((pd.to_numeric(sdf["id"], errors="coerce").max() or 0) + 1)
            srow = {"id": sid, "类别": row.get("品类") or classify_category(row.get("sheet",""), row.get("品类",""), row.get("品名","")), "供货商": row.get("供应商"), "联系人": "", "电话": "", "地址": "", "备注": "新增报价时自动加入", "启用": "是", "更新时间": today_str()}
            write_csv(SUPPLIERS_CSV, pd.concat([sdf, pd.DataFrame([srow])], ignore_index=True)[SUPPLIER_COLS])
    return redirect("/library")


@app.post("/library/update")
def library_update():
    df = library_df()
    idv = str(request.form.get("id"))
    idx = df.index[df["id"].astype(str) == idv]
    if len(idx):
        i = idx[0]
        old = df.loc[i].to_dict()
        for c in ["品类", "品名", "规格", "单位", "供应商", "启用"]:
            df.at[i, c] = request.form.get(c, "").strip()
        df.at[i, "单价"] = float(request.form.get("单价") or 0)
        df.at[i, "更新时间"] = today_str()
        write_csv(LIB_CSV, df[LIB_COLS])
        log_change("修改价格/资料", json.dumps({"旧": old, "新": df.loc[i].to_dict()}, ensure_ascii=False))
    return redirect("/library")


@app.route("/suppliers")
def suppliers():
    init_suppliers()
    sdf = suppliers_df()
    lib = library_df()
    supplier_names = sorted(set([x for x in lib["供应商"].astype(str).unique() if x] + [x for x in sdf["供货商"].astype(str).unique() if x]))
    category_tables = {}
    for cat in SUPPLIER_CATEGORIES:
        rows = sdf[(sdf["类别"].astype(str) == cat) & (sdf["启用"].astype(str) != "否")].sort_values("供货商").to_dict("records")
        # 给每个供货商补上成本库 SKU 数量，方便点名改价。
        for r in rows:
            r['SKU数量'] = int((lib['供应商'].astype(str) == str(r.get('供货商',''))).sum()) if not lib.empty else 0
        category_tables[cat] = rows
    all_rows = sdf.sort_values(["类别", "供货商"]).to_dict("records")
    content = render_template_string("""
<div class='card'><h3 class='section-title'><span class='step'>2</span>更新已有供货商价格</h3><p class='muted'>梦洁姐，这里就是你要的：选一个供货商 → 粘贴它最新报价 → 系统自动匹配成本库 → 先预览旧价/新价/差额 → 点确认后实时批量更新价格。</p>
<form method='post' action='/suppliers/price_preview'>
<div class='grid'><div><label>要更新哪个供货商</label><select name='old'>{% for s in supplier_names %}<option>{{s}}</option>{% endfor %}</select></div><div><label>更新后供货商名称</label><input name='new_supplier' placeholder='不换名字就留空；换供货商就填新名字'></div></div>
<p class='muted'>粘贴格式：一行一个，支持：品名 单价；品名 单位 单价；品名 规格 单位 单价。例：圆包菜 斤 0.95 / 花菜 斤 2.6 / 鸡腿XL 件 86 / 蚝油 6kg/桶 桶 38</p>
<textarea name='price_text' placeholder='把这个供货商的新报价整段粘贴到这里\n圆包菜 斤 0.95\n花菜 斤 2.6\n鸡腿XL 件 86'></textarea><br><br><button class='btn green'>预览旧价/新价并批量更新</button>
</form></div>
<div class='card'><h3 class='section-title'><span class='step'>1</span>新增新的供货商和价格</h3><p class='muted'>新供货商第一次给整份报价时，用这里：填供货商名称 + 类别，整段粘贴价格，系统会把这些品项批量新增到成本库。以后价格变动，再用上面的“实时批量更新供货商价格”。</p>
<form method='post' action='/suppliers/new_price_preview'>
<div class='grid'><div><label>供货商名称</label><input name='supplier' placeholder='例如：新蔬菜供货商'></div><div><label>类别</label><select name='category'>{% for c in cats %}<option>{{c}}</option>{% endfor %}</select></div></div>
<p class='muted'>粘贴格式：一行一个，例：圆包菜 斤 0.95 / 花菜 斤 2.6 / 蚝油 6kg/桶 桶 38。系统会先预览，已存在的显示“可更新”，不存在的显示“可新增”。</p>
<textarea name='price_text' placeholder='把新供货商报价整段粘贴到这里
圆包菜 斤 0.95
花菜 斤 2.6
蚝油 6kg/桶 桶 38'></textarea><br><br><button class='btn green'>预览并批量新增/更新价格</button>
</form></div>
<div class='card'><h3 class='section-title'><span class='step'>3</span>只新增供货商名录</h3><p class='muted'>先把供货商加进来，后面就可以按供货商批量更新价格。</p>
<form method='post' action='/suppliers/add' class='grid'><select name='类别'>{% for c in cats %}<option>{{c}}</option>{% endfor %}</select><input name='供货商' placeholder='供货商名称'><input name='联系人' placeholder='联系人，可空'><input name='电话' placeholder='电话，可空'><input name='地址' placeholder='地址，可空'><input name='备注' placeholder='备注，可空'><button class='btn green'>新增供货商</button></form></div>
<div class='grid'>{% for cat, rows in category_tables.items() %}<div class='card'><h3>{{cat}}供货商</h3>{% if rows %}<table><tr><th>供货商</th><th>SKU数</th><th>电话</th></tr>{% for r in rows %}<tr><td><b>{{r['供货商']}}</b></td><td>{{r['SKU数量']}}</td><td>{{r['电话']}}</td></tr>{% endfor %}</table>{% else %}<p class='muted'>暂未添加</p>{% endif %}</div>{% endfor %}</div>
<div class='card'><h3 class='section-title'><span class='step'>4</span>供货商名录编辑</h3><div class='scroll'><table><tr><th>ID</th><th>类别</th><th>供货商</th><th>联系人</th><th>电话</th><th>地址</th><th>备注</th><th>启用</th><th>操作</th></tr>{% for r in all_rows %}<tr><form method='post' action='/suppliers/update'><input type='hidden' name='id' value='{{r['id']}}'><td>{{r['id']}}</td><td><select name='类别'>{% for c in cats %}<option {% if c==r['类别'] %}selected{% endif %}>{{c}}</option>{% endfor %}</select></td><td><input name='供货商' value='{{r['供货商']}}'></td><td><input name='联系人' value='{{r['联系人']}}'></td><td><input name='电话' value='{{r['电话']}}'></td><td><input name='地址' value='{{r['地址']}}'></td><td><input name='备注' value='{{r['备注']}}'></td><td><select name='启用'><option {% if r['启用']=='是' %}selected{% endif %}>是</option><option {% if r['启用']=='否' %}selected{% endif %}>否</option></select></td><td><button class='btn'>保存</button></td></form></tr>{% endfor %}</table></div></div>
<div class='card'><h3 class='section-title'><span class='step'>5</span>辅助：只替换供应商名称 / 停用旧供应商</h3><p class='muted'>这个只用于供应商改名或停用旧报价；不会改价格。要改价格，用页面最上面的“实时批量更新供货商价格”。</p><form method='post' action='/suppliers/replace' class='grid'><select name='old'>{% for s in supplier_names %}<option>{{s}}</option>{% endfor %}</select><input name='new' placeholder='新供应商名称'><select name='mode'><option value='replace'>只替换供应商名称，不改价格</option><option value='disable'>停用旧供应商报价</option></select><button class='btn danger'>执行名称替换/停用</button></form></div>
""", cats=SUPPLIER_CATEGORIES, category_tables=category_tables, all_rows=all_rows, supplier_names=supplier_names)
    return page(content)


@app.post('/suppliers/add')
def supplier_add():
    sdf = suppliers_df()
    name = request.form.get('供货商', '').strip()
    cat = request.form.get('类别', '').strip() or '蔬菜'
    if name:
        exists = ((sdf['类别'].astype(str) == cat) & (sdf['供货商'].astype(str) == name)).any() if not sdf.empty else False
        if not exists:
            sid = str((pd.to_numeric(sdf['id'], errors='coerce').max() or 0) + 1)
            row = {c: '' for c in SUPPLIER_COLS}
            row.update({'id': sid, '类别': cat, '供货商': name, '联系人': request.form.get('联系人','').strip(), '电话': request.form.get('电话','').strip(), '地址': request.form.get('地址','').strip(), '备注': request.form.get('备注','').strip(), '启用': '是', '更新时间': today_str()})
            write_csv(SUPPLIERS_CSV, pd.concat([sdf, pd.DataFrame([row])], ignore_index=True)[SUPPLIER_COLS])
            log_change('新增供货商', json.dumps(row, ensure_ascii=False))
    return redirect('/suppliers')


@app.post('/suppliers/update')
def supplier_update():
    sdf = suppliers_df()
    idv = str(request.form.get('id',''))
    idx = sdf.index[sdf['id'].astype(str) == idv]
    if len(idx):
        i = idx[0]
        old = sdf.loc[i].to_dict()
        for c in ['类别','供货商','联系人','电话','地址','备注','启用']:
            sdf.at[i, c] = request.form.get(c, '').strip()
        sdf.at[i, '更新时间'] = today_str()
        write_csv(SUPPLIERS_CSV, sdf[SUPPLIER_COLS])
        log_change('修改供货商', json.dumps({'旧': old, '新': sdf.loc[i].to_dict()}, ensure_ascii=False))
    return redirect('/suppliers')


def parse_supplier_price_text(text):
    rows = []
    units = {'斤','件','包','桶','瓶','袋','盒','个','板','只','条','整项'}
    for raw in str(text or '').splitlines():
        line = raw.strip()
        if not line or line.startswith('#'):
            continue
        line = line.replace('：', ':').replace('，', ',').replace('\t', ' ')
        parts = [p for p in re.split(r'[,，\s]+', line) if p]
        price = None
        price_idx = None
        for i in range(len(parts)-1, -1, -1):
            v = parts[i].replace('元','')
            if re.fullmatch(r'\d+(?:\.\d+)?', v):
                price = float(v)
                price_idx = i
                break
        if price is None:
            rows.append({'原文': raw, '品名': line, '规格': '', '单位': '', '新价': '', '状态': '❌ 未识别单价'})
            continue
        before = parts[:price_idx]
        unit = ''
        spec = ''
        if before and before[-1] in units:
            unit = before[-1]
            before = before[:-1]
        if len(before) >= 2 and (('/' in before[-1]) or ('×' in before[-1]) or ('*' in before[-1]) or any(u in before[-1] for u in units)):
            spec = before[-1]
            before = before[:-1]
        name = ''.join(before).strip() if before else line[:line.rfind(str(parts[price_idx]))].strip()
        rows.append({'原文': raw, '品名': name, '规格': spec, '单位': unit, '新价': price, '状态': '待匹配'})
    return rows


def match_library_for_supplier_price(name, old_supplier=''):
    import difflib
    df = library_df()
    if old_supplier:
        scoped = df[df['供应商'].astype(str) == str(old_supplier)].copy()
    else:
        scoped = df.copy()
    if scoped.empty:
        scoped = df.copy()
    nk = norm(name)
    scoped['_n'] = scoped['品名'].map(norm)
    exact = scoped[scoped['_n'] == nk]
    if not exact.empty:
        return exact.iloc[0].to_dict(), '精确匹配'
    scored = []
    for _, r in scoped.iterrows():
        rn = norm(str(r.get('品名','')) + str(r.get('规格','')))
        pn = norm(str(r.get('品名','')))
        if not nk:
            continue
        if nk in rn or pn in nk:
            ratio = difflib.SequenceMatcher(None, nk, pn).ratio()
            # 防止“蚝油”误匹配到“蚝油肉片”这类不同 SKU；圆包菜→圆包菜去老叶、鸡腿XL→圣农鸡腿XL仍可过。
            if ratio >= 0.55:
                scored.append((ratio, float(r.get('单价') or 0), r.to_dict()))
    if scored:
        scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return scored[0][2], '模糊匹配'
    return None, '未匹配'


@app.post('/suppliers/price_preview')
def supplier_price_preview():
    old = request.form.get('old','')
    new_supplier = request.form.get('new_supplier','').strip()
    text = request.form.get('price_text','')
    parsed = parse_supplier_price_text(text)
    preview = []
    for r in parsed:
        hit, mtype = match_library_for_supplier_price(r['品名'], old)
        if hit and r.get('新价') != '':
            old_price = float(hit.get('单价') or 0)
            new_price = float(r['新价'])
            diff = new_price - old_price
            status = '✅ 可替换' if abs(diff) > 1e-9 else '⚠️ 价格相同'
            preview.append({**r, '状态': status, '匹配方式': mtype, '库ID': hit['id'], '库品名': hit['品名'], '旧规格': hit.get('规格',''), '旧单位': hit.get('单位',''), '旧价': old_price, '差额': round(diff,2)})
        else:
            preview.append({**r, '状态': '❌ 未匹配', '匹配方式': mtype, '库ID': '', '库品名': '', '旧规格': '', '旧单位': '', '旧价': '', '差额': ''})
    token = datetime.now().strftime('%Y%m%d%H%M%S')
    tmp = DATA_DIR / f'supplier_price_preview_{token}.csv'
    pd.DataFrame(preview).to_csv(tmp, index=False, encoding='utf-8-sig')
    rows = preview
    content = render_template_string("""
<div class='card'><h3>供应商价格替换预览</h3><p>更新供货商：<b>{{old}}</b>{% if new_supplier %} → 新名称：<b>{{new_supplier}}</b>{% endif %}</p><p class='muted'>请先看匹配是否对。确认后，系统会把匹配到的 SKU 改成新供应商名称 + 新价格；未匹配的不会乱改。</p>
<form method='post' action='/suppliers/price_apply'><input type='hidden' name='token' value='{{token}}'><input type='hidden' name='new_supplier' value='{{new_supplier}}'><button class='btn green'>确认批量更新价格</button> <a class='btn light' href='/suppliers'>返回重填</a></form></div>
<div class='card'><div class='scroll'><table><tr><th>状态</th><th>新报价原文</th><th>新报价品名</th><th>匹配库品名</th><th>单位</th><th>旧价</th><th>新价</th><th>差额</th><th>库ID</th></tr>{% for r in rows %}<tr><td>{{r['状态']}}</td><td>{{r['原文']}}</td><td>{{r['品名']}}</td><td>{{r['库品名']}}</td><td>{{r['单位'] or r['旧单位']}}</td><td>{{r['旧价']}}</td><td>{{r['新价']}}</td><td>{{r['差额']}}</td><td>{{r['库ID']}}</td></tr>{% endfor %}</table></div></div>
""", rows=rows, old=old, new_supplier=new_supplier, token=token)
    return page(content)


@app.post('/suppliers/price_apply')
def supplier_price_apply():
    token = request.form.get('token','')
    new_supplier = request.form.get('new_supplier','').strip()
    tmp = DATA_DIR / f'supplier_price_preview_{token}.csv'
    if not tmp.exists():
        return page("<div class='card'><h3>预览已失效，请返回重新粘贴报价。</h3><a class='btn' href='/suppliers'>返回</a></div>")
    preview = pd.read_csv(tmp).fillna('')
    df = library_df()
    changed = []
    for _, r in preview.iterrows():
        if '可替换' not in str(r.get('状态','')) and '价格相同' not in str(r.get('状态','')):
            continue
        idv = str(r.get('库ID',''))
        idx = df.index[df['id'].astype(str) == idv]
        if not len(idx):
            continue
        i = idx[0]
        old_row = df.loc[i].to_dict()
        df.at[i, '单价'] = float(r['新价'])
        if new_supplier:
            df.at[i, '供应商'] = new_supplier
        if str(r.get('单位','')).strip():
            df.at[i, '单位'] = str(r.get('单位')).strip()
        if str(r.get('规格','')).strip():
            df.at[i, '规格'] = str(r.get('规格')).strip()
        df.at[i, '启用'] = '是'
        df.at[i, '更新时间'] = today_str()
        changed.append({'id': idv, '品名': df.at[i,'品名'], '旧价': old_row.get('单价'), '新价': float(r['新价']), '旧供应商': old_row.get('供应商'), '新供应商': df.at[i,'供应商']})
    write_csv(LIB_CSV, df[LIB_COLS])
    log_change('供应商价格批量替换', json.dumps({'新供应商': new_supplier, '修改条数': len(changed), '明细': changed[:80]}, ensure_ascii=False))
    content = render_template_string("""
<div class='card'><h3>✅ 已批量更新供货商价格</h3><p>成功修改 <b>{{changed|length}}</b> 条。未匹配项没有改。</p><a class='btn' href='/library'>去成本库查看</a> <a class='btn light' href='/suppliers'>继续替换</a></div>
<div class='card'><table><tr><th>ID</th><th>品名</th><th>旧价</th><th>新价</th><th>旧供应商</th><th>新供应商</th></tr>{% for r in changed %}<tr><td>{{r['id']}}</td><td>{{r['品名']}}</td><td>{{r['旧价']}}</td><td>{{r['新价']}}</td><td>{{r['旧供应商']}}</td><td>{{r['新供应商']}}</td></tr>{% endfor %}</table></div>
""", changed=changed)
    return page(content)


@app.post('/suppliers/new_price_preview')
def supplier_new_price_preview():
    supplier = request.form.get('supplier','').strip()
    category = request.form.get('category','').strip() or '蔬菜'
    text = request.form.get('price_text','')
    parsed = parse_supplier_price_text(text)
    preview = []
    for r in parsed:
        # 新供货商新增报价：全库匹配，用于判断“已有项改价”还是“新项新增”。
        hit, mtype = match_library_for_supplier_price(r['品名'], '')
        if hit and r.get('新价') != '':
            old_price = float(hit.get('单价') or 0)
            new_price = float(r['新价'])
            preview.append({**r, '状态': '🔄 已存在，可更新价格', '匹配方式': mtype, '库ID': hit['id'], '库品名': hit['品名'], '旧规格': hit.get('规格',''), '旧单位': hit.get('单位',''), '旧价': old_price, '差额': round(new_price-old_price,2)})
        elif r.get('新价') != '':
            preview.append({**r, '状态': '🆕 新品项，可新增', '匹配方式': '新增', '库ID': '', '库品名': r.get('品名',''), '旧规格': '', '旧单位': r.get('单位',''), '旧价': '', '差额': ''})
        else:
            preview.append({**r, '状态': '❌ 未识别单价', '匹配方式': '失败', '库ID': '', '库品名': '', '旧规格': '', '旧单位': '', '旧价': '', '差额': ''})
    token = datetime.now().strftime('%Y%m%d%H%M%S')
    tmp = DATA_DIR / f'new_supplier_price_preview_{token}.csv'
    pd.DataFrame(preview).to_csv(tmp, index=False, encoding='utf-8-sig')
    content = render_template_string("""
<div class='card'><h3>新增供货商报价预览</h3><p>供货商：<b>{{supplier}}</b>；类别：<b>{{category}}</b></p><p class='muted'>请先核对。确认后：已存在的品项会实时改价；不存在的品项会批量新增到成本库；供货商也会自动加入供货商名录。</p>
<form method='post' action='/suppliers/new_price_apply'><input type='hidden' name='token' value='{{token}}'><input type='hidden' name='supplier' value='{{supplier}}'><input type='hidden' name='category' value='{{category}}'><button class='btn green'>确认批量新增/更新价格</button> <a class='btn light' href='/suppliers'>返回重填</a></form></div>
<div class='card'><div class='scroll'><table><tr><th>状态</th><th>报价原文</th><th>品名</th><th>匹配/新增品名</th><th>规格</th><th>单位</th><th>旧价</th><th>新价</th><th>差额</th><th>库ID</th></tr>{% for r in rows %}<tr><td>{{r['状态']}}</td><td>{{r['原文']}}</td><td>{{r['品名']}}</td><td>{{r['库品名']}}</td><td>{{r['规格'] or r['旧规格']}}</td><td>{{r['单位'] or r['旧单位']}}</td><td>{{r['旧价']}}</td><td>{{r['新价']}}</td><td>{{r['差额']}}</td><td>{{r['库ID']}}</td></tr>{% endfor %}</table></div></div>
""", rows=preview, supplier=supplier, category=category, token=token)
    return page(content)


@app.post('/suppliers/new_price_apply')
def supplier_new_price_apply():
    token = request.form.get('token','')
    supplier = request.form.get('supplier','').strip()
    category = request.form.get('category','').strip() or '蔬菜'
    tmp = DATA_DIR / f'new_supplier_price_preview_{token}.csv'
    if not tmp.exists():
        return page("<div class='card'><h3>预览已失效，请返回重新粘贴报价。</h3><a class='btn' href='/suppliers'>返回</a></div>")
    preview = pd.read_csv(tmp).fillna('')
    df = library_df()
    changed, added = [], []
    next_id_num = int(pd.to_numeric(df['id'], errors='coerce').max() or 0) + 1
    for _, r in preview.iterrows():
        status = str(r.get('状态',''))
        if r.get('新价') == '' or '未识别' in status:
            continue
        if str(r.get('库ID','')).strip():
            idv = str(r.get('库ID',''))
            idx = df.index[df['id'].astype(str) == idv]
            if not len(idx):
                continue
            i = idx[0]
            old_row = df.loc[i].to_dict()
            df.at[i, '单价'] = float(r['新价'])
            df.at[i, '供应商'] = supplier or df.at[i, '供应商']
            df.at[i, '品类'] = category or df.at[i, '品类']
            if str(r.get('单位','')).strip():
                df.at[i, '单位'] = str(r.get('单位')).strip()
            if str(r.get('规格','')).strip():
                df.at[i, '规格'] = str(r.get('规格')).strip()
            df.at[i, '启用'] = '是'
            df.at[i, '更新时间'] = today_str()
            changed.append({'id': idv, '品名': df.at[i,'品名'], '旧价': old_row.get('单价'), '新价': float(r['新价'])})
        else:
            row = {c: '' for c in LIB_COLS}
            row.update({'id': str(next_id_num), 'sheet': '新增报价', '品类': category, '品名': str(r.get('品名','')).strip(), '规格': str(r.get('规格','')).strip(), '单位': str(r.get('单位','')).strip(), '单价': float(r['新价']), '供应商': supplier, '启用': '是', '来源': '新供货商批量新增', '更新时间': today_str()})
            next_id_num += 1
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            added.append({'id': row['id'], '品名': row['品名'], '单价': row['单价'], '供应商': supplier})
    write_csv(LIB_CSV, df[LIB_COLS])
    # 同步供货商名录
    if supplier:
        sdf = suppliers_df()
        exists = ((sdf['类别'].astype(str) == category) & (sdf['供货商'].astype(str) == supplier)).any() if not sdf.empty else False
        if not exists:
            sid = str((pd.to_numeric(sdf['id'], errors='coerce').max() or 0) + 1)
            srow = {'id': sid, '类别': category, '供货商': supplier, '联系人': '', '电话': '', '地址': '', '备注': '批量新增报价时自动加入', '启用': '是', '更新时间': today_str()}
            write_csv(SUPPLIERS_CSV, pd.concat([sdf, pd.DataFrame([srow])], ignore_index=True)[SUPPLIER_COLS])
    log_change('新供货商报价批量新增/更新', json.dumps({'供货商': supplier, '类别': category, '新增条数': len(added), '改价条数': len(changed), '新增': added[:80], '改价': changed[:80]}, ensure_ascii=False))
    content = render_template_string("""
<div class='card'><h3>✅ 新供货商报价已实时写入</h3><p>新增 <b>{{added|length}}</b> 条；更新价格 <b>{{changed|length}}</b> 条。</p><a class='btn' href='/library'>去成本库查看</a> <a class='btn light' href='/suppliers'>继续新增/改价</a></div>
<div class='grid'><div class='card'><h3>新增品项</h3><table><tr><th>ID</th><th>品名</th><th>单价</th><th>供货商</th></tr>{% for r in added %}<tr><td>{{r['id']}}</td><td>{{r['品名']}}</td><td>{{r['单价']}}</td><td>{{r['供应商']}}</td></tr>{% endfor %}</table></div><div class='card'><h3>改价品项</h3><table><tr><th>ID</th><th>品名</th><th>旧价</th><th>新价</th></tr>{% for r in changed %}<tr><td>{{r['id']}}</td><td>{{r['品名']}}</td><td>{{r['旧价']}}</td><td>{{r['新价']}}</td></tr>{% endfor %}</table></div></div>
""", added=added, changed=changed)
    return page(content)


@app.post("/suppliers/replace")
def suppliers_replace():
    old = request.form.get("old", "")
    new = request.form.get("new", "")
    mode = request.form.get("mode", "replace")
    df = library_df()
    mask = df["供应商"].astype(str) == old
    count = int(mask.sum())
    if mode == "disable":
        df.loc[mask, "启用"] = "否"
        action = "停用供应商报价"
    else:
        df.loc[mask, "供应商"] = new
        df.loc[mask, "启用"] = "是"
        action = "一键替换供应商"
    df.loc[mask, "更新时间"] = today_str()
    write_csv(LIB_CSV, df[LIB_COLS])
    log_change(action, json.dumps({"旧供应商": old, "新供应商": new, "数量": count, "模式": mode}, ensure_ascii=False))
    return redirect("/suppliers")


@app.route("/logs")
def logs():
    df = read_csv(LOG_CSV, ["时间", "动作", "详情"]).tail(200)
    rows = df.iloc[::-1].to_dict("records")
    content = render_template_string("<div class='card'><h3>修改记录</h3><table><tr><th>时间</th><th>动作</th><th>详情</th></tr>{% for r in rows %}<tr><td>{{r['时间']}}</td><td>{{r['动作']}}</td><td class='small'>{{r['详情']}}</td></tr>{% endfor %}</table></div>", rows=rows)
    return page(content)


@app.route("/api/health")
def health():
    return jsonify({"ok": True, "sku": len(library_df()), "db": "supabase" if USE_SUPABASE else "csv"})


@app.post("/api/bootstrap_supabase")
def bootstrap_supabase():
    if not USE_SUPABASE:
        return jsonify({"ok": False, "error": "SUPABASE_URL/SUPABASE_SERVICE_KEY not configured"}), 400
    result = {}
    for path, cols in [(LIB_CSV, LIB_COLS), (ORDERS_CSV, ORDER_COLS), (SALES_CSV, ["日期", "营业额"]), (LOG_CSV, ["时间", "动作", "详情"]), (SUPPLIERS_CSV, SUPPLIER_COLS)]:
        table = TABLE_MAP[Path(path).name]
        df = pd.read_csv(path, dtype={"id": str}).fillna("") if path.exists() else pd.DataFrame(columns=cols)
        write_csv(path, df[cols] if all(c in df.columns for c in cols) else df)
        result[table] = len(df)
    return jsonify({"ok": True, "result": result})


if __name__ == "__main__":
    init_library()
    app.run(host="0.0.0.0", port=8088, debug=False)



