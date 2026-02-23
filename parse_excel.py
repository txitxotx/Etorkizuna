#!/usr/bin/env python3
"""
parse_excel.py  —  Portfolio Dashboard Data Extractor
Lee portfolio_cuadro_mandos.xlsx y genera public/data.json

USO:
    python parse_excel.py                    # Excel en la misma carpeta
    python parse_excel.py mi_cartera.xlsx    # ruta personalizada
"""
import json, sys, os
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl -q")
    import openpyxl

EXCEL_FILE  = sys.argv[1] if len(sys.argv) > 1 else "portfolio_cuadro_mandos.xlsx"
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public", "data.json")

SHEET_ASSETS   = "📋 ACTIVOS"
SHEET_INPUTS   = "⚙️ INPUTS"
SHEET_ANALISIS = "🔍 ANÁLISIS"

def to_float(v, default=0.0):
    if v is None or str(v).strip() in ("", "-", "—", "#N/A", "#REF!", "#VALUE!"):
        return default
    try:
        return float(str(v).replace(",",".").replace("€","").replace("%","").replace(" ","").strip())
    except:
        return default

def to_pct(v, default=0.0):
    f = to_float(v, default)
    return f / 100 if abs(f) > 1.5 else f

def to_str(v, default=""):
    return str(v).strip() if v is not None else default

def parse():
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: No se encuentra '{EXCEL_FILE}'")
        sys.exit(1)

    print(f"Leyendo {EXCEL_FILE}...")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws_act = wb[SHEET_ASSETS]
    ws_inp = wb[SHEET_INPUTS]
    ws_ana = wb[SHEET_ANALISIS]

    # ── Activos ────────────────────────────────────────────────────────────────
    assets = []
    for r in range(5, 30):
        row = [ws_act.cell(row=r, column=c).value for c in range(1, 15)]
        _, cat, name, titles, buy_px, invested, price_now, curr_val, gp, rt, ytd, mtd, rt21, weight = row
        if name is None:
            continue
        assets.append({
            "name": to_str(name), "cat": to_str(cat),
            "titles": to_float(titles), "buy_px": to_float(buy_px),
            "invested": to_float(invested), "price_now": to_float(price_now),
            "val": to_float(curr_val), "gp": to_float(gp),
            "rt": to_pct(rt), "ytd": to_pct(ytd),
            "mtd": to_pct(mtd), "rt21": to_pct(rt21),
            "weight": to_pct(weight),
        })

    # ── Inputs ─────────────────────────────────────────────────────────────────
    def inp_pct(row, col=2): return to_pct(ws_inp.cell(row=row, column=col).value)
    def inp_val(row, col=2): return to_float(ws_inp.cell(row=row, column=col).value)
    def inp_str(row, col=6): return to_str(ws_inp.cell(row=row, column=col).value)

    inputs = {
        "rf": inp_pct(5), "market_premium": inp_pct(6),
        "inflation": inp_pct(7), "tax_rate": inp_pct(8),
        "fee_rf": inp_pct(9), "fee_rv": inp_pct(10), "fee_cr": inp_pct(11),
        "target_return": inp_pct(12), "target_vol": inp_pct(13),
        "target_sharpe": inp_val(14),
        "update_date": inp_str(5, 6), "horizon_years": inp_val(6, 6),
        "rebalance_freq": inp_str(7, 6),
        "target_weight_rf": inp_pct(18), "target_weight_rv": inp_pct(19), "target_weight_cr": inp_pct(20),
        "exp_ret_rf": inp_pct(18, 6), "exp_ret_rv": inp_pct(19, 6), "exp_ret_cr": inp_pct(20, 6),
        "exp_vol_rf": inp_pct(18, 7), "exp_vol_rv": inp_pct(19, 7), "exp_vol_cr": inp_pct(20, 7),
        "exp_return_portfolio": inp_pct(25), "exp_vol_portfolio": inp_pct(26),
        "sharpe_portfolio": inp_val(27),
    }

    # ── Escenarios ─────────────────────────────────────────────────────────────
    scenario_labels = ["🟢 Favorable","⚪ Base","🟡 Corrección moderada","🔴 Mercado bajista","🚨 Crisis severa"]
    scenarios = []
    for i, r in enumerate(range(26, 31)):
        v = [ws_ana.cell(row=r, column=c).value for c in range(1, 8)]
        scenarios.append({
            "label": to_str(v[0]) or scenario_labels[i],
            "shock_rf": to_pct(v[1]), "shock_rv": to_pct(v[2]), "shock_cr": to_pct(v[3]),
            "impact": to_pct(v[4]), "val_est": to_float(v[5]), "loss_est": to_float(v[6]),
        })

    # ── Resumen ────────────────────────────────────────────────────────────────
    total_inv = sum(a["invested"] for a in assets)
    total_val = sum(a["val"] for a in assets)
    total_gp  = total_val - total_inv
    total_rt  = (total_gp / total_inv) if total_inv else 0

    def cat_data(code):
        grp = [a for a in assets if a["cat"] == code]
        inv = sum(a["invested"] for a in grp)
        val = sum(a["val"] for a in grp)
        gp  = val - inv
        return {
            "inv": round(inv,2), "val": round(val,2), "gp": round(gp,2),
            "rt":  round((gp/inv) if inv else 0, 6),
            "ytd": round(sum(a["ytd"]*a["val"] for a in grp)/val if val else 0, 6),
            "mtd": round(sum(a["mtd"]*a["val"] for a in grp)/val if val else 0, 6),
            "weight": round(val/total_val if total_val else 0, 6),
            "count": len(grp),
        }

    sorted_rt = sorted(assets, key=lambda a: a["rt"])
    output = {
        "generated": datetime.now().isoformat(),
        "source": os.path.basename(EXCEL_FILE),
        "assets": assets,
        "inputs": inputs,
        "summary": {
            "total_inv": round(total_inv,2), "total_val": round(total_val,2),
            "total_gp": round(total_gp,2), "total_rt": round(total_rt,6),
            "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "best_asset":  {"name": sorted_rt[-1]["name"] if assets else "", "rt": sorted_rt[-1]["rt"] if assets else 0},
            "worst_asset": {"name": sorted_rt[0]["name"]  if assets else "", "rt": sorted_rt[0]["rt"]  if assets else 0},
            "cats": {"RF": cat_data("RF"), "RV": cat_data("RV"), "CR": cat_data("CR")},
        },
        "scenarios": scenarios,
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    s = output["summary"]
    print(f"\n✅  data.json generado en: {OUTPUT_FILE}")
    print(f"   Activos:     {len(assets)}")
    print(f"   Total inv:   €{total_inv:>12,.2f}")
    print(f"   Valor actual:€{total_val:>12,.2f}")
    print(f"   G/P total:   €{total_gp:>+12,.2f}  ({total_rt*100:+.2f}%)")
    print(f"   RF: €{s['cats']['RF']['val']:>10,.0f} ({s['cats']['RF']['weight']*100:.1f}%)")
    print(f"   RV: €{s['cats']['RV']['val']:>10,.0f} ({s['cats']['RV']['weight']*100:.1f}%)")
    print(f"   CR: €{s['cats']['CR']['val']:>10,.0f} ({s['cats']['CR']['weight']*100:.1f}%)")
    print(f"   Rf: {inputs['rf']*100:.2f}%  |  Sharpe: {inputs['sharpe_portfolio']:.2f}\n")

if __name__ == "__main__":
    parse()
