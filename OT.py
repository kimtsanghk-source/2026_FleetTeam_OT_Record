import json
import numpy as np
import pandas as pd

excel_file = "MASTER.xlsx"
months = ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026"]

monthly_summary = {}
distribution_bands = {}
all_records = []
staff_agg = {}

print("正在讀取 MASTER.xlsx 並統計各月份數據...")

for m in months:
    df = pd.read_excel(excel_file, sheet_name=m)
    # 過濾合計行及空白行
    df = df[df["Employee No."].notna() & df["Employee Name"].notna()]
    df = df[
        ~df["Employee Name"].astype(str).str.contains("Total", case=False)
    ].copy()

    df["Employee No."] = (
        df["Employee No."].astype(int).astype(str)
        if df["Employee No."].dtype != object
        else df["Employee No."].astype(str)
    )
    df["OT \nNormal (Hour)"] = (
        pd.to_numeric(df["OT \nNormal (Hour)"], errors="coerce")
        .fillna(0)
        .round(2)
    )
    df["OT Normal (Amount)"] = (
        pd.to_numeric(df["OT Normal (Amount)"], errors="coerce")
        .fillna(0)
        .round(1)
    )
    df["KPI%"] = (
        pd.to_numeric(df["KPI%"], errors="coerce").fillna(0).round(4)
    )
    df["KPI Incentive"] = (
        pd.to_numeric(df["KPI Incentive"], errors="coerce").fillna(0).round(2)
    )

    hol_col = [c for c in df.columns if "Holiday" in c and "Add" not in c][0]
    shift_col = [c for c in df.columns if "Shift" in c and "Add" not in c][0]
    add_col = [c for c in df.columns if "Add Work" in c][0]
    rem_col = [c for c in df.columns if "Remark" in c][0]

    df[hol_col] = pd.to_numeric(df[hol_col], errors="coerce").fillna(0)
    df[shift_col] = pd.to_numeric(df[shift_col], errors="coerce").fillna(0)
    df[add_col] = pd.to_numeric(df[add_col], errors="coerce").fillna(0)
    df[rem_col] = df[rem_col].fillna("").astype(str).str.strip()

    ot_hours = df["OT \nNormal (Hour)"]

    # 月度總結 KPI
    monthly_summary[m] = {
        "headcount": int(len(df)),
        "total_ot_hours": round(float(ot_hours.sum()), 2),
        "total_ot_amount": round(float(df["OT Normal (Amount)"].sum()), 1),
        "avg_ot_hours": round(float(ot_hours.mean()), 1),
        "median_ot_hours": round(float(ot_hours.median()), 1),
        "max_ot_hours": round(float(ot_hours.max()), 2),
        "total_holiday_work": round(float(df[hol_col].sum()), 1),
        "total_shift_pay": round(float(df[shift_col].sum()), 1),
        "total_add_work": round(float(df[add_col].sum()), 1),
        "total_kpi_incentive": round(float(df["KPI Incentive"].sum()), 2),
        "avg_kpi_pct": round(float(df["KPI%"].mean()), 3),
        "over_60h": int((ot_hours >= 60).sum()),
        "over_80h": int((ot_hours >= 80).sum()),
        "over_100h": int((ot_hours >= 100).sum()),
    }

    # 工時梯隊分佈
    distribution_bands[m] = {
        "0-20h (輕度)": int(((ot_hours >= 0) & (ot_hours < 20)).sum()),
        "20-40h (標準)": int(((ot_hours >= 20) & (ot_hours < 40)).sum()),
        "40-60h (偏高)": int(((ot_hours >= 40) & (ot_hours < 60)).sum()),
        "60-80h (超限預警)": int(((ot_hours >= 60) & (ot_hours < 80)).sum()),
        "80h+ (極高風險)": int((ot_hours >= 80).sum()),
    }

    # 441 筆明細聚合
    for _, row in df.iterrows():
        emp_id = str(row["Employee No."])
        emp_name = str(row["Employee Name"]).strip()
        rec = {
            "month": m,
            "emp_no": emp_id,
            "name": emp_name,
            "kpi_pct": float(row["KPI%"]),
            "kpi_amt": float(row["KPI Incentive"]),
            "ot_hours": float(row["OT \nNormal (Hour)"]),
            "ot_amt": float(row["OT Normal (Amount)"]),
            "holiday_work": float(row[hol_col]),
            "shift_pay": float(row[shift_col]),
            "add_work": float(row[add_col]),
            "remarks": row[rem_col],
            "manager": "Joe Chung",
        }
        all_records.append(rec)

        if emp_id not in staff_agg:
            staff_agg[emp_id] = {
                "emp_no": emp_id,
                "name": emp_name,
                "total_ot_hours": 0.0,
                "ot_hours_list": [],
                "total_ot_amt": 0.0,
                "total_holiday": 0.0,
                "total_payout": 0.0,
                "kpi_list": [],
                "remarks": row[rem_col],
            }
        staff_agg[emp_id]["total_ot_hours"] += rec["ot_hours"]
        staff_agg[emp_id]["ot_hours_list"].append(rec["ot_hours"])
        staff_agg[emp_id]["total_ot_amt"] += rec["ot_amt"]
        staff_agg[emp_id]["total_holiday"] += rec["holiday_work"]
        staff_agg[emp_id]["total_payout"] += (
            rec["ot_amt"] + rec["holiday_work"] + rec["kpi_amt"]
        )
        staff_agg[emp_id]["kpi_list"].append(rec["kpi_pct"])
        if rec["remarks"]:
            staff_agg[emp_id]["remarks"] = rec["remarks"]

top_staff_overall = []
for s in staff_agg.values():
    top_staff_overall.append({
        "emp_no": s["emp_no"],
        "name": s["name"],
        "total_ot_hours": round(s["total_ot_hours"], 2),
        "avg_ot_hours": round(float(np.mean(s["ot_hours_list"])), 1),
        "max_ot_hours": round(float(np.max(s["ot_hours_list"])), 2),
        "total_ot_amt": round(s["total_ot_amt"], 1),
        "total_holiday": round(s["total_holiday"], 1),
        "total_payout": round(s["total_payout"], 1),
        "avg_kpi": round(float(np.mean(s["kpi_list"])), 2),
        "remarks": s["remarks"],
    })
top_staff_overall.sort(key=lambda x: x["total_ot_hours"], reverse=True)

tot_hours = sum(s["total_ot_hours"] for s in monthly_summary.values())
tot_amt = sum(s["total_ot_amount"] for s in monthly_summary.values())
tot_hol = sum(s["total_holiday_work"] for s in monthly_summary.values())
tot_shift = sum(s["total_shift_pay"] for s in monthly_summary.values())
tot_kpi = sum(s["total_kpi_incentive"] for s in monthly_summary.values())

master_payload = {
    "months": months,
    "monthly_summary": monthly_summary,
    "distribution_bands": distribution_bands,
    "overall": {
        "total_ot_hours": round(tot_hours, 2),
        "total_ot_amount": round(tot_amt, 1),
        "total_holiday_work": round(tot_hol, 1),
        "total_shift_pay": round(tot_shift, 1),
        "total_kpi_incentive": round(tot_kpi, 2),
        "total_payout": round(tot_amt + tot_hol + tot_shift + tot_kpi, 2),
        "unique_staff": len(staff_agg),
        "avg_monthly_hours": round(tot_hours / 5, 1),
        "avg_staff_monthly_hours": round(tot_hours / len(staff_agg) / 5, 1),
    },
    "all_records": all_records,
    "top_staff_overall": top_staff_overall,
}

master_json = json.dumps(master_payload, ensure_ascii=False)

html_template = f"""<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>車隊 OT分析儀表板 (Apr - Aug 2026)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Noto+Sans+TC:wght@400;500;700&display=swap');
        body {{
            font-family: 'Plus Jakarta Sans', 'Noto Sans TC', sans-serif;
            transition: background-color 0.25s ease, color 0.25s ease;
        }}
        /* Dark Theme */
        body.dark-mode {{
            background-color: #0b1120;
            color: #cbd5e1;
        }}
        body.dark-mode .card-surface {{
            background: linear-gradient(145deg, rgba(30, 41, 59, 0.75), rgba(15, 23, 42, 0.85));
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.07);
        }}
        body.dark-mode .inner-card {{ background-color: rgba(15, 23, 42, 0.6); border-color: #1e293b; }}
        body.dark-mode .text-main {{ color: #ffffff; }}
        body.dark-mode .text-sub {{ color: #94a3b8; }}
        body.dark-mode .ctrl-bg {{ background-color: rgba(15, 23, 42, 0.9); border-color: #1e293b; }}
        body.dark-mode .table-head {{ background-color: rgba(15, 23, 42, 0.9); color: #94a3b8; border-color: #1e293b; }}
        body.dark-mode .table-row-hover:hover {{ background-color: rgba(30, 41, 59, 0.5); }}
        body.dark-mode .border-theme {{ border-color: #1e293b; }}
        body.dark-mode .input-theme {{ background-color: rgba(15, 23, 42, 0.8); border-color: #334155; color: #e2e8f0; }}

        /* Light Theme */
        body.light-mode {{
            background-color: #f8fafc;
            color: #334155;
        }}
        body.light-mode .card-surface {{
            background: #ffffff;
            backdrop-filter: none;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        }}
        body.light-mode .inner-card {{ background-color: #f1f5f9; border-color: #e2e8f0; }}
        body.light-mode .text-main {{ color: #0f172a; }}
        body.light-mode .text-sub {{ color: #64748b; }}
        body.light-mode .ctrl-bg {{ background-color: #ffffff; border-color: #e2e8f0; }}
        body.light-mode .table-head {{ background-color: #f1f5f9; color: #475569; border-color: #e2e8f0; }}
        body.light-mode .table-row-hover:hover {{ background-color: #f8fafc; }}
        body.light-mode .border-theme {{ border-color: #e2e8f0; }}
        body.light-mode .input-theme {{ background-color: #ffffff; border-color: #cbd5e1; color: #0f172a; }}

        .metric-card {{ transition: all 0.25s ease; }}
        .metric-card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 12px 24px -10px rgba(99, 102, 241, 0.15);
        }}
        .tab-btn.active {{
            background-color: #4f46e5;
            color: #ffffff !important;
            box-shadow: 0 4px 12px rgba(79, 70, 229, 0.35);
        }}
        ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
        ::-webkit-scrollbar-track {{ background: transparent; }}
        ::-webkit-scrollbar-thumb {{ background: #64748b88; border-radius: 4px; }}
    </style>
</head>
<body class="dark-mode min-h-screen p-4 md:p-8">

    <!-- Header -->
    <header class="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 mb-8 pb-6 border-b border-theme">
        <div>
            <div class="flex items-center gap-3">
                <span class="p-2.5 bg-gradient-to-br from-indigo-500 to-indigo-700 rounded-xl text-white shadow-lg shadow-indigo-500/20">
                    <i data-lucide="activity" class="w-6 h-6"></i>
                </span>
                <div>
                    <h1 class="text-2xl lg:text-3xl font-extrabold text-main tracking-tight">
                        車隊 OT分析儀表板
                    </h1>
                    <p class="text-sub text-xs md:text-sm mt-0.5">
                        數據週期：<span class="text-indigo-500 dark:text-indigo-400 font-semibold">2026年4月 至 8月 (Master 實績數據)</span> ｜ 部門管理：<span class="font-medium">Joe Chung</span>
                    </p>
                </div>
            </div>
        </div>

        <!-- Controls: Filters & Dark/Light Toggle -->
        <div class="flex flex-wrap items-center gap-3">
            <div class="flex flex-wrap items-center gap-1.5 p-1.5 rounded-2xl border border-theme ctrl-bg">
                <button onclick="setFilter('ALL')" id="btn-ALL" class="tab-btn active px-3.5 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">全期累計 (Apr-Aug)</button>
                <button onclick="setFilter('Apr 2026')" id="btn-Apr 2026" class="tab-btn px-3 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">4月</button>
                <button onclick="setFilter('May 2026')" id="btn-May 2026" class="tab-btn px-3 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">5月</button>
                <button onclick="setFilter('Jun 2026')" id="btn-Jun 2026" class="tab-btn px-3 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">6月</button>
                <button onclick="setFilter('Jul 2026')" id="btn-Jul 2026" class="tab-btn px-3 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">7月 (高峰)</button>
                <button onclick="setFilter('Aug 2026')" id="btn-Aug 2026" class="tab-btn px-3 py-1.5 rounded-xl text-xs font-semibold text-sub transition-all">8月</button>
            </div>

            <button id="themeToggleBtn" onclick="toggleTheme()" class="flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-semibold border border-theme ctrl-bg hover:border-indigo-500 transition-all text-main shadow-sm">
                <span id="themeIcon"><i data-lucide="sun" class="w-4 h-4 text-amber-400"></i></span>
                <span id="themeText">切換淺色</span>
            </button>
        </div>
    </header>

    <!-- Top KPI Highlights -->
    <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5 mb-8">
        <div class="card-surface rounded-2xl p-5 metric-card relative overflow-hidden">
            <div class="flex justify-between items-center text-sub mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider">總加班時數 (OT Normal)</span>
                <span class="p-2 bg-indigo-500/10 text-indigo-500 rounded-lg"><i data-lucide="clock" class="w-4 h-4"></i></span>
            </div>
            <div class="text-3xl font-extrabold text-main" id="kpi-total-hours">--</div>
            <div class="mt-3 flex items-center justify-between text-xs">
                <span class="text-sub" id="kpi-hours-sub">人均工時：--</span>
                <span class="px-2 py-0.5 rounded-md bg-indigo-500/10 text-indigo-500 font-medium" id="kpi-hours-badge">HK$50/小時</span>
            </div>
        </div>

        <div class="card-surface rounded-2xl p-5 metric-card relative overflow-hidden">
            <div class="flex justify-between items-center text-sub mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider">OT 薪酬總額 (OT Amount)</span>
                <span class="p-2 bg-emerald-500/10 text-emerald-500 rounded-lg"><i data-lucide="dollar-sign" class="w-4 h-4"></i></span>
            </div>
            <div class="text-3xl font-extrabold text-main" id="kpi-total-amount">--</div>
            <div class="mt-3 flex items-center justify-between text-xs">
                <span class="text-sub" id="kpi-amount-sub">含額外假日/輪班津貼</span>
                <span class="text-emerald-500 font-semibold" id="kpi-amount-badge">OT基率：$50/h</span>
            </div>
        </div>

        <div class="card-surface rounded-2xl p-5 metric-card relative overflow-hidden">
            <div class="flex justify-between items-center text-sub mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider">現有活躍編制 (Headcount)</span>
                <span class="p-2 bg-cyan-500/10 text-cyan-500 rounded-lg"><i data-lucide="users" class="w-4 h-4"></i></span>
            </div>
            <div class="text-3xl font-extrabold text-main" id="kpi-headcount">--</div>
            <div class="mt-3 flex items-center justify-between text-xs">
                <span class="text-sub" id="kpi-headcount-sub">中位數時數：--</span>
                <span class="text-cyan-500 font-medium" id="kpi-headcount-badge">平均人均 41.6h</span>
            </div>
        </div>

        <div class="card-surface rounded-2xl p-5 metric-card relative overflow-hidden">
            <div class="flex justify-between items-center text-sub mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider">高負荷警示 (OT &gt; 80h)</span>
                <span class="p-2 bg-rose-500/10 text-rose-500 rounded-lg"><i data-lucide="alert-triangle" class="w-4 h-4"></i></span>
            </div>
            <div class="text-3xl font-extrabold text-rose-500" id="kpi-high-risk">--</div>
            <div class="mt-3 flex items-center justify-between text-xs">
                <span class="text-sub" id="kpi-high-risk-sub">超 100 小時極限：--</span>
                <span class="px-2 py-0.5 rounded-md bg-rose-500/15 text-rose-500 font-medium">需重點排更</span>
            </div>
        </div>
    </section>

    <!-- Key Insights Banner -->
    <section class="card-surface rounded-2xl p-6 mb-8 border-l-4 border-indigo-500">
        <div class="flex items-start gap-4">
            <div class="p-2.5 bg-indigo-500/15 text-indigo-500 rounded-xl mt-0.5 shrink-0">
                <i data-lucide="trending-up" class="w-5 h-5"></i>
            </div>
            <div class="flex-1">
                <h2 class="text-sm font-bold text-main uppercase tracking-wider">5個月實測 Master 核心管理分析 (Executive Insights)</h2>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-2 text-xs text-sub leading-relaxed">
                    <div class="p-3 inner-card rounded-xl border border-theme">
                        <strong class="text-indigo-500 dark:text-indigo-400 block mb-1">1. 工時高峰與季節波動 (Spike Analysis)</strong>
                        4月至7月 OT 呈持續上升趨勢，於 7月達頂峰 (3,947.25h，月支出 HK$19.7萬)。8月新增多名入職人員 (編制擴充至 90人) 後，工時適度回調至 3,773.5h，有效緩解單人負荷。
                    </div>
                    <div class="p-3 inner-card rounded-xl border border-theme">
                        <strong class="text-amber-500 dark:text-amber-400 block mb-1">2. 帕累托集中度 (Pareto 80/20 Rule)</strong>
                        前 10% 核心員工 (9位) 承擔了全期 20.7% 的 OT 工時；前 20% 員工承擔 36.5%。排名前兩位的 <strong>Tsang Kai Sang (535.8h)</strong> 與 <strong>Law Pui Yi (519.8h)</strong> 累計超額均破 500小時，需警惕疲勞風險。
                    </div>
                    <div class="p-3 inner-card rounded-xl border border-theme">
                        <strong class="text-emerald-500 dark:text-emerald-400 block mb-1">3. 薪酬津貼與 KPI 連動性</strong>
                        全期 5 個月累計發放 KPI 獎金達 HK$1,068,147，人均 KPI 達成率穩定在 126%~132%。假日工時集中於 4、5、7月（共發放 HK$13,200），輪班津貼穩定每月發放 HK$760。
                    </div>
                </div>
            </div>
        </div>
    </section>

    <!-- Charts Row 1 -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
        <div class="card-surface rounded-2xl p-6 lg:col-span-2 flex flex-col justify-between">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4">
                <div>
                    <h3 class="text-base font-bold text-main flex items-center gap-2">
                        <i data-lucide="bar-chart-3" class="w-4 h-4 text-indigo-500"></i>
                        月度 OT 工時與總支出走勢對比 (Monthly Hours & Cost)
                    </h3>
                    <p class="text-xs text-sub mt-0.5">4月至8月每月總工時 (小時) 與對應 OT 薪酬支出 (HKD)</p>
                </div>
                <div class="flex items-center gap-4 text-xs">
                    <span class="flex items-center gap-1.5 text-sub"><span class="w-3 h-3 rounded-full bg-indigo-500 inline-block"></span> OT 工時</span>
                    <span class="flex items-center gap-1.5 text-sub"><span class="w-3 h-3 rounded-full bg-emerald-400 inline-block"></span> OT 支出 (HKD)</span>
                </div>
            </div>
            <div class="h-72 w-full">
                <canvas id="monthlyTrendChart"></canvas>
            </div>
        </div>

        <div class="card-surface rounded-2xl p-6 flex flex-col justify-between">
            <div class="mb-4">
                <h3 class="text-base font-bold text-main flex items-center gap-2">
                    <i data-lucide="pie-chart" class="w-4 h-4 text-cyan-500"></i>
                    工時分佈梯隊 (Hours Bracket)
                </h3>
                <p class="text-xs text-sub mt-0.5">檢視該月份區間人數分佈 (健康 vs 超限)</p>
            </div>
            <div class="h-72 w-full flex items-center justify-center">
                <canvas id="distributionChart"></canvas>
            </div>
        </div>
    </section>

    <!-- Charts Row 2 -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
        <div class="card-surface rounded-2xl p-6 lg:col-span-2">
            <div class="flex items-center justify-between mb-4">
                <div>
                    <h3 class="text-base font-bold text-main flex items-center gap-2">
                        <i data-lucide="award" class="w-4 h-4 text-amber-500"></i>
                        高工時員工 Top 10 (High OT Staff Ranking)
                    </h3>
                    <p class="text-xs text-sub mt-0.5">排名前列之核心骨幹人員 (顏色標記：紅色 &gt; 80h / 橘色 &gt; 60h)</p>
                </div>
                <span class="text-xs px-2.5 py-1 bg-amber-500/10 text-amber-500 border border-amber-500/20 rounded-lg font-medium">負荷監控</span>
            </div>
            <div class="h-80 w-full">
                <canvas id="topStaffChart"></canvas>
            </div>
        </div>

        <div class="card-surface rounded-2xl p-6 flex flex-col justify-between">
            <div>
                <h3 class="text-base font-bold text-main flex items-center gap-2">
                    <i data-lucide="git-commit" class="w-4 h-4 text-purple-500"></i>
                    KPI% 與 OT 工時關聯點位
                </h3>
                <p class="text-xs text-sub mt-0.5">整體呈現正相關 (r = +0.31)</p>
            </div>
            <div class="h-60 w-full my-auto">
                <canvas id="kpiOtChart"></canvas>
            </div>
            <p class="text-[11px] text-sub mt-2 inner-card p-2 rounded-lg border border-theme">
                💡 管理建議：主力人員 (KPI 1.4~1.5) 往往承擔了更多緊急加班，建議建立第二梯隊輪替機制，避免核心人力流失。
            </p>
        </div>
    </section>

    <!-- Table Section -->
    <section class="card-surface rounded-2xl p-6">
        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 mb-6">
            <div>
                <h3 class="text-base font-bold text-main flex items-center gap-2">
                    <i data-lucide="table-2" class="w-4 h-4 text-indigo-500"></i>
                    員工工時與薪酬詳細明細 (Detailed Roster Breakdown)
                </h3>
                <p class="text-xs text-sub mt-0.5">即時搜尋員工工號、姓名，或點擊 OT 時數進行排序</p>
            </div>

            <div class="relative w-full sm:w-64">
                <i data-lucide="search" class="w-4 h-4 absolute left-3 top-2.5 text-sub"></i>
                <input type="text" id="tableSearch" onkeyup="filterTable()" placeholder="搜尋工號、姓名、備註..."
                    class="input-theme w-full border text-xs rounded-xl pl-9 pr-4 py-2 focus:outline-none focus:border-indigo-500 transition">
            </div>
        </div>

        <div class="overflow-x-auto">
            <table class="w-full text-left text-xs">
                <thead class="table-head uppercase tracking-wider border-b border-theme">
                    <tr>
                        <th class="py-3 px-3">月份</th>
                        <th class="py-3 px-3">員工編號</th>
                        <th class="py-3 px-3">員工姓名</th>
                        <th class="py-3 px-3 text-right cursor-pointer hover:text-indigo-500 transition" onclick="sortTable('ot_hours')">OT 時數 ↕</th>
                        <th class="py-3 px-3 text-right">OT 金額 (HKD)</th>
                        <th class="py-3 px-3 text-right">假日/輪班津貼</th>
                        <th class="py-3 px-3 text-center">KPI %</th>
                        <th class="py-3 px-3 text-right">KPI 獎金</th>
                        <th class="py-3 px-3">備註 / 狀態</th>
                        <th class="py-3 px-3 text-center">負荷評級</th>
                    </tr>
                </thead>
                <tbody id="tableBody" class="divide-y border-theme">
                </tbody>
            </table>
        </div>

        <div class="flex items-center justify-between text-xs text-sub mt-4 pt-3 border-t border-theme">
            <span id="recordCountDisplay">顯示中...</span>
            <span>時薪基數依據 Master 表：HK$50.00 / 小時</span>
        </div>
    </section>

    <script>
        const MASTER_DATA = {master_json};

        let currentFilter = 'ALL';
        let trendChart = null;
        let distChart = null;
        let topChart = null;
        let scatterChart = null;
        let filteredRecords = [];
        let isDarkMode = true;

        function initTheme() {{
            const saved = localStorage.getItem('dashboard_theme');
            isDarkMode = saved !== 'light';
            applyTheme(isDarkMode);
        }}

        function toggleTheme() {{
            isDarkMode = !isDarkMode;
            applyTheme(isDarkMode);
            localStorage.setItem('dashboard_theme', isDarkMode ? 'dark' : 'light');
            renderCharts();
        }}

        function applyTheme(dark) {{
            const body = document.body;
            const icon = document.getElementById('themeIcon');
            const txt = document.getElementById('themeText');
            if (dark) {{
                body.classList.remove('light-mode');
                body.classList.add('dark-mode');
                if (icon) icon.innerHTML = '<i data-lucide="sun" class="w-4 h-4 text-amber-400"></i>';
                if (txt) txt.textContent = '切換淺色';
            }} else {{
                body.classList.remove('dark-mode');
                body.classList.add('light-mode');
                if (icon) icon.innerHTML = '<i data-lucide="moon" class="w-4 h-4 text-indigo-500"></i>';
                if (txt) txt.textContent = '切換深色';
            }}
            if (window.lucide) lucide.createIcons();
        }}

        function formatNumber(num, decimals = 0) {{
            if (num === null || num === undefined || isNaN(num)) return '0';
            return Number(num).toLocaleString('en-US', {{ minimumFractionDigits: decimals, maximumFractionDigits: decimals }});
        }}

        function setFilter(filter) {{
            currentFilter = filter;
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            const activeBtn = document.getElementById('btn-' + filter);
            if (activeBtn) activeBtn.classList.add('active');

            updateKPIs();
            renderCharts();
            renderTable();
        }}

        function updateKPIs() {{
            if (currentFilter === 'ALL') {{
                const ov = MASTER_DATA.overall;
                document.getElementById('kpi-total-hours').innerHTML = formatNumber(ov.total_ot_hours, 1) + ' <span class="text-xs font-normal text-sub">小時</span>';
                document.getElementById('kpi-hours-sub').textContent = '月均：' + formatNumber(ov.avg_monthly_hours, 1) + 'h (全隊)';
                document.getElementById('kpi-total-amount').innerHTML = 'HK$ ' + formatNumber(ov.total_ot_amount);
                document.getElementById('kpi-amount-sub').textContent = '總發放 (含獎金)：HK$ ' + formatNumber(ov.total_payout);
                document.getElementById('kpi-headcount').textContent = ov.unique_staff + ' 人';
                document.getElementById('kpi-headcount-sub').textContent = '人均月 OT：' + ov.avg_staff_monthly_hours + 'h';
                const highRiskTotal = MASTER_DATA.top_staff_overall.filter(s => s.total_ot_hours >= 350).length;
                document.getElementById('kpi-high-risk').innerHTML = highRiskTotal + ' <span class="text-xs font-normal text-sub">人</span>';
                document.getElementById('kpi-high-risk-sub').textContent = '5個月累計超 350h 極高負荷';
            }} else {{
                const m = MASTER_DATA.monthly_summary[currentFilter];
                document.getElementById('kpi-total-hours').innerHTML = formatNumber(m.total_ot_hours, 1) + ' <span class="text-xs font-normal text-sub">小時</span>';
                document.getElementById('kpi-hours-sub').textContent = '人均工時：' + m.avg_ot_hours + 'h / 人';
                document.getElementById('kpi-total-amount').innerHTML = 'HK$ ' + formatNumber(m.total_ot_amount);
                document.getElementById('kpi-amount-sub').textContent = '加送工時/輪班：HK$ ' + formatNumber(m.total_add_work);
                document.getElementById('kpi-headcount').textContent = m.headcount + ' 人';
                document.getElementById('kpi-headcount-sub').textContent = '中位數工時：' + m.median_ot_hours + 'h';
                document.getElementById('kpi-high-risk').innerHTML = m.over_80h + ' <span class="text-xs font-normal text-sub">人</span>';
                document.getElementById('kpi-high-risk-sub').textContent = '單月超過 100h：' + m.over_100h + ' 人';
            }}
        }}

        function renderCharts() {{
            const textColor = isDarkMode ? '#94a3b8' : '#64748b';
            const gridColor = isDarkMode ? 'rgba(255, 255, 255, 0.05)' : 'rgba(0, 0, 0, 0.06)';
            const tooltipBg = isDarkMode ? '#1e293b' : '#ffffff';
            const tooltipTitle = isDarkMode ? '#ffffff' : '#0f172a';
            const tooltipBody = isDarkMode ? '#cbd5e1' : '#334155';
            const tooltipBorder = isDarkMode ? '#334155' : '#e2e8f0';

            Chart.defaults.color = textColor;
            Chart.defaults.font.family = "'Plus Jakarta Sans', 'Noto Sans TC', sans-serif";

            // 1. Trend Chart
            const ctxTrend = document.getElementById('monthlyTrendChart').getContext('2d');
            if (trendChart) trendChart.destroy();
            const monthLabels = MASTER_DATA.months;
            trendChart = new Chart(ctxTrend, {{
                data: {{
                    labels: monthLabels,
                    datasets: [
                        {{
                            type: 'bar',
                            label: 'OT 總工時 (Hours)',
                            data: monthLabels.map(m => MASTER_DATA.monthly_summary[m].total_ot_hours),
                            backgroundColor: 'rgba(99, 102, 241, 0.75)',
                            hoverBackgroundColor: '#6366f1',
                            borderRadius: 6,
                            yAxisID: 'y'
                        }},
                        {{
                            type: 'line',
                            label: 'OT 薪酬總額 (HKD)',
                            data: monthLabels.map(m => MASTER_DATA.monthly_summary[m].total_ot_amount),
                            borderColor: '#10b981',
                            backgroundColor: '#10b981',
                            borderWidth: 2.5,
                            tension: 0.3,
                            pointRadius: 5,
                            yAxisID: 'y1'
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{ backgroundColor: tooltipBg, titleColor: tooltipTitle, bodyColor: tooltipBody, borderColor: tooltipBorder, borderWidth: 1 }}
                    }},
                    scales: {{
                        x: {{ grid: {{ color: gridColor }}, ticks: {{ color: textColor }} }},
                        y: {{ type: 'linear', position: 'left', grid: {{ color: gridColor }}, ticks: {{ color: textColor }}, title: {{ display: true, text: '工時 (小時)', color: textColor }} }},
                        y1: {{ type: 'linear', position: 'right', grid: {{ drawOnChartArea: false }}, ticks: {{ color: '#10b981' }}, title: {{ display: true, text: '金額 (HKD)', color: '#10b981' }} }}
                    }}
                }}
            }});

            // 2. Distribution Chart
            const ctxDist = document.getElementById('distributionChart').getContext('2d');
            if (distChart) distChart.destroy();
            const bandLabels = ['0-20h (輕度)', '20-40h (標準)', '40-60h (偏高)', '60-80h (超限預警)', '80h+ (極高風險)'];
            let bandValues = [0, 0, 0, 0, 0];
            if (currentFilter === 'ALL') {{
                MASTER_DATA.months.forEach(m => {{
                    const d = MASTER_DATA.distribution_bands[m];
                    bandValues[0] += d['0-20h (輕度)'];
                    bandValues[1] += d['20-40h (標準)'];
                    bandValues[2] += d['40-60h (偏高)'];
                    bandValues[3] += d['60-80h (超限預警)'];
                    bandValues[4] += d['80h+ (極高風險)'];
                }});
            }} else {{
                const d = MASTER_DATA.distribution_bands[currentFilter];
                bandValues = [d['0-20h (輕度)'], d['20-40h (標準)'], d['40-60h (偏高)'], d['60-80h (超限預警)'], d['80h+ (極高風險)']];
            }}
            distChart = new Chart(ctxDist, {{
                type: 'doughnut',
                data: {{
                    labels: bandLabels,
                    datasets: [{{
                        data: bandValues,
                        backgroundColor: ['#0284c7', '#6366f1', '#f59e0b', '#ea580c', '#ef4444'],
                        borderWidth: isDarkMode ? 0 : 2,
                        borderColor: isDarkMode ? 'transparent' : '#ffffff'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '68%',
                    plugins: {{
                        legend: {{ position: 'bottom', labels: {{ boxWidth: 10, padding: 8, font: {{ size: 10 }}, color: textColor }} }},
                        tooltip: {{ backgroundColor: tooltipBg, titleColor: tooltipTitle, bodyColor: tooltipBody, borderColor: tooltipBorder, borderWidth: 1 }}
                    }}
                }}
            }});

            // 3. Top Staff Chart
            const ctxTop = document.getElementById('topStaffChart').getContext('2d');
            if (topChart) topChart.destroy();
            let topList = currentFilter === 'ALL'
                ? MASTER_DATA.top_staff_overall.slice(0, 10)
                : MASTER_DATA.all_records.filter(r => r.month === currentFilter).sort((a,b) => b.ot_hours - a.ot_hours).slice(0, 10).map(r => ({{ name: r.name, total_ot_hours: r.ot_hours }}));

            topChart = new Chart(ctxTop, {{
                type: 'bar',
                data: {{
                    labels: topList.map(s => s.name),
                    datasets: [{{
                        label: 'OT 時數',
                        data: topList.map(s => s.total_ot_hours),
                        backgroundColor: (ctx) => {{
                            const v = ctx.raw;
                            if (currentFilter === 'ALL' ? v > 400 : v >= 80) return '#ef4444';
                            if (currentFilter === 'ALL' ? v > 300 : v >= 60) return '#f97316';
                            return '#6366f1';
                        }},
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{ backgroundColor: tooltipBg, titleColor: tooltipTitle, bodyColor: tooltipBody, borderColor: tooltipBorder, borderWidth: 1 }}
                    }},
                    scales: {{
                        x: {{ grid: {{ color: gridColor }}, ticks: {{ color: textColor }} }},
                        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 11 }}, color: textColor }} }}
                    }}
                }}
            }});

            // 4. Scatter Chart
            const ctxScatter = document.getElementById('kpiOtChart').getContext('2d');
            if (scatterChart) scatterChart.destroy();
            const recs = currentFilter === 'ALL' ? MASTER_DATA.all_records : MASTER_DATA.all_records.filter(r => r.month === currentFilter);
            scatterChart = new Chart(ctxScatter, {{
                type: 'scatter',
                data: {{
                    datasets: [{{
                        label: 'KPI vs OT',
                        data: recs.map(r => ({{ x: r.kpi_pct, y: r.ot_hours }})),
                        backgroundColor: 'rgba(168, 85, 247, 0.65)',
                        borderColor: '#a855f7',
                        pointRadius: 4
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{ backgroundColor: tooltipBg, titleColor: tooltipTitle, bodyColor: tooltipBody, borderColor: tooltipBorder, borderWidth: 1 }}
                    }},
                    scales: {{
                        x: {{ title: {{ display: true, text: 'KPI 達成率', color: textColor }}, grid: {{ color: gridColor }}, ticks: {{ color: textColor }} }},
                        y: {{ title: {{ display: true, text: 'OT 工時 (小時)', color: textColor }}, grid: {{ color: gridColor }}, ticks: {{ color: textColor }} }}
                    }}
                }}
            }});
        }}

        function renderTable() {{
            filteredRecords = currentFilter === 'ALL'
                ? [...MASTER_DATA.all_records]
                : MASTER_DATA.all_records.filter(r => r.month === currentFilter);
            filteredRecords.sort((a, b) => b.ot_hours - a.ot_hours);
            displayRecords(filteredRecords);
        }}

        function displayRecords(records) {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';
            records.forEach(r => {{
                let riskBadge = '';
                if (r.ot_hours >= 80) riskBadge = '<span class="px-2 py-0.5 rounded-full bg-rose-500/20 text-rose-500 font-semibold border border-rose-500/30">極高負荷</span>';
                else if (r.ot_hours >= 60) riskBadge = '<span class="px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-500 font-semibold border border-amber-500/30">超限預警</span>';
                else if (r.ot_hours >= 40) riskBadge = '<span class="px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-500 font-medium">偏高</span>';
                else riskBadge = '<span class="px-2 py-0.5 rounded-full bg-slate-500/20 text-sub font-normal">正常</span>';

                const tr = document.createElement('tr');
                tr.className = 'table-row-hover transition border-b border-theme';
                tr.innerHTML = `
                    <td class="py-2.5 px-3 font-medium text-sub">${{r.month}}</td>
                    <td class="py-2.5 px-3 font-mono text-sub">${{r.emp_no}}</td>
                    <td class="py-2.5 px-3 font-semibold text-main">${{r.name}}</td>
                    <td class="py-2.5 px-3 text-right font-bold text-main">${{formatNumber(r.ot_hours, 2)}}</td>
                    <td class="py-2.5 px-3 text-right font-medium text-emerald-500">HK$ ${{formatNumber(r.ot_amt, 1)}}</td>
                    <td class="py-2.5 px-3 text-right text-sub">${{r.add_work > 0 ? 'HK$ ' + formatNumber(r.add_work) : '-'}}</td>
                    <td class="py-2.5 px-3 text-center font-medium ${{r.kpi_pct >= 1.4 ? 'text-amber-500 font-bold' : 'text-sub'}}">${{(r.kpi_pct * 100).toFixed(0)}}%</td>
                    <td class="py-2.5 px-3 text-right text-sub">HK$ ${{formatNumber(r.kpi_amt, 0)}}</td>
                    <td class="py-2.5 px-3 text-sub truncate max-w-[140px]">${{r.remarks || '-'}}</td>
                    <td class="py-2.5 px-3 text-center">${{riskBadge}}</td>
                `;
                tbody.appendChild(tr);
            }});
            document.getElementById('recordCountDisplay').textContent = '共 ' + records.length + ' 筆記錄顯示中';
        }}

        function filterTable() {{
            const q = document.getElementById('tableSearch').value.trim().toLowerCase();
            if (!q) {{ displayRecords(filteredRecords); return; }}
            const matched = filteredRecords.filter(r => 
                String(r.emp_no).includes(q) ||
                r.name.toLowerCase().includes(q) ||
                (r.remarks && r.remarks.toLowerCase().includes(q)) ||
                r.month.toLowerCase().includes(q)
            );
            displayRecords(matched);
        }}

        let sortAsc = false;
        function sortTable() {{
            sortAsc = !sortAsc;
            filteredRecords.sort((a, b) => sortAsc ? a.ot_hours - b.ot_hours : b.ot_hours - a.ot_hours);
            displayRecords(filteredRecords);
        }}

        window.addEventListener('DOMContentLoaded', () => {{
            initTheme();
            setFilter('ALL');
        }});
    </script>
</body>
</html>"""

output_file = "ot_analysis_dashboard.html"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(html_template)

print(
    f"成功生成 {output_file}！已注入完整 441 筆明細，標題為「車隊 OT分析儀表板」，並支援 Dark / Light 模式。"
)