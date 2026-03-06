import pandas as pd
import numpy as np
import os
import warnings
import gc
from datetime import datetime

# =====================================================
# 0. CLEAN SLATE, FOLDER SETUP & DATA LOADING
# =====================================================
# try:
#     del df_d, master_data
#     gc.collect()
# except NameError:
#     pass

timestamp = datetime.now().strftime("%d%b_%H%M")
output_folder = f"Output_{timestamp}"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

file_path = r"\\192.168.50.100\DataScience\Projects\5.Pilot5_Funds\Output Prathmesh\27-02-2026_v1\March_Data_05_03_2026_updated.xlsx"
df_raw = pd.read_excel(file_path, sheet_name='Sheet1')

# Data Cleaning & Type Casting
df_raw = df_raw[df_raw["MCAP"].notna()]
df_filtered = df_raw[~df_raw["Sector"].str.contains("Bank|Insurance", case=False, na=False)].copy()

df_d = df_filtered.rename(columns={'CAGR_2Y':'CAGR2Y', 'CAGR_3Y':'CAGR3Y'}).copy()
df_d.columns = df_d.columns.str.strip()

# CRITICAL: Convert Net Debt to numeric to prevent comparison errors
df_d["Net Debt"] = pd.to_numeric(df_d["Net Debt"], errors='coerce')

# =====================================================
# 1. GLOBAL SETTINGS
# =====================================================
warnings.filterwarnings('ignore')
date_col = next((c for c in df_d.columns if c.lower() == 'date'), None)
MONTH_NAME = pd.to_datetime(df_d[date_col], errors='coerce').dropna().iloc[0].strftime("%b") if date_col else "Analysis"

MCAP_OCF_COL = "MCAP/OCF"
REV_GROWTH_MIN, PAT_GROWTH_MIN = 0.15, 0.15
DEBT_LIMIT_SAFE = 0.3  # Standardized debt limit for Portfolio Picks
YEARS = [2021, 2022, 2023, 2024, 2025]
METRICS_TO_RUN = ["PS", "PB"]

MCAP_SCENARIOS = [{"min": 0, "max": 20}, {"min": 0, "max": 50}, {"min": -np.inf, "max": np.inf}]

# SYNCED DEBT SCENARIOS
DEBT_SCENARIOS = [
    {"label": "NetDebt_LT0",    "cond": lambda df: df["Net Debt"] < 0},
    # {"label": "NetDebt_0.3",   "cond": lambda df: df["Net Debt"] <= 0.3},
    # {"label": "NetDebt_LEneg0.3", "cond": lambda df: df["Net Debt"] <= -0.3}, # Fixed: Cash Rich
    {"label": "Debt_NoFilter",  "cond": lambda df: df["Company"].notna()}
]

VAL_MAPPING = {
    "Fwd_Return": {"ps": "Fwd_PS", "pb": "Fwd_PB", "mcap_ocf": "Fwd_MCAP_OCF"},
    "CAGR2Y": {"ps": "PS_Fwd_2Y", "pb": "PB_Fwd_2Y", "mcap_ocf": "MCAP_OCF_Fwd_2Y"},
    "CAGR3Y": {"ps": "PS_Fwd_3Y", "pb": "PB_Fwd_3Y", "mcap_ocf": "MCAP_OCF_Fwd_3Y"}
}

PERCENT_ROWS = {"Avg Ret","Min","Max","Avg -ve","Avg of +ve","%_ve","%+ve", "Avg Revenue Growth", "Avg PAT Growth"}
BASE_PARAMS_START = ["Count", "Avg Ret", "Count of SR>100"]
BASE_PARAMS_END = [
    "Min","Max","Count of -ve","Avg -ve","Count of +ve","Avg of +ve","%_ve","%+ve",
    "Avg Revenue Growth", "Avg PAT Growth",
    "Avg PS for +ve","Avg PS for -ve","Avg PB for +ve","Avg PB for -ve",
    "Avg MCAP/OCF for +ve","Avg MCAP/OCF for -ve",
    "Avg Fwd PS for +ve","Avg Fwd PS for -ve","Avg Fwd PB for +ve","Avg Fwd PB for -ve",
    "Avg Fwd MCAP/OCF for +ve","Avg Fwd MCAP/OCF for -ve"
]

# =====================================================
# 2. HELPER FUNCTIONS (compute_block & apply_formatting)
# =====================================================
def compute_block(pivot_df, return_type_name):
    if len(pivot_df) == 0: return [np.nan] * 27
    ret = pivot_df["Avg_Return"]
    visible = ret.notna()
    results = [len(pivot_df), ret[visible].mean(), ((ret > 1) & visible).sum()]
    if return_type_name == "Fwd_Return":
        results.extend([((pivot_df.get("CAGR2Y", 0) > 0.50) & visible).sum(), 
                        ((pivot_df.get("CAGR3Y", 0) > 0.25) & visible).sum()])
    else:
        results.append(((pivot_df.get("Avg_Rev_Gr_Fwd_1Y", 0) >= 0.15) & (pivot_df.get("Avg_PAT_Gr_Fwd_1Y", 0) >= 0.15)).sum())
        results.append("") 
    c_neg, c_pos = ((ret < 0) & visible).sum(), ((ret > 0) & visible).sum()
    denom = c_neg + c_pos
    results.extend([ret[visible].min(), ret[visible].max(), c_neg, ret[(ret < 0) & visible].mean(), c_pos, ret[(ret > 0) & visible].mean(),
                    c_neg / denom if denom > 0 else np.nan, c_pos / denom if denom > 0 else np.nan])
    results.append(pivot_df["Avg_Rev_Growth"].mean())
    results.append(pivot_df["Avg_PAT_Growth"].mean())
    for col in ["Avg_PS", "Avg_PB", "Avg_MCAP_OCF", "Avg_Fwd_PS", "Avg_Fwd_PB", "Avg_Fwd_MCAP_OCF"]:
        results.append(pivot_df.loc[(ret > 0) & visible, col].replace([np.inf, -np.inf], np.nan).mean())
        results.append(pivot_df.loc[(ret < 0) & visible, col].replace([np.inf, -np.inf], np.nan).mean())
    return results

def apply_formatting(df):
    out = df.astype(object)
    for r in out.index:
        for c in out.columns:
            v = out.loc[r, c]
            if pd.isna(v) or v == "" or v == np.inf or v == -np.inf: out.loc[r, c] = ""
            elif r in PERCENT_ROWS: 
                try: out.loc[r, c] = f"{float(v):.1%}"
                except: out.loc[r, c] = v
            else:
                try: out.loc[r, c] = f"{float(v):.2f}"
                except: out.loc[r, c] = v
    return out

# =====================================================
# 3. MAIN LOOP
# =====================================================
master_data = {"Fwd_Return": [], "CAGR2Y": [], "CAGR3Y": []}

for BIN_COL in METRICS_TO_RUN:
    BINS = {
        f"All_{BIN_COL}": lambda df, bc=BIN_COL: np.isfinite(df[bc]),
        "0_3": lambda df, bc=BIN_COL: (df[bc] >= 0) & (df[bc] < 3) & (np.isfinite(df[bc])),
        "0_4": lambda df, bc=BIN_COL: (df[bc] >= 0) & (df[bc] < 4) & (np.isfinite(df[bc])),
        "0_5": lambda df, bc=BIN_COL: (df[bc] >= 0) & (df[bc] < 5) & (np.isfinite(df[bc])),
        "0_6": lambda df, bc=BIN_COL: (df[bc] >= 0) & (df[bc] < 6) & (np.isfinite(df[bc])),
        "0_10": lambda df, bc=BIN_COL: (df[bc] >= 0) & (df[bc] < 10) & (np.isfinite(df[bc])),
        "GT10": lambda df, bc=BIN_COL: (df[bc] >= 10) & (np.isfinite(df[bc])),
        "GT15": lambda df, bc=BIN_COL: (df[bc] >= 15) & (np.isfinite(df[bc])),
        "6_10": lambda df, bc=BIN_COL: (df[bc] >= 6) & (df[bc] < 10) & (np.isfinite(df[bc])),
        "10_15": lambda df, bc=BIN_COL: (df[bc] >= 10) & (df[bc] < 15) & (np.isfinite(df[bc])),
        "15_20": lambda df, bc=BIN_COL: (df[bc] >= 15) & (df[bc] < 20) & (np.isfinite(df[bc])),
        "GT20": lambda df, bc=BIN_COL: (df[bc] >= 20) & (np.isfinite(df[bc]))
    }

    for mcap_scen in MCAP_SCENARIOS:
        m_min, m_max = mcap_scen["min"], mcap_scen["max"]
        mcap_label = "OCF_All" if np.isinf(m_max) else f"OCF_{int(m_min)}_{int(m_max)}"
        
        for debt_scen in DEBT_SCENARIOS:
            debt_label = debt_scen["label"]
            file_name = os.path.join(output_folder, f"{BIN_COL}_{mcap_label}_{debt_label}_{MONTH_NAME}.xlsx")
            
            with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
                for bin_name, bin_cond in BINS.items():
                    start_row = 0
                    sheet_name = bin_name[:31]
                    filter_key = f"{BIN_COL}_{bin_name}_{mcap_label}_{debt_label}"

                    for ret_name, ret_func in {
                        "Fwd_Return": lambda df: np.where(df["Year"] == 2025, df["Feb26Ret"], df["Fwd1Y"]),
                        "CAGR2Y": lambda df: df["CAGR2Y"],
                        "CAGR3Y": lambda df: df["CAGR3Y"]
                    }.items():
                        
                        master_row_g = {"Filter_Key": filter_key, "Type": "Growth"}
                        master_row_ng = {"Filter_Key": filter_key, "Type": "Non-Growth"}
                        v_cols = VAL_MAPPING[ret_name]
                        curr_params = BASE_PARAMS_START + (["CAGR_2Y>50", "CAGR_3Y>25"] if ret_name == "Fwd_Return" else [f"Growth_{ret_name}", "Spacer"]) + BASE_PARAMS_END
                        
                        g_blocks, ng_blocks = {}, {}
                        for y in YEARS:
                            col_h = "Feb_26" if y == 2025 else str(y)
                            for is_growth, container in [(True, g_blocks), (False, ng_blocks)]:
                                df_y = df_d[(df_d["Year"] == y) & (df_d[MCAP_OCF_COL] >= m_min) & (df_d[MCAP_OCF_COL] <= m_max)].copy()
                                df_y = df_y[debt_scen["cond"](df_y)]
                                
                                if is_growth: 
                                    df_y = df_y[(df_y["Revenue_Growth"] >= REV_GROWTH_MIN) & (df_y["PAT_Growth"] >= PAT_GROWTH_MIN)]
                                
                                df_y = df_y[bin_cond(df_y)]
                                df_y["Effective_Return"] = ret_func(df_y)
                                
                                target_row = master_row_g if is_growth else master_row_ng
                                target_row[f"{col_h}_Avg_Ret"] = df_y["Effective_Return"].mean()
                                target_row[f"{col_h}_%+ve"] = (df_y["Effective_Return"] > 0).sum() / len(df_y) if not df_y.empty else np.nan
                                target_row[f"{col_h}_Count"] = len(df_y)

                                pivot = df_y.groupby("Company").agg(
                                    Avg_Return=("Effective_Return", "mean"),
                                    Avg_PS=("PS", "mean"), Avg_PB=("PB", "mean"), Avg_MCAP_OCF=(MCAP_OCF_COL, "mean"),
                                    Avg_Fwd_PS=(v_cols["ps"], "mean"), Avg_Fwd_PB=(v_cols["pb"], "mean"), Avg_Fwd_MCAP_OCF=(v_cols["mcap_ocf"], "mean"),
                                    Avg_Rev_Growth=("Revenue_Growth", "mean"), Avg_PAT_Growth=("PAT_Growth", "mean"),
                                    Avg_Rev_Gr_Fwd_1Y=("Rev_Fwd_GR1Y", "mean"), Avg_PAT_Gr_Fwd_1Y=("PAT_Fwd_GR1Y", "mean"),
                                    CAGR2Y=("CAGR2Y", "mean"), CAGR3Y=("CAGR3Y", "mean")
                                )
                                container[col_h] = compute_block(pivot, ret_name)

                        apply_formatting(pd.DataFrame(g_blocks, index=curr_params)).to_excel(writer, sheet_name=sheet_name, startrow=start_row+2, startcol=0)
                        apply_formatting(pd.DataFrame(ng_blocks, index=curr_params)).to_excel(writer, sheet_name=sheet_name, startrow=start_row+2, startcol=len(YEARS)+2)
                        start_row += 35 
                        master_data[ret_name].extend([master_row_g, master_row_ng])

# --- SUMMARY EXPORT ---
master_file_name = f"Master_Summary_{timestamp}.xlsx"
with pd.ExcelWriter(master_file_name, engine="xlsxwriter") as master_writer:
    for tab_name, rows in master_data.items():
        df_tab = pd.DataFrame(rows)
        ret_cols = [c for c in df_tab.columns if "_Avg_Ret" in c]
        pos_cols = [c for c in df_tab.columns if "_%+ve" in c]
        cnt_cols = [c for c in df_tab.columns if "_Count" in c]
        df_tab = df_tab[["Filter_Key", "Type"] + ret_cols + pos_cols + cnt_cols]
        for col in df_tab.columns:
            if any(x in col for x in ["_Avg_Ret", "_%+ve"]):
                df_tab[col] = df_tab[col].apply(lambda x: f"{x:.1%}" if pd.notna(x) and x != "" else "")
        df_tab.to_excel(master_writer, sheet_name=tab_name, index=False)

    # PORTFOLIO PICKS - NOW WITH DEBT FILTER
    df_2025_growth = df_d[
        (df_d["Year"] == 2025) & 
        (df_d["Revenue_Growth"] >= REV_GROWTH_MIN) & 
        (df_d["PAT_Growth"] >= PAT_GROWTH_MIN) &
        (df_d["Net Debt"] <= DEBT_LIMIT_SAFE) # GLOBAL DEBT SYNC
    ].copy()
    
    p1 = df_2025_growth[
        (df_2025_growth["PB"] <= 3) & 
        (df_2025_growth[MCAP_OCF_COL] >= 0) & 
        (df_2025_growth[MCAP_OCF_COL] <= 50)
    ].copy()
    p1["Portfolio_Tag"] = "PB_0-3_OCF_0-50_SafeDebt"
    
    p2 = df_2025_growth[
        (df_2025_growth["PB"] <= 5) & 
        (df_2025_growth[MCAP_OCF_COL] >= 0) & 
        (df_2025_growth[MCAP_OCF_COL] <= 20)
    ].copy()
    p2["Portfolio_Tag"] = "PB_0-5_OCF_0-20_SafeDebt"
    
    final_picks = pd.concat([p1, p2]).drop_duplicates(subset=["Company"])
    final_picks.to_excel(master_writer, sheet_name="Selected_Picks", index=False)

print(f"\n✅ FULL SYNC COMPLETE!")
print(f"📁 Folder: {output_folder}")
print(f"📉 Selected Picks now only includes companies with Net Debt <= {DEBT_LIMIT_SAFE}")