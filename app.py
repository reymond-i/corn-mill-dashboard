import math
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO
import hashlib
import plotly.graph_objects as go

st.set_page_config(page_title="Corn Mill Performance Dashboard", layout="wide")

DATA_PATH_DEFAULT = "Corn Mill Data Request (1) edited - AENG 236.xlsx"

def is_nan(x):
    return isinstance(x, float) and math.isnan(x)

def to_float(x):
    if x is None or is_nan(x):
        return np.nan
    if isinstance(x, str):
        x = x.strip()
        if x in ["", "-"]:
            return np.nan
        try:
            return float(x)
        except:
            return np.nan
    try:
        return float(x)
    except:
        return np.nan

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes, file_hash: str):
    bio = BytesIO(file_bytes)
    wb = openpyxl.load_workbook(bio, data_only=True)
    datasets = []

    def parse_sheet(sheet_name: str):
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        ar = None
        for r in range(len(df)):
            if (df.iloc[r, :] == "Aperture, mm").any():
                ar = r
                break
        if ar is None:
            return []

        col_ap = int(np.where(df.iloc[ar, :] == "Aperture, mm")[0][0])

        def get_outlets(r):
            outs = []
            for c in range(col_ap + 1, df.shape[1]):
                v = df.iloc[r, c]
                if isinstance(v, str) and v.strip() and v.strip().lower() not in [
                    "percentage of grits collected at each outlet"
                ]:
                    outs.append(v.strip())
            return outs

        outlets = get_outlets(ar)
        if not outlets and ar + 1 < len(df):
            outlets = get_outlets(ar + 1)

        # dataset rows = numeric sample no in col 0
        data_rows = []
        for i, v in enumerate(df.iloc[:, 0].tolist()):
            if isinstance(v, (int, float)) and not is_nan(v):
                data_rows.append(i)
        data_rows = sorted(data_rows)

        out = []
        for k, start in enumerate(data_rows):
            end = data_rows[k + 1] if k + 1 < len(data_rows) else len(df)
            metrics = df.iloc[start, 0:9].tolist()

            sieve_rows = []
            for r in range(start, end):
                sieve = df.iloc[r, 9]
                ap = df.iloc[r, col_ap]
                if (not is_nan(ap) and ap is not None) or (isinstance(sieve, str) and sieve.strip()):
                    row = {"sieve": sieve, "aperture_mm": ap}
                    for j, o in enumerate(outlets):
                        row[o] = df.iloc[r, col_ap + 1 + j]
                    sieve_rows.append(row)

            out.append(
                dict(
                    sheet=sheet_name,
                    sample_no=int(metrics[0]) if not is_nan(metrics[0]) else None,
                    variety=metrics[1],
                    moisture_pct=metrics[2],
                    bulk_density_kgm3=metrics[3],
                    purity_pct=metrics[4],
                    input_kg=metrics[5],
                    main_product_recovery_pct=metrics[6],
                    byproduct_recovery_pct=metrics[7],
                    losses_pct=metrics[8],
                    outlets=outlets,
                    sieve_rows=sieve_rows,
                )
            )
        return out

    for s in wb.sheetnames:
        datasets.extend(parse_sheet(s))

    # ---- Build tables ----
    metrics_rows = []
    sieve_long = []

    for d in datasets:
        group_id = " | ".join(d["outlets"]) if d["outlets"] else "Unknown outlets"
        run_id = f"Run {d['sample_no']:02d}" if d["sample_no"] is not None else f"Sheet {d['sheet']}"

        metrics_rows.append(
            dict(
                Run=run_id,
                Sample_No=d["sample_no"],
                Sheet_Tab=d["sheet"],
                Outlet_Group=group_id,
                Variety=d["variety"],
                Moisture_Content_pct=to_float(d["moisture_pct"]),
                Bulk_Density_kgm3=to_float(d["bulk_density_kgm3"]),
                Purity_pct=to_float(d["purity_pct"]),
                Input_kg=to_float(d["input_kg"]),
                Main_Product_Recovery_pct=to_float(d["main_product_recovery_pct"]),
                Byproduct_Recovery_pct=to_float(d["byproduct_recovery_pct"]),
                Losses_pct=to_float(d["losses_pct"]),
            )
        )

        for row in d["sieve_rows"]:
            sieve = row["sieve"]
            ap = to_float(row["aperture_mm"])
            for o in d["outlets"]:
                sieve_long.append(
                    dict(
                        Run=run_id,
                        Sample_No=d["sample_no"],
                        Sheet_Tab=d["sheet"],
                        Outlet_Group=group_id,
                        Outlet=o,
                        Sieve=sieve,
                        Aperture_mm=ap,
                        Percent_Retained=to_float(row.get(o)),
                    )
                )

    metrics_df = pd.DataFrame(metrics_rows).sort_values("Sample_No").reset_index(drop=True)
    sieve_df = pd.DataFrame(sieve_long).dropna(subset=["Percent_Retained"]).copy()

    # PSD cumulative pass per run/outlet (exclude pan/NaN apertures)
    psd_rows = []
    for (run, outlet), g in sieve_df.groupby(["Run", "Outlet"]):
        g2 = g.dropna(subset=["Aperture_mm"]).copy()
        if g2.empty:
            continue
        g2 = g2.sort_values("Aperture_mm", ascending=False)
        g2["Cum_Retained"] = g2["Percent_Retained"].cumsum()
        g2["Cum_Pass_pct"] = 100 - g2["Cum_Retained"]
        for _, r in g2.iterrows():
            psd_rows.append(
                dict(
                    Run=run,
                    Sample_No=int(r.get("Sample_No")) if not pd.isna(r.get("Sample_No")) else None,
                    Outlet=outlet,
                    Aperture_mm=float(r["Aperture_mm"]),
                    Cum_Pass_pct=float(r["Cum_Pass_pct"]),
                )
            )
    psd_df = pd.DataFrame(psd_rows)

    # ---- Year-based compliance (as requested) ----
    # Run 01–08 => 2018 standard: Main Product Recovery ≥64% and Losses ≤5%
    # Run 09+  => 2021 standard: Main Product Recovery ≥55% and Losses ≤5%
    def compute_standard(sample_no: int) -> str:
        if sample_no is None or pd.isna(sample_no):
            return "Unknown"
        return "2018" if int(sample_no) <= 8 else "2021"

    metrics_df["Standard_Used"] = metrics_df["Sample_No"].apply(compute_standard)

    def compliant(row) -> bool:
        loss_ok = row["Losses_pct"] <= 5
        if row["Standard_Used"] == "2018":
            main_ok = row["Main_Product_Recovery_pct"] >= 64
        elif row["Standard_Used"] == "2021":
            main_ok = row["Main_Product_Recovery_pct"] >= 55
        else:
            return False
        return bool(main_ok and loss_ok)

    metrics_df["Compliance"] = metrics_df.apply(lambda r: "PASS" if compliant(r) else "FAIL", axis=1)

    # Failure reason (Recovery / Losses / Both) for easier reporting
    def fail_reason(row) -> str:
        loss_fail = row["Losses_pct"] > 5
        if row["Standard_Used"] == "2018":
            rec_fail = row["Main_Product_Recovery_pct"] < 64
        elif row["Standard_Used"] == "2021":
            rec_fail = row["Main_Product_Recovery_pct"] < 55
        else:
            rec_fail = True
        if rec_fail and loss_fail:
            return "Recovery + Losses"
        if rec_fail:
            return "Recovery"
        if loss_fail:
            return "Losses"
        return "—"
    metrics_df["Fail_Reason"] = metrics_df.apply(fail_reason, axis=1)

    # Batch-level compliance score (0–100)
    # 50% points = recovery criterion, 50% points = loss criterion
    def compliance_score(row) -> float:
        # Recovery component scaled to threshold (cap at 100)
        if row["Standard_Used"] == "2018":
            rec_thr = 64.0
        else:
            rec_thr = 55.0
        rec = row["Main_Product_Recovery_pct"]
        loss = row["Losses_pct"]

        # Handle missing
        if pd.isna(rec) or pd.isna(loss):
            return float("nan")

        rec_ratio = max(0.0, min(1.0, float(rec) / rec_thr))   # 0..1
        loss_ratio = max(0.0, min(1.0, (5.0 - float(loss)) / 5.0))  # 1 at 0% losses, 0 at 5% losses, negative clipped

        return 100.0 * (0.5 * rec_ratio + 0.5 * loss_ratio)

    metrics_df["Compliance_Score_%"] = metrics_df.apply(compliance_score, axis=1)

    return metrics_df, sieve_df, psd_df

def pill(text: str, ok: bool):
    bg = "#d4edda" if ok else "#f8d7da"
    fg = "#155724" if ok else "#721c24"
    st.markdown(
        f"<span style='background:{bg};color:{fg};padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700;'>{text}</span>",
        unsafe_allow_html=True,
    )

def fig_psd_single_run(psd_df: pd.DataFrame, run: str):
    d = psd_df[psd_df["Run"] == run].copy()
    fig = go.Figure()
    for outlet in sorted(d["Outlet"].unique().tolist()):
        g = d[d["Outlet"] == outlet].sort_values("Aperture_mm", ascending=True)
        fig.add_trace(go.Scatter(x=g["Aperture_mm"], y=g["Cum_Pass_pct"], mode="lines+markers", line_shape="spline", marker=dict(symbol="circle", size=9, line=dict(width=1, color="black")), line=dict(width=2), name=outlet))
    fig.update_layout(
        title=f"Particle Size Distribution — {run}",
        xaxis=dict(title="Aperture (mm) — log scale", type="log"),
        yaxis=dict(title="Cumulative % Passing", range=[0, 100]),
        height=480,
        margin=dict(l=60, r=20, t=60, b=50),
        legend_title="Outlet",
    )
    return fig

def fig_psd_compare_runs(psd_df: pd.DataFrame, outlet: str, run_order: list[str]):
    d = psd_df[psd_df["Outlet"] == outlet].copy()
    fig = go.Figure()
    # Preserve user-visible run order (01..20)
    for run in run_order:
        g = d[d["Run"] == run].sort_values("Aperture_mm", ascending=True)
        if g.empty:
            continue
        fig.add_trace(go.Scatter(x=g["Aperture_mm"], y=g["Cum_Pass_pct"], mode="lines+markers", line_shape="spline", marker=dict(symbol="circle", size=9, line=dict(width=1, color="black")), line=dict(width=2), name=run))
    fig.update_layout(
        title=f"PSD Comparison Across Runs — Outlet: {outlet}",
        xaxis=dict(title="Aperture (mm) — log scale", type="log"),
        yaxis=dict(title="Cumulative % Passing", range=[0, 100]),
        height=480,
        margin=dict(l=60, r=20, t=60, b=50),
        legend_title="Run",
    )
    return fig

def fig_bar_sieve(sieve_df: pd.DataFrame, run: str, outlet: str):
    d = sieve_df[(sieve_df["Run"] == run) & (sieve_df["Outlet"] == outlet)].dropna(subset=["Aperture_mm"]).copy()
    d = d.sort_values("Aperture_mm", ascending=False)
    x = [f"{s} ({a:g} mm)" for s, a in zip(d["Sieve"], d["Aperture_mm"])]
    fig = go.Figure(data=[go.Bar(x=x, y=d["Percent_Retained"])])
    fig.update_layout(
        title=f"Sieve Size Distribution — {run} | Outlet: {outlet}",
        xaxis=dict(title="Sieve (Aperture mm)", tickangle=30),
        yaxis=dict(title="Percent Retained (%)"),
        height=480,
        margin=dict(l=60, r=20, t=60, b=120),
    )
    return fig

def fig_compliance_pie(metrics_df: pd.DataFrame):
    counts = metrics_df["Compliance"].value_counts().reindex(["PASS", "FAIL"]).fillna(0).astype(int)
    fig = go.Figure(data=[go.Pie(labels=counts.index.tolist(), values=counts.values.tolist(), hole=0.55)])
    fig.update_layout(title="Compliance Summary (PASS vs FAIL)", height=320, margin=dict(l=20, r=20, t=60, b=10))
    return fig

def fig_pass_gauge(metrics_df: pd.DataFrame):
    total = len(metrics_df)
    pass_pct = 0.0 if total == 0 else 100.0 * float((metrics_df["Compliance"] == "PASS").sum()) / float(total)
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=pass_pct,
            number={"suffix": "%"},
            title={"text": "PASS rate"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"thickness": 0.25},
                "steps": [
                    {"range": [0, 50]},
                    {"range": [50, 80]},
                    {"range": [80, 100]},
                ],
            },
        )
    )
    fig.update_layout(height=320, margin=dict(l=20, r=20, t=60, b=10))
    return fig

def year_group_label(sample_no: int) -> str:
    if sample_no <= 8:
        return "2018 group (Run 01–08)"
    return "2021 group (Run 09+)"
    
st.title("Corn Mill Performance Dashboard")

with st.sidebar:
    st.header("Upload data")
    st.write("Upload the Excel workbook (.xlsx) to generate the dashboard.")
    up = st.file_uploader("Upload .xlsx", type=["xlsx"], key="uploader")
    if up is None:
        st.info("Please upload the Excel file to start.")
        st.stop()
    file_bytes = up.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()

metrics_df, sieve_df, psd_df = load_workbook(file_bytes, file_hash)

# Derived lists
all_runs = metrics_df["Run"].tolist()
all_outlets = sorted(sieve_df["Outlet"].unique().tolist())
all_groups = sorted(metrics_df["Outlet_Group"].unique().tolist())

# ---- Year filter toggle ----
with st.sidebar:
    st.header("Year / Standard filter")
    year_choice = st.radio(
        "Show runs from:",
        options=["All (2018 + 2021)", "2018 group (Run 01–08)", "2021 group (Run 09+)"],
        index=0,
    )

if year_choice == "2018 group (Run 01–08)":
    metrics_f = metrics_df[metrics_df["Sample_No"] <= 8].copy()
elif year_choice == "2021 group (Run 09+)":
    metrics_f = metrics_df[metrics_df["Sample_No"] >= 9].copy()
else:
    metrics_f = metrics_df.copy()

runs_f = metrics_f["Run"].tolist()
if not runs_f:
    st.warning("No runs match the selected year filter.")
    st.stop()

# Filter sieve/psd to matching runs
sieve_f = sieve_df[sieve_df["Run"].isin(runs_f)].copy()
psd_f = psd_df[psd_df["Run"].isin(runs_f)].copy()

groups_f = sorted(metrics_f["Outlet_Group"].unique().tolist())

with st.sidebar:
    st.header("Filters")
    group_sel = st.selectbox("Outlet group (same outlets grouped together)", options=groups_f)
    runs_in_group = metrics_f.loc[metrics_f["Outlet_Group"] == group_sel, "Run"].tolist()
    run_sel = st.selectbox("Run", options=runs_in_group)
    outlets_in_run = sorted(sieve_f.loc[sieve_f["Run"] == run_sel, "Outlet"].unique().tolist())
    outlet_sel = st.selectbox("Outlet (for bar chart / comparison)", options=outlets_in_run if outlets_in_run else all_outlets)

# ---- Top summary cards ----
row = metrics_df[metrics_df["Run"] == run_sel].iloc[0]
std = row["Standard_Used"]
year_group = year_group_label(int(row["Sample_No"]))

st.caption(f"**Compliance basis:** Year-based standard. {year_group} uses **{std}** thresholds (Recovery + Losses). Losses limit = **≤5%** for both.")

c0, c1, c2, c3, c4, c5 = st.columns([1.2, 1.1, 1.1, 1.1, 1.2, 1.4])
with c0:
    st.metric("Moisture Content (%)", f"{row['Moisture_Content_pct']:.2f}")
    st.metric("Bulk Density (kg/m³)", f"{row['Bulk_Density_kgm3']:.2f}")
with c1:
    st.metric("Purity (%)", f"{row['Purity_pct']:.2f}")
    st.metric("Input (kg)", f"{row['Input_kg']:.2f}")
with c2:
    st.metric("Main Product Recovery (%)", f"{row['Main_Product_Recovery_pct']:.2f}")
    st.metric("By-product Recovery (%)", f"{row['Byproduct_Recovery_pct']:.2f}")
with c3:
    st.metric("Losses (%)", f"{row['Losses_pct']:.2f}")
    st.metric("Compliance Score (%)", f"{row['Compliance_Score_%']:.1f}")
with c4:
    st.write("Compliance")
    pill(f"{std}: {row['Compliance']}", row["Compliance"] == "PASS")
    st.write("")
    st.write("Standard used")
    st.write(f"**{std}**")
with c5:
    st.write("Run info")
    st.write(f"**{run_sel}**")
    st.write(f"Tab: **{row['Sheet_Tab']}**")
    st.write(f"Variety: **{row['Variety']}**")
    st.write(f"Outlet group: **{row['Outlet_Group']}**")

st.divider()

# ---- Compliance summary pie + key stats ----
p1, p2, p3 = st.columns([1.0, 1.0, 1.3])
with p1:
    st.plotly_chart(fig_compliance_pie(metrics_f), use_container_width=True)
with p2:
    st.plotly_chart(fig_pass_gauge(metrics_f), use_container_width=True)
with p3:
    # Small KPI table for filtered subset
    total = len(metrics_f)
    pass_n = int((metrics_f["Compliance"] == "PASS").sum())
    fail_n = total - pass_n
    avg_score = float(metrics_f["Compliance_Score_%"].mean())
    st.subheader("Filtered summary")
    st.write(f"- Runs shown: **{total}**")
    st.write(f"- PASS: **{pass_n}**  |  FAIL: **{fail_n}**")
    st.write(f"- Average compliance score: **{avg_score:.1f}%**")
    st.write("")
    st.write("Runs that PASSED")
    passed = metrics_f[metrics_f["Compliance"] == "PASS"].copy()
    if passed.empty:
        st.info("No runs passed under the selected filter.")
    else:
        passed_sorted = passed.sort_values("Compliance_Score_%", ascending=False)[
            ["Run","Compliance_Score_%","Compliance","Standard_Used"]
        ]
        st.dataframe(passed_sorted, use_container_width=True, hide_index=True)

    st.write("Failed runs (with reason)")
    failed = metrics_f[metrics_f["Compliance"] == "FAIL"].copy()
    if failed.empty:
        st.info("No failed runs under the selected filter.")
    else:
        failed_sorted = failed.sort_values("Compliance_Score_%", ascending=True)[
            ["Run","Compliance_Score_%","Standard_Used","Fail_Reason","Main_Product_Recovery_pct","Losses_pct"]
        ].rename(columns={
            "Main_Product_Recovery_pct":"Main Recovery (%)",
            "Losses_pct":"Losses (%)"
        })
        st.dataframe(failed_sorted, use_container_width=True, hide_index=True)

st.divider()

t1, t2 = st.tabs(["Charts", "All Data Table"])

with t1:
    left, right = st.columns(2)
    with left:
        st.plotly_chart(fig_psd_single_run(psd_f, run_sel), use_container_width=True)
    with right:
        st.plotly_chart(fig_bar_sieve(sieve_f, run_sel, outlet_sel), use_container_width=True)

    st.plotly_chart(fig_psd_compare_runs(psd_f, outlet_sel, runs_f), use_container_width=True)

with t2:
    show = metrics_f.copy()
    st.dataframe(
        show[
            [
                "Run","Sample_No","Standard_Used","Compliance","Fail_Reason","Compliance_Score_%",
                "Sheet_Tab","Outlet_Group","Variety",
                "Moisture_Content_pct","Bulk_Density_kgm3","Purity_pct","Input_kg",
                "Main_Product_Recovery_pct","Byproduct_Recovery_pct","Losses_pct",
            ]
        ].rename(columns={
            "Moisture_Content_pct":"Moisture (%)",
            "Bulk_Density_kgm3":"Bulk Density (kg/m³)",
            "Purity_pct":"Purity (%)",
            "Main_Product_Recovery_pct":"Main Product Recovery (%)",
            "Byproduct_Recovery_pct":"By-product Recovery (%)",
            "Losses_pct":"Losses (%)",
            "Fail_Reason":"Fail Reason",
            "Compliance_Score_%":"Compliance Score (%)"
        }),
        use_container_width=True,
        hide_index=True,
    )

st.caption("Compliance score (0–100) = average of: (recovery performance vs threshold) and (losses performance vs 5% max).")
