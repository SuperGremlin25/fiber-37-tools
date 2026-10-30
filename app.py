import gradio as gr
import pandas as pd
import numpy as np
from io import BytesIO
import tempfile
import os

# ==================== LONGHAUL OSP LOGIC ====================
def calculate_longhaul(
    total_miles,
    fiber_cable_rate,
    duct_rate,
    plow_rate,
    bore_rate,
    rock_premium,
    restoration_rate,
    permitting_rate,
    mobilization,
    contingency_base,
    plow_ft,
    bore_ft,
    rock_pct_of_bore,
    plow_daily_rate,
    bore_daily_rate
):
    total_ft = total_miles * 5280
    duct_ft = bore_ft * 0.4
    rock_ft = bore_ft * (rock_pct_of_bore / 100)

    costs = {
        "Fiber Cable": (plow_ft + bore_ft) * fiber_cable_rate,
        "Duct / Conduit": duct_ft * duct_rate,
        "Plow Installation": plow_ft * plow_rate,
        "Bore Installation": bore_ft * bore_rate,
        "Rock Premium": rock_ft * rock_premium,
        "Restoration": (plow_ft + bore_ft) * restoration_rate,
        "Permitting & Locates": total_miles * permitting_rate,
        "Mobilization": mobilization * 2,
    }
    subtotal = sum(costs.values())
    contingency_pct = max(contingency_base, 0.12)
    contingency_amt = subtotal * contingency_pct
    total_cost = subtotal + contingency_amt
    cost_per_mile = total_cost / total_miles if total_miles else 0

    plow_days = plow_ft / plow_daily_rate if plow_daily_rate else 0
    bore_days = bore_ft / bore_daily_rate if bore_daily_rate else 0
    calendar_days = max(plow_days, bore_days) * 1.15

    summary = {
        "Total Estimated Cost": f"${total_cost:,.0f}",
        "Cost per Mile": f"${cost_per_mile:,.0f}",
        "Contingency Amount": f"${contingency_amt:,.0f}",
        "Contingency %": f"{contingency_pct:.0%}",
        "Estimated Duration (days)": f"{calendar_days:.0f}",
        "Plow Footage": f"{plow_ft:,.0f}",
        "Bore Footage": f"{bore_ft:,.0f}",
    }
    cost_df = pd.DataFrame([{"Category": k, "Extended": v} for k, v in costs.items()])
    return summary, cost_df, total_cost, cost_per_mile

def export_longhaul_excel(summary, cost_df, total_miles, plow_ft, bore_ft):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
            cost_df.to_excel(writer, sheet_name="Cost Breakdown", index=False)
            pd.DataFrame([summary]).T.reset_index().rename(columns={"index": "Metric", 0: "Value"}).to_excel(writer, sheet_name="Summary", index=False)
            pd.DataFrame({"Input": ["Total Miles", "Plow Ft", "Bore Ft"], "Value": [total_miles, plow_ft, bore_ft]}).to_excel(writer, sheet_name="Inputs", index=False)
        return tmp.name

def export_longhaul_md(summary, cost_df):
    md = "# Longhaul OSP Assessment Summary\n\n"
    md += "## Key Metrics\n"
    for k, v in summary.items():
        md += f"- **{k}**: {v}\n"
    md += "\n## Cost Breakdown\n\n| Category | Extended |\n|----------|----------|\n"
    for _, row in cost_df.iterrows():
        md += f"| {row['Category']} | ${row['Extended']:,.0f} |\n"
    md += "\n*Generated for NotebookLM import. Paste into a source or upload the file.*\n"
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w") as tmp:
        tmp.write(md)
        return tmp.name

# ==================== MARKET / MDU LOGIC ====================
def calculate_market(
    passings, route_miles, pct_aerial, pct_ug, mdu_pct, mdu_units_per_bldg,
    take_rate, ug_cost_per_ft, aerial_cost_per_ft, drop_ug, drop_aerial,
    eng_pct, permit_pct, contingency_pct, mdu_interior, mdu_common, arpu
):
    aerial_ft = route_miles * 5280 * (pct_aerial / 100)
    ug_ft = route_miles * 5280 * (pct_ug / 100)
    aerial_cost = aerial_ft * aerial_cost_per_ft
    ug_cost = ug_ft * ug_cost_per_ft
    drop_cost = passings * ((pct_aerial/100)*drop_aerial + (pct_ug/100)*drop_ug)
    eng = (aerial_cost + ug_cost) * (eng_pct / 100)
    permit = (aerial_cost + ug_cost) * (permit_pct / 100)
    base = aerial_cost + ug_cost + drop_cost + eng + permit
    contingency = base * (contingency_pct / 100)
    mdu_passings = passings * (mdu_pct / 100)
    bldgs = mdu_passings / mdu_units_per_bldg if mdu_units_per_bldg else 0
    mdu_cost = mdu_passings * mdu_interior + bldgs * mdu_common
    total_capex = base + contingency + mdu_cost
    capex_per_passing = total_capex / passings if passings else 0
    adj_take = take_rate + 0.05 * (mdu_pct / 100)
    yr1_rev = passings * (adj_take / 100) * arpu * 12

    summary = {
        "Total Capex": f"${total_capex:,.0f}",
        "Capex / Passing": f"${capex_per_passing:,.0f}",
        "Adjusted Take Rate": f"{adj_take:.1f}%",
        "Est. Year 1 Revenue": f"${yr1_rev:,.0f}",
        "MDU Incremental": f"${mdu_cost:,.0f}",
    }
    return summary, total_capex, capex_per_passing

def export_market_excel(summary, passings, route_miles, mdu_pct):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
            pd.DataFrame([summary]).T.reset_index().rename(columns={"index": "Metric", 0: "Value"}).to_excel(writer, sheet_name="Summary", index=False)
            pd.DataFrame({"Input": ["Passings", "Route Miles", "MDU %"], "Value": [passings, route_miles, mdu_pct]}).to_excel(writer, sheet_name="Inputs", index=False)
        return tmp.name

def export_market_md(summary):
    md = "# Market / MDU Cost Model Summary\n\n"
    md += "## Key Metrics\n"
    for k, v in summary.items():
        md += f"- **{k}**: {v}\n"
    md += "\n*Generated for NotebookLM import. Use this as a source to create data-rich slides or audio overviews.*\n"
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w") as tmp:
        tmp.write(md)
        return tmp.name

# ==================== GRADIO UI ====================
with gr.Blocks(title="Fiber OSP Cost Estimator", theme=gr.themes.Soft()) as demo:
    gr.Markdown("# Fiber OSP Cost Estimator\nTwo distinct tools: Longhaul OSP Assessment + Market / MDU Cost Model\n\n**Exports:** Excel workbooks + NotebookLM-ready Markdown")
    
    with gr.Tabs():
        # ---------- LONGHAUL TAB ----------
        with gr.Tab("Longhaul OSP Assessment"):
            with gr.Row():
                with gr.Column():
                    gr.Markdown("### Route & Production Inputs")
                    lh_miles = gr.Number(label="Total Route Miles", value=85)
                    lh_plow_ft = gr.Number(label="Plow Footage", value=339120)
                    lh_bore_ft = gr.Number(label="Bore Footage", value=109680)
                    lh_rock_pct = gr.Slider(0, 50, value=15, label="Rock % of Bore Footage")
                    lh_plow_daily = gr.Number(label="Plow Rate (ft/day)", value=4500)
                    lh_bore_daily = gr.Number(label="Bore Rate (ft/day)", value=250)
                with gr.Column():
                    gr.Markdown("### Unit Rates (editable)")
                    lh_fiber_rate = gr.Number(label="Fiber Cable ($/ft)", value=3.5)
                    lh_duct_rate = gr.Number(label="Duct ($/ft)", value=0.65)
                    lh_plow_rate = gr.Number(label="Plow Install ($/ft)", value=4.5)
                    lh_bore_rate = gr.Number(label="Bore Install ($/ft)", value=28)
                    lh_rock_rate = gr.Number(label="Rock Premium ($/ft)", value=45)
                    lh_rest_rate = gr.Number(label="Restoration ($/ft)", value=1.2)
                    lh_permit_rate = gr.Number(label="Permitting ($/mile)", value=450)
                    lh_mob = gr.Number(label="Mobilization (lump)", value=15000)
                    lh_cont_base = gr.Slider(0.10, 0.40, value=0.25, label="Contingency Base %", step=0.01)
            
            lh_btn = gr.Button("Calculate Longhaul", variant="primary")
            lh_summary = gr.JSON(label="Longhaul Summary")
            lh_table = gr.Dataframe(label="Cost Breakdown")
            with gr.Row():
                lh_xlsx = gr.File(label="Download Excel Workbook")
                lh_md = gr.File(label="Download NotebookLM Markdown")

            def run_longhaul(*args):
                summary, cost_df, total, cpm = calculate_longhaul(*args)
                xlsx_path = export_longhaul_excel(summary, cost_df, args[0], args[10], args[11])
                md_path = export_longhaul_md(summary, cost_df)
                return summary, cost_df, xlsx_path, md_path

            lh_btn.click(
                run_longhaul,
                inputs=[lh_miles, lh_fiber_rate, lh_duct_rate, lh_plow_rate, lh_bore_rate,
                        lh_rock_rate, lh_rest_rate, lh_permit_rate, lh_mob, lh_cont_base,
                        lh_plow_ft, lh_bore_ft, lh_rock_pct, lh_plow_daily, lh_bore_daily],
                outputs=[lh_summary, lh_table, lh_xlsx, lh_md]
            )

        # ---------- MARKET / MDU TAB ----------
        with gr.Tab("Market / MDU Cost Model"):
            with gr.Row():
                with gr.Column():
                    gr.Markdown("### Market Inputs")
                    m_passings = gr.Number(label="Passings", value=4500)
                    m_route_mi = gr.Number(label="Route Miles", value=15)
                    m_pct_aerial = gr.Slider(0, 100, value=70, label="% Aerial")
                    m_pct_ug = gr.Slider(0, 100, value=30, label="% Underground")
                    m_mdu_pct = gr.Slider(0, 100, value=15, label="MDU % of Passings")
                    m_mdu_units = gr.Number(label="MDU Units per Building", value=24)
                    m_take = gr.Slider(10, 70, value=35, label="Base Take Rate %")
                with gr.Column():
                    gr.Markdown("### Cost Benchmarks")
                    m_ug_ft = gr.Number(label="UG Cost ($/ft)", value=18)
                    m_aerial_ft = gr.Number(label="Aerial Cost ($/ft)", value=8)
                    m_drop_ug = gr.Number(label="Drop Cost UG ($)", value=600)
                    m_drop_aerial = gr.Number(label="Drop Cost Aerial ($)", value=370)
                    m_eng = gr.Slider(0, 20, value=8, label="Engineering %")
                    m_permit = gr.Slider(0, 25, value=12, label="Permitting / Make-Ready %")
                    m_cont = gr.Slider(5, 25, value=10, label="Contingency %")
                    m_mdu_int = gr.Number(label="MDU Interior ($/unit)", value=450)
                    m_mdu_com = gr.Number(label="MDU Common Area ($/bldg)", value=8500)
                    m_arpu = gr.Number(label="ARPU Monthly ($)", value=70)
            
            m_btn = gr.Button("Calculate Market", variant="primary")
            m_summary = gr.JSON(label="Market Summary")
            with gr.Row():
                m_xlsx = gr.File(label="Download Excel Workbook")
                m_md = gr.File(label="Download NotebookLM Markdown")

            def run_market(*args):
                summary, total, cpp = calculate_market(*args)
                xlsx_path = export_market_excel(summary, args[0], args[1], args[4])
                md_path = export_market_md(summary)
                return summary, xlsx_path, md_path

            m_btn.click(
                run_market,
                inputs=[m_passings, m_route_mi, m_pct_aerial, m_pct_ug, m_mdu_pct, m_mdu_units,
                        m_take, m_ug_ft, m_aerial_ft, m_drop_ug, m_drop_aerial, m_eng, m_permit,
                        m_cont, m_mdu_int, m_mdu_com, m_arpu],
                outputs=[m_summary, m_xlsx, m_md]
            )

        # ---------- COMBINED TAB ----------
        with gr.Tab("Combined View"):
            gr.Markdown("Run both calculators above, then review side-by-side here.")
            with gr.Row():
                comb_lh = gr.JSON(label="Longhaul OSP Summary")
                comb_mkt = gr.JSON(label="Market / MDU Summary")

    gr.Markdown("""
    ---
    **Deliverables**  
    • **Excel Workbooks** — Full calculated models ready for client or internal use.  
    • **Markdown** — Structured for direct import into Google NotebookLM to generate slide decks or audio overviews.  
    • **Gamma decks** — Best path for polished slides. Generate the MD or summary, then use your connected Gamma account (or paste the outline).  

    All rates are preliminary. Validate against current vendor quotes and local conditions.
    """)

if __name__ == "__main__":
    demo.launch()
