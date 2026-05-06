import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import json
import os

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
st.set_page_config(page_title="DCRS Dashboard", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in dir() else os.getcwd()
TASK_FILE = os.path.join(BASE_DIR, "user_tasks.json")
ADHOC_FILE = os.path.join(BASE_DIR, "adhoc_tasks.json")

EY_YELLOW = "#FFE600"
EY_DARK = "#2E2E38"
EY_GRAY = "#747480"
EY_WHITE = "#FFFFFF"
EY_GREEN = "#2C9C6A"
EY_RED = "#C4122F"
EY_BLUE = "#188CE5"
EY_ORANGE = "#F0AB00"
EY_AMBER = "#F0AB00"

RAG_COLORS = {"Green": EY_GREEN, "Amber": EY_AMBER, "Red": EY_RED}

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main {background-color: #F5F5F5;}
    .stMetric {background: white; border-radius: 8px; padding: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.12);}
    div[data-testid="stMetricValue"] {font-size: 2rem; font-weight: 700; color: #2E2E38;}
    div[data-testid="stMetricLabel"] {font-size: 0.85rem; color: #747480;}
    .dashboard-header {
        background: linear-gradient(135deg, #2E2E38 0%, #3D3D4E 100%);
        padding: 20px 30px; border-radius: 10px; margin-bottom: 20px;
    }
    .dashboard-header h1 {color: #FFE600; margin: 0; font-size: 2rem;}
    .dashboard-header p {color: #FFFFFF; margin: 4px 0 0 0; font-size: 0.9rem;}
    .rag-badge {
        display: inline-block; padding: 4px 14px; border-radius: 12px;
        font-weight: 700; font-size: 0.85rem; color: white; text-align: center;
    }
    .rag-green {background-color: #2C9C6A;}
    .rag-amber {background-color: #F0AB00; color: #2E2E38;}
    .rag-red {background-color: #C4122F;}
    .traffic-card {
        background: white; border-radius: 10px; padding: 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.10); text-align: center; margin-bottom: 10px;
    }
    .traffic-card .value {font-size: 1.8rem; font-weight: 700; color: #2E2E38;}
    .traffic-card .label {font-size: 0.82rem; color: #747480; margin-bottom: 6px;}
    .adhoc-card {
        background: white; padding: 12px; border-radius: 8px;
        margin-bottom: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

np.random.seed(42)

# ---------------------------------------------------------------------------
# RAG helpers
# ---------------------------------------------------------------------------
def rag_status(value, green_min, amber_min, invert=False):
    if invert:
        if value <= green_min: return "Green"
        if value <= amber_min: return "Amber"
        return "Red"
    else:
        if value >= green_min: return "Green"
        if value >= amber_min: return "Amber"
        return "Red"

def rag_html(status, text=None):
    css = {"Green": "rag-green", "Amber": "rag-amber", "Red": "rag-red"}[status]
    label = text or status
    return f'<span class="rag-badge {css}">{label}</span>'

def traffic_card(label, value, status):
    color = RAG_COLORS[status]
    return f"""
    <div class="traffic-card" style="border-top: 5px solid {color};">
        <div class="label">{label}</div>
        <div class="value">{value}</div>
        {rag_html(status)}
    </div>"""

# ---------------------------------------------------------------------------
# Data Loading from Excel files (auto-refreshes every 5 minutes)
# ---------------------------------------------------------------------------
DATA_DIR = os.path.join(BASE_DIR, "data")
CACHE_TTL = 300  # seconds

@st.cache_data(ttl=CACHE_TTL)
def load_monthly_cases():
    path = os.path.join(DATA_DIR, "monthly_cases.xlsx")
    df = pd.read_excel(path)
    df["Month"] = pd.to_datetime(df["Month"])
    df["Task"] = df["Task"].astype(str)
    df["Actual"] = df["Actual"].astype(int)
    df["Target"] = df["Target"].astype(int)
    return df

@st.cache_data(ttl=CACHE_TTL)
def load_task_breakdown():
    path = os.path.join(DATA_DIR, "task_breakdown.xlsx")
    df = pd.read_excel(path)
    df.rename(columns={"Not_Started": "Not started", "In_Progress": "In Progress"}, inplace=True)
    return df

@st.cache_data(ttl=CACHE_TTL)
def load_hybrid_status():
    path = os.path.join(DATA_DIR, "hybrid_status.xlsx")
    df = pd.read_excel(path)
    df.rename(columns={"In_Progress": "In Progress", "Not_Started": "Not started"}, inplace=True)
    return df

@st.cache_data(ttl=CACHE_TTL)
def load_availability():
    path = os.path.join(DATA_DIR, "resource_availability.xlsx")
    return pd.read_excel(path)

@st.cache_data(ttl=CACHE_TTL)
def load_weekly_throughput():
    path = os.path.join(DATA_DIR, "weekly_throughput.xlsx")
    df = pd.read_excel(path)
    df.rename(columns={"Week_Starting": "Week", "Cases_Handled": "Cases_Handled",
                        "Cases_Incoming": "Cases_Incoming", "FTE_Available": "FTE_Available",
                        "Avg_Handling_Time_Hrs": "Avg_Handling_Hrs"}, inplace=True)
    df["Week"] = pd.to_datetime(df["Week"])
    return df

@st.cache_data(ttl=CACHE_TTL)
def load_fte_roster():
    path = os.path.join(DATA_DIR, "fte_roster.xlsx")
    return pd.read_excel(path)

# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------
def load_adhoc():
    if os.path.exists(ADHOC_FILE):
        with open(ADHOC_FILE, "r") as f:
            return json.load(f)
    return [
        {"id": 1, "fte": "FTE_01", "desc": "Investigate legacy settlement",
         "priority": "High", "status": "In Progress", "hours": 8,
         "assigned": "2026-04-10", "due": "2026-04-17"},
        {"id": 2, "fte": "FTE_03", "desc": "Client escalation call",
         "priority": "Medium", "status": "Not Started", "hours": 4,
         "assigned": "2026-04-12", "due": "2026-04-19"},
        {"id": 3, "fte": "FTE_07", "desc": "Data quality review",
         "priority": "Low", "status": "Completed", "hours": 6,
         "assigned": "2026-04-08", "due": "2026-04-15"},
    ]

def save_adhoc(tasks):
    with open(ADHOC_FILE, "w") as f:
        json.dump(tasks, f, indent=2)

def load_tasks():
    if os.path.exists(TASK_FILE):
        with open(TASK_FILE, "r") as f:
            return json.load(f)
    return []

def save_tasks(tasks):
    with open(TASK_FILE, "w") as f:
        json.dump(tasks, f, indent=2)

# ---------------------------------------------------------------------------
# Load all data
# ---------------------------------------------------------------------------
monthly_cases = load_monthly_cases()
task_breakdown = load_task_breakdown()
hybrid_status = load_hybrid_status()
availability = load_availability()
weekly_tp = load_weekly_throughput()
fte_roster = load_fte_roster()

total_cases = int(monthly_cases["Actual"].sum())
total_completed = int(task_breakdown["Completed"].sum())
total_all = int(task_breakdown[["Not started", "In Progress", "Completed"]].sum().sum())
completion_rate = total_completed / total_all * 100 if total_all else 0
error_rate = 1.90
ftes = 11
handling_time = 5.45
resource_availability = 100.0

remaining_cases = total_all - total_completed
avg_weekly_tp = weekly_tp["Cases_Handled"].mean()
avg_cases_per_fte = weekly_tp["Cases_Handled"].sum() / weekly_tp["FTE_Available"].sum()

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("""
<div class="dashboard-header">
    <h1>DCRS DASHBOARD</h1>
    <p>Latest refresh: {}</p>
</div>
""".format(datetime.now().strftime("%d %B %Y, %H:%M")), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
if st.sidebar.button("Refresh Data"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.caption(f"Data folder: {DATA_DIR}")
st.sidebar.markdown("---")
st.sidebar.header("Filters")
selected_priority = st.sidebar.selectbox("Priority", ["All", "High", "Medium", "Low"])
selected_entity = st.sidebar.selectbox("Entity", ["All", "Entity A", "Entity B", "Entity C"])
selected_tasks = st.sidebar.multiselect(
    "Task Types", monthly_cases["Task"].unique().tolist(),
    default=monthly_cases["Task"].unique().tolist()
)

st.sidebar.markdown("---")
st.sidebar.header("Navigation")
page = st.sidebar.radio(
    "Go to", ["Overview", "Traffic Light", "Ad Hoc Tasks", "Forecasting", "Task Board"]
)

# ===================================================================== #
# PAGE: Overview                                                         #
# ===================================================================== #
if page == "Overview":
    st.subheader("Resource Summary")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("W13 Resource Availability", f"{resource_availability:.0f}%")
    k2.metric("FTEs", ftes)
    k3.metric("Total Cases", f"{total_cases:,}")
    k4.metric("Cases Completed", f"{total_completed:,}")
    k5.metric("Handling Time", f"{handling_time:.2f}")

    st.markdown("---")

    with st.expander("Resource Availability Detail", expanded=False):
        st.dataframe(availability, use_container_width=True, hide_index=True)

    col_chart, col_rates = st.columns([3, 1])

    with col_chart:
        st.subheader("Actual Cases Handled vs Targets")
        latest_month = monthly_cases["Month"].max()
        latest_data = monthly_cases[
            (monthly_cases["Month"] == latest_month) &
            (monthly_cases["Task"].isin(selected_tasks))
        ]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name="Actual", x=latest_data["Task"], y=latest_data["Actual"], marker_color=EY_BLUE))
        fig_bar.add_trace(go.Bar(name="Target", x=latest_data["Task"], y=latest_data["Target"], marker_color=EY_ORANGE))
        fig_bar.update_layout(barmode="group", height=350, margin=dict(l=20,r=20,t=30,b=30),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1), plot_bgcolor="white")
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_rates:
        st.subheader("Rates")
        fig_comp = go.Figure(go.Pie(values=[completion_rate, 100-completion_rate], hole=0.7,
            marker_colors=[EY_GREEN,"#E0E0E0"], textinfo="none", labels=["Completed","Remaining"]))
        fig_comp.update_layout(height=180, margin=dict(l=10,r=10,t=10,b=10), showlegend=False,
            annotations=[dict(text=f"<b>{completion_rate:.1f}%</b>",x=0.5,y=0.5,font_size=18,showarrow=False)])
        st.markdown("**Completion Rate**")
        st.plotly_chart(fig_comp, use_container_width=True)

        fig_err = go.Figure(go.Pie(values=[error_rate, 100-error_rate], hole=0.7,
            marker_colors=[EY_RED,"#E0E0E0"], textinfo="none", labels=["Errors","Clean"]))
        fig_err.update_layout(height=180, margin=dict(l=10,r=10,t=10,b=10), showlegend=False,
            annotations=[dict(text=f"<b>{error_rate:.1f}%</b>",x=0.5,y=0.5,font_size=18,showarrow=False)])
        st.markdown("**Error Rate**")
        st.plotly_chart(fig_err, use_container_width=True)

    st.markdown("---")
    col_tbl, col_hybrid = st.columns(2)

    with col_tbl:
        st.subheader("Task Breakdown")
        tb_display = task_breakdown.copy()
        tb_display["Total"] = tb_display["Not started"] + tb_display["In Progress"] + tb_display["Completed"]
        tb_display["Completion %"] = (tb_display["Completed"] / tb_display["Total"].replace(0, 1) * 100).round(1)
        tb_display["RAG"] = tb_display["Completion %"].apply(lambda x: rag_status(x, 80, 50))
        st.dataframe(tb_display, use_container_width=True, hide_index=True)

    with col_hybrid:
        st.subheader("Status of Hybrid Data Collection")
        fig_hybrid = go.Figure()
        for status, color in [("Completed",EY_GREEN),("In Progress",EY_ORANGE),("Not started",EY_RED)]:
            fig_hybrid.add_trace(go.Bar(y=hybrid_status["Dimension"], x=hybrid_status[status],
                name=status, orientation="h", marker_color=color))
        fig_hybrid.update_layout(barmode="stack", height=350, margin=dict(l=20,r=20,t=10,b=30),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1), plot_bgcolor="white")
        st.plotly_chart(fig_hybrid, use_container_width=True)

    st.subheader("Monthly Case Trend")
    filtered = monthly_cases[monthly_cases["Task"].isin(selected_tasks)]
    monthly_agg = filtered.groupby("Month")["Actual"].sum().reset_index()
    fig_trend = px.line(monthly_agg, x="Month", y="Actual", markers=True)
    fig_trend.update_traces(line_color=EY_BLUE, line_width=3)
    fig_trend.update_layout(height=300, margin=dict(l=20,r=20,t=10,b=30), plot_bgcolor="white")
    st.plotly_chart(fig_trend, use_container_width=True)

# ===================================================================== #
# PAGE: Traffic Light                                                     #
# ===================================================================== #
elif page == "Traffic Light":
    st.subheader("Traffic Light Dashboard")
    st.markdown("RAG status across all key performance indicators.")

    rag_comp = rag_status(completion_rate, 80, 60)
    rag_err = rag_status(error_rate, 2, 5, invert=True)
    rag_avail = rag_status(resource_availability, 90, 70)
    rag_ht = rag_status(handling_time, 5, 8, invert=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(traffic_card("Completion Rate", f"{completion_rate:.1f}%", rag_comp), unsafe_allow_html=True)
    c2.markdown(traffic_card("Error Rate", f"{error_rate:.1f}%", rag_err), unsafe_allow_html=True)
    c3.markdown(traffic_card("FTE Availability", f"{resource_availability:.0f}%", rag_avail), unsafe_allow_html=True)
    c4.markdown(traffic_card("Avg Handling Time", f"{handling_time:.2f} hrs", rag_ht), unsafe_allow_html=True)

    st.markdown("---")

    st.subheader("Task-Level Progress")
    tb = task_breakdown.copy()
    tb["Total"] = tb["Not started"] + tb["In Progress"] + tb["Completed"]
    tb["Completion %"] = (tb["Completed"] / tb["Total"].replace(0, 1) * 100).round(1)
    tb["RAG"] = tb["Completion %"].apply(lambda x: rag_status(x, 80, 50))

    for _, row in tb.iterrows():
        status = row["RAG"]
        color = RAG_COLORS[status]
        pct = row["Completion %"]
        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.markdown(
                f"<div style='background:white; padding:12px; border-radius:8px; "
                f"border-left:5px solid {color}; margin-bottom:6px; "
                f"box-shadow:0 1px 3px rgba(0,0,0,0.08);'>"
                f"<b>{row['Task']}</b> &nbsp; {rag_html(status, f'{pct}%')}"
                f"<div style='margin-top:6px; background:#E0E0E0; border-radius:4px; height:10px;'>"
                f"<div style='width:{min(pct,100)}%; background:{color}; height:10px; border-radius:4px;'></div>"
                f"</div>"
                f"<small style='color:#747480;'>Not started: {row['Not started']:,} | "
                f"In Progress: {row['In Progress']:,} | Completed: {row['Completed']:,}</small>"
                f"</div>", unsafe_allow_html=True
            )
        with col_b:
            fig_mini = go.Figure(go.Pie(
                values=[row["Completed"], row["In Progress"], row["Not started"]],
                hole=0.6, marker_colors=[EY_GREEN, EY_ORANGE, EY_RED],
                textinfo="none", hoverinfo="label+value",
                labels=["Completed", "In Progress", "Not started"]
            ))
            fig_mini.update_layout(height=100, margin=dict(l=0,r=0,t=0,b=0), showlegend=False)
            st.plotly_chart(fig_mini, use_container_width=True)

    st.markdown("---")

    st.subheader("Hybrid Data Collection RAG")
    hs = hybrid_status.copy()
    hs["Total"] = hs["Completed"] + hs["In Progress"] + hs["Not started"]
    hs["Completion %"] = (hs["Completed"] / hs["Total"] * 100).round(1)
    hs["RAG"] = hs["Completion %"].apply(lambda x: rag_status(x, 70, 50))

    cols = st.columns(3)
    for i, (_, row) in enumerate(hs.iterrows()):
        with cols[i % 3]:
            st.markdown(traffic_card(row["Dimension"], f"{row['Completion %']}%", row["RAG"]),
                        unsafe_allow_html=True)

    st.markdown("---")

    st.subheader("Forecast RAG")
    weeks_to_complete = int(np.ceil(remaining_cases / avg_weekly_tp)) if avg_weekly_tp > 0 else 999
    est_date = datetime.now() + timedelta(weeks=weeks_to_complete)
    fc_rag = rag_status(weeks_to_complete, 4, 12, invert=True)

    fc1, fc2, fc3 = st.columns(3)
    fc1.markdown(traffic_card("Remaining Cases", f"{remaining_cases:,}", fc_rag), unsafe_allow_html=True)
    fc2.markdown(traffic_card("Weeks to Complete", str(weeks_to_complete), fc_rag), unsafe_allow_html=True)
    fc3.markdown(traffic_card("Est. Completion", est_date.strftime("%d %b %Y"), fc_rag), unsafe_allow_html=True)

# ===================================================================== #
# PAGE: Ad Hoc Tasks                                                      #
# ===================================================================== #
elif page == "Ad Hoc Tasks":
    st.subheader("Ad Hoc Task Management")
    st.markdown("Assign one-off tasks to individual FTEs. This reduces their effective capacity in the forecast.")

    if "adhoc" not in st.session_state:
        st.session_state.adhoc = load_adhoc()

    adhoc = st.session_state.adhoc

    open_tasks = len([t for t in adhoc if t["status"] != "Completed"])
    overdue = len([t for t in adhoc if t["status"] != "Completed" and t["due"] < datetime.now().strftime("%Y-%m-%d")])
    total_hrs = sum(t["hours"] for t in adhoc if t["status"] != "Completed")
    total_capacity = ftes * 40
    effective_cap = total_capacity - total_hrs

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Open Tasks", open_tasks)
    m2.metric("Overdue", overdue, delta=f"{-overdue}" if overdue else "0", delta_color="inverse")
    m3.metric("Ad Hoc Hrs Allocated", f"{total_hrs} hrs")
    m4.metric("Effective Capacity", f"{effective_cap} / {total_capacity} hrs/wk")

    st.markdown("---")

    with st.form("add_adhoc", clear_on_submit=True):
        st.markdown("**Assign New Ad Hoc Task**")
        ac1, ac2, ac3, ac4 = st.columns([3, 1, 1, 1])
        ad_desc = ac1.text_input("Task description")
        ad_fte = ac2.selectbox("Assign to FTE", fte_roster["FTE_Name"].tolist())
        ad_priority = ac3.selectbox("Priority", ["High", "Medium", "Low"])
        ad_hours = ac4.number_input("Est. hours", min_value=1, max_value=80, value=4)
        ac5, ac6, _ = st.columns([1, 1, 2])
        ad_due = ac5.date_input("Due date", value=datetime.now() + timedelta(days=7))
        ad_submit = st.form_submit_button("Add Task")
        if ad_submit and ad_desc.strip():
            new_id = max([t["id"] for t in adhoc], default=0) + 1
            adhoc.append({
                "id": new_id, "fte": ad_fte, "desc": ad_desc.strip(),
                "priority": ad_priority, "status": "Not Started", "hours": ad_hours,
                "assigned": datetime.now().strftime("%Y-%m-%d"),
                "due": ad_due.strftime("%Y-%m-%d"),
            })
            save_adhoc(adhoc)
            st.rerun()

    st.markdown("---")

    st.subheader("Tasks by FTE")
    ftes_with_tasks = sorted(set(t["fte"] for t in adhoc))

    for fte_name in ftes_with_tasks:
        fte_tasks = [t for t in adhoc if t["fte"] == fte_name]
        fte_open_hrs = sum(t["hours"] for t in fte_tasks if t["status"] != "Completed")
        load_pct = fte_open_hrs / 40 * 100
        load_rag = rag_status(load_pct, 25, 60, invert=True)

        with st.expander(f"{fte_name} -- {len(fte_tasks)} task(s), {fte_open_hrs}h allocated", expanded=False):
            for task in fte_tasks:
                is_overdue = task["status"] != "Completed" and task["due"] < datetime.now().strftime("%Y-%m-%d")
                p_colors = {"High": EY_RED, "Medium": EY_ORANGE, "Low": EY_GREEN}
                border = p_colors.get(task["priority"], EY_GRAY)
                overdue_badge = ' <span style="color:#C4122F; font-weight:700;">OVERDUE</span>' if is_overdue else ""

                st.markdown(
                    f"<div class='adhoc-card' style='border-left:4px solid {border};'>"
                    f"<b>{task['desc']}</b>{overdue_badge}<br>"
                    f"<small>Priority: {task['priority']} | Status: {task['status']} | "
                    f"Hours: {task['hours']} | Due: {task['due']}</small></div>",
                    unsafe_allow_html=True
                )

                uc1, uc2, uc3 = st.columns([2, 1, 1])
                new_st = uc1.selectbox("Status", ["Not Started", "In Progress", "Completed"],
                    index=["Not Started", "In Progress", "Completed"].index(task["status"]),
                    key=f"adhoc_st_{task['id']}")
                if uc2.button("Delete", key=f"adhoc_del_{task['id']}"):
                    st.session_state.adhoc = [t for t in adhoc if t["id"] != task["id"]]
                    save_adhoc(st.session_state.adhoc)
                    st.rerun()
                if new_st != task["status"]:
                    task["status"] = new_st
                    save_adhoc(adhoc)

    st.markdown("---")

    st.subheader("FTE Ad Hoc Load")
    load_data = []
    for _, fte in fte_roster.iterrows():
        hrs = sum(t["hours"] for t in adhoc if t["fte"] == fte["FTE_Name"] and t["status"] != "Completed")
        load_data.append({"FTE": fte["FTE_Name"], "Ad Hoc Hours": hrs, "Free Hours": 40 - hrs})
    load_df = pd.DataFrame(load_data)

    fig_load = go.Figure()
    fig_load.add_trace(go.Bar(name="Ad Hoc Hours", x=load_df["FTE"], y=load_df["Ad Hoc Hours"], marker_color=EY_RED))
    fig_load.add_trace(go.Bar(name="Free Hours", x=load_df["FTE"], y=load_df["Free Hours"], marker_color=EY_GREEN))
    fig_load.update_layout(barmode="stack", height=350, plot_bgcolor="white",
        margin=dict(l=20,r=20,t=30,b=30),
        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
        yaxis_title="Hours per Week")
    fig_load.add_hline(y=40, line_dash="dash", line_color=EY_GRAY, annotation_text="Full Capacity (40h)")
    st.plotly_chart(fig_load, use_container_width=True)

# ===================================================================== #
# PAGE: Forecasting                                                       #
# ===================================================================== #
elif page == "Forecasting":
    st.subheader("Case Handling Forecast")
    st.markdown("Predicts when all cases will be completed based on FTE availability, throughput history, and ad hoc task load.")

    if "adhoc" not in st.session_state:
        st.session_state.adhoc = load_adhoc()
    adhoc = st.session_state.adhoc
    adhoc_hrs = sum(t["hours"] for t in adhoc if t["status"] != "Completed")

    st.markdown("#### Scenario Planning")
    sc1, sc2 = st.columns(2)
    scenario_ftes = sc1.slider("FTEs Available", min_value=3, max_value=25, value=ftes)
    scenario_adhoc_hrs = sc2.slider("Ad Hoc Hours (total/week)", min_value=0, max_value=200, value=int(adhoc_hrs))

    effective_ftes = scenario_ftes - (scenario_adhoc_hrs / 40)
    effective_weekly_cap = max(effective_ftes * avg_cases_per_fte, 1)
    weeks_needed = int(np.ceil(remaining_cases / effective_weekly_cap))
    est_completion = datetime.now() + timedelta(weeks=weeks_needed)
    fc_rag = rag_status(weeks_needed, 4, 12, invert=True)

    st.markdown("---")

    fc1, fc2, fc3, fc4, fc5 = st.columns(5)
    fc1.markdown(traffic_card("Remaining Cases", f"{remaining_cases:,}", fc_rag), unsafe_allow_html=True)
    fc2.markdown(traffic_card("Avg Weekly Throughput", f"{avg_weekly_tp:.0f}", "Green"), unsafe_allow_html=True)
    fc3.markdown(traffic_card("Effective Capacity/wk", f"{effective_weekly_cap:.0f}",
        rag_status(effective_weekly_cap, 150, 100)), unsafe_allow_html=True)
    fc4.markdown(traffic_card("Weeks to Complete", str(weeks_needed), fc_rag), unsafe_allow_html=True)
    fc5.markdown(traffic_card("Est. Completion", est_completion.strftime("%d %b %Y"), fc_rag), unsafe_allow_html=True)

    st.markdown("---")

    st.subheader("Weekly Throughput Trend")
    fig_tp = go.Figure()
    fig_tp.add_trace(go.Scatter(x=weekly_tp["Week"], y=weekly_tp["Cases_Handled"],
        mode="lines+markers", name="Cases Handled", line=dict(color=EY_BLUE, width=2)))
    fig_tp.add_trace(go.Scatter(x=weekly_tp["Week"], y=weekly_tp["Cases_Incoming"],
        mode="lines+markers", name="Cases Incoming", line=dict(color=EY_ORANGE, width=2)))
    fig_tp.add_hline(y=avg_weekly_tp, line_dash="dash", line_color=EY_GREEN,
        annotation_text=f"Avg throughput: {avg_weekly_tp:.0f}/wk")
    fig_tp.update_layout(height=350, plot_bgcolor="white", margin=dict(l=20,r=20,t=30,b=30),
        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
    st.plotly_chart(fig_tp, use_container_width=True)

    st.subheader("Cumulative Progress & Projected Completion")
    cumulative = weekly_tp[["Week", "Cases_Handled"]].copy()
    cumulative["Cumulative"] = cumulative["Cases_Handled"].cumsum()

    future_weeks = pd.date_range(
        cumulative["Week"].max() + timedelta(weeks=1),
        periods=weeks_needed, freq="W-MON"
    )
    last_cum = cumulative["Cumulative"].iloc[-1]
    future_rows = []
    for i, w in enumerate(future_weeks):
        last_cum += effective_weekly_cap
        future_rows.append({"Week": w, "Cumulative": int(last_cum)})
    future_df = pd.DataFrame(future_rows)

    fig_cum = go.Figure()
    fig_cum.add_trace(go.Scatter(x=cumulative["Week"], y=cumulative["Cumulative"],
        mode="lines+markers", name="Actual (cumulative)", line=dict(color=EY_BLUE, width=3)))
    if len(future_df) > 0:
        fig_cum.add_trace(go.Scatter(x=future_df["Week"], y=future_df["Cumulative"],
            mode="lines+markers", name="Projected", line=dict(color=EY_ORANGE, width=3, dash="dash")))
    fig_cum.add_hline(y=total_all, line_dash="dot", line_color=EY_RED,
        annotation_text=f"Total cases: {total_all:,}")
    fig_cum.update_layout(height=400, plot_bgcolor="white", margin=dict(l=20,r=20,t=30,b=30),
        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
        xaxis_title="Week", yaxis_title="Cumulative Cases")
    st.plotly_chart(fig_cum, use_container_width=True)

    st.subheader("FTE Scenario Comparison")
    scenario_rows = []
    for n in [5, 7, 9, 11, 13, 15, 18, 20]:
        eff = max(n - (scenario_adhoc_hrs / 40), 0.5)
        cap = eff * avg_cases_per_fte
        wks = int(np.ceil(remaining_cases / cap)) if cap > 0 else 999
        dt = datetime.now() + timedelta(weeks=wks)
        r = rag_status(wks, 4, 12, invert=True)
        current = " (current)" if n == ftes else ""
        scenario_rows.append({
            "FTEs": f"{n}{current}", "Effective Capacity/wk": f"{cap:.0f} cases",
            "Weeks to Complete": wks, "Est. Completion": dt.strftime("%d %b %Y"), "RAG": r
        })
    st.dataframe(pd.DataFrame(scenario_rows), use_container_width=True, hide_index=True)

# ===================================================================== #
# PAGE: Task Board                                                        #
# ===================================================================== #
elif page == "Task Board":
    st.subheader("User Task Board")
    st.markdown("Manage your personal tasks below. Tasks are saved locally.")

    if "tasks" not in st.session_state:
        st.session_state.tasks = load_tasks()

    with st.form("add_task", clear_on_submit=True):
        st.markdown("**Add New Task**")
        c1, c2, c3 = st.columns([3, 1, 1])
        new_title = c1.text_input("Task description")
        new_assignee = c2.text_input("Assignee", value="Me")
        new_priority = c3.selectbox("Priority", ["High", "Medium", "Low"])
        submitted = st.form_submit_button("Add Task")
        if submitted and new_title.strip():
            st.session_state.tasks.append({
                "id": int(datetime.now().timestamp() * 1000),
                "title": new_title.strip(),
                "assignee": new_assignee.strip(),
                "priority": new_priority,
                "status": "Not started",
                "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
            save_tasks(st.session_state.tasks)
            st.rerun()

    if not st.session_state.tasks:
        st.info("No tasks yet. Add one above.")
    else:
        statuses = ["Not started", "In Progress", "Completed"]
        cols = st.columns(3)
        priority_colors = {"High": "red", "Medium": "orange", "Low": "green"}

        for idx, status in enumerate(statuses):
            with cols[idx]:
                st.markdown(f"### {status}")
                status_tasks = [t for t in st.session_state.tasks if t["status"] == status]
                if not status_tasks:
                    st.caption("No tasks")
                for task in status_tasks:
                    p_color = priority_colors.get(task["priority"], "gray")
                    st.markdown(
                        f"<div style='background:white; padding:10px; border-radius:8px; "
                        f"margin-bottom:8px; border-left:4px solid {p_color}; "
                        f"box-shadow:0 1px 3px rgba(0,0,0,0.1);'>"
                        f"<b>{task['title']}</b><br>"
                        f"<small>Assignee: {task['assignee']} | "
                        f"Priority: {task['priority']} | "
                        f"Created: {task['created']}</small></div>",
                        unsafe_allow_html=True
                    )

        st.markdown("---")
        st.markdown("**Update Tasks**")
        for i, task in enumerate(st.session_state.tasks):
            with st.expander(f"{task['title']} ({task['status']})"):
                c1, c2, c3 = st.columns([2, 1, 1])
                new_status = c1.selectbox("Status", statuses,
                    index=statuses.index(task["status"]), key=f"status_{task['id']}")
                new_pr = c2.selectbox("Priority", ["High", "Medium", "Low"],
                    index=["High", "Medium", "Low"].index(task["priority"]), key=f"pri_{task['id']}")
                if c3.button("Delete", key=f"del_{task['id']}"):
                    st.session_state.tasks = [t for t in st.session_state.tasks if t["id"] != task["id"]]
                    save_tasks(st.session_state.tasks)
                    st.rerun()
                if new_status != task["status"] or new_pr != task["priority"]:
                    task["status"] = new_status
                    task["priority"] = new_pr
                    save_tasks(st.session_state.tasks)

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("---")
st.caption("DCRS Dashboard | Nordea - Project Aragorn | Built with Streamlit")
