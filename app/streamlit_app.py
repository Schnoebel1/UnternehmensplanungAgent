# File: app/streamlit_app.py

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# --- Page config ------------------------------------------------------------
st.set_page_config(
    page_title="Forecast Simulator",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Sidebar: Steuerung + Navigation --------------------------------------
st.sidebar.header("🔧 Steuerung")
szenario = st.sidebar.slider(
    "🔮 Szenariostärke", min_value=0, max_value=10, value=5
)
if st.sidebar.button("🚀 Simulation starten"):
    st.sidebar.success("Simulation gestartet!")

st.sidebar.markdown("---")
page = st.sidebar.radio(
    "📊 Seiten",
    ["Übersicht", "Umsatz", "EBIT-Marge", "Cashflow", "FCF", "Kapitalrendite"]
)

# --- Dummy-Daten für Charts ------------------------------------------------
def make_dummy_timeseries(base: float):
    t = [1, 2, 3]
    vals = np.array([base * (1 + 0.1 * i) for i in range(len(t))])
    return pd.DataFrame({"Periode": t, "Wert": vals})

# --- Seite: Übersicht ------------------------------------------------------
if page == "Übersicht":
    st.title("🚀 Forecast Simulator – Übersicht")
    st.markdown(
        "Dieses Dashboard gibt dir eine Schnellübersicht über die wichtigsten KPIs deiner Bilanz-Forecast-Simulation."
    )

    # KPI-Metriken
    kpis = {
        "Umsatz t1": 1_200,
        "Umsatz t2": 1_440,
        "Umsatz t3": 1_680,
        "EBIT-Marge": 12.0,
        "Free Cashflow": 250
    }
    cols = st.columns(len(kpis))
    for col, (label, value) in zip(cols, kpis.items()):
        col.metric(label, f"{value:,}" + (" %" if "Marge" in label else ""))

    st.markdown("---")

    # Umsatz-Verlauf als Beispielgrafik
    df_over = make_dummy_timeseries(1200)
    chart = (
        alt.Chart(df_over)
        .mark_line(point=True)
        .encode(
            x=alt.X("Periode:O", title="Periode"),
            y=alt.Y("Wert:Q", title="Wert"),
        )
        .properties(height=300)  # width-Parameter entfernt!
    )
    st.altair_chart(chart, use_container_width=True)

# --- Seite: Umsatz ---------------------------------------------------------
elif page == "Umsatz":
    st.title("📈 Umsatz-Forecast")
    st.markdown("**Umsatz-KPIs** und grafischer Verlauf für t1–t3.")

    df = make_dummy_timeseries(1200)
    st.metric("Umsatz t1", f"{df.loc[0,'Wert']:.0f}")
    st.metric("Umsatz t2", f"{df.loc[1,'Wert']:.0f}")
    st.metric("Umsatz t3", f"{df.loc[2,'Wert']:.0f}")

    st.subheader("Verlauf")
    chart1 = (
        alt.Chart(df)
        .mark_line(point=True)
        .encode(x="Periode:O", y="Wert:Q")
        .properties(height=300)
    )
    st.altair_chart(chart1, use_container_width=True)

    st.subheader("Balken-Ansicht")
    chart2 = (
        alt.Chart(df)
        .mark_bar()
        .encode(x="Periode:O", y="Wert:Q")
        .properties(height=200)
    )
    st.altair_chart(chart2, use_container_width=True)

# --- Seite: EBIT-Marge -----------------------------------------------------
elif page == "EBIT-Marge":
    st.title("💹 EBIT-Marge-Forecast")
    df = make_dummy_timeseries(12.0)
    st.metric("EBIT-Marge t1", f"{df.loc[0,'Wert']:.1f}%")
    st.metric("EBIT-Marge t2", f"{df.loc[1,'Wert']:.1f}%")
    st.metric("EBIT-Marge t3", f"{df.loc[2,'Wert']:.1f}%")

    st.subheader("Marge-Verlauf")
    chart3 = (
        alt.Chart(df)
        .mark_line(point=True)
        .encode(x="Periode:O", y=alt.Y("Wert:Q", title="Marge (%)"))
        .properties(height=300)
    )
    st.altair_chart(chart3, use_container_width=True)

# --- Seite: Cashflow -------------------------------------------------------
elif page == "Cashflow":
    st.title("💧 Cashflow-Forecast")
    df = make_dummy_timeseries(250)
    st.metric("Cashflow t1", f"{df.loc[0,'Wert']:.0f}")
    st.metric("Cashflow t2", f"{df.loc[1,'Wert']:.0f}")
    st.metric("Cashflow t3", f"{df.loc[2,'Wert']:.0f}")

    st.subheader("Verlauf")
    chart4 = (
        alt.Chart(df)
        .mark_area(opacity=0.5)
        .encode(x="Periode:O", y="Wert:Q")
        .properties(height=300)
    )
    st.altair_chart(chart4, use_container_width=True)

# --- Seite: FCF ------------------------------------------------------------
elif page == "FCF":
    st.title("🎯 Free Cashflow (FCF)-Forecast")
    df = make_dummy_timeseries(200)
    st.metric("FCF t1", f"{df.loc[0,'Wert']:.0f}")
    st.metric("FCF t2", f"{df.loc[1,'Wert']:.0f}")
    st.metric("FCF t3", f"{df.loc[2,'Wert']:.0f}")

    st.subheader("Verlauf")
    chart5 = (
        alt.Chart(df)
        .mark_line(point=True)
        .encode(x="Periode:O", y="Wert:Q")
        .properties(height=300)
    )
    st.altair_chart(chart5, use_container_width=True)

# --- Seite: Kapitalrendite ------------------------------------------------
elif page == "Kapitalrendite":
    st.title("📊 Kapitalrendite-Forecast")
    df = make_dummy_timeseries(8.0)
    st.metric("RoI t1", f"{df.loc[0,'Wert']:.1f}%")
    st.metric("RoI t2", f"{df.loc[1,'Wert']:.1f}%")
    st.metric("RoI t3", f"{df.loc[2,'Wert']:.1f}%")

    st.subheader("Verlauf")
    chart6 = (
        alt.Chart(df)
        .mark_bar()
        .encode(x="Periode:O", y=alt.Y("Wert:Q", title="ROI (%)"))
        .properties(height=300)
    )
    st.altair_chart(chart6, use_container_width=True)
