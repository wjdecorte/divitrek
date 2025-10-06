import os

import pandas as pd
import streamlit as st

API_BASE = os.environ.get("API_BASE", "http://localhost:8000")


st.set_page_config(page_title="DiviTrek", layout="wide")
st.title("DiviTrek - Dividend Tracker")


@st.cache_data(show_spinner=False)
def fetch_assets() -> pd.DataFrame:
    import httpx

    with httpx.Client() as client:
        r = client.get(f"{API_BASE}/assets/")
        r.raise_for_status()
        return pd.DataFrame(r.json())


def create_asset(symbol: str, name: str, kind: str) -> None:
    import httpx

    payload = {"symbol": symbol.upper(), "name": name, "type": kind}
    with httpx.Client() as client:
        r = client.post(f"{API_BASE}/assets/", json=payload)
        if r.status_code >= 400:
            st.error(r.json().get("detail", "Error creating asset"))
        else:
            st.success("Asset created")
            fetch_assets.clear()


tab1, tab2 = st.tabs(["Assets", "Data Entry"])

with tab1:
    assets_df = fetch_assets()
    st.dataframe(assets_df, use_container_width=True)

with tab2:
    st.subheader("Add Asset")
    c1, c2, c3, c4 = st.columns([2, 3, 2, 1])
    with c1:
        symbol = st.text_input("Symbol", placeholder="VTI")
    with c2:
        name = st.text_input("Name", placeholder="Vanguard Total Stock Market ETF")
    with c3:
        kind = st.selectbox("Type", ["stock", "etf"], index=1)
    with c4:
        st.write("")
        if st.button("Create"):
            if symbol and name:
                create_asset(symbol, name, kind)
            else:
                st.warning("Provide both symbol and name")
