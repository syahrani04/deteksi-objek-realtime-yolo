import numpy as np
import pandas as pd
import requests
import joblib
import streamlit as st
from datetime import datetime

API_URL_TEMPLATE = "https://api.bmkg.go.id/publik/prakiraan-cuaca?adm4={adm4}"

# Preset lokasi (Bekasi - Rawalumbu) + opsi custom
PRESET_ADM4 = {
    "Bekasi - Bojong Rawalumbu (32.75.05.1001)": "32.75.05.1001",
    "Bekasi - Sepanjangjaya (32.75.05.1003)": "32.75.05.1003",
    "Custom (ketik kode adm4 sendiri)": "CUSTOM",
}

# -------------------------
# THEME PRESETS (DINAMIS)
# -------------------------
THEMES = {
    "sunny": {
        "name": "Cerah",
        "icon": "☀️",
        "accent": "#FDBA2D",
        "bg1": "#0B1220",
        "bg2": "#1D2A4A",
        "hero1": "rgba(253,186,45,0.18)",
        "hero2": "rgba(29,42,74,0.55)",
        "chip": "rgba(253,186,45,0.18)",
        "chip_border": "rgba(253,186,45,0.35)",
    },
    "cloudy": {
        "name": "Berawan",
        "icon": "☁️",
        "accent": "#7AA7FF",
        "bg1": "#0B1220",
        "bg2": "#1B2B3E",
        "hero1": "rgba(122,167,255,0.18)",
        "hero2": "rgba(27,43,62,0.60)",
        "chip": "rgba(122,167,255,0.16)",
        "chip_border": "rgba(122,167,255,0.30)",
    },
    "rainy": {
        "name": "Hujan",
        "icon": "🌧️",
        "accent": "#22C55E",
        "bg1": "#071A2B",
        "bg2": "#0F2A3A",
        "hero1": "rgba(34,197,94,0.16)",
        "hero2": "rgba(15,42,58,0.70)",
        "chip": "rgba(34,197,94,0.14)",
        "chip_border": "rgba(34,197,94,0.30)",
    },
    "storm": {
        "name": "Badai/Petir",
        "icon": "⛈️",
        "accent": "#A78BFA",
        "bg1": "#0B0F1A",
        "bg2": "#241B3B",
        "hero1": "rgba(167,139,250,0.16)",
        "hero2": "rgba(36,27,59,0.70)",
        "chip": "rgba(167,139,250,0.14)",
        "chip_border": "rgba(167,139,250,0.30)",
    },
    "fog": {
        "name": "Kabut",
        "icon": "🌫️",
        "accent": "#94A3B8",
        "bg1": "#0B1220",
        "bg2": "#111827",
        "hero1": "rgba(148,163,184,0.14)",
        "hero2": "rgba(17,24,39,0.70)",
        "chip": "rgba(148,163,184,0.12)",
        "chip_border": "rgba(148,163,184,0.28)",
    },
}

def choose_theme_key(pred_label: str | None, bmkg_desc: str | None, tcc: float | None) -> str:
    """
    Memilih tema berdasarkan:
    - deskripsi BMKG (jika ada)
    - prediksi model (hujan/tidak)
    - tcc (awan) sebagai fallback
    """
    desc = (bmkg_desc or "").lower()

    # prioritas: petir/badai
    if any(k in desc for k in ["petir", "guntur", "badai", "thunder"]):
        return "storm"
    # kabut
    if "kabut" in desc or "fog" in desc or "mist" in desc:
        return "fog"
    # hujan
    if "hujan" in desc:
        return "rainy"
    if pred_label == "Hujan":
        return "rainy"

    # fallback dari tutupan awan
    try:
        if tcc is not None and float(tcc) >= 75:
            return "cloudy"
    except Exception:
        pass

    return "sunny"

def apply_dynamic_css(theme_key: str):
    th = THEMES.get(theme_key, THEMES["cloudy"])

    st.markdown(
        f"""
<style>
:root {{
  --accent: {th["accent"]};
  --bg1: {th["bg1"]};
  --bg2: {th["bg2"]};
  --hero1: {th["hero1"]};
  --hero2: {th["hero2"]};
  --chip: {th["chip"]};
  --chipBorder: {th["chip_border"]};
}}

/* App background */
[data-testid="stAppViewContainer"] {{
  background: linear-gradient(135deg, var(--bg1), var(--bg2));
}}

/* Sidebar */
section[data-testid="stSidebar"] {{
  background: rgba(255,255,255,0.03);
  border-right: 1px solid rgba(255,255,255,0.07);
}}

/* Metrics as cards */
div[data-testid="stMetric"] {{
  background: rgba(255,255,255,0.045);
  border: 1px solid rgba(255,255,255,0.08);
  padding: 14px;
  border-radius: 14px;
}}

/* Dataframe container */
div[data-testid="stDataFrame"] {{
  background: rgba(255,255,255,0.02);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 14px;
  padding: 8px;
}}

/* Buttons */
.stButton button {{
  border-radius: 12px;
  padding: 0.6rem 1rem;
  font-weight: 650;
  border: 1px solid rgba(255,255,255,0.12);
}}
.stButton button:hover {{
  border: 1px solid rgba(255,255,255,0.20);
}}

/* Hero card */
.hero {{
  background: linear-gradient(135deg, var(--hero1), var(--hero2));
  border: 1px solid rgba(255,255,255,0.10);
  border-radius: 18px;
  padding: 18px 18px;
}}
.hero-title {{
  font-size: 26px;
  font-weight: 800;
  margin: 0;
}}
.hero-sub {{
  opacity: 0.90;
  margin-top: 6px;
  margin-bottom: 0;
}}
.chip {{
  display: inline-block;
  margin-top: 10px;
  padding: 6px 10px;
  border-radius: 999px;
  background: var(--chip);
  border: 1px solid var(--chipBorder);
  font-weight: 650;
}}
.small-muted {{
  opacity: 0.75;
  font-size: 12px;
}}
</style>
""",
        unsafe_allow_html=True,
    )

def render_hero(theme_key: str, lokasi_text: str, waktu_text: str, bmkg_desc: str, model_label: str, prob: float):
    th = THEMES.get(theme_key, THEMES["cloudy"])
    icon = th["icon"]

    st.markdown(
        f"""
<div class="hero">
  <div style="display:flex; align-items:center; gap:14px;">
    <div style="font-size:44px; line-height:1;">{icon}</div>
    <div style="flex:1;">
      <p class="hero-title">Prediksi Cuaca: {model_label}</p>
      <p class="hero-sub">{lokasi_text}</p>
      <p class="small-muted">Waktu prakiraan: {waktu_text}</p>
      <span class="chip">BMKG: {bmkg_desc if bmkg_desc else "-" } • Prob: {prob:.2f}</span>
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

# -------------------------
# STREAMLIT CONFIG
# -------------------------
st.set_page_config(
    page_title="Prediksi Cuaca BMKG + SVM",
    page_icon="🌦️",
    layout="wide",
)

# -------------------------
# LOAD MODEL + SCALER
# -------------------------
@st.cache_resource
def load_model_and_scaler():
    model = joblib.load("bmkg_svm_model.pkl")
    scaler = joblib.load("bmkg_svm_scaler.pkl")
    return model, scaler

# -------------------------
# BMKG FETCH
# -------------------------
@st.cache_data(ttl=300)  # cache 5 menit
def fetch_bmkg_forecasts(adm4_code: str):
    url = API_URL_TEMPLATE.format(adm4=adm4_code)
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    data = resp.json()

    lokasi = data.get("lokasi", {}) or {}
    meta = {
        "provinsi": lokasi.get("provinsi"),
        "kotkab": lokasi.get("kotkab"),
        "kecamatan": lokasi.get("kecamatan"),
        "desa": lokasi.get("desa"),
        "adm4": adm4_code,
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    cuaca_all = data.get("data", [])
    if not cuaca_all:
        raise ValueError("Response BMKG tidak memiliki field 'data'.")

    cuaca_harian = (cuaca_all[0] or {}).get("cuaca", [])
    if not cuaca_harian:
        raise ValueError("Response BMKG tidak memiliki field 'cuaca'.")

    records = []
    for day_list in cuaca_harian:
        if not isinstance(day_list, list):
            continue
        for entry in day_list:
            records.append(
                {
                    "local_datetime": entry.get("local_datetime"),
                    "t": entry.get("t"),
                    "hu": entry.get("hu"),
                    "ws": entry.get("ws"),
                    "tcc": entry.get("tcc"),
                    "weather_desc": entry.get("weather_desc"),
                }
            )

    df = pd.DataFrame(records)
    for col in ["t", "hu", "ws", "tcc"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    try:
        df["_dt"] = pd.to_datetime(df["local_datetime"], errors="coerce")
        df = df.sort_values("_dt").drop(columns=["_dt"])
    except Exception:
        pass

    df = df.dropna(subset=["t", "hu", "ws", "tcc"]).reset_index(drop=True)
    if df.empty:
        raise ValueError("Data prakiraan kosong setelah pembersihan (NaN).")

    return df, meta

def predict_label(model, scaler, t, hu, ws, tcc):
    X = np.array([[t, hu, ws, tcc]], dtype=float)
    Xs = scaler.transform(X)
    pred = int(model.predict(Xs)[0])
    proba = model.predict_proba(Xs)[0]
    label = "Hujan" if pred == 1 else "Tidak Hujan"
    prob = float(proba[1] if pred == 1 else proba[0])
    return label, prob

# -------------------------
# SIDEBAR
# -------------------------
st.sidebar.header("Kontrol Aplikasi")
mode = st.sidebar.radio("Mode", ["Ambil dari BMKG", "Input Manual"])

st.sidebar.markdown("---")
st.sidebar.subheader("Lokasi (BMKG)")
preset = st.sidebar.selectbox("Pilih lokasi", list(PRESET_ADM4.keys()))
adm4_code = PRESET_ADM4[preset]
if adm4_code == "CUSTOM":
    adm4_code = st.sidebar.text_input("Kode adm4", value="32.75.05.1003")

st.sidebar.markdown("---")
max_rows = st.sidebar.slider("Jumlah baris tabel prakiraan", 8, 40, 24, 1)
show_chart = st.sidebar.checkbox("Tampilkan grafik ringkas", value=True)
refresh = st.sidebar.button("Ambil / Refresh Data BMKG")

# -------------------------
# MAIN HEADER
# -------------------------
st.title("Aplikasi Prediksi Kategori Cuaca (BMKG + SVM)")
st.write(
    "Aplikasi ini memprediksi **Hujan** atau **Tidak Hujan** menggunakan model **SVM** "
    "dari fitur: **t (suhu), hu (kelembapan), ws (angin), tcc (awan)**."
)
st.caption("Sumber data prakiraan: BMKG (Prakiraan Cuaca).")

# Load model
try:
    model, scaler = load_model_and_scaler()
except Exception as e:
    st.error(
        "Model/scaler belum tersedia. Jalankan training dulu (bmkg_train_svm.py) "
        "hingga file `bmkg_svm_model.pkl` dan `bmkg_svm_scaler.pkl` terbentuk.\n\n"
        f"Detail error: {e}"
    )
    st.stop()

# -------------------------
# CONTENT
# -------------------------
if mode == "Input Manual":
    st.subheader("Input Manual Parameter Cuaca")

    col1, col2 = st.columns(2)
    with col1:
        t = st.slider("🌡️ Suhu (°C)", 0.0, 40.0, 28.0, 0.5)
        hu = st.slider("💧 Kelembapan (%)", 0.0, 100.0, 80.0, 1.0)
    with col2:
        ws = st.slider("🌬️ Kecepatan Angin (km/jam)", 0.0, 50.0, 6.0, 0.5)
        tcc = st.slider("☁️ Tutupan Awan (%)", 0.0, 100.0, 70.0, 1.0)

    # Prediksi langsung (agar tema dinamis juga mengikuti slider)
    model_label, prob = predict_label(model, scaler, t, hu, ws, tcc)
    theme_key = choose_theme_key(model_label, bmkg_desc=None, tcc=tcc)
    apply_dynamic_css(theme_key)

    # Hero banner
    render_hero(
        theme_key=theme_key,
        lokasi_text="Mode: Input Manual",
        waktu_text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        bmkg_desc="(manual)",
        model_label=model_label,
        prob=prob,
    )

    # Metric cards
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🌡️ Suhu (t)", f"{t:.1f} °C")
    m2.metric("💧 Kelembapan (hu)", f"{hu:.0f} %")
    m3.metric("🌬️ Angin (ws)", f"{ws:.1f} km/jam")
    m4.metric("☁️ Tutupan Awan (tcc)", f"{tcc:.0f} %")

    st.info(
        "Catatan: Hasil prediksi adalah keluaran model pembelajaran berbasis pola data. "
        "Untuk informasi resmi, tetap rujuk BMKG."
    )

else:
    st.subheader("Prakiraan BMKG (3 hari, per 3 jam)")

    if (
        refresh
        or ("forecast_df" not in st.session_state)
        or (st.session_state.get("adm4_last") != adm4_code)
    ):
        with st.spinner("Mengambil data dari BMKG..."):
            df_forecast, meta = fetch_bmkg_forecasts(adm4_code)
            st.session_state["forecast_df"] = df_forecast
            st.session_state["meta"] = meta
            st.session_state["adm4_last"] = adm4_code

    df_forecast = st.session_state["forecast_df"]
    meta = st.session_state["meta"]

    lokasi_text = f'{meta.get("desa")}, {meta.get("kecamatan")}, {meta.get("kotkab")}, {meta.get("provinsi")} (adm4={meta.get("adm4")})'
    st.info(f"Lokasi: {lokasi_text} • Last fetch: {meta.get('fetched_at')}")

    # Pilih waktu prakiraan
    time_options = df_forecast["local_datetime"].astype(str).tolist()
    chosen_time = st.selectbox("Pilih waktu prakiraan", time_options)

    chosen_row = df_forecast[df_forecast["local_datetime"].astype(str) == str(chosen_time)].iloc[0]
    t = float(chosen_row["t"])
    hu = float(chosen_row["hu"])
    ws = float(chosen_row["ws"])
    tcc = float(chosen_row["tcc"])
    bmkg_desc = str(chosen_row.get("weather_desc", ""))

    # Prediksi model
    model_label, prob = predict_label(model, scaler, t, hu, ws, tcc)

    # Tema dinamis: gabungkan BMKG desc + prediksi + tcc
    theme_key = choose_theme_key(model_label, bmkg_desc=bmkg_desc, tcc=tcc)
    apply_dynamic_css(theme_key)

    # Hero banner
    render_hero(
        theme_key=theme_key,
        lokasi_text=lokasi_text,
        waktu_text=str(chosen_time),
        bmkg_desc=bmkg_desc,
        model_label=model_label,
        prob=prob,
    )

    # Metric cards
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🌡️ Suhu (t)", f"{t:.1f} °C")
    m2.metric("💧 Kelembapan (hu)", f"{hu:.0f} %")
    m3.metric("🌬️ Angin (ws)", f"{ws:.1f} km/jam")
    m4.metric("☁️ Tutupan Awan (tcc)", f"{tcc:.0f} %")

    # Grafik ringkas (opsional)
    if show_chart:
        st.markdown("---")
        st.subheader("Grafik Ringkas Prakiraan")
        df_plot = df_forecast.copy()
        # coba parse waktu agar rapi di chart
        try:
            df_plot["waktu"] = pd.to_datetime(df_plot["local_datetime"], errors="coerce")
        except Exception:
            df_plot["waktu"] = df_plot["local_datetime"]

        df_plot = df_plot[["waktu", "t", "hu", "ws", "tcc"]].dropna()
        df_plot = df_plot.set_index("waktu")

        cA, cB = st.columns(2)
        with cA:
            st.line_chart(df_plot[["t", "hu"]], use_container_width=True)
        with cB:
            st.line_chart(df_plot[["ws", "tcc"]], use_container_width=True)

    # Tabel
    st.markdown("---")
    st.subheader("Tabel Prakiraan (BMKG)")
    show_cols = ["local_datetime", "t", "hu", "ws", "tcc", "weather_desc"]
    df_show = df_forecast[show_cols].head(max_rows).copy()
    df_show = df_show.rename(
        columns={
            "local_datetime": "Waktu",
            "t": "Suhu (°C)",
            "hu": "Kelembapan (%)",
            "ws": "Angin (km/jam)",
            "tcc": "Awan (%)",
            "weather_desc": "Deskripsi",
        }
    )
    st.dataframe(df_show, use_container_width=True)

    st.caption(
        "Catatan: Model ini dibuat untuk pembelajaran dan meniru pola dari data prakiraan. "
        "BMKG tetap menjadi rujukan resmi."
    )
