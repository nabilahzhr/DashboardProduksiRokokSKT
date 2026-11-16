import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

# ============================================================
# KONFIGURASI HALAMAN & TEMA
# ============================================================
st.set_page_config(
    page_title="Dashboard Kualitas SKT",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

BROWN = "#8B6B52"
DARK = "#594638"
LATTE = "#C5A58A"
BEIGE = "#E8D8C8"
CREAM = "#F7F2EC"
INK = "#493A30"
MUTED = "#8A796B"
RED = "#B85C50"
GREEN = "#66836A"

st.markdown(f"""
<style>
    .stApp {{ background-color: {CREAM}; color: {INK}; }}
    [data-testid="stSidebar"] {{ background-color: {DARK}; }}
    [data-testid="stSidebar"] * {{ color: #FFF8F0 !important; }}
    h1, h2, h3 {{ color: {DARK}; }}
    div[data-testid="stMetric"] {{
        background: #FFFFFF; border: 1px solid {BEIGE};
        padding: 16px; border-radius: 12px;
    }}
    div.stButton > button {{
        background: {BROWN}; color: white; border: 0; border-radius: 8px;
    }}
    div[data-testid="stAlert"] {{ border-radius: 10px; }}
    .section-note {{
        background: #EFE5DA; padding: 12px 16px; border-radius: 10px;
        border-left: 4px solid {LATTE}; margin-bottom: 12px;
    }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# PEMBACAAN WORKBOOK
# ============================================================
DEFAULT_FILE = Path(__file__).parent / "data" / "KLIK-KLIKAN FIXXXX(1).xlsx"

@st.cache_data
def load_workbook(file_bytes=None):
    if file_bytes is None:
        path = DEFAULT_FILE
        if not path.exists():
            raise FileNotFoundError(
                "File Excel tidak ditemukan. Letakkan workbook di folder data "
                "dengan nama KLIK-KLIKAN FIXXXX(1).xlsx."
            )
        excel = pd.ExcelFile(path)
    else:
        excel = pd.ExcelFile(file_bytes)

    # Baris ketiga pada sheet data adalah header tabel.
    p1 = pd.read_excel(excel, sheet_name="Peta Kendali 1", header=2)
    p2 = pd.read_excel(excel, sheet_name="Peta Kendali 2", header=2)
    run = pd.read_excel(excel, sheet_name="Uji Keacakan", header=2)
    cap = pd.read_excel(excel, sheet_name="Kapabilitas", header=2)
    sigma = pd.read_excel(excel, sheet_name="Level Sigma ", header=2)
    fmea = pd.read_excel(excel, sheet_name="FMEA", header=0)

    # Membersihkan spasi nama kolom dan baris ringkasan "Jumlah".
    for df in (p1, p2, run, cap, sigma):
        df.columns = [str(c).strip() for c in df.columns]

    # Filter baris data secara aman. pd.notna dipakai agar hasil konversi
    # numerik tidak bergantung pada tipe Series atau scalar dari workbook.
    def keep_numeric_rows(frame, column):
        if column not in frame.columns:
            raise ValueError(f"Kolom '{column}' tidak ditemukan pada salah satu sheet workbook.")
        values = frame[column]
        if isinstance(values, pd.DataFrame):  # antisipasi nama kolom duplikat
            values = values.iloc[:, 0]
        numeric_values = pd.to_numeric(values, errors="coerce")
        return frame.loc[pd.notna(numeric_values)].copy()

    p1 = keep_numeric_rows(p1, "Subgrup")
    p2 = keep_numeric_rows(p2, "Subgrup")
    run = keep_numeric_rows(run, "Subgrup")
    cap = keep_numeric_rows(cap, "No")
    sigma = keep_numeric_rows(sigma, "Subgrup")

    for df in (p1, p2, run, cap, sigma):
        date_col = "Day"
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    fmea.columns = [str(c).strip() for c in fmea.columns]
    for c in ["S", "O", "D", "RPN"]:
        if c in fmea.columns:
            fmea[c] = pd.to_numeric(fmea[c], errors="coerce")

    return p1, p2, run, cap, sigma, fmea

st.sidebar.markdown("## QUALITY CONTROL")
st.sidebar.caption("MONITORING PRODUKSI ROKOK SKT")
uploaded = st.sidebar.file_uploader("Unggah workbook Excel (.xlsx)", type=["xlsx"])

try:
    if uploaded is not None:
        p1, p2, run_df, cap_df, sigma_df, fmea = load_workbook(uploaded.getvalue())
        source_label = uploaded.name
    else:
        p1, p2, run_df, cap_df, sigma_df, fmea = load_workbook()
        source_label = "KLIK-KLIKAN FIXXXX(1).xlsx"
except Exception as e:
    st.error(f"Gagal membaca workbook: {e}")
    st.info("Pastikan file Excel berada di folder data atau unggah melalui sidebar.")
    st.stop()

def numeric(df, col):
    if col not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return pd.to_numeric(df[col], errors="coerce")

def get_defect_columns(df):
    candidates = [
        "Gembos", "Lem ambri tengah", "Lem ambri ujung",
        "Guntingan", "Korep", "Besar kepala"
    ]
    return [c for c in candidates if c in df.columns]

def filtered_dates(df, start_date, end_date):
    out = df.copy()
    if "Day" in out.columns:
        dates = pd.to_datetime(out["Day"], errors="coerce")
        out = out[(dates.dt.date >= start_date) & (dates.dt.date <= end_date)]
    return out

# Periode mengikuti seluruh tanggal yang tersedia pada workbook.
date_values = pd.to_datetime(p1["Day"], errors="coerce").dropna()
if date_values.empty:
    st.error("Tidak ada tanggal produksi yang valid pada sheet Peta Kendali 1.")
    st.stop()
min_date = date_values.min().date()
max_date = date_values.max().date()

st.sidebar.markdown("---")
st.sidebar.markdown("### Filter Periode")
date_range = st.sidebar.date_input(
    "Tanggal produksi",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

data = filtered_dates(p1, start_date, end_date)
data2 = filtered_dates(p2, start_date, end_date)
run_data = filtered_dates(run_df, start_date, end_date)

defects = get_defect_columns(data)
total_sample = numeric(data, "Total Sampel").sum()
total_defect = numeric(data, "Jumlah Rokok cacat").sum()
defect_rate = total_defect / total_sample if total_sample else np.nan

# ============================================================
# NAVIGASI
# ============================================================
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "NAVIGASI",
    ["Beranda", "Measure", "Analyze", "Improve", "Hasil Akhir"],
    label_visibility="visible"
)
st.sidebar.markdown("---")
st.sidebar.caption(f"Sumber data: {source_label}")
st.sidebar.caption(f"Periode tersedia: {min_date:%d %b %Y} – {max_date:%d %b %Y}")

# Header utama
st.markdown(
    f"<div style='background:{DARK};padding:22px 26px;border-radius:14px;margin-bottom:20px'>"
    f"<div style='color:#E8D8C8;font-size:12px;letter-spacing:2px'>QUALITY CONTROL DASHBOARD</div>"
    f"<div style='color:#FFF8F0;font-size:27px;font-weight:700'>Monitoring Kualitas Produksi Rokok SKT</div>"
    f"<div style='color:#E8D8C8;font-size:14px'>Pengendalian Kualitas menggunakan Metode Six Sigma · PT. XYZ</div>"
    f"</div>",
    unsafe_allow_html=True
)

# ============================================================
# BERANDA
# ============================================================
if page == "Beranda":
    st.subheader("Ringkasan Kualitas Produksi")
    st.markdown(
        "<div class='section-note'>Informasi berikut dihitung dari data produksi pada periode yang dipilih di sidebar.</div>",
        unsafe_allow_html=True
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sampel", f"{total_sample:,.0f}".replace(",", "."))
    c2.metric("Jumlah Produk Cacat", f"{total_defect:,.0f}".replace(",", "."))
    c3.metric("Proporsi Cacat", f"{defect_rate:.2%}" if pd.notna(defect_rate) else "—")
    c4.metric("Jumlah Subgrup", f"{len(data):,}".replace(",", "."))

    left, right = st.columns([1.2, 1])
    with left:
        st.markdown("#### Tren Proporsi Cacat")
        if not data.empty:
            fig = px.line(
                data, x="Day", y="Proporsi Cacat", markers=True,
                labels={"Day": "Tanggal Produksi", "Proporsi Cacat": "Proporsi Cacat"},
                color_discrete_sequence=[BROWN]
            )
            fig.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                              font_color=INK, margin=dict(l=10,r=10,t=25,b=10))
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown("#### Komposisi Jenis Kecacatan")
        if defects:
            sums = data[defects].apply(pd.to_numeric, errors="coerce").sum().sort_values(ascending=False)
            fig = px.pie(
                values=sums.values, names=sums.index, hole=0.55,
                color_discrete_sequence=["#594638", "#8B6B52", "#C5A58A",
                                         "#D8BFA8", "#A88970", "#E8D8C8"]
            )
            fig.update_layout(paper_bgcolor="white", font_color=INK,
                              margin=dict(l=10,r=10,t=25,b=10), legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### Data Produksi")
    show_cols = [c for c in ["Subgrup", "Day", *defects, "Jumlah Rokok cacat", "Total Sampel", "Proporsi Cacat"] if c in data.columns]
    st.dataframe(data[show_cols], use_container_width=True, hide_index=True)

# ============================================================
# MEASURE
# ============================================================
elif page == "Measure":
    st.subheader("Tahap Measure")
    st.markdown("Pengukuran kualitas proses berdasarkan P-Chart, Run Test, kapabilitas atribut, dan metrik Six Sigma.")

    st.markdown("### 1. Peta Kendali P")
    st.warning(
        "Workbook memuat dua versi P-Chart dengan garis tengah berbeda: "
        "Peta Kendali 1 sekitar 0,060470 dan Peta Kendali 2 sekitar 0,057468. "
        "Keduanya ditampilkan sebagai pilihan pembanding; dashboard belum menetapkan salah satunya sebagai hasil final."
    )
    chart_choice = st.selectbox(
        "Pilih versi P-Chart untuk ditampilkan",
        ["Bandingkan kedua versi", "Peta Kendali 1", "Peta Kendali 2"],
        index=0
    )

    def draw_pchart(df, title):
        required = ["Day", "Proporsi Cacat", "BKA", "BKB", "GT"]
        if not all(c in df.columns for c in required):
            st.error(f"Kolom P-Chart tidak lengkap pada {title}.")
            return
        d = df.copy()
        for c in required[1:]:
            d[c] = pd.to_numeric(d[c], errors="coerce")
        d = d.dropna(subset=["Day", "Proporsi Cacat", "BKA", "BKB", "GT"])
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=d["Day"], y=d["Proporsi Cacat"], mode="lines+markers",
                                 name="Proporsi Cacat", line=dict(color=BROWN, width=2),
                                 marker=dict(size=6, color=BROWN)))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["BKA"], mode="lines",
                                 name="BKA", line=dict(color=RED, dash="dash")))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["BKB"], mode="lines",
                                 name="BKB", line=dict(color=GREEN, dash="dash")))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["GT"], mode="lines",
                                 name="Garis Tengah (GT)", line=dict(color=LATTE, dash="dot", width=2)))
        ooc = d[d["Proporsi Cacat"] > d["BKA"]]
        ooc2 = d[d["Proporsi Cacat"] < d["BKB"]]
        ooc = pd.concat([ooc, ooc2])
        if not ooc.empty:
            fig.add_trace(go.Scatter(x=ooc["Day"], y=ooc["Proporsi Cacat"],
                                     mode="markers", name="Di luar batas kendali",
                                     marker=dict(color=RED, size=10, symbol="circle-open", line=dict(width=2))))
        fig.update_layout(title=title, plot_bgcolor="white", paper_bgcolor="white",
                          font_color=INK, hovermode="x unified",
                          legend=dict(orientation="h", y=-0.25),
                          margin=dict(l=10,r=10,t=50,b=70))
        fig.update_yaxes(tickformat=".0%", title="Proporsi Cacat")
        fig.update_xaxes(title="Tanggal Produksi")
        st.plotly_chart(fig, use_container_width=True)
        outside = ((d["Proporsi Cacat"] > d["BKA"]) | (d["Proporsi Cacat"] < d["BKB"])).sum()
        st.caption(f"Jumlah titik di luar batas kendali pada versi ini: {outside} dari {len(d)} subgrup.")

    if chart_choice == "Bandingkan kedua versi":
        a, b = st.columns(2)
        with a:
            draw_pchart(data, "Peta Kendali 1")
        with b:
            draw_pchart(data2, "Peta Kendali 2")
    elif chart_choice == "Peta Kendali 1":
        draw_pchart(data, "Peta Kendali 1")
    else:
        draw_pchart(data2, "Peta Kendali 2")

    st.markdown("---")
    st.markdown("### 2. Uji Keacakan (Run Test)")
    run_values = {
        "Median Proporsi Cacat": 0.06,
        "Jumlah data di atas median (n1)": 33,
        "Jumlah data di bawah median (n2)": 33,
        "Jumlah run (V)": 33,
        "Nilai Z": -0.2480988317,
        "P-value": 0.8040579395
    }
    # Nilai di atas dibaca dari ringkasan sheet Uji Keacakan workbook.
    cols = st.columns(3)
    for col, (label, value) in zip(cols, run_values.items()):
        col.metric(label, f"{value:.4f}" if isinstance(value, float) else str(value))
    st.info("Keputusan pada workbook: H0 Tidak Ditolak (p-value 0,8041 > 0,05). Interpretasi: data proporsi cacat tidak menunjukkan penyimpangan dari pola acak berdasarkan uji ini. Hasil uji keacakan tidak dengan sendirinya menyatakan proses terkendali.")

    st.markdown("---")
    st.markdown("### 3. Kapabilitas Proses Atribut")
    cap_mean = 0.05746815901196449
    cap_z = 1.5763884088598492
    cap_index = 0.5254628029532831
    c1, c2, c3 = st.columns(3)
    c1.metric("Proporsi Cacat Rata-rata (sheet)", f"{cap_mean:.5f}")
    c2.metric("Nilai Zp̄ (sheet)", f"{cap_z:.4f}")
    c3.metric("Indeks Kapabilitas (sheet)", f"{cap_index:.4f}")
    st.warning("Nilai kapabilitas di atas ditampilkan sesuai ringkasan sheet Kapabilitas. Workbook juga memuat nilai proporsi rata-rata lain; verifikasi konsistensi perhitungan perlu dilakukan sebelum angka ini dijadikan kesimpulan final.")

    st.markdown("---")
    st.markdown("### 4. Metrik Six Sigma")
    # Nilai ringkasan mengikuti sheet Level Sigma; bukan dihitung ulang diam-diam.
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("DPU", "0,06047")
    s2.metric("DPO", "0,01008")
    s3.metric("DPMO", "10.078,41")
    s4.metric("Level Sigma", "3,8234")
    st.caption("Nilai ditampilkan sesuai ringkasan sheet 'Level Sigma' pada workbook.")

# ============================================================
# ANALYZE
# ============================================================
elif page == "Analyze":
    st.subheader("Tahap Analyze")
    st.markdown("### 1. Diagram Pareto")
    defect_totals = data[defects].apply(pd.to_numeric, errors="coerce").sum().sort_values(ascending=False)
    pareto = pd.DataFrame({"Jenis Cacat": defect_totals.index, "Frekuensi": defect_totals.values})
    grand_total = pareto["Frekuensi"].sum()
    pareto["Persentase"] = pareto["Frekuensi"] / grand_total if grand_total else 0
    pareto["Kumulatif"] = pareto["Persentase"].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=pareto["Jenis Cacat"], y=pareto["Frekuensi"],
                         name="Frekuensi", marker_color=BROWN))
    fig.add_trace(go.Scatter(x=pareto["Jenis Cacat"], y=pareto["Kumulatif"] * grand_total,
                             name="Persentase Kumulatif", mode="lines+markers",
                             yaxis="y2", line=dict(color=LATTE, width=3)))
    fig.update_layout(
        yaxis=dict(title="Frekuensi"),
        yaxis2=dict(title="Persentase Kumulatif", overlaying="y", side="right",
                    tickformat=".0%", range=[0, 1]),
        plot_bgcolor="white", paper_bgcolor="white", font_color=INK,
        legend=dict(orientation="h", y=-0.25), margin=dict(l=10,r=10,t=35,b=80)
    )
    st.plotly_chart(fig, use_container_width=True)
    pareto_display = pareto.copy()
    pareto_display["Persentase"] = (pareto_display["Persentase"] * 100).round(2)
    pareto_display["Kumulatif"] = (pareto_display["Kumulatif"] * 100).round(2)
    st.dataframe(pareto_display, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 2. Fishbone Diagram")
    st.caption("Kategori penyebab mengikuti enam faktor yang digunakan dalam penelitian.")
    fishbone = {
        "Manusia (Man)": "Ketelitian pekerja berbeda-beda; kelelahan dapat memengaruhi hasil pelintingan.",
        "Mesin (Machine)": "Alat pemotong kurang tajam atau kondisi peralatan kurang baik.",
        "Material": "Kadar air tembakau, kualitas kertas ambri, atau daya rekat lem tidak sesuai standar.",
        "Metode (Method)": "Prosedur pelintingan dan pemberian lem belum diterapkan secara konsisten.",
        "Pengukuran (Measurement)": "Pemeriksaan kurang teliti dan standar kecacatan belum dipahami seragam.",
        "Lingkungan (Environment)": "Suhu, kelembapan, atau pencahayaan ruang produksi kurang sesuai."
    }
    for name, desc in fishbone.items():
        with st.expander(name):
            st.write(desc)

# ============================================================
# IMPROVE
# ============================================================
elif page == "Improve":
    st.subheader("Tahap Improve")
    st.markdown("### 1. Failure Mode and Effect Analysis (FMEA)")
    if "RPN" in fmea.columns:
        fmea_sorted = fmea.sort_values("RPN", ascending=False)
        st.dataframe(fmea_sorted, use_container_width=True, hide_index=True)
        if not fmea_sorted.empty:
            top = fmea_sorted.iloc[0]
            st.warning(f"Prioritas RPN tertinggi pada workbook: {top.get('faktor', 'Faktor tidak diketahui')} — RPN {top.get('RPN', '—')}.")
    else:
        st.error("Kolom RPN tidak ditemukan pada sheet FMEA.")

    st.markdown("---")
    st.markdown("### 2. K-Means Clustering")
    st.markdown("Pengelompokan dilakukan berdasarkan enam variabel jenis kecacatan pada setiap subgrup produksi.")
    if len(defects) == 6 and len(data) >= 3:
        x = data[defects].apply(pd.to_numeric, errors="coerce").fillna(0)
        scaler = StandardScaler()
        x_scaled = scaler.fit_transform(x)
        max_k = min(5, len(x) - 1)
        scores = []
        for k in range(2, max_k + 1):
            model = KMeans(n_clusters=k, random_state=42, n_init=10)
            labels = model.fit_predict(x_scaled)
            if len(set(labels)) > 1:
                scores.append((k, silhouette_score(x_scaled, labels)))
        if scores:
            best_k, best_score = max(scores, key=lambda z: z[1])
            model = KMeans(n_clusters=best_k, random_state=42, n_init=10)
            labels = model.fit_predict(x_scaled)
            result = data[["Subgrup", "Day"]].copy()
            result["Cluster"] = labels + 1
            st.metric("Jumlah cluster terpilih berdasarkan Silhouette Score", best_k)
            st.metric("Silhouette Score", f"{best_score:.4f}")
            fig = px.scatter(
                result, x="Subgrup", y="Cluster", color="Cluster",
                hover_data=["Day"], color_discrete_sequence=["#594638", "#8B6B52", "#C5A58A", "#D8BFA8", "#A88970"]
            )
            fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", font_color=INK)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(result, use_container_width=True, hide_index=True)
            st.caption("Hasil clustering dihitung dari data pada periode filter aktif. Periksa kesesuaian jumlah cluster dan interpretasinya dengan rancangan analisis Bab 3 sebelum digunakan sebagai hasil final.")
        else:
            st.warning("Silhouette Score tidak dapat dihitung untuk data terpilih.")
    else:
        st.warning("Clustering memerlukan enam kolom jenis kecacatan dan minimal tiga subgrup data.")

# ============================================================
# HASIL AKHIR
# ============================================================
elif page == "Hasil Akhir":
    st.subheader("Hasil Akhir Analisis Kualitas")
    st.markdown(
        "<div class='section-note'>Halaman ini merangkum indikator dari workbook dan hasil analisis pada periode yang dipilih. "
        "Perbedaan antarversi P-Chart masih perlu diverifikasi sebelum kesimpulan final ditetapkan.</div>",
        unsafe_allow_html=True
    )
    a, b, c = st.columns(3)
    a.metric("Total Sampel", f"{total_sample:,.0f}".replace(",", "."))
    b.metric("Produk Cacat", f"{total_defect:,.0f}".replace(",", "."))
    c.metric("Proporsi Cacat", f"{defect_rate:.2%}" if pd.notna(defect_rate) else "—")
    st.markdown("#### Prioritas Kecacatan")
    st.write("Jenis kecacatan dengan frekuensi tertinggi pada periode terpilih:")
    if defects:
        top_defect = data[defects].apply(pd.to_numeric, errors="coerce").sum().sort_values(ascending=False)
        st.success(f"{top_defect.index[0]} — {top_defect.iloc[0]:,.0f} kejadian.")
    st.markdown("#### Prioritas Perbaikan FMEA")
    if "RPN" in fmea.columns and not fmea.empty:
        top = fmea.sort_values("RPN", ascending=False).iloc[0]
        st.write(f"Faktor dengan RPN tertinggi pada workbook adalah **{top.get('faktor', '—')}**, dengan RPN **{top.get('RPN', '—')}**.")
    st.warning("Kesimpulan akhir mengenai kestabilan dan kapabilitas proses belum ditetapkan otomatis karena terdapat perbedaan nilai pada beberapa sheet workbook. Verifikasi hasil analisis terlebih dahulu sebelum memasukkannya ke Bab 4.")
