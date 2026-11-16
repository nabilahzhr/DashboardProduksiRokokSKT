import io
import re
from pathlib import Path
from statistics import NormalDist

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

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

# Header utama (ditampilkan lebih dulu agar pesan error muncul di bawahnya)
st.markdown(
    f"<div style='background:{DARK};padding:22px 26px;border-radius:14px;margin-bottom:20px'>"
    f"<div style='color:#E8D8C8;font-size:12px;letter-spacing:2px'>QUALITY CONTROL DASHBOARD</div>"
    f"<div style='color:#FFF8F0;font-size:27px;font-weight:700'>Monitoring Kualitas Produksi Rokok SKT</div>"
    f"<div style='color:#E8D8C8;font-size:14px'>Pengendalian Kualitas menggunakan Metode Six Sigma · PT. XYZ</div>"
    f"</div>",
    unsafe_allow_html=True
)

# ============================================================
# KONSTANTA STRUKTUR WORKBOOK (sesuai Olah_Data_Skripsi.xlsx)
# ============================================================
DATA_DIR = Path(__file__).parent / "data"
NORM = NormalDist()

DEFECTS = ["Gembos", "Lem ambri tengah", "Lem ambri ujung", "Guntingan", "Korep", "Besar kepala"]
# Nama tampilan saja (di Excel tidak diubah); mengikuti penulisan pada tabel Pareto di workbook.
DISPLAY = {
    "Gembos": "Gembos", "Lem ambri tengah": "Lem Ambri Tengah",
    "Lem ambri ujung": "Lem Ambri Ujung", "Guntingan": "Guntingan",
    "Korep": "Korep", "Besar kepala": "Besar Kepala",
}
DEF, TOT, GOOD, PROP = "Jumlah Rokok cacat", "Total Sampel", "Rokok Baik (n)", "Proporsi Cacat"
CHART_COLS = ["Subgrup", "Day", *DEFECTS, DEF, GOOD, TOT, PROP, "BKA", "BKB", "GT", "Keterangan"]
REQUIRED_CHART = [c for c in CHART_COLS if c not in (GOOD, "Keterangan")]
FMEA_COLS = ["No.", "Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan",
             "S", "O", "D", "RPN", "Usulan Perbaikan"]
# Kolom catatan pribadi di sheet FMEA ('pertanyaanku', 'jawaban masnya') sengaja tidak dibaca.

SHEET_NAMES = {
    "p1": "Peta Kendali 1", "p2": "Peta Kendali 2", "fmea": "FMEA",
    "run": "Uji Keacakan", "cap": "Kapabilitas", "sigma": "Level Sigma",
}
REQUIRED_SHEETS = ("p1", "p2", "fmea")  # sheet lain hanya dipakai untuk verifikasi angka

# Label ringkasan di sheet -> nama internal (nilainya berada tepat di bawah label)
REF_LABELS = {
    "run": {"Median Proporsi Cacat": "median", "Jumlah Data Atas Median (n1)": "n1",
            "Jumlah Data Bawah Median (n2)": "n2", "Jumlah Run (V)": "V", "μr": "mu",
            "σr": "sigma", "Z": "z", "P-Value": "p", "Keputusan": "keputusan"},
    "cap": {"Proporsi Cacat Rata-Rata": "pbar", "Nilai Zpbar": "z", "Indeks Kapabilitas": "index"},
    "sigma": {"DPU": "dpu", "DPO": "dpo", "DPMO": "dpmo", "Level Sigma": "sigma"},
}


class WorkbookError(Exception):
    """Error dengan pesan yang aman dan jelas untuk ditampilkan ke pengguna."""


def norm(text):
    """Seragamkan teks: hilangkan spasi berlebih/di ujung dan huruf besar-kecil."""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip().lower()


CANON_CHART = {norm(c): c for c in CHART_COLS}
CANON_FMEA = {norm(c): c for c in FMEA_COLS}


# ============================================================
# PEMBACAAN WORKBOOK
# ============================================================
def table_from_sheet(raw, canon, must_have, sheet_name):
    """Cari baris header secara otomatis, lalu ambil hanya kolom yang dikenali."""
    header_row = None
    for i in range(min(len(raw), 15)):
        cells = {norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)}
        if all(norm(m) in cells for m in must_have):
            header_row = i
            break
    if header_row is None:
        raise WorkbookError(
            f"Sheet '{sheet_name}': baris header tidak ditemukan. Dashboard mencari baris yang "
            f"memuat kolom {', '.join(must_have)} di 15 baris pertama."
        )
    names = [canon.get(norm(v)) if pd.notna(v) else None for v in raw.iloc[header_row].tolist()]
    keep, seen = [], set()
    for j, name in enumerate(names):
        if name and name not in seen:
            keep.append(j)
            seen.add(name)
    df = raw.iloc[header_row + 1:, keep].copy()
    df.columns = [names[j] for j in keep]
    return df.reset_index(drop=True)


def clean_chart(raw, sheet_name):
    df = table_from_sheet(raw, CANON_CHART, ["Subgrup", TOT], sheet_name)
    missing = [c for c in REQUIRED_CHART if c not in df.columns]
    if missing:
        raise WorkbookError(
            f"Sheet '{sheet_name}': kolom berikut tidak ditemukan: {', '.join(missing)}. "
            f"Kolom yang terbaca: {', '.join(df.columns)}."
        )
    df["Subgrup"] = pd.to_numeric(df["Subgrup"], errors="coerce")
    df = df[df["Subgrup"].notna()].copy()  # baris 'Jumlah' dan baris kosong otomatis terbuang
    df["Day"] = pd.to_datetime(df["Day"], errors="coerce")
    for c in [*DEFECTS, DEF, GOOD, TOT, PROP, "BKA", "BKB", "GT"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df["Day"].notna()]
    if df.empty:
        raise WorkbookError(f"Sheet '{sheet_name}': tidak ada baris data dengan nomor subgrup dan tanggal yang valid.")
    return df.sort_values("Subgrup").reset_index(drop=True)


def clean_fmea(raw, sheet_name):
    df = table_from_sheet(raw, CANON_FMEA, ["Faktor", "RPN"], sheet_name)
    for c in ["No.", "S", "O", "D", "RPN"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df["RPN"].notna()].reset_index(drop=True)
    if df.empty:
        raise WorkbookError(f"Sheet '{sheet_name}': tidak ada baris FMEA dengan nilai RPN.")
    return df


def side_value(raw, label):
    """Ambil nilai yang terletak tepat di bawah sebuah label pada tabel ringkasan sheet."""
    target = norm(label)
    for r in range(raw.shape[0]):
        for c in range(raw.shape[1]):
            v = raw.iat[r, c]
            if isinstance(v, str) and norm(v) == target:
                for rr in (r + 1, r + 2):
                    if rr < raw.shape[0] and pd.notna(raw.iat[rr, c]) and str(raw.iat[rr, c]).strip():
                        val = raw.iat[rr, c]
                        try:
                            return float(val)
                        except (TypeError, ValueError):
                            return str(val).strip()
    return None


@st.cache_data(show_spinner=False)
def parse_workbook(raw_bytes: bytes):
    """Terima isi file (bytes), kembalikan tabel yang sudah dibersihkan."""
    try:
        # Bytes WAJIB dibungkus BytesIO; pandas tidak menerima bytes mentah.
        xls = pd.ExcelFile(io.BytesIO(raw_bytes), engine="openpyxl")
    except Exception as exc:
        raise WorkbookError(
            "File tidak dapat dibuka sebagai workbook Excel (.xlsx). Kemungkinan penyebab: file rusak, "
            "dilindungi password, atau berasal dari format lain (.xls/.csv) yang hanya diganti ekstensinya. "
            f"Detail teknis: {type(exc).__name__}: {exc}"
        )
    found = {}
    for key, wanted in SHEET_NAMES.items():
        found[key] = next((n for n in xls.sheet_names if norm(n) == norm(wanted)), None)
    missing = [SHEET_NAMES[k] for k in REQUIRED_SHEETS if not found[k]]
    if missing:
        raise WorkbookError(
            f"Sheet wajib tidak ditemukan: {', '.join(missing)}. "
            f"Sheet yang ada di file ini: {', '.join(repr(n) for n in xls.sheet_names)}."
        )
    raw = {k: pd.read_excel(xls, sheet_name=found[k], header=None) for k in found if found[k]}
    ref = {}
    for key, labels in REF_LABELS.items():
        ref[key] = {name: side_value(raw[key], lab) for lab, name in labels.items()} if key in raw else {}
    return clean_chart(raw["p1"], found["p1"]), clean_chart(raw["p2"], found["p2"]), \
        clean_fmea(raw["fmea"], found["fmea"]), ref


def list_excel_files():
    if not DATA_DIR.is_dir():
        return []
    return sorted(p for p in DATA_DIR.iterdir()
                  if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$"))


# ---------- Sidebar: sumber data ----------
st.sidebar.markdown("## QUALITY CONTROL")
st.sidebar.caption("MONITORING PRODUKSI ROKOK SKT")
uploaded = st.sidebar.file_uploader("Unggah workbook Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    source_bytes, source_label = uploaded.getvalue(), uploaded.name
else:
    excel_files = list_excel_files()
    if not excel_files:
        if not DATA_DIR.is_dir():
            reason = "Folder `data/` tidak ditemukan di repository."
        else:
            others = [p.name for p in DATA_DIR.iterdir() if p.is_file()]
            reason = "Folder `data/` ada, tetapi tidak berisi file berekstensi `.xlsx`."
            reason += f" Isi folder saat ini: {', '.join(others)}." if others else " Folder kosong."
        st.error(f"Belum ada workbook yang bisa dibaca. {reason}")
        st.info(
            "Cara memperbaiki: (1) unggah file Excel lewat sidebar di kiri, atau (2) tambahkan file `.xlsx` "
            "ke folder `data/` di GitHub, lalu buka menu aplikasi → Reboot app. Nama file bebas."
        )
        st.stop()
    if len(excel_files) > 1:
        chosen = st.sidebar.selectbox("Pilih file di folder data/", [p.name for p in excel_files])
        path = DATA_DIR / chosen
    else:
        path = excel_files[0]
    try:
        source_bytes, source_label = path.read_bytes(), f"data/{path.name}"
    except OSError as exc:
        st.error(f"File `{path.name}` ditemukan tetapi tidak bisa dibaca: {exc}")
        st.stop()

try:
    p1, p2, fmea, ref = parse_workbook(source_bytes)
except WorkbookError as exc:
    st.error(f"Gagal membaca workbook `{source_label}`.")
    st.warning(str(exc))
    st.info(
        "Cara memperbaiki: pastikan file adalah workbook .xlsx asli dengan sheet Peta Kendali 1, "
        "Peta Kendali 2, dan FMEA, serta nama kolom seperti pada file Olah_Data_Skripsi.xlsx. "
        "Anda juga dapat mengunggah file yang benar lewat sidebar."
    )
    st.stop()
except Exception as exc:  # kegagalan tak terduga: tampilkan penyebabnya, bukan layar kosong
    st.error(f"Terjadi kesalahan tak terduga saat membaca `{source_label}`.")
    st.warning(f"{type(exc).__name__}: {exc}")
    st.info("Coba unggah ulang file. Jika masih gagal, kirim pesan di atas kepada pengembang dashboard.")
    st.stop()


# ============================================================
# FUNGSI BANTU
# ============================================================
def fmt_int(x):
    return f"{x:,.0f}".replace(",", ".")


def fmt_dec(x, n=4):
    if x is None or pd.isna(x):
        return "—"
    return f"{x:,.{n}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def show_plot(fig):
    st.plotly_chart(fig, width="stretch")


def show_table(df):
    st.dataframe(df, width="stretch", hide_index=True)


def filtered_dates(df, start_date, end_date):
    dates = pd.to_datetime(df["Day"], errors="coerce").dt.date
    return df[(dates >= start_date) & (dates <= end_date)].copy()


def for_display(df):
    out = df.copy()
    if "Day" in out.columns:
        out["Day"] = pd.to_datetime(out["Day"]).dt.date
    return out.rename(columns=DISPLAY)


def six_sigma(df, n_ctq):
    """DPU = total cacat / total sampel; DPO = DPU / jumlah CTQ; DPMO = DPO x 1.000.000;
    Level Sigma = NORMSINV(1 - DPO) + 1,5."""
    tot_n, tot_d = df[TOT].sum(), df[DEF].sum()
    if not tot_n:
        return None
    dpu = tot_d / tot_n
    dpo = dpu / n_ctq
    sigma = NORM.inv_cdf(1 - dpo) + 1.5 if 0 < dpo < 1 else np.nan
    return {"dpu": dpu, "dpo": dpo, "dpmo": dpo * 1e6, "sigma": sigma}


def capability(df):
    """Kapabilitas atribut: Zp̄ = NORMSINV(1 - p̄); Indeks kapabilitas = Zp̄ / 3."""
    tot_n = df[TOT].sum()
    if not tot_n:
        return None
    pbar = df[DEF].sum() / tot_n
    if not 0 < pbar < 1:
        return None
    z = NORM.inv_cdf(1 - pbar)
    return {"pbar": pbar, "z": z, "index": z / 3}


def run_test(df):
    """Run Test terhadap median proporsi cacat (data di atas median vs sisanya)."""
    if len(df) < 2:
        return None
    # Tanpa pembulatan, agar klasifikasi atas/bawah median identik dengan perhitungan di workbook.
    p = df.sort_values("Subgrup")[PROP].astype(float)
    med = float(p.median())
    above = (p > med).to_numpy()
    n1, n2 = int(above.sum()), int((~above).sum())
    if n1 == 0 or n2 == 0:
        return None
    v = 1 + int((above[1:] != above[:-1]).sum())
    n = n1 + n2
    mu = 2 * n1 * n2 / n + 1
    var = 2 * n1 * n2 * (2 * n1 * n2 - n) / (n ** 2 * (n - 1))
    z = (v - mu) / np.sqrt(var)
    return {"median": med, "n1": n1, "n2": n2, "V": v, "mu": mu, "sigma": float(np.sqrt(var)),
            "z": float(z), "p": float(2 * (1 - NORM.cdf(abs(z))))}


def quality_issues(df):
    """Pemeriksaan konsistensi data. Dashboard hanya memberi peringatan, tidak mengubah angka."""
    issues = []
    if GOOD in df.columns:
        diff = df[DEF] + df[GOOD] - df[TOT]
        for _, r in df[diff.abs() > 0.5].iterrows():
            issues.append(
                f"Subgrup {int(r['Subgrup'])} ({r['Day']:%d %b %Y}): cacat {r[DEF]:.0f} + rokok baik "
                f"{r[GOOD]:.0f} = {r[DEF] + r[GOOD]:.0f}, tetapi Total Sampel tertulis {r[TOT]:.0f}."
            )
    diff2 = df[DEFECTS].sum(axis=1) - df[DEF]
    for _, r in df[diff2.abs() > 0.5].iterrows():
        issues.append(
            f"Subgrup {int(r['Subgrup'])} ({r['Day']:%d %b %Y}): jumlah enam jenis cacat tidak sama "
            f"dengan Jumlah Rokok cacat ({r[DEF]:.0f})."
        )
    return issues


def compare_table(rows):
    out = []
    for label, mine, theirs, digits in rows:
        if theirs is None or isinstance(theirs, str) or mine is None or pd.isna(mine):
            status = "—"
        else:
            status = "Sesuai" if np.isclose(mine, theirs, rtol=1e-6, atol=1e-9) else "Berbeda"
        out.append({"Indikator": label, "Dashboard": fmt_dec(mine, digits),
                    "Workbook": fmt_dec(theirs, digits) if isinstance(theirs, float) else "—",
                    "Status": status})
    return pd.DataFrame(out)


# ============================================================
# FILTER PERIODE & NAVIGASI
# ============================================================
date_values = p1["Day"].dropna()
min_date, max_date = date_values.min().date(), date_values.max().date()

st.sidebar.markdown("---")
st.sidebar.markdown("### Filter Periode")
date_range = st.sidebar.date_input(
    "Tanggal produksi", value=(min_date, max_date), min_value=min_date, max_value=max_date
)
if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date
full_period = (start_date, end_date) == (min_date, max_date)

data = filtered_dates(p1, start_date, end_date)
data2 = filtered_dates(p2, start_date, end_date)

st.sidebar.markdown("---")
page = st.sidebar.radio(
    "NAVIGASI", ["Beranda", "Measure", "Analyze", "Improve", "Hasil Akhir"],
    label_visibility="visible"
)
st.sidebar.markdown("---")
st.sidebar.caption(f"Sumber data: {source_label}")
st.sidebar.caption(f"Periode tersedia: {min_date:%d %b %Y} – {max_date:%d %b %Y}")

if data.empty:
    st.warning("Tidak ada data produksi pada rentang tanggal yang dipilih. Ubah filter periode di sidebar.")
    st.stop()

total_sample = data[TOT].sum()
total_defect = data[DEF].sum()
defect_rate = total_defect / total_sample if total_sample else np.nan

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
    c1.metric("Total Sampel", fmt_int(total_sample))
    c2.metric("Jumlah Produk Cacat", fmt_int(total_defect))
    c3.metric("Proporsi Cacat", f"{defect_rate:.2%}" if pd.notna(defect_rate) else "—")
    c4.metric("Jumlah Subgrup", fmt_int(len(data)))

    issues = quality_issues(data)
    if issues:
        with st.expander(f"⚠️ {len(issues)} data perlu diperiksa pada file Excel", expanded=False):
            for msg in issues:
                st.write("• " + msg)
            st.caption("Dashboard tidak mengubah angka. Periksa data asli di Excel, perbaiki bila perlu, lalu unggah ulang.")

    left, right = st.columns([1.2, 1])
    with left:
        st.markdown("#### Tren Proporsi Cacat")
        fig = px.line(
            data, x="Day", y=PROP, markers=True,
            labels={"Day": "Tanggal Produksi", PROP: "Proporsi Cacat"},
            color_discrete_sequence=[BROWN]
        )
        fig.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                          font_color=INK, margin=dict(l=10, r=10, t=25, b=10))
        fig.update_yaxes(tickformat=".0%")
        show_plot(fig)
    with right:
        st.markdown("#### Komposisi Jenis Kecacatan")
        sums = data[DEFECTS].sum().sort_values(ascending=False)
        fig = px.pie(
            values=sums.values, names=[DISPLAY[c] for c in sums.index], hole=0.55,
            color_discrete_sequence=["#594638", "#8B6B52", "#C5A58A",
                                     "#D8BFA8", "#A88970", "#E8D8C8"]
        )
        fig.update_layout(paper_bgcolor="white", font_color=INK,
                          margin=dict(l=10, r=10, t=25, b=10), legend_title_text="")
        show_plot(fig)
    st.markdown("#### Data Produksi")
    show_table(for_display(data[["Subgrup", "Day", *DEFECTS, DEF, TOT, PROP]]))

# ============================================================
# MEASURE
# ============================================================
elif page == "Measure":
    st.subheader("Tahap Measure")
    st.markdown("Pengukuran kualitas proses berdasarkan P-Chart, Run Test, kapabilitas atribut, dan metrik Six Sigma.")

    # ---------- 1. P-Chart ----------
    st.markdown("### 1. Peta Kendali P")
    gt1 = f"{data['GT'].iloc[0]:.6f}"
    gt2 = f"{data2['GT'].iloc[0]:.6f}" if not data2.empty else "—"
    st.info(
        f"Workbook memuat dua versi P-Chart. Peta Kendali 1 memuat {len(data)} subgrup dengan garis tengah "
        f"{gt1}. Peta Kendali 2 memuat {len(data2)} subgrup dengan garis tengah {gt2} "
        "(versi revisi setelah subgrup di luar batas kendali pada Peta Kendali 1 dikeluarkan)."
    )
    chart_choice = st.selectbox(
        "Pilih versi P-Chart untuk ditampilkan",
        ["Bandingkan kedua versi", "Peta Kendali 1", "Peta Kendali 2"], index=0
    )

    def draw_pchart(df, title):
        d = df.dropna(subset=["Day", PROP, "BKA", "BKB", "GT"])
        if d.empty:
            st.warning(f"Tidak ada data {title} pada periode yang dipilih.")
            return
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=d["Day"], y=d[PROP], mode="lines+markers",
                                 name="Proporsi Cacat", line=dict(color=BROWN, width=2),
                                 marker=dict(size=6, color=BROWN)))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["BKA"], mode="lines",
                                 name="BKA", line=dict(color=RED, dash="dash")))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["BKB"], mode="lines",
                                 name="BKB", line=dict(color=GREEN, dash="dash")))
        fig.add_trace(go.Scatter(x=d["Day"], y=d["GT"], mode="lines",
                                 name="Garis Tengah (GT)", line=dict(color=LATTE, dash="dot", width=2)))
        ooc = d[(d[PROP] > d["BKA"]) | (d[PROP] < d["BKB"])]
        if not ooc.empty:
            fig.add_trace(go.Scatter(x=ooc["Day"], y=ooc[PROP],
                                     mode="markers", name="Di luar batas kendali",
                                     marker=dict(color=RED, size=10, symbol="circle-open", line=dict(width=2))))
        fig.update_layout(title=title, plot_bgcolor="white", paper_bgcolor="white",
                          font_color=INK, hovermode="x unified",
                          legend=dict(orientation="h", y=-0.25),
                          margin=dict(l=10, r=10, t=50, b=70))
        fig.update_yaxes(tickformat=".0%", title="Proporsi Cacat")
        fig.update_xaxes(title="Tanggal Produksi")
        show_plot(fig)
        st.caption(f"Jumlah titik di luar batas kendali pada versi ini: {len(ooc)} dari {len(d)} subgrup.")

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

    # ---------- 2. Run Test ----------
    st.markdown("---")
    st.markdown("### 2. Uji Keacakan (Run Test)")
    rt = run_test(data)
    if rt is None:
        st.warning("Run Test memerlukan minimal dua subgrup dengan data di atas dan di bawah median.")
    else:
        r1 = st.columns(3)
        r1[0].metric("Median Proporsi Cacat", fmt_dec(rt["median"], 4))
        r1[1].metric("Jumlah data di atas median (n1)", str(rt["n1"]))
        r1[2].metric("Jumlah data di bawah median (n2)", str(rt["n2"]))
        r2 = st.columns(3)
        r2[0].metric("Jumlah run (V)", str(rt["V"]))
        r2[1].metric("Nilai Z", fmt_dec(rt["z"], 4))
        r2[2].metric("P-value", fmt_dec(rt["p"], 4))
        st.caption("Klasifikasi mengikuti cara hitung workbook: proporsi > median dihitung 'di atas', selainnya 'di bawah'.")
        if rt["p"] > 0.05:
            st.info(
                f"Keputusan: H0 Tidak Ditolak (p-value {fmt_dec(rt['p'], 4)} > 0,05). Interpretasi: data proporsi cacat "
                "tidak menunjukkan penyimpangan dari pola acak berdasarkan uji ini. Hasil uji keacakan tidak "
                "dengan sendirinya menyatakan proses terkendali."
            )
        else:
            st.warning(
                f"Keputusan: H0 Ditolak (p-value {fmt_dec(rt['p'], 4)} ≤ 0,05). Interpretasi: data proporsi cacat "
                "menunjukkan penyimpangan dari pola acak berdasarkan uji ini."
            )

    # ---------- 3. Kapabilitas ----------
    st.markdown("---")
    st.markdown("### 3. Kapabilitas Proses Atribut")
    cap = capability(data2) if not data2.empty else None
    if cap is None:
        st.warning("Kapabilitas tidak dapat dihitung pada periode yang dipilih.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Proporsi Cacat Rata-rata", fmt_dec(cap["pbar"], 5))
        c2.metric("Nilai Zp̄", fmt_dec(cap["z"], 4))
        c3.metric("Indeks Kapabilitas", fmt_dec(cap["index"], 4))
        st.caption(
            f"Dihitung dari Peta Kendali 2 ({len(data2)} subgrup): Zp̄ = NORMSINV(1 − p̄) dan "
            "Indeks Kapabilitas = Zp̄ ÷ 3. Perhatikan bahwa p̄ ini berbeda dengan p̄ Peta Kendali 1 "
            "karena jumlah subgrup yang dipakai berbeda."
        )

    # ---------- 4. Six Sigma ----------
    st.markdown("---")
    st.markdown("### 4. Metrik Six Sigma")
    ss = six_sigma(data, len(DEFECTS))
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("DPU", fmt_dec(ss["dpu"], 5))
    s2.metric("DPO", fmt_dec(ss["dpo"], 5))
    s3.metric("DPMO", fmt_dec(ss["dpmo"], 2))
    s4.metric("Level Sigma", fmt_dec(ss["sigma"], 4))
    st.caption(
        f"Dihitung dari Peta Kendali 1 ({len(data)} subgrup): DPU = total cacat ÷ total sampel; "
        f"DPO = DPU ÷ {len(DEFECTS)} jenis cacat (CTQ); DPMO = DPO × 1.000.000; "
        "Level Sigma = NORMSINV(1 − DPO) + 1,5."
    )

    # ---------- Verifikasi terhadap workbook ----------
    with st.expander("Verifikasi: bandingkan hasil dashboard dengan nilai di workbook"):
        if not full_period:
            st.info("Perbandingan hanya bermakna jika filter periode mencakup seluruh tanggal. "
                    "Atur filter ke periode penuh di sidebar.")
        else:
            rows = []
            rr, rc, rs = ref.get("run", {}), ref.get("cap", {}), ref.get("sigma", {})
            if rt:
                rows += [("Run Test: median", rt["median"], rr.get("median"), 4),
                         ("Run Test: n1", rt["n1"], rr.get("n1"), 0),
                         ("Run Test: n2", rt["n2"], rr.get("n2"), 0),
                         ("Run Test: V", rt["V"], rr.get("V"), 0),
                         ("Run Test: Z", rt["z"], rr.get("z"), 4),
                         ("Run Test: P-value", rt["p"], rr.get("p"), 4)]
            if cap:
                rows += [("Kapabilitas: p̄", cap["pbar"], rc.get("pbar"), 5),
                         ("Kapabilitas: Zp̄", cap["z"], rc.get("z"), 4),
                         ("Kapabilitas: indeks", cap["index"], rc.get("index"), 4)]
            rows += [("DPU", ss["dpu"], rs.get("dpu"), 5), ("DPO", ss["dpo"], rs.get("dpo"), 5),
                     ("DPMO", ss["dpmo"], rs.get("dpmo"), 2), ("Level Sigma", ss["sigma"], rs.get("sigma"), 4)]
            show_table(compare_table(rows))
            st.caption("'—' berarti nilai tersebut tidak ditemukan di sheet ringkasan workbook.")

# ============================================================
# ANALYZE
# ============================================================
elif page == "Analyze":
    st.subheader("Tahap Analyze")
    st.markdown("### 1. Diagram Pareto")
    defect_totals = data[DEFECTS].sum().sort_values(ascending=False)
    pareto = pd.DataFrame({"Jenis Cacat": [DISPLAY[c] for c in defect_totals.index],
                           "Frekuensi": defect_totals.values})
    grand_total = pareto["Frekuensi"].sum()
    pareto["Persentase"] = pareto["Frekuensi"] / grand_total if grand_total else 0
    pareto["Kumulatif"] = pareto["Persentase"].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=pareto["Jenis Cacat"], y=pareto["Frekuensi"],
                         name="Frekuensi", marker_color=BROWN))
    fig.add_trace(go.Scatter(x=pareto["Jenis Cacat"], y=pareto["Kumulatif"],
                             name="Persentase Kumulatif", mode="lines+markers",
                             yaxis="y2", line=dict(color=LATTE, width=3)))
    fig.update_layout(
        yaxis=dict(title="Frekuensi"),
        yaxis2=dict(title="Persentase Kumulatif", overlaying="y", side="right",
                    tickformat=".0%", range=[0, 1]),
        plot_bgcolor="white", paper_bgcolor="white", font_color=INK,
        legend=dict(orientation="h", y=-0.25), margin=dict(l=10, r=10, t=35, b=80)
    )
    show_plot(fig)
    pareto_display = pareto.copy()
    pareto_display["Persentase (%)"] = (pareto_display.pop("Persentase") * 100).round(2)
    pareto_display["Persentase Kumulatif (%)"] = (pareto_display.pop("Kumulatif") * 100).round(2)
    show_table(pareto_display)

    st.markdown("---")
    st.markdown("### 2. Fishbone Diagram")
    st.caption("Kategori penyebab mengikuti enam faktor pada sheet FMEA di workbook.")
    if "Penyebab Kegagalan" in fmea.columns:
        fish = fmea.sort_values("No.") if "No." in fmea.columns else fmea
        for _, row in fish.iterrows():
            with st.expander(str(row["Faktor"])):
                st.write(row["Penyebab Kegagalan"])
    else:
        st.info("Kolom 'Penyebab Kegagalan' tidak ditemukan pada sheet FMEA, sehingga Fishbone tidak dapat ditampilkan.")

# ============================================================
# IMPROVE
# ============================================================
elif page == "Improve":
    st.subheader("Tahap Improve")
    st.markdown("### 1. Failure Mode and Effect Analysis (FMEA)")
    fmea_sorted = fmea.sort_values("RPN", ascending=False)
    show_table(fmea_sorted)
    top = fmea_sorted.iloc[0]
    st.warning(f"Prioritas RPN tertinggi pada workbook: {top['Faktor']} — RPN {fmt_int(top['RPN'])}.")

# ============================================================
# HASIL AKHIR
# ============================================================
elif page == "Hasil Akhir":
    st.subheader("Hasil Akhir Analisis Kualitas")
    st.markdown(
        "<div class='section-note'>Halaman ini merangkum indikator hasil analisis pada periode yang dipilih. "
        "Interpretasi akhir mengikuti pembahasan pada naskah Tugas Akhir.</div>",
        unsafe_allow_html=True
    )
    a, b, c = st.columns(3)
    a.metric("Total Sampel", fmt_int(total_sample))
    b.metric("Produk Cacat", fmt_int(total_defect))
    c.metric("Proporsi Cacat", f"{defect_rate:.2%}" if pd.notna(defect_rate) else "—")

    ss = six_sigma(data, len(DEFECTS))
    cap = capability(data2) if not data2.empty else None
    rt = run_test(data)
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("DPMO", fmt_dec(ss["dpmo"], 2))
    d2.metric("Level Sigma", fmt_dec(ss["sigma"], 4))
    d3.metric("Indeks Kapabilitas", fmt_dec(cap["index"], 4) if cap else "—")
    d4.metric("P-value Run Test", fmt_dec(rt["p"], 4) if rt else "—")
    ooc1 = int(((data[PROP] > data["BKA"]) | (data[PROP] < data["BKB"])).sum())
    st.caption(f"Titik di luar batas kendali pada Peta Kendali 1: {ooc1} dari {len(data)} subgrup.")

    st.markdown("#### Prioritas Kecacatan")
    st.write("Jenis kecacatan dengan frekuensi tertinggi pada periode terpilih:")
    top_defect = data[DEFECTS].sum().sort_values(ascending=False)
    st.success(f"{DISPLAY[top_defect.index[0]]} — {fmt_int(top_defect.iloc[0])} kejadian.")

    st.markdown("#### Prioritas Perbaikan FMEA")
    top = fmea.sort_values("RPN", ascending=False).iloc[0]
    st.write(f"Faktor dengan RPN tertinggi pada workbook adalah **{top['Faktor']}**, dengan RPN **{fmt_int(top['RPN'])}**.")
