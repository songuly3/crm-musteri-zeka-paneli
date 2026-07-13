from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Optional
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "final_musteri_skorlari_FINAL_3MODEL.xlsx"
MAIN_SHEET = "Tum_Musteriler"
TURKEY_TIMEZONE = ZoneInfo("Europe/Istanbul")

def current_turkey_date():
    """Her çalıştırmada Türkiye saatine göre güncel tarihi döndürür."""
    return datetime.now(TURKEY_TIMEZONE).date()

def schedule_midnight_refresh():
    """Tarayıcı açık bırakılırsa bir sonraki gece yarısında sayfayı yeniler."""
    components.html(
        """
        <script>
        const now = new Date();
        const nextMidnight = new Date(now);
        nextMidnight.setHours(24, 0, 5, 0);
        const waitMs = nextMidnight.getTime() - now.getTime();
        setTimeout(() => window.parent.location.reload(), waitMs);
        </script>
        """,
        height=0,
        width=0,
    )


st.set_page_config(
    page_title="CRM Müşteri Zekâ Paneli",
    page_icon="◉",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Tasarım
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --bg: #F6F8FC;
        --card: #FFFFFF;
        --ink: #172033;
        --muted: #6B7280;
        --line: #E7EAF0;
        --primary: #5B5BD6;
        --primary-soft: #EEEEFF;
        --success: #1C9A6C;
        --warning: #E7A11A;
        --danger: #D9534F;
    }
    .stApp { background: var(--bg); }
    [data-testid="stSidebar"] {
        background: #111827;
        border-right: 1px solid #1F2937;
    }
    [data-testid="stSidebar"] * { color: #F9FAFB; }
    [data-testid="stSidebar"] .stCaption { color: #9CA3AF !important; }
    .block-container { padding-top: 1.3rem; padding-bottom: 3rem; max-width: 1500px; }
    h1, h2, h3 { color: var(--ink); letter-spacing: -0.02em; }
    .eyebrow {
        font-size: 0.76rem; font-weight: 700; letter-spacing: 0.12em;
        text-transform: uppercase; color: var(--primary); margin-bottom: 0.2rem;
    }
    .page-title { font-size: 2.15rem; font-weight: 760; color: var(--ink); margin: 0; }
    .page-subtitle { color: var(--muted); margin-top: 0.25rem; margin-bottom: 1.25rem; }
    .hero {
        background: linear-gradient(135deg, #FFFFFF 0%, #F4F1FF 100%);
        border: 1px solid var(--line); border-radius: 22px; padding: 1.35rem 1.5rem;
        box-shadow: 0 10px 30px rgba(23,32,51,0.05); margin-bottom: 1rem;
    }
    .metric-card {
        background: var(--card); border: 1px solid var(--line); border-radius: 18px;
        padding: 1rem 1.05rem; min-height: 118px;
        box-shadow: 0 8px 24px rgba(23,32,51,0.045);
    }
    .metric-label { color: var(--muted); font-size: 0.82rem; font-weight: 600; }
    .metric-value { color: var(--ink); font-size: 1.65rem; font-weight: 760; margin-top: 0.25rem; }
    .metric-foot { color: var(--muted); font-size: 0.75rem; margin-top: 0.2rem; }
    .panel {
        background: var(--card); border: 1px solid var(--line); border-radius: 20px;
        padding: 1.1rem 1.2rem; box-shadow: 0 8px 24px rgba(23,32,51,0.04);
        margin-bottom: 0.8rem;
    }
    .badge {
        display: inline-block; padding: 0.34rem 0.68rem; border-radius: 999px;
        font-size: 0.78rem; font-weight: 700; margin-right: 0.35rem;
    }
    .badge-a { background: #EDEBFF; color: #4B47B8; }
    .badge-b { background: #E9F8F2; color: #197A58; }
    .badge-high { background: #FDECEC; color: #B63E3A; }
    .badge-mid { background: #FFF5DC; color: #9B6900; }
    .badge-low { background: #EAF5FF; color: #2C6D9E; }
    .badge-wait { background: #E7F6EC; color: #067647; }
    .action-box {
        border-radius: 18px; padding: 1rem 1.1rem; border-left: 5px solid var(--primary);
        background: #F6F5FF; color: var(--ink); margin-top: 0.5rem;
    }
    .action-title { font-weight: 760; font-size: 1rem; margin-bottom: 0.25rem; }
    .action-text { color: #4B5563; font-size: 0.9rem; }
    .small-note { color: var(--muted); font-size: 0.79rem; }
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px; font-weight: 700; min-height: 42px;
    }
    div[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { border-radius: 12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Veri yardımcıları
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_workbook(source, source_version: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(source)
    wanted = [
        MAIN_SHEET,
        "Model_A_Churn",
        "Model_B_Reaktivasyon",
        "Model_Performansi",
        "Model_Karsilastirma",
        "Validation_Tum_Denemeler",
        "Threshold_Analizi",
        "Lift_Analizi",
        "Feature_Importance",
        "Veri_Kalitesi",
    ]
    frames: dict[str, pd.DataFrame] = {}
    for sheet in wanted:
        if sheet in xls.sheet_names:
            df = pd.read_excel(source, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]
            frames[sheet] = df
    if MAIN_SHEET not in frames:
        raise ValueError(f"'{MAIN_SHEET}' sayfası bulunamadı.")
    return frames


def parse_excel_date_series(series: pd.Series) -> pd.Series:
    """Excel seri tarihlerini ve normal tarihleri güvenli biçimde çevirir."""
    numeric = pd.to_numeric(series, errors="coerce")
    numeric_ratio = float(numeric.notna().mean()) if len(series) else 0.0

    if (
        numeric_ratio >= 0.80
        and not numeric.dropna().empty
        and numeric.dropna().between(20000, 80000).all()
    ):
        return pd.to_datetime(
            numeric,
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    return pd.to_datetime(series, errors="coerce")


def detect_id_column(df: pd.DataFrame) -> Optional[str]:
    for col in [
        "CustomerID", "customer_id", "CRM_ID", "crm_id",
        "MusteriID", "musteri_id", "pipeline_musteri_id",
    ]:
        if col in df.columns:
            return col
    return None


def enrich_main(
    frames: dict[str, pd.DataFrame],
    id_col: str,
    as_of_date,
) -> pd.DataFrame:
    """Veriyi güncel tarihe göre yeniden hesaplar.

    Model B'de iki farklı kavram üretilir:
    1) Yaşam döngüsü segmenti
    2) Kampanya önceliği

    Böylece uzun süredir kayıp bir müşteri, reaktivasyon ihtiyacı
    çok yüksek olsa da düşük maliyetli kampanya önceliğinde kalabilir.
    """
    df = frames[MAIN_SHEET].copy()

    # Sürüm uyumluluğu:
    # Eski dosyalarda "segment", yeni v7 dosyasında "kampanya_onceligi"
    # alanı bulunabilir. Uygulama her iki yapıyı da destekler.
    if "segment" not in df.columns:
        if "kampanya_onceligi" in df.columns:
            df["segment"] = df["kampanya_onceligi"]
        else:
            df["segment"] = "—"

    if "kampanya_onceligi" not in df.columns:
        df["kampanya_onceligi"] = df["segment"]

    if "yasam_dongusu_segmenti" not in df.columns:
        if "musteri_durumu" in df.columns:
            df["yasam_dongusu_segmenti"] = df["musteri_durumu"]
        else:
            df["yasam_dongusu_segmenti"] = "—"

    if "reaktivasyon_ihtiyaci" not in df.columns:
        df["reaktivasyon_ihtiyaci"] = "—"

    for date_col in [
        "FirstPurchaseDate",
        "LastPurchaseDate",
        "snapshot_tarihi",
        "skor_tarihi",
    ]:
        if date_col in df.columns:
            df[date_col] = parse_excel_date_series(df[date_col])

    if "recency_snapshot" not in df.columns:
        df["recency_snapshot"] = pd.to_numeric(
            df.get("recency"),
            errors="coerce",
        )

    score_date = pd.Timestamp(as_of_date).normalize()
    df["skor_tarihi"] = score_date

    if "LastPurchaseDate" in df.columns:
        df["recency"] = (
            score_date - df["LastPurchaseDate"].dt.normalize()
        ).dt.days.clip(lower=0)
    else:
        df["recency"] = pd.to_numeric(
            df.get("recency"),
            errors="coerce",
        )

    # Ortak feature'lar
    df["avg_basket_value"] = np.where(
        pd.to_numeric(df["frequency"], errors="coerce") > 0,
        pd.to_numeric(df["monetary"], errors="coerce")
        / pd.to_numeric(df["frequency"], errors="coerce"),
        np.nan,
    )
    df["avg_qty_per_order"] = np.where(
        pd.to_numeric(df["frequency"], errors="coerce") > 0,
        pd.to_numeric(df["Qty"], errors="coerce")
        / pd.to_numeric(df["frequency"], errors="coerce"),
        np.nan,
    )

    if {"FirstPurchaseDate", "LastPurchaseDate"}.issubset(df.columns):
        df["active_days"] = (
            df["LastPurchaseDate"] - df["FirstPurchaseDate"]
        ).dt.days

        df["inter_purchase_time"] = np.where(
            pd.to_numeric(df["frequency"], errors="coerce") > 1,
            df["active_days"]
            / (
                pd.to_numeric(df["frequency"], errors="coerce")
                - 1
            ),
            np.nan,
        )
        df["monetary_per_day"] = np.where(
            df["active_days"] > 0,
            pd.to_numeric(df["monetary"], errors="coerce")
            / df["active_days"],
            pd.to_numeric(df["monetary"], errors="coerce"),
        )
        df["qty_per_day"] = np.where(
            df["active_days"] > 0,
            pd.to_numeric(df["Qty"], errors="coerce")
            / df["active_days"],
            pd.to_numeric(df["Qty"], errors="coerce"),
        )

    frequency = pd.to_numeric(df["frequency"], errors="coerce")
    b_mask = frequency.eq(1)
    a_mask = frequency.gt(1)

    # Model A alanları
    if a_mask.any():
        df.loc[a_mask, "yasam_dongusu_segmenti"] = (
            "Model A - Churn Risk İzleme"
        )
        df.loc[a_mask, "reaktivasyon_ihtiyaci"] = (
            "Model A ile değerlendirilir"
        )
        # Model A'nın mevcut öncelik segmentini koru.
        # Dosyada yalnızca kampanya_onceligi varsa alias üzerinden okunur.
        df.loc[a_mask, "kampanya_onceligi"] = df.loc[
            a_mask, "segment"
        ]

    # Model B alanları
    if b_mask.any():
        monetary = pd.to_numeric(
            df.loc[b_mask, "monetary"],
            errors="coerce",
        )
        recency = pd.to_numeric(
            df.loc[b_mask, "recency"],
            errors="coerce",
        )

        df.loc[b_mask, "value_percentile"] = (
            monetary.rank(pct=True, method="average") * 100
        )
        df.loc[b_mask, "freshness_score"] = np.clip(
            100 - (recency / 730) * 100,
            0,
            100,
        )
        df.loc[b_mask, "ham_model_skoru"] = (
            0.50 * df.loc[b_mask, "value_percentile"]
            + 0.50 * df.loc[b_mask, "freshness_score"]
        )

        df.loc[b_mask, "model_tipi"] = (
            "Model B - Tek Seferlik Müşteri"
        )
        df.loc[b_mask, "model_skoru"] = np.nan
        df.loc[b_mask, "grup_ici_yuzdelik"] = np.nan

        new_mask = b_mask & df["recency"].lt(30)
        second_mask = (
            b_mask
            & df["recency"].ge(30)
            & df["recency"].lt(90)
        )
        hot_mask = (
            b_mask
            & df["recency"].ge(90)
            & df["recency"].le(180)
        )
        react_mask = (
            b_mask
            & df["recency"].gt(180)
            & df["recency"].le(365)
        )
        sleeping_mask = (
            b_mask
            & df["recency"].gt(365)
            & df["recency"].le(540)
        )
        lost_mask = b_mask & df["recency"].gt(540)

        # Yaşam döngüsü segmenti
        df.loc[new_mask, "yasam_dongusu_segmenti"] = (
            "Yeni Müşteri"
        )
        df.loc[second_mask, "yasam_dongusu_segmenti"] = (
            "İkinci Satın Alma Adayı"
        )
        df.loc[hot_mask, "yasam_dongusu_segmenti"] = (
            "Sıcak Reaktivasyon"
        )
        df.loc[react_mask, "yasam_dongusu_segmenti"] = (
            "Reaktivasyon Adayı"
        )
        df.loc[sleeping_mask, "yasam_dongusu_segmenti"] = (
            "Uykuda Müşteri"
        )
        df.loc[lost_mask, "yasam_dongusu_segmenti"] = (
            "Uzun Süredir Kayıp"
        )

        # Reaktivasyon ihtiyacı
        df.loc[new_mask, "reaktivasyon_ihtiyaci"] = (
            "Yok - Bekleme"
        )
        df.loc[second_mask, "reaktivasyon_ihtiyaci"] = "Düşük"
        df.loc[hot_mask, "reaktivasyon_ihtiyaci"] = "Orta"
        df.loc[react_mask, "reaktivasyon_ihtiyaci"] = "Yüksek"
        df.loc[sleeping_mask, "reaktivasyon_ihtiyaci"] = (
            "Çok Yüksek"
        )
        df.loc[lost_mask, "reaktivasyon_ihtiyaci"] = (
            "Çok Yüksek"
        )

        # Yeni müşteri
        df.loc[new_mask, "segment"] = "Bekleme Dönemi"
        df.loc[new_mask, "kampanya_onceligi"] = (
            "Bekleme Dönemi"
        )
        df.loc[new_mask, "musteri_durumu"] = (
            "Yeni Müşteri - Bekleme Dönemi"
        )
        df.loc[new_mask, "kampanya_uygun"] = "Hayır"
        df.loc[new_mask, "skor_tipi"] = (
            "Skor hesaplanmadı (bekleme dönemi)"
        )
        df.loc[new_mask, "aksiyon_tipi"] = (
            "Yeni müşteri yönetimi"
        )
        df.loc[new_mask, "aksiyon_karari"] = (
            "Reaktivasyon mesajı gönderme; "
            "30. güne kadar bekle"
        )

        # İkinci satın alma grubu
        df.loc[second_mask, "model_skoru"] = df.loc[
            second_mask,
            "ham_model_skoru",
        ]
        df.loc[second_mask, "musteri_durumu"] = (
            "İkinci Satın Alma Adayı"
        )
        df.loc[second_mask, "kampanya_uygun"] = "Evet"
        df.loc[second_mask, "skor_tipi"] = (
            "İkinci satın alma öncelik puanı "
            "(olasılık değildir)"
        )
        df.loc[second_mask, "aksiyon_tipi"] = (
            "İkinci satın alma"
        )

        # 90+ gün reaktivasyon havuzu
        reactivation_pool = (
            hot_mask
            | react_mask
            | sleeping_mask
            | lost_mask
        )
        df.loc[reactivation_pool, "model_skoru"] = df.loc[
            reactivation_pool,
            "ham_model_skoru",
        ]
        df.loc[reactivation_pool, "musteri_durumu"] = df.loc[
            reactivation_pool,
            "yasam_dongusu_segmenti",
        ]
        df.loc[reactivation_pool, "kampanya_uygun"] = "Evet"
        df.loc[reactivation_pool, "skor_tipi"] = (
            "Reaktivasyon kampanya öncelik puanı "
            "(olasılık değildir)"
        )
        df.loc[reactivation_pool, "aksiyon_tipi"] = (
            "Reaktivasyon"
        )

        # Kampanya önceliği:
        # 30-89 gün ve 90+ gün ayrı havuzlarda sıralanır.
        for mask, group_name in [
            (second_mask, "ikinci_satin_alma"),
            (reactivation_pool, "reaktivasyon"),
        ]:
            if not mask.any():
                continue

            df.loc[mask, "grup_ici_yuzdelik"] = (
                df.loc[mask, "model_skoru"]
                .rank(pct=True, method="average")
                * 100
            )

            high = mask & df["grup_ici_yuzdelik"].ge(80)
            medium = (
                mask
                & df["grup_ici_yuzdelik"].ge(50)
                & df["grup_ici_yuzdelik"].lt(80)
            )
            low = mask & df["grup_ici_yuzdelik"].lt(50)

            df.loc[high, "segment"] = "Yüksek Öncelik"
            df.loc[medium, "segment"] = "Orta Öncelik"
            df.loc[low, "segment"] = "Düşük Öncelik"

            df.loc[high, "kampanya_onceligi"] = (
                "Yüksek Öncelik"
            )
            df.loc[medium, "kampanya_onceligi"] = (
                "Orta Öncelik"
            )
            df.loc[low, "kampanya_onceligi"] = (
                "Düşük Öncelik"
            )

            if group_name == "ikinci_satin_alma":
                df.loc[high, "aksiyon_karari"] = (
                    "İkinci satın alma kampanyasına al"
                )
                df.loc[medium, "aksiyon_karari"] = (
                    "Hafif ikinci alışveriş teşviki uygula"
                )
                df.loc[low, "aksiyon_karari"] = (
                    "İzlemeye devam et"
                )
            else:
                # Yaşam döngüsüne göre aksiyon açıklaması
                df.loc[high, "aksiyon_karari"] = (
                    "Öncelikli geri kazanım kampanyasına al"
                )
                df.loc[medium, "aksiyon_karari"] = (
                    "Standart geri kazanım akışına al"
                )
                df.loc[low, "aksiyon_karari"] = (
                    "Düşük maliyetli/toplu kampanyada izle"
                )

    return df

def fmt_number(value, digits: int = 1) -> str:
    if pd.isna(value):
        return "—"
    return f"{float(value):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_currency(value) -> str:
    return "—" if pd.isna(value) else f"₺{fmt_number(value, 2)}"


def fmt_int(value) -> str:
    if pd.isna(value):
        return "—"
    return f"{int(round(float(value))):,}".replace(",", ".")


def fmt_date(value) -> str:
    if pd.isna(value):
        return "—"
    return pd.to_datetime(value).strftime("%d.%m.%Y")


def safe_value(row: pd.Series, col: str, default=np.nan):
    return row[col] if col in row.index else default


def segment_badge(segment: str) -> str:
    s = str(segment)
    if "Bekleme" in s:
        cls = "badge-wait"
    elif "Yüksek" in s:
        cls = "badge-high"
    elif "Orta" in s:
        cls = "badge-mid"
    else:
        cls = "badge-low"
    return f'<span class="badge {cls}">{s}</span>'


def model_badge(model: str) -> str:
    cls = "badge-a" if "Model A" in str(model) else "badge-b"
    return f'<span class="badge {cls}">{model}</span>'


def metric_card(label: str, value: str, foot: str = "") -> str:
    return (
        '<div class="metric-card">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-foot">{foot}</div>'
        '</div>'
    )


def action_explanation(row: pd.Series) -> tuple[str, str]:
    model = str(safe_value(row, "model_tipi", ""))
    lifecycle = str(
        safe_value(
            row,
            "yasam_dongusu_segmenti",
            safe_value(row, "musteri_durumu", "—"),
        )
    )
    need = str(
        safe_value(row, "reaktivasyon_ihtiyaci", "—")
    )
    priority = str(
        safe_value(
            row,
            "kampanya_onceligi",
            safe_value(row, "segment", "—"),
        )
    )
    karar = str(safe_value(row, "aksiyon_karari", "—"))

    if "Model A" in model:
        detail = (
            "Müşteri birden fazla alışveriş yaptığı için "
            "Model A churn risk modeliyle değerlendiriliyor. "
            f"Kampanya önceliği: {priority}. "
            f"Önerilen işlem: {karar}. "
            "Yeni alışveriş geldiğinde Model A içinde kalır."
        )
    elif lifecycle == "Yeni Müşteri":
        detail = (
            "Müşteri yeni olduğu için reaktivasyon listesine alınmaz. "
            f"Reaktivasyon ihtiyacı: {need}. "
            f"Önerilen işlem: {karar}."
        )
    elif lifecycle == "İkinci Satın Alma Adayı":
        detail = (
            "Müşteri bekleme dönemini tamamladı. "
            f"Kampanya önceliği: {priority}. "
            f"Önerilen işlem: {karar}. "
            "İkinci alışverişten sonra Model A'ya geçer."
        )
    else:
        detail = (
            f"Yaşam döngüsü: {lifecycle}. "
            f"Reaktivasyon ihtiyacı: {need}. "
            f"Kampanya önceliği: {priority}. "
            f"Önerilen işlem: {karar}. "
            "İkinci alışverişten sonra Model A'ya geçer."
        )

    return f"{lifecycle} için önerilen CRM aksiyonu", detail

def export_customer_excel(row: pd.Series) -> bytes:
    output = BytesIO()
    pd.DataFrame({"Alan": row.index.astype(str), "Değer": row.values}).to_excel(output, index=False)
    return output.getvalue()


def group_comparison_chart(df: pd.DataFrame, row: pd.Series) -> go.Figure:
    group = df[df["model_tipi"] == row["model_tipi"]].copy()
    fields = {
        "frequency": "Alışveriş sayısı",
        "monetary": "Toplam harcama",
        "Qty": "Ürün adedi",
        "recency": "Recency",
        "avg_basket_value": "Ort. sepet",
    }
    labels, values = [], []
    for col, label in fields.items():
        if col not in group.columns or pd.isna(row.get(col)):
            continue
        median = pd.to_numeric(group[col], errors="coerce").median()
        ratio = 100 if median == 0 or pd.isna(median) else min(float(row[col]) / median * 100, 300)
        labels.append(label)
        values.append(ratio)
    fig = go.Figure(
        go.Bar(
            x=values,
            y=labels,
            orientation="h",
            text=[f"%{v:.0f}" for v in values],
            textposition="outside",
            marker=dict(color="#5B5BD6", cornerradius=7),
            hovertemplate="Grup medyanına göre %{x:.0f}%<extra></extra>",
        )
    )
    fig.add_vline(x=100, line_dash="dash", line_color="#9CA3AF")
    fig.update_layout(
        height=320,
        margin=dict(l=10, r=35, t=15, b=10),
        xaxis_title="Grup medyanı = %100",
        yaxis_title="",
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(gridcolor="#EEF0F4"),
    )
    return fig


def gauge_chart(score: float, label: str) -> go.Figure:
    score = float(np.clip(score if pd.notna(score) else 0, 0, 100))
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=score,
            number={"suffix": "", "font": {"size": 38, "color": "#172033"}},
            title={"text": label, "font": {"size": 14, "color": "#6B7280"}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 0, "tickcolor": "white"},
                "bar": {"color": "#5B5BD6", "thickness": 0.28},
                "bgcolor": "#EEF0F4",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 40], "color": "#EAF5FF"},
                    {"range": [40, 70], "color": "#FFF5DC"},
                    {"range": [70, 100], "color": "#FDECEC"},
                ],
            },
        )
    )
    fig.update_layout(height=280, margin=dict(l=20, r=20, t=35, b=10), paper_bgcolor="white")
    return fig

schedule_midnight_refresh()

# -----------------------------------------------------------------------------
# Sidebar ve veri kaynağı
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ◉ CRM Intelligence")
    st.caption("Churn ve reaktivasyon karar paneli")
    st.markdown("---")
    page = st.radio(
        "Menü",
        ["Müşteri Profili", "Portföy Özeti", "Müşteri Listesi", "Model Performansı"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    uploaded = st.file_uploader(
        "Güncel skor dosyası",
        type=["xlsx"],
    )
    # Tarih kullanıcıdan alınmaz; uygulama her çalıştırmada Türkiye saatine
    # göre bugünün tarihini otomatik kullanır.
    analysis_date = current_turkey_date()

    if uploaded is not None:
        source = uploaded
        source_name = uploaded.name
        source_version = (
            f"{uploaded.name}-"
            f"{getattr(uploaded, 'size', 0)}-"
            f"{analysis_date.isoformat()}"
        )
    elif DEFAULT_FILE.exists():
        source = DEFAULT_FILE
        source_name = DEFAULT_FILE.name
        stat = DEFAULT_FILE.stat()
        source_version = (
            f"{stat.st_mtime_ns}-{stat.st_size}-"
            f"{analysis_date.isoformat()}"
        )
    else:
        source = None
        source_name = "Dosya bulunamadı"
        source_version = "none"

    st.caption(f"Aktif kaynak: {source_name}")
    st.caption("Panel sürümü: v8 · XGBoost + 3 model karşılaştırması")
    st.success(
        "Otomatik değerlendirme tarihi: "
        f"{analysis_date.strftime('%d.%m.%Y')}"
    )

if source is None:
    st.error("Skor Excel dosyası bulunamadı. Sol menüden dosya yükleyin.")
    st.stop()

try:
    frames = load_workbook(source, source_version)
    main_raw = frames[MAIN_SHEET]
except Exception as exc:
    st.error(f"Dosya okunamadı: {exc}")
    st.stop()

id_col = detect_id_column(main_raw)
if id_col is None:
    st.error("Müşteri ID kolonu bulunamadı.")
    st.stop()

df = enrich_main(frames, id_col, analysis_date)
id_text = df[id_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

# -----------------------------------------------------------------------------
# Sayfa: Müşteri Profili
# -----------------------------------------------------------------------------
if page == "Müşteri Profili":
    st.markdown('<div class="eyebrow">Tek müşteri görünümü</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-title">Müşteri Profili</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="page-subtitle">XGBoost tabanlı Model A riskini, Model B yaşam döngüsünü, kampanya önceliğini ve önerilen aksiyonu tek ekranda inceleyin. Recency her gün otomatik güncellenir.</div>',
        unsafe_allow_html=True,
    )

    with st.form("customer_search", clear_on_submit=False):
        c1, c2 = st.columns([4, 1])
        with c1:
            customer_id = st.text_input(
                "Müşteri ID",
                placeholder=f"Örnek: {id_text.iloc[0] if len(id_text) else ''}",
                label_visibility="collapsed",
            ).strip()
        with c2:
            submitted = st.form_submit_button("Müşteriyi Aç", type="primary", use_container_width=True)

    if not customer_id:
        st.markdown(
            """
            <div class="hero">
                <div class="eyebrow">Başlamak için</div>
                <div style="font-size:1.25rem;font-weight:740;color:#172033;">Bir müşteri ID girin</div>
                <div style="color:#6B7280;margin-top:.35rem;">Skor, segment, satın alma davranışı ve CRM önerisi otomatik olarak gösterilecektir.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        preview = df[[
            id_col,
            "model_tipi",
            "yasam_dongusu_segmenti",
            "kampanya_onceligi",
            "model_skoru",
        ]].head(8).copy()
        preview.columns = [
            "Müşteri ID",
            "Model",
            "Yaşam Döngüsü",
            "Kampanya Önceliği",
            "Skor",
        ]
        st.caption("Örnek müşteri kayıtları")
        st.dataframe(preview, hide_index=True, use_container_width=True)
        st.stop()

    matches = df[id_text == customer_id]
    if matches.empty:
        try:
            matches = df[pd.to_numeric(df[id_col], errors="coerce") == float(customer_id)]
        except ValueError:
            pass
    if matches.empty:
        st.error(f"'{customer_id}' ID'li müşteri bulunamadı.")
        st.stop()

    row = matches.iloc[0]
    model = str(safe_value(row, "model_tipi", "—"))
    segment = str(
        safe_value(
            row,
            "kampanya_onceligi",
            safe_value(row, "segment", "—"),
        )
    )
    lifecycle = str(
        safe_value(
            row,
            "yasam_dongusu_segmenti",
            safe_value(row, "musteri_durumu", "—"),
        )
    )
    reactivation_need = str(
        safe_value(row, "reaktivasyon_ihtiyaci", "—")
    )

    st.markdown(
        f"""
        <div class="hero">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:1rem;flex-wrap:wrap;">
                <div>
                    <div class="eyebrow">Müşteri kimliği</div>
                    <div style="font-size:2rem;font-weight:780;color:#172033;">{safe_value(row, id_col)}</div>
                    <div style="margin-top:.55rem;">{model_badge(model)}{segment_badge(segment)}</div>
                    <div style="margin-top:.45rem;color:#4B5563;font-size:.88rem;font-weight:650;">
                        {lifecycle} · Reaktivasyon ihtiyacı: {reactivation_need} · Kampanya uygunluğu: {safe_value(row, 'kampanya_uygun', '—')}
                    </div>
                </div>
                <div style="text-align:right;">
                    <div class="small-note">Son alışveriş</div>
                    <div style="font-size:1.05rem;font-weight:700;color:#172033;">{fmt_date(safe_value(row, 'LastPurchaseDate'))}</div>
                    <div class="small-note" style="margin-top:.45rem;">Değerlendirme tarihi</div>
                    <div style="font-size:.95rem;font-weight:700;color:#172033;">{fmt_date(safe_value(row, 'skor_tarihi'))}</div>
                    <div class="small-note" style="margin-top:.45rem;">Skor türü</div>
                    <div style="font-size:.9rem;font-weight:650;color:#4B5563;">{safe_value(row, 'skor_tipi', '—')}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cards = st.columns(5)
    card_data = [
        ("Model skoru", fmt_number(safe_value(row, "model_skoru"), 1),
         "Bekleme döneminde hesaplanmaz" if pd.isna(safe_value(row, "model_skoru")) else "0–100 arası model çıktısı"),
        ("Grup yüzdeliği", f"%{fmt_number(safe_value(row, 'grup_ici_yuzdelik'), 1)}", "Kendi model grubu içindeki sıra"),
        ("Alışveriş", fmt_int(safe_value(row, "frequency")), "Toplam sipariş sayısı"),
        (
            "Güncel recency",
            f"{fmt_int(safe_value(row, 'recency'))} gün",
            (
                f"Değerlendirme: "
                f"{fmt_date(safe_value(row, 'skor_tarihi'))} "
                f"· Veri kesitinde "
                f"{fmt_int(safe_value(row, 'recency_snapshot'))} gün"
            ),
        ),
        ("Toplam harcama", fmt_currency(safe_value(row, "monetary")), "Müşteri yaşam boyu değeri"),
    ]
    for col, (label, value, foot) in zip(cards, card_data):
        col.markdown(metric_card(label, value, foot), unsafe_allow_html=True)

    decision_cards = st.columns(3)
    decision_data = [
        (
            "Yaşam döngüsü",
            lifecycle,
            "Müşterinin zamana dayalı davranış durumu",
        ),
        (
            "Reaktivasyon ihtiyacı",
            reactivation_need,
            "Ne kadar geri kazanım ihtiyacı olduğunu gösterir",
        ),
        (
            "Kampanya önceliği",
            segment,
            "CRM bütçesinde ele alınma sırası",
        ),
    ]
    for col, (label, value, foot) in zip(
        decision_cards,
        decision_data,
    ):
        col.markdown(
            metric_card(label, value, foot),
            unsafe_allow_html=True,
        )

    title, detail = action_explanation(row)
    st.markdown(
        f'<div class="action-box"><div class="action-title">{title}</div><div class="action-text">{detail}</div></div>',
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3 = st.tabs(["Karar Özeti", "Davranış Analizi", "Tüm Sütunlar"])

    with tab1:
        left, right = st.columns([1, 1.25])
        with left:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            current_score = safe_value(row, "model_skoru")
            if pd.isna(current_score):
                st.markdown(
                    """
                    <div style="height:250px;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;">
                        <div style="font-size:2.1rem;">⏳</div>
                        <div style="font-size:1.25rem;font-weight:760;color:#172033;margin-top:.5rem;">Bekleme Dönemi</div>
                        <div style="color:#6B7280;margin-top:.35rem;max-width:360px;">Bu müşteri için henüz reaktivasyon skoru hesaplanmaz ve kampanya mesajı gönderilmez.</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.plotly_chart(
                    gauge_chart(float(current_score), str(safe_value(row, "skor_tipi", "Model skoru"))),
                    use_container_width=True,
                    config={"displayModeBar": False},
                )
            st.markdown('</div>', unsafe_allow_html=True)
        with right:
            st.markdown("### Satın alma özeti")
            summary = pd.DataFrame(
                {
                    "Gösterge": [
                        "Toplam ürün adedi", "Ortalama sepet", "Sipariş başına ürün",
                        "Aktif müşteri süresi", "Ortalama alışveriş aralığı", "İlk alışveriş",
                    ],
                    "Değer": [
                        fmt_int(safe_value(row, "Qty")),
                        fmt_currency(safe_value(row, "avg_basket_value")),
                        fmt_number(safe_value(row, "avg_qty_per_order"), 2),
                        f"{fmt_int(safe_value(row, 'active_days'))} gün",
                        f"{fmt_number(safe_value(row, 'inter_purchase_time'), 1)} gün" if pd.notna(safe_value(row, "inter_purchase_time")) else "Tek alışveriş",
                        fmt_date(safe_value(row, "FirstPurchaseDate")),
                    ],
                }
            )
            st.dataframe(summary, hide_index=True, use_container_width=True, height=250)

            if "Model B" in model:
                st.info("Bu müşteri ikinci alışverişini yaptığında frequency 2 olur ve pipeline tekrar çalıştığında Model A'ya geçer.")
            else:
                st.info("Bu müşteri yeni alışveriş yaptığında Model A'da kalır; satın alma değişkenleri ve churn skoru yeniden hesaplanır.")

    with tab2:
        c1, c2 = st.columns([1.35, 1])
        with c1:
            st.markdown("### Kendi grubuyla karşılaştırma")
            st.caption("%100 çizgisi, aynı model grubundaki medyan müşteriyi temsil eder.")
            st.plotly_chart(group_comparison_chart(df, row), use_container_width=True, config={"displayModeBar": False})
        with c2:
            st.markdown("### Model özel göstergeleri")
            if "Model B" in model:
                details = pd.DataFrame(
                    {
                        "Gösterge": [
                            "Yaşam döngüsü",
                            "Reaktivasyon ihtiyacı",
                            "Kampanya önceliği",
                            "Değer yüzdeliği",
                            "Tazelik skoru",
                            "Kampanya öncelik skoru",
                        ],
                        "Değer": [
                            lifecycle,
                            reactivation_need,
                            segment,
                            f"%{fmt_number(safe_value(row, 'value_percentile'), 1)}",
                            fmt_number(safe_value(row, "freshness_score"), 1),
                            fmt_number(safe_value(row, "model_skoru"), 1),
                        ],
                    }
                )
            else:
                details = pd.DataFrame(
                    {
                        "Gösterge": ["Aktif gün", "Alışveriş aralığı", "Günlük harcama", "Günlük ürün"],
                        "Değer": [
                            f"{fmt_int(safe_value(row, 'active_days'))} gün",
                            f"{fmt_number(safe_value(row, 'inter_purchase_time'), 1)} gün",
                            fmt_currency(safe_value(row, "monetary_per_day")),
                            fmt_number(safe_value(row, "qty_per_day"), 3),
                        ],
                    }
                )
            st.dataframe(details, hide_index=True, use_container_width=True)

    with tab3:
        long_df = pd.DataFrame({"Sütun": row.index.astype(str), "Değer": [str(v) if pd.notna(v) else "—" for v in row.values]})
        st.dataframe(long_df, hide_index=True, use_container_width=True, height=480)
        st.download_button(
            "Müşteri profilini Excel olarak indir",
            data=export_customer_excel(row),
            file_name=f"musteri_{safe_value(row, id_col)}_profil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# -----------------------------------------------------------------------------
# Sayfa: Portföy Özeti
# -----------------------------------------------------------------------------
elif page == "Portföy Özeti":
    st.markdown(
        '<div class="eyebrow">Yönetim görünümü</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-title">Portföy Özeti</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-subtitle">'
        'Yaşam döngüsü ve kampanya önceliği ayrı ayrı gösterilir.'
        '</div>',
        unsafe_allow_html=True,
    )

    total = len(df)
    a_count = int(
        df["model_tipi"]
        .astype(str)
        .str.contains("Model A")
        .sum()
    )
    b_count = total - a_count
    high = int(
        df["kampanya_onceligi"]
        .astype(str)
        .str.contains("Yüksek")
        .sum()
    )
    lost = int(
        df["yasam_dongusu_segmenti"]
        .astype(str)
        .eq("Uzun Süredir Kayıp")
        .sum()
    )

    cards = st.columns(5)
    for col, data in zip(
        cards,
        [
            (
                "Toplam müşteri",
                fmt_int(total),
                "Tüm müşteri portföyü",
            ),
            (
                "Model A",
                fmt_int(a_count),
                "Çoklu alışveriş müşterisi",
            ),
            (
                "Model B",
                fmt_int(b_count),
                "Tek alışveriş müşterisi",
            ),
            (
                "Yüksek kampanya önceliği",
                fmt_int(high),
                "CRM aksiyonunda ilk grup",
            ),
            (
                "Uzun süredir kayıp",
                fmt_int(lost),
                "541+ gündür alışveriş yok",
            ),
        ],
    ):
        col.markdown(
            metric_card(*data),
            unsafe_allow_html=True,
        )

    c1, c2 = st.columns(2)

    with c1:
        life = (
            df[
                df["model_tipi"]
                .astype(str)
                .str.contains("Model B")
            ]["yasam_dongusu_segmenti"]
            .value_counts()
            .reset_index()
        )
        life.columns = ["Yaşam döngüsü", "Müşteri"]
        fig = px.bar(
            life,
            x="Yaşam döngüsü",
            y="Müşteri",
            text="Müşteri",
            title="Model B yaşam döngüsü dağılımı",
        )
        fig.update_traces(
            marker_color="#5B5BD6",
            textposition="outside",
        )
        fig.update_layout(
            height=420,
            margin=dict(l=10, r=10, t=55, b=100),
            xaxis_title="",
            yaxis_title="Müşteri",
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

    with c2:
        priority = (
            df["kampanya_onceligi"]
            .value_counts()
            .reset_index()
        )
        priority.columns = ["Kampanya önceliği", "Müşteri"]
        fig = px.pie(
            priority,
            names="Kampanya önceliği",
            values="Müşteri",
            hole=0.62,
            title="Kampanya önceliği dağılımı",
        )
        fig.update_traces(
            textposition="inside",
            textinfo="percent+label",
        )
        fig.update_layout(
            height=420,
            margin=dict(l=10, r=10, t=55, b=10),
            legend_title="",
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

    scored_df = df[df["model_skoru"].notna()].copy()
    fig = px.histogram(
        scored_df,
        x="model_skoru",
        color="model_tipi",
        nbins=25,
        barmode="overlay",
        opacity=0.7,
        title="Model skorlarının dağılımı",
        labels={
            "model_skoru": "Model skoru",
            "count": "Müşteri",
        },
    )
    fig.update_layout(
        height=410,
        margin=dict(l=10, r=10, t=55, b=10),
        legend_title="",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )

# -----------------------------------------------------------------------------
# Sayfa: Müşteri Listesi
# -----------------------------------------------------------------------------
elif page == "Müşteri Listesi":
    st.markdown(
        '<div class="eyebrow">Operasyon görünümü</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-title">Müşteri Listesi</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-subtitle">'
        'Yaşam döngüsü ve kampanya önceliğine göre müşterileri filtreleyin.'
        '</div>',
        unsafe_allow_html=True,
    )

    f1, f2, f3, f4 = st.columns(4)

    model_options = ["Tümü"] + sorted(
        df["model_tipi"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    lifecycle_options = ["Tümü"] + sorted(
        df["yasam_dongusu_segmenti"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    priority_options = ["Tümü"] + sorted(
        df["kampanya_onceligi"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    eligibility_options = ["Tümü"] + sorted(
        df["kampanya_uygun"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    with f1:
        selected_model = st.selectbox(
            "Model",
            model_options,
        )
    with f2:
        selected_lifecycle = st.selectbox(
            "Yaşam döngüsü",
            lifecycle_options,
        )
    with f3:
        selected_priority = st.selectbox(
            "Kampanya önceliği",
            priority_options,
        )
    with f4:
        selected_eligibility = st.selectbox(
            "Kampanya uygunluğu",
            eligibility_options,
        )

    filtered = df.copy()

    if selected_model != "Tümü":
        filtered = filtered[
            filtered["model_tipi"] == selected_model
        ]
    if selected_lifecycle != "Tümü":
        filtered = filtered[
            filtered["yasam_dongusu_segmenti"]
            == selected_lifecycle
        ]
    if selected_priority != "Tümü":
        filtered = filtered[
            filtered["kampanya_onceligi"]
            == selected_priority
        ]
    if selected_eligibility != "Tümü":
        filtered = filtered[
            filtered["kampanya_uygun"]
            == selected_eligibility
        ]

    filtered = filtered.sort_values(
        ["grup_ici_yuzdelik", "model_skoru"],
        ascending=False,
        na_position="last",
    )

    st.markdown(
        metric_card(
            "Filtrelenen müşteri",
            fmt_int(len(filtered)),
            "Seçili koşullara uyan kayıt",
        ),
        unsafe_allow_html=True,
    )

    show_cols = [
        id_col,
        "yasam_dongusu_segmenti",
        "reaktivasyon_ihtiyaci",
        "kampanya_onceligi",
        "kampanya_uygun",
        "model_tipi",
        "model_skoru",
        "grup_ici_yuzdelik",
        "frequency",
        "monetary",
        "Qty",
        "recency",
        "aksiyon_karari",
    ]
    show_cols = [
        c for c in show_cols
        if c in filtered.columns
    ]

    st.dataframe(
        filtered[show_cols],
        hide_index=True,
        use_container_width=True,
        height=520,
    )

    st.download_button(
        "Filtrelenmiş listeyi CSV olarak indir",
        filtered[show_cols]
        .to_csv(index=False)
        .encode("utf-8-sig"),
        file_name="crm_musteri_listesi_v8.csv",
        mime="text/csv",
        use_container_width=True,
    )

# -----------------------------------------------------------------------------
# Sayfa: Model Performansı
# -----------------------------------------------------------------------------
else:
    st.markdown(
        '<div class="eyebrow">Analitik görünüm</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-title">Model Performansı</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-subtitle">'
        'Random Forest, XGBoost ve LightGBM validation setinde karşılaştırıldı. '
        'Threshold, validation F1 skorunu maksimum yapan değer olarak seçildi.'
        '</div>',
        unsafe_allow_html=True,
    )

    perf = frames.get("Model_Performansi", pd.DataFrame())
    comparison = frames.get("Model_Karsilastirma", pd.DataFrame())
    selected = pd.Series(dtype=object)

    if not perf.empty:
        if "degerlendirme" in perf.columns:
            selected_mask = (
                perf["degerlendirme"]
                .astype(str)
                .str.contains(
                    "Seçilen F1 threshold",
                    case=False,
                    na=False,
                )
            )
            selected = (
                perf.loc[selected_mask].iloc[0]
                if selected_mask.any()
                else perf.iloc[0]
            )
        else:
            selected = perf.iloc[0]

        model_name = str(
            selected.get("model_name", "Model A")
        )

        cards = st.columns(5)
        metrics = [
            (
                "Kazanan model",
                model_name,
                "Validation F1'e göre seçildi",
            ),
            (
                "ROC-AUC",
                fmt_number(selected.get("roc_auc"), 3),
                "Test ayırt etme gücü",
            ),
            (
                "Precision",
                fmt_number(
                    selected.get("precision_churn"),
                    3,
                ),
                "Hedeflenen churn doğruluğu",
            ),
            (
                "Recall",
                fmt_number(
                    selected.get("recall_churn"),
                    3,
                ),
                "Yakalanan churn oranı",
            ),
            (
                "Threshold",
                fmt_number(selected.get("threshold"), 2),
                "Validation F1 maksimum",
            ),
        ]

        for col, data in zip(cards, metrics):
            col.markdown(
                metric_card(*data),
                unsafe_allow_html=True,
            )

        split_cards = st.columns(3)
        split_data = [
            (
                "Train",
                fmt_int(selected.get("n_train")),
                "Modelin öğrendiği veri",
            ),
            (
                "Validation",
                fmt_int(selected.get("n_validation")),
                "Model ve threshold seçimi",
            ),
            (
                "Test",
                fmt_int(selected.get("n_test")),
                "Nihai bağımsız değerlendirme",
            ),
        ]

        for col, data in zip(split_cards, split_data):
            col.markdown(
                metric_card(*data),
                unsafe_allow_html=True,
            )

    if not comparison.empty:
        st.markdown("### Model karşılaştırması")

        compare_cols = [
            c
            for c in [
                "model_name",
                "selected_threshold",
                "val_precision",
                "val_recall",
                "val_f1",
                "val_roc_auc",
            ]
            if c in comparison.columns
        ]

        st.dataframe(
            comparison[compare_cols],
            hide_index=True,
            use_container_width=True,
        )

        if {"model_name", "val_f1"}.issubset(
            comparison.columns
        ):
            compare_plot = comparison.sort_values(
                "val_f1",
                ascending=True,
            )

            fig = px.bar(
                compare_plot,
                x="val_f1",
                y="model_name",
                orientation="h",
                text="val_f1",
                title="Validation F1 karşılaştırması",
            )
            fig.update_traces(
                marker_color="#5B5BD6",
                texttemplate="%{text:.3f}",
                textposition="outside",
            )
            fig.update_layout(
                height=330,
                margin=dict(l=10, r=10, t=55, b=10),
                xaxis_title="Validation F1",
                yaxis_title="",
            )
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )

    threshold = frames.get(
        "Threshold_Analizi",
        pd.DataFrame(),
    )
    lift = frames.get(
        "Lift_Analizi",
        pd.DataFrame(),
    )
    feature = frames.get(
        "Feature_Importance",
        pd.DataFrame(),
    )

    c1, c2 = st.columns(2)

    with c1:
        if not threshold.empty:
            metric_cols = [
                c
                for c in ["precision", "recall", "f1"]
                if c in threshold.columns
            ]

            long = threshold.melt(
                id_vars="threshold",
                value_vars=metric_cols,
                var_name="Metrik",
                value_name="Değer",
            )

            fig = px.line(
                long,
                x="threshold",
                y="Değer",
                color="Metrik",
                markers=True,
                title="Validation threshold dengesi",
            )

            if not selected.empty:
                chosen = pd.to_numeric(
                    pd.Series(
                        [selected.get("threshold")]
                    ),
                    errors="coerce",
                ).iloc[0]

                if pd.notna(chosen):
                    fig.add_vline(
                        x=float(chosen),
                        line_dash="dash",
                        line_color="#D9534F",
                        annotation_text=(
                            f"Seçilen: {float(chosen):.2f}"
                        ),
                    )

            fig.update_layout(
                height=390,
                margin=dict(l=10, r=10, t=55, b=10),
                yaxis_range=[0, 1],
            )
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )

    with c2:
        if not lift.empty:
            fig = px.line(
                lift,
                x="decile",
                y="lift",
                markers=True,
                title="Test setinde decile bazında lift",
            )
            fig.add_hline(
                y=1,
                line_dash="dash",
                line_color="#9CA3AF",
            )
            fig.update_layout(
                height=390,
                margin=dict(l=10, r=10, t=55, b=10),
                xaxis=dict(autorange="reversed"),
            )
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )

    if not feature.empty:
        feature_sorted = feature.sort_values(
            "importance",
            ascending=True,
        )
        fig = px.bar(
            feature_sorted,
            x="importance",
            y="feature",
            orientation="h",
            title="Kazanan Model A değişken önemleri",
        )
        fig.update_traces(marker_color="#5B5BD6")
        fig.update_layout(
            height=430,
            margin=dict(l=10, r=10, t=55, b=10),
            xaxis_title="Önem",
            yaxis_title="",
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

if id_col == "pipeline_musteri_id":
    st.caption("Not: pipeline_musteri_id satır sırasından üretilmiş geçici bir kimliktir. Canlı kullanımda sabit CustomerID/CRM_ID alanı kullanılmalıdır.")
