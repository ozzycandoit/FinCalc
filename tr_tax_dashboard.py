"""
tr_tax_dashboard.py
Turkey Foreign Securities Tax Calculator — web front-end

Two tiers:
  - Quick Check (free): stocks/ETFs only, no options, no Yİ-ÜFE indexation.
    Fast approximate estimate, on-screen only, no filing-ready PDF.
  - Full Calculation (paid once STRIPE_SECRET_KEY / CALC_PRICE_TRY are set):
    the full engine — options, Yİ-ÜFE indexation, GİB Hazır Beyan PDF.

Until STRIPE_SECRET_KEY is configured in Secrets, Full Calculation stays free
too (PAID_MODE below), so the app is fully usable before monetization is
switched on.
"""

import io
import os

import streamlit as st
import pandas as pd

import tr_tax_core
import tr_tax_report

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG  (all overridable via environment variables / Streamlit secrets)
# ─────────────────────────────────────────────────────────────────────────────
STRIPE_SECRET_KEY  = os.getenv("STRIPE_SECRET_KEY", "")          # blank  = free
CALC_PRICE_TRY     = int(os.getenv("CALC_PRICE_TRY", "0"))       # kuruş, e.g. 49900 = ₺499
PAID_MODE          = bool(STRIPE_SECRET_KEY)

if PAID_MODE:
    import stripe
    stripe.api_key = STRIPE_SECRET_KEY

# ─────────────────────────────────────────────────────────────────────────────
# i18n strings
# ─────────────────────────────────────────────────────────────────────────────
TR = {
    "page_title": "Yurt Dışı Menkul Kıymet Vergi Hesaplama",
    "page_icon": "📊",
    "hero": "📊 Yurt Dışı Menkul Kıymet Vergi Hesaplama",
    "hero_sub": (
        "Yurt dışı hisse, ETF ve opsiyon işlemleriniz için 2026 Türkiye gelir vergisi tahmini. "
        "**USD ve EUR** cinsinden varlıklar desteklenir. Sonuç **GİB Hazır Beyan** formatında üretilir."
    ),
    "lang_label": "🌐 Dil / Language",

    # ── Tier selection ──────────────────────────────────────────────────────
    "tier_head": "Nasıl hesaplamak istersiniz?",
    "tier_free_badge": "ÜCRETSİZ",
    "tier_free_title": "🆓 Hızlı Kontrol",
    "tier_free_desc": (
        "Sadece hisse/ETF alım-satımları için hızlı, yaklaşık kazanç ve vergi tahmini. "
        "Opsiyon yok, Yİ-ÜFE endekslemesi yok, dosyalanabilir PDF yok — sadece hızlı bir fikir."
    ),
    "tier_free_cta": "Ücretsiz Kullan →",
    "tier_full_badge_priced": "Hesaplama başına {price}",
    "tier_full_badge_free": "ŞİMDİLİK ÜCRETSİZ",
    "tier_full_title": "💎 Tam Hesaplama",
    "tier_full_desc": (
        "Opsiyonlar dahil, Yİ-ÜFE endekslemesi ile tam FIFO hesaplaması. "
        "**GİB Hazır Beyan** formatında indirilebilir PDF ve işlenmiş Excel içerir."
    ),
    "tier_full_cta": "Tam Hesaplamaya Geç →",
    "back_to_tiers": "‹ Hesaplama türünü değiştir",

    # ── Quick check tier ─────────────────────────────────────────────────────
    "quick_title": "🆓 Hızlı Kontrol",
    "quick_step1_head": "1️⃣ Boş şablonu indirin",
    "quick_step1_body": "Her satıra bir ALIM veya SATIM işlemi girin. Sadece hisse ve ETF'ler.",
    "quick_step1_btn": "⬇️ Hızlı_Kontrol_Sablonu.xlsx indir",
    "quick_step2_head": "2️⃣ Yükleyin ve hesaplayın",
    "quick_step2_body": "Doldurduğunuz dosyayı yükleyin ve Hesapla butonuna basın.",
    "quick_upload_label": "Doldurulmuş şablon",
    "quick_calc_btn": "🧮 Hızlı Hesapla (ücretsiz)",
    "quick_spinner": "TCMB kurları alınıyor, işlemler hesaplanıyor…",
    "quick_success": "Hızlı hesaplama tamamlandı ✅",
    "quick_gain": "Toplam kazanç (TL)",
    "quick_tax": "Tahmini 2026 gelir vergisi",
    "quick_instal": "Taksitler (Mart / Temmuz)",
    "quick_trades": "İşlenen işlem sayısı",
    "quick_upgrade_head": "Opsiyonlarınız mı var? Kesin ve dosyalanabilir bir sonuç mu istiyorsunuz?",
    "quick_upgrade_body": "Tam Hesaplama; Yİ-ÜFE endekslemesi, opsiyonlar ve indirilebilir GİB Hazır Beyan PDF'i içerir.",
    "quick_upgrade_btn": "💎 Tam Hesaplamaya Geç",

    "step1_head": "1️⃣ Boş çalışma kitabını indirin",
    "step1_body": (
        "İşlemlerinizi gireceğiniz Excel dosyasını indirin. ASSET sayfalarına "
        "BUY/SELL satırlarınızı, OPTION sayfalarına opsiyon işlemlerinizi girin. "
        "Kılavuz için README sayfasına bakın."
    ),
    "step1_btn": "⬇️ Turkey_Tax_Tracker.xlsx indir",

    "step2_head": "2️⃣ Hesaplayın",
    "step2_body": "Doldurduğunuz dosyayı yükleyin ve Hesapla butonuna basın.",
    "evds_label": "EVDS anahtarı (Yİ-ÜFE, opsiyonel)",
    "evds_help": "Boş bırakılırsa yerleşik Yİ-ÜFE tablosu kullanılır. evds2.tcmb.gov.tr ücretsiz.",
    "upload_label": "Doldurulmuş Turkey_Tax_Tracker.xlsx",
    "div_head": "Temettü girişi (opsiyonel) — 3.C Menkul Sermaye İradı",
    "div_cap": "Yurt dışı temettüleriniz varsa ekleyin; TCMB döviz alış ile TL'ye çevrilir.",
    "div_date": "Tarih (YYYY-AA-GG)", "div_ccy": "Döviz",
    "div_gross": "Brüt temettü", "div_wh": "Yurt dışı stopaj", "div_exp": "Gider",
    "calc_btn_free": "🧮 Hesapla (ücretsiz)",
    "calc_btn_paid": f"🧮 Hesapla (₺{CALC_PRICE_TRY/100:.0f})",
    "spinner": "TCMB kurları ve Yİ-ÜFE alınıyor, işlemler hesaplanıyor…",
    "success": "Hesaplama tamamlandı ✅",
    "status_label": "Durum",
    "ufe_label": "Yİ-ÜFE kaynağı",
    "matrah": "Matrah (vergiye esas)",
    "tax": "2026 Gelir Vergisi",
    "instal": "Taksitler (Mart / Temmuz)",
    "tab_beyan": "Hazır Beyan Sistemi Özeti",
    "tab_detail": "Yatırım İşlemleri Detaylı Raporu",
    "3c_head": "3.C BEYAN EDİLECEK MENKUL SERMAYE İRADI GELİRLERİNİZ",
    "3d_head": "3.D BEYAN EDİLECEK DİĞER KAZANÇ VE İRAT GELİRLERİNİZ",
    "irat_turu": "İradın Türü", "gayrisafi_irat": "Gayrisafi İrat",
    "indirilebilir": "İndirilecek Giderler", "safi_irat": "Safi İrat",
    "kesilen": "Kesilen Gelir Vergisi", "kaz_turu": "Kazancın Türü",
    "gayrisafi_tutar": "Gayrisafi Tutar", "gider_indirim": "Gider / İndirim",
    "safi_kaz": "Safi Kazanç",
    "col_inst": "Enstrüman", "col_date": "Tarih", "col_action": "İşlem",
    "col_qty": "Adet", "col_price": "Fiyat", "col_tl": "TL Tutar",
    "col_gross": "Brüt K/Z", "col_taxable": "Vergiye Esas", "col_status": "Durum",
    "no_trades": "İşlem bulunamadı.",
    "dl_pdf": "⬇️ Hazır Beyan PDF indir",
    "dl_xlsx": "⬇️ İşlenmiş Excel indir",
    "err": "Hata: ",
    "disclaimer_head": "⚠️ Sorumluluk Reddi",
    "disclaimer": (
        "Bu araç yalnızca **tahmini hesaplama** amaçlıdır; resmi vergi beyannamesi değildir. "
        "Yİ-ÜFE endekslemesi ve türev işlemler gibi sınır durumlarda hesaplamalar farklılık "
        "gösterebilir. Beyan etmeden önce sonuçları bir **mali müşavir** ile teyit edin. "
        "Bu araç aracılığıyla elde edilen sonuçlardan doğacak vergi cezaları veya mali "
        "kayıplardan sorumluluk kabul edilmez."
    ),
    "payment_head": "💳 Ödeme",
    "payment_body": "Hesaplama başlatmak için ödeme gerekiyor.",
    "payment_btn": "Ödemeye git",
}

EN = {
    "page_title": "Foreign Securities Tax Calculator — Turkey",
    "page_icon": "📊",
    "hero": "📊 Foreign Securities Tax Calculator — Turkey",
    "hero_sub": (
        "Estimate your 2026 Turkey income tax on foreign stocks, ETFs and options. "
        "Both **USD and EUR** assets are supported. "
        "Results are produced in **GİB Hazır Beyan** (Turkish tax return) format."
    ),
    "lang_label": "🌐 Dil / Language",

    # ── Tier selection ──────────────────────────────────────────────────────
    "tier_head": "How would you like to calculate?",
    "tier_free_badge": "FREE",
    "tier_free_title": "🆓 Quick Check",
    "tier_free_desc": (
        "A fast, approximate gain/tax estimate for stock & ETF trades only. "
        "No options, no Yİ-ÜFE indexation, no filing-ready PDF — just a quick read."
    ),
    "tier_free_cta": "Use for Free →",
    "tier_full_badge_priced": "{price} per calculation",
    "tier_full_badge_free": "FREE FOR NOW",
    "tier_full_title": "💎 Full Calculation",
    "tier_full_desc": (
        "Full FIFO calculation with Yİ-ÜFE indexation, including options. "
        "Includes a downloadable **GİB Hazır Beyan** format PDF and processed Excel."
    ),
    "tier_full_cta": "Go to Full Calculation →",
    "back_to_tiers": "‹ Change calculation type",

    # ── Quick check tier ─────────────────────────────────────────────────────
    "quick_title": "🆓 Quick Check",
    "quick_step1_head": "1️⃣ Download the blank template",
    "quick_step1_body": "Enter one BUY or SELL per row. Stocks & ETFs only.",
    "quick_step1_btn": "⬇️ Download Quick_Check_Template.xlsx",
    "quick_step2_head": "2️⃣ Upload and calculate",
    "quick_step2_body": "Upload your filled template and press Calculate.",
    "quick_upload_label": "Filled template",
    "quick_calc_btn": "🧮 Quick Calculate (free)",
    "quick_spinner": "Fetching TCMB rates, processing trades…",
    "quick_success": "Quick check complete ✅",
    "quick_gain": "Total gain (TL)",
    "quick_tax": "Estimated 2026 income tax",
    "quick_instal": "Instalments (March / July)",
    "quick_trades": "Trades processed",
    "quick_upgrade_head": "Have options? Want a precise, filing-ready result?",
    "quick_upgrade_body": "Full Calculation adds Yİ-ÜFE indexation, options, and a downloadable GİB Hazır Beyan PDF.",
    "quick_upgrade_btn": "💎 Go to Full Calculation",

    "step1_head": "1️⃣ Download the blank workbook",
    "step1_body": (
        "Download the Excel file and fill in your trades. Enter BUY/SELL rows on "
        "ASSET sheets and option legs on OPTION sheets. See the README sheet for "
        "instructions."
    ),
    "step1_btn": "⬇️ Download Turkey_Tax_Tracker.xlsx",

    "step2_head": "2️⃣ Calculate",
    "step2_body": "Upload your filled workbook and press Calculate.",
    "evds_label": "EVDS key for Yİ-ÜFE indexation (optional)",
    "evds_help": "Leave blank to use the built-in Yİ-ÜFE table. Free key at evds2.tcmb.gov.tr.",
    "upload_label": "Filled Turkey_Tax_Tracker.xlsx",
    "div_head": "Dividend input (optional) — 3.C Menkul Sermaye İradı",
    "div_cap": "Add foreign dividends; they are converted to TL at the official TCMB rate.",
    "div_date": "Date (YYYY-MM-DD)", "div_ccy": "Currency",
    "div_gross": "Gross dividend", "div_wh": "Foreign withholding", "div_exp": "Expense",
    "calc_btn_free": "🧮 Calculate (free)",
    "calc_btn_paid": f"🧮 Calculate (₺{CALC_PRICE_TRY/100:.0f})",
    "spinner": "Fetching TCMB rates and Yİ-ÜFE, processing stocks + options…",
    "success": "Calculation complete ✅",
    "status_label": "Status",
    "ufe_label": "Yİ-ÜFE source",
    "matrah": "Taxable base",
    "tax": "2026 income tax",
    "instal": "Instalments (March / July)",
    "tab_beyan": "Hazır Beyan Summary",
    "tab_detail": "Investment Transactions Detail",
    "3c_head": "3.C FOREIGN DIVIDEND INCOME (Menkul Sermaye İradı)",
    "3d_head": "3.D OTHER CAPITAL GAINS (Değer Artışı Kazancı — GVK Mük. 80/1)",
    "irat_turu": "Income type", "gayrisafi_irat": "Gross income",
    "indirilebilir": "Deductible expenses", "safi_irat": "Net income",
    "kesilen": "Tax withheld", "kaz_turu": "Gain type",
    "gayrisafi_tutar": "Gross amount", "gider_indirim": "Deduction / indexation",
    "safi_kaz": "Net gain",
    "col_inst": "Instrument", "col_date": "Date", "col_action": "Action",
    "col_qty": "Qty", "col_price": "Price", "col_tl": "TL Amount",
    "col_gross": "Gross P/L", "col_taxable": "Taxable", "col_status": "Status",
    "no_trades": "No transactions found.",
    "dl_pdf": "⬇️ Download Hazır Beyan PDF",
    "dl_xlsx": "⬇️ Download processed Excel",
    "err": "Error: ",
    "disclaimer_head": "⚠️ Disclaimer",
    "disclaimer": (
        "This tool provides **estimates only** and does not constitute an official tax filing. "
        "Edge cases — particularly Yİ-ÜFE inflation indexation and derivative instruments — may "
        "produce figures that differ from a professional assessment. **Verify all results with a "
        "licensed tax advisor (mali müşavir) before filing.** No liability is accepted for tax "
        "penalties or financial losses arising from use of this tool."
    ),
    "payment_head": "💳 Payment",
    "payment_body": "A payment is required to start the calculation.",
    "payment_btn": "Go to payment",
}

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def tl(x):
    try:
        return f"{float(x):,.2f} ₺"
    except Exception:
        return "-"


@st.cache_data(show_spinner=False)
def get_blank_workbook_bytes(lang="TR"):
    """Build a fresh blank full workbook in memory and return its bytes."""
    import tax_tool as tt
    wb = tt.build_blank_workbook(lang=lang)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def get_quick_template_bytes(lang="TR"):
    """Blank Quick Check template (free tier)."""
    return tr_tax_core.build_quick_template_bytes(lang=lang)


def create_stripe_checkout(price_try_kurus):
    """Create a Stripe PaymentIntent and return client_secret.
    Raises on failure."""
    intent = stripe.PaymentIntent.create(
        amount=price_try_kurus,
        currency="try",
        automatic_payment_methods={"enabled": True},
    )
    return intent.client_secret


# ─────────────────────────────────────────────────────────────────────────────
# Page setup
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Turkey Tax Calculator", page_icon="📊", layout="wide")

# Editorial "Classical" look (matches the app's Claude-design mockup):
# serif headings, small uppercase kicker labels, outlined terracotta buttons,
# a numbered step indicator, and stat rows with a hairline rule instead of a
# filled tile. Kept as CSS only (no layout-critical HTML swapped in) so a
# future Streamlit version can't silently break the page.
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=Inter:wght@400;500;600&display=swap');

    h1, h2, h3, [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3 {
        font-family: 'Source Serif 4', Georgia, serif !important;
        font-weight: 600 !important;
        color: #3D3929 !important;
        letter-spacing: -0.01em;
    }

    /* Small uppercase kicker label, used above step/section headings */
    .eyebrow {
        font-family: 'Inter', sans-serif;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        color: #D97757;
        margin: 0 0 2px 0;
    }

    /* Status / tier badge pill */
    .badge-pill {
        display: inline-block;
        background: #F0EEE5;
        color: #B85C38;
        font-family: 'Inter', sans-serif;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        padding: 4px 12px;
        border-radius: 20px;
        margin-bottom: 8px;
    }

    /* Numbered step indicator (Trade -> Exemptions -> Result style) */
    .step-wizard { display: flex; align-items: center; margin: 10px 0 22px 0; font-family: 'Inter', sans-serif; }
    .step-wizard .step { display: flex; align-items: center; gap: 8px; white-space: nowrap; }
    .step-wizard .circle {
        width: 24px; height: 24px; border-radius: 50%;
        border: 1.5px solid #D8D4C8; color: #8A8578;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.72rem; font-weight: 700; flex-shrink: 0;
    }
    .step-wizard .circle.active { border-color: #D97757; color: #D97757; }
    .step-wizard .circle.done { background: #D97757; border-color: #D97757; color: #FAF9F5; }
    .step-wizard .label { font-size: 0.72rem; letter-spacing: 0.06em; text-transform: uppercase; color: #8A8578; }
    .step-wizard .label.active { color: #D97757; font-weight: 700; }
    .step-wizard .line { flex: 1; height: 1px; background: #E5E1D6; min-width: 20px; margin: 0 10px; }

    /* Outlined buttons instead of solid fills, to match the mockup */
    div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {
        border-radius: 8px;
        font-weight: 600;
        border: 1.5px solid #D97757;
        background: transparent;
        color: #B85C38;
        transition: background 0.15s ease, color 0.15s ease;
    }
    div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {
        background: #D97757;
        color: #FFFFFF;
        border-color: #D97757;
    }
    div[data-testid="stButton"] > button[kind="secondary"] {
        border-color: #D8D4C8;
        color: #3D3929;
    }
    div[data-testid="stButton"] > button[kind="secondary"]:hover {
        background: #F0EEE5;
        color: #3D3929;
        border-color: #D8D4C8;
    }

    /* Stat rows: hairline rule + serif number, closer to the mockup's RESULT panel */
    div[data-testid="stMetric"] {
        background: transparent;
        border-top: 1px solid #E5E1D6;
        border-radius: 0;
        padding: 10px 4px 0 4px;
    }
    div[data-testid="stMetricLabel"] p {
        font-family: 'Inter', sans-serif !important;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        font-size: 0.68rem !important;
        color: #8A8578 !important;
    }
    div[data-testid="stMetricValue"] {
        font-family: 'Source Serif 4', Georgia, serif !important;
        color: #3D3929 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def eyebrow(text: str):
    st.markdown(f'<div class="eyebrow">{text}</div>', unsafe_allow_html=True)


def badge_pill(text: str):
    st.markdown(f'<span class="badge-pill">{text}</span>', unsafe_allow_html=True)


def step_indicator(labels, current_index: int):
    """Numbered circle-and-line step tracker, e.g. labels=['Download','Calculate','Result']."""
    parts = ['<div class="step-wizard">']
    for i, label in enumerate(labels):
        state = "done" if i < current_index else ("active" if i == current_index else "")
        label_state = "active" if i == current_index else ""
        parts.append(
            f'<div class="step"><div class="circle {state}">{i + 1}</div>'
            f'<div class="label {label_state}">{label}</div></div>'
        )
        if i < len(labels) - 1:
            parts.append('<div class="line"></div>')
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)

if "lang" not in st.session_state:
    st.session_state.lang = "TR"
if "tier" not in st.session_state:
    st.session_state.tier = None  # None -> show tier picker; "quick" or "full"

with st.sidebar:
    lang_choice = st.radio("🌐 Dil / Language", ["Türkçe", "English"],
                           index=0 if st.session_state.lang == "TR" else 1,
                           horizontal=True)
    st.session_state.lang = "TR" if lang_choice == "Türkçe" else "EN"
    st.divider()
    if st.button("🔄 " + ("Sıfırla" if st.session_state.lang == "TR" else "Reset"), use_container_width=True):
        for k in ["calc_result", "calc_beyan", "calc_detail", "calc_pdf",
                  "quick_result", "tier", "quick_step", "full_step"]:
            st.session_state.pop(k, None)
        st.rerun()

T = TR if st.session_state.lang == "TR" else EN

# ─────────────────────────────────────────────────────────────────────────────
# Hero
# ─────────────────────────────────────────────────────────────────────────────
st.title(T["hero"])
st.markdown(T["hero_sub"])
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# TIER PICKER  — shown until the user chooses Quick Check or Full Calculation
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.tier is None:
    st.subheader(T["tier_head"])
    c1, c2 = st.columns(2)

    with c1:
        with st.container(border=True):
            badge_pill(T["tier_free_badge"])
            st.markdown(f"### {T['tier_free_title']}")
            st.write(T["tier_free_desc"])
            if st.button(T["tier_free_cta"], use_container_width=True, type="secondary"):
                st.session_state.tier = "quick"
                st.rerun()

    with c2:
        with st.container(border=True):
            if PAID_MODE:
                badge_pill(T["tier_full_badge_priced"].format(price=f"₺{CALC_PRICE_TRY/100:.0f}"))
            else:
                badge_pill(T["tier_full_badge_free"])
            st.markdown(f"### {T['tier_full_title']}")
            st.write(T["tier_full_desc"])
            if st.button(T["tier_full_cta"], use_container_width=True, type="primary"):
                st.session_state.tier = "full"
                st.rerun()

    st.divider()
    with st.expander(T["disclaimer_head"], expanded=False):
        st.markdown(T["disclaimer"])
    st.stop()

# Back-to-picker link, shown on both tiers once chosen
if st.button(T["back_to_tiers"]):
    st.session_state.tier = None
    st.rerun()
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# QUICK CHECK TIER  (free — stocks/ETFs only, no options, no Yİ-ÜFE)
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.tier == "quick":
    st.header(T["quick_title"])

    if "quick_step" not in st.session_state:
        st.session_state.quick_step = 0

    _quick_step_labels = ["Şablon", "Hesapla"] if T is TR else ["Template", "Calculate"]
    step_indicator(_quick_step_labels, st.session_state.quick_step)

    # ── STEP 1 OF 2 — download the blank template ───────────────────────────
    if st.session_state.quick_step == 0:
        eyebrow(("ADIM 1 / 2" if T is TR else "STEP 1 OF 2"))
        st.subheader(T["quick_step1_head"])
        st.markdown(T["quick_step1_body"])
        quick_blank = get_quick_template_bytes(lang=st.session_state.lang)
        st.download_button(
            T["quick_step1_btn"],
            data=quick_blank,
            file_name="Quick_Check_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.write("")
        if st.button(("Devam →" if T is TR else "Continue →"), type="primary"):
            st.session_state.quick_step = 1
            st.rerun()

    # ── STEP 2 OF 2 — upload, calculate, and see the result ─────────────────
    elif st.session_state.quick_step == 1:
        if st.button(("‹ Şablon adımına dön" if T is TR else "‹ Back to template")):
            st.session_state.quick_step = 0
            st.rerun()

        eyebrow(("ADIM 2 / 2" if T is TR else "STEP 2 OF 2"))
        st.subheader(T["quick_step2_head"])
        st.markdown(T["quick_step2_body"])
        quick_up = st.file_uploader(T["quick_upload_label"], type=["xlsx"], key="quick_uploader")
        quick_go = st.button(T["quick_calc_btn"], type="primary", disabled=(quick_up is None))

        if quick_go and quick_up is not None:
            with st.spinner(T["quick_spinner"]):
                try:
                    st.session_state["quick_result"] = tr_tax_core.calculate_turkish_taxes(
                        quick_up, lang=st.session_state.lang
                    )
                except Exception as e:
                    st.error(f"{T['err']}{e}")
                    st.stop()

        if "quick_result" in st.session_state:
            qr = st.session_state["quick_result"]
            st.divider()
            st.success(T["quick_success"])
            badge_pill(f"{T['status_label']}: {qr['status']}")

            m1, m2, m3 = st.columns(3)
            m1.metric(T["quick_gain"], tl(qr["total_gains"]))
            m2.metric(T["quick_tax"], tl(qr["estimated_tax"]))
            m3.metric(T["quick_instal"],
                      f"{tl(qr['instalment_1'])} + {tl(qr['instalment_2'])}")
            st.caption(f"{T['quick_trades']}: {qr['trades_processed']}")

            for w in qr.get("warnings", []):
                st.warning(w)

            st.divider()
            with st.container(border=True):
                st.markdown(f"**{T['quick_upgrade_head']}**")
                st.write(T["quick_upgrade_body"])
                if st.button(T["quick_upgrade_btn"], type="primary"):
                    st.session_state.tier = "full"
                    st.session_state.full_step = 0
                    for k in ["quick_result"]:
                        st.session_state.pop(k, None)
                    st.rerun()

    st.divider()
    with st.expander(T["disclaimer_head"], expanded=False):
        st.markdown(T["disclaimer"])

# ─────────────────────────────────────────────────────────────────────────────
# FULL CALCULATION TIER  (paid once STRIPE_SECRET_KEY/CALC_PRICE_TRY are set)
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.tier == "full":
    st.header(T["tier_full_title"])

    if "full_step" not in st.session_state:
        st.session_state.full_step = 0

    _full_step_labels = ["Şablon", "Hesapla"] if T is TR else ["Template", "Calculate"]
    step_indicator(_full_step_labels, st.session_state.full_step)

    # ── STEP 1 OF 2 — download the blank workbook ───────────────────────────
    if st.session_state.full_step == 0:
        eyebrow(("ADIM 1 / 2" if T is TR else "STEP 1 OF 2"))
        st.subheader(T["step1_head"])
        st.markdown(T["step1_body"])
        blank = get_blank_workbook_bytes(lang=st.session_state.lang)
        st.download_button(
            T["step1_btn"],
            data=blank,
            file_name="Turkey_Tax_Tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.write("")
        if st.button(("Devam →" if T is TR else "Continue →"), type="primary"):
            st.session_state.full_step = 1
            st.rerun()

    # ── STEP 2 OF 2 — upload, calculate, and see the result ─────────────────
    elif st.session_state.full_step == 1:
        if st.button(("‹ Şablon adımına dön" if T is TR else "‹ Back to template")):
            st.session_state.full_step = 0
            st.rerun()

        eyebrow(("ADIM 2 / 2" if T is TR else "STEP 2 OF 2"))
        st.subheader(T["step2_head"])
        st.markdown(T["step2_body"])

        up = st.file_uploader(T["upload_label"], type=["xlsx"], key="full_uploader")

        # EVDS key is server-side only — read from Streamlit secrets or env, never shown in UI
        try:
            evds_key = st.secrets.get("EVDS_KEY", "") or os.getenv("EVDS_KEY", "")
        except Exception:
            evds_key = os.getenv("EVDS_KEY", "")

        # Paid mode: show payment notice before the button
        if PAID_MODE:
            st.info(f"{T['payment_head']}: {T['payment_body']} — ₺{CALC_PRICE_TRY/100:.0f}")

        calc_label = T["calc_btn_paid"] if PAID_MODE else T["calc_btn_free"]
        go = st.button(calc_label, type="primary", disabled=(up is None))

        # ── CALCULATION  — runs on button click, results persist in session_state
        if go and up is not None:
            # ── Payment gate (skipped when PAID_MODE is False) ──────────────
            if PAID_MODE:
                try:
                    client_secret = create_stripe_checkout(CALC_PRICE_TRY)
                    st.session_state["stripe_intent"] = client_secret
                    st.warning(
                        "⚙️ **Stripe integration point.** In production, redirect the user "
                        "to `stripe.com/checkout` here. Set `STRIPE_SECRET_KEY` and "
                        "`CALC_PRICE_TRY` env vars and plug in a success webhook to "
                        "proceed automatically after payment. Skipping payment for now."
                    )
                except Exception as e:
                    st.error(f"Stripe error: {e}")
                    st.stop()

            with st.spinner(T["spinner"]):
                try:
                    results = tr_tax_core.calculate_from_workbook(
                        up, evds_key or None, lang=st.session_state.lang
                    )
                    beyan = tr_tax_core.build_gib_beyan(results)
                    detail = tr_tax_core.detailed_transactions(results)
                    pdf = tr_tax_report.build_beyan_pdf(beyan, detail)
                    # Store everything so downloads survive page reruns
                    st.session_state["calc_result"] = results
                    st.session_state["calc_beyan"]  = beyan
                    st.session_state["calc_detail"] = detail
                    st.session_state["calc_pdf"]    = pdf
                except Exception as e:
                    st.error(f"{T['err']}{e}")
                    st.stop()

        # ── RESULTS  — rendered from session_state so downloads never vanish ─
        if "calc_result" in st.session_state:
            results = st.session_state["calc_result"]
            beyan   = st.session_state["calc_beyan"]
            detail  = st.session_state["calc_detail"]
            pdf     = st.session_state["calc_pdf"]

            st.divider()

            # ── Status bar ────────────────────────────────────────────────
            # Matched on status_code (language-neutral), not the display text
            # in results["status"] - that text is now localized (TR/EN) so
            # matching on it directly would silently fall through to the
            # generic icon whenever the app is running in Turkish.
            badge = {"FINAL": "✅", "PROVISIONAL": "⚠️", "INCOMPLETE": "⛔"}
            b = badge.get(results.get("status_code"), "ℹ️")
            st.success(T["success"])
            badge_pill(
                f"{b} {T['status_label']}: {results['status']}  ·  "
                f"{T['ufe_label']}: {results['ufe_source']}"
            )

            # ── Key metrics ───────────────────────────────────────────────
            m1, m2, m3 = st.columns(3)
            m1.metric(T["matrah"], tl(beyan["tax_base"]))
            m2.metric(T["tax"], tl(beyan["tax"]))
            m3.metric(T["instal"],
                      f"{tl(beyan['instalment_1'])} + {tl(beyan['instalment_2'])}")

            # ── Two-tab results ──────────────────────────────────────────────
            tab1, tab2 = st.tabs([T["tab_beyan"], T["tab_detail"]])

            with tab1:
                # 3.D capital gains
                cg = beyan["capital_gains"]
                st.markdown(f"**{T['3d_head']}**")
                st.table(pd.DataFrame([{
                    T["kaz_turu"]:         f"{cg['code']} — {cg['label']}",
                    T["gayrisafi_tutar"]:  tl(cg["gayrisafi"]),
                    T["gider_indirim"]:    tl(cg["gider_indirim"]),
                    T["safi_kaz"]:         tl(cg["safi"]),
                    T["kesilen"]:          tl(cg["kesilen"]),
                }]))

                # Per-instrument breakdown
                if results["lines"]:
                    st.markdown(f"**{'Enstrüman bazında' if T is TR else 'Per instrument'}**")
                    rows = [{
                        T["col_inst"]: ln["name"],
                        "Tür" if T is TR else "Type": ln["kind"],
                        "Döviz" if T is TR else "Ccy": ln["currency"],
                        T["col_gross"]: round(ln["gross"], 2),
                        T["col_taxable"]: round(ln["taxable_raw"], 2),
                        "Flag": ln["flag"],
                    } for ln in results["lines"]]
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            with tab2:
                if detail:
                    tx_rows = []
                    for t in detail:
                        d = t.get("date")
                        tx_rows.append({
                            T["col_inst"]:    t.get("asset", ""),
                            T["col_date"]:    d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d),
                            T["col_action"]:  t.get("type", ""),
                            T["col_qty"]:     t.get("qty", ""),
                            T["col_price"]:   t.get("price", ""),
                            T["col_tl"]:      round(t["tl_amount"], 2) if t.get("tl_amount") is not None else None,
                            T["col_gross"]:   round(t["gross"], 2) if t.get("gross") is not None else None,
                            T["col_taxable"]: round(t["taxable"], 2) if t.get("taxable") is not None else None,
                            T["col_status"]:  t.get("status", ""),
                        })
                    st.dataframe(pd.DataFrame(tx_rows), use_container_width=True, hide_index=True)
                else:
                    st.info(T["no_trades"])

            # ── Warnings ─────────────────────────────────────────────────────
            for w in results["warnings"]:
                st.warning(w)

            # ── Downloads — always visible once calculated ────────────────────
            st.divider()
            dc1, dc2 = st.columns(2)
            dc1.download_button(
                T["dl_pdf"], data=pdf,
                file_name="Hazir_Beyan_Ozeti.pdf",
                mime="application/pdf",
            )
            dc2.download_button(
                T["dl_xlsx"],
                data=results["workbook_bytes"],
                file_name="Turkey_Tax_Tracker_processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.divider()
    with st.expander(T["disclaimer_head"], expanded=False):
        st.markdown(T["disclaimer"])
