"""
tr_tax_dashboard.py
Turkey Foreign Securities Tax Calculator — web front-end

FREE now. Stripe-ready: set STRIPE_SECRET_KEY and CALC_PRICE_TRY env vars to
enable paid mode with zero code changes.
"""

import io
import os

import streamlit as st
import pandas as pd

import tr_tax_core
import tr_tax_report

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG  (all overridable via environment variables)
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
        "Yurt dışı hisse, ETF ve opsiyon işlemleriniz için 2026 Türkiye gelir vergisi "
        "tahmini. Sonuç **GİB Hazır Beyan** formatında üretilir."
    ),
    "lang_label": "🌐 Dil / Language",

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
        "Results are produced in **GİB Hazır Beyan** (Turkish tax return) format."
    ),
    "lang_label": "🌐 Dil / Language",

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
def get_blank_workbook_bytes():
    """Build a fresh blank workbook in memory and return its bytes."""
    import tax_tool as tt
    import tempfile, os
    tmp = tempfile.mktemp(suffix=".xlsx")
    orig = tt.FILENAME
    tt.FILENAME = tmp
    try:
        tt.cmd_init([])
    finally:
        tt.FILENAME = orig
    with open(tmp, "rb") as f:
        data = f.read()
    os.unlink(tmp)
    return data


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

if "lang" not in st.session_state:
    st.session_state.lang = "TR"

with st.sidebar:
    lang_choice = st.radio("🌐 Dil / Language", ["Türkçe", "English"],
                           index=0 if st.session_state.lang == "TR" else 1,
                           horizontal=True)
    st.session_state.lang = "TR" if lang_choice == "Türkçe" else "EN"

T = TR if st.session_state.lang == "TR" else EN

# ─────────────────────────────────────────────────────────────────────────────
# Hero
# ─────────────────────────────────────────────────────────────────────────────
st.title(T["hero"])
st.markdown(T["hero_sub"])
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1  — download blank workbook
# ─────────────────────────────────────────────────────────────────────────────
st.subheader(T["step1_head"])
st.markdown(T["step1_body"])
blank = get_blank_workbook_bytes()
st.download_button(
    T["step1_btn"],
    data=blank,
    file_name="Turkey_Tax_Tracker.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2  — upload + calculate
# ─────────────────────────────────────────────────────────────────────────────
st.subheader(T["step2_head"])
st.markdown(T["step2_body"])

col_up, col_evds = st.columns([2, 1])
with col_up:
    up = st.file_uploader(T["upload_label"], type=["xlsx"])
with col_evds:
    evds_key = st.text_input(T["evds_label"], help=T["evds_help"], type="password")

st.caption(T["div_head"])
st.caption(T["div_cap"])
div_default = pd.DataFrame(
    [{"date": "", "currency": "USD", "gross": 0.0, "withholding": 0.0, "expense": 0.0}]
)
div_table = st.data_editor(
    div_default, num_rows="dynamic", use_container_width=True,
    column_config={
        "date":        st.column_config.TextColumn(T["div_date"]),
        "currency":    st.column_config.SelectboxColumn(T["div_ccy"], options=["USD", "EUR"]),
        "gross":       st.column_config.NumberColumn(T["div_gross"]),
        "withholding": st.column_config.NumberColumn(T["div_wh"]),
        "expense":     st.column_config.NumberColumn(T["div_exp"]),
    },
    key="divs",
)

# Paid mode: show payment notice before the button
if PAID_MODE:
    st.info(f"{T['payment_head']}: {T['payment_body']} — ₺{CALC_PRICE_TRY/100:.0f}")

calc_label = T["calc_btn_paid"] if PAID_MODE else T["calc_btn_free"]
go = st.button(calc_label, type="primary", disabled=(up is None))

# ─────────────────────────────────────────────────────────────────────────────
# CALCULATION
# ─────────────────────────────────────────────────────────────────────────────
if go and up is not None:

    # ── Payment gate (skipped when PAID_MODE is False) ──────────────────────
    if PAID_MODE:
        try:
            client_secret = create_stripe_checkout(CALC_PRICE_TRY)
            # In production, redirect user to a hosted Stripe Checkout page or
            # embed the Payment Element. For now, store the intent and note the
            # integration point.
            st.session_state["stripe_intent"] = client_secret
            st.warning(
                "⚙️ **Stripe integration point.** In production, redirect the user "
                "to `stripe.com/checkout` here. Set `STRIPE_SECRET_KEY` and "
                "`CALC_PRICE_TRY` env vars and plug in a success webhook to "
                "proceed automatically after payment. Skipping payment for now."
            )
            # IMPORTANT: in production, do NOT continue here — wait for
            # webhook confirmation. For development, we fall through.
        except Exception as e:
            st.error(f"Stripe error: {e}")
            st.stop()

    # ── Run the full engine ──────────────────────────────────────────────────
    with st.spinner(T["spinner"]):
        try:
            results = tr_tax_core.calculate_from_workbook(up, evds_key or None)

            # Dividends
            div_entries = []
            for _, row in div_table.iterrows():
                if str(row.get("date") or "").strip() and float(row.get("gross") or 0) > 0:
                    div_entries.append({
                        "date": row["date"], "currency": row["currency"],
                        "gross": row["gross"], "withholding": row["withholding"],
                        "expense": row["expense"],
                    })
            divs = tr_tax_core.compute_dividends(div_entries) if div_entries else None
            beyan = tr_tax_core.build_gib_beyan(results, dividends=divs)
            detail = tr_tax_core.detailed_transactions(results)

        except Exception as e:
            st.error(f"{T['err']}{e}")
            st.stop()

    # ── Status bar ──────────────────────────────────────────────────────────
    badge = {
        "FINAL from available data.": "✅",
        "PROVISIONAL - missing YI-UFE affected at least one realised sale.": "⚠️",
        "INCOMPLETE - do not file until the flagged rows are fixed.": "⛔",
    }
    b = badge.get(results["status"], "ℹ️")
    st.success(T["success"])
    st.caption(
        f"{b} {T['status_label']}: {results['status']}  ·  "
        f"{T['ufe_label']}: {results['ufe_source']}"
    )

    # ── Key metrics ─────────────────────────────────────────────────────────
    m1, m2, m3 = st.columns(3)
    m1.metric(T["matrah"], tl(beyan["tax_base"]))
    m2.metric(T["tax"], tl(beyan["tax"]))
    m3.metric(T["instal"],
              f"{tl(beyan['instalment_1'])} + {tl(beyan['instalment_2'])}")

    # ── Two-tab results ──────────────────────────────────────────────────────
    tab1, tab2 = st.tabs([T["tab_beyan"], T["tab_detail"]])

    with tab1:
        # 3.C dividends
        if beyan["dividends"]:
            dv = beyan["dividends"]
            st.markdown(f"**{T['3c_head']}**")
            st.table(pd.DataFrame([{
                T["irat_turu"]:      dv["code"],
                T["gayrisafi_irat"]: tl(dv["gayrisafi"]),
                T["indirilebilir"]:  tl(dv["indirilecek"]),
                T["safi_irat"]:      tl(dv["safi"]),
                T["kesilen"]:        tl(dv["kesilen"]),
            }]))
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
            rows = []
            for t in detail:
                d = t.get("date")
                rows.append({
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
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info(T["no_trades"])

    # ── Warnings ─────────────────────────────────────────────────────────────
    for w in results["warnings"]:
        st.warning(w)

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.divider()
    dc1, dc2 = st.columns(2)
    pdf = tr_tax_report.build_beyan_pdf(beyan, detail)
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

# ─────────────────────────────────────────────────────────────────────────────
# Disclaimer  (always visible)
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
with st.expander(T["disclaimer_head"], expanded=False):
    st.markdown(T["disclaimer"])
