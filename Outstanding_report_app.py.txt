import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime


# ---------- הגדרות עמוד + סטייל ----------

st.set_page_config(
    page_title="Altshare - Option Report Tool",
    page_icon="💼",
    layout="wide",
)

CUSTOM_CSS = """
<style>
/* רקע כללי */
.stApp {
    background-color: #f5f7fb;
    font-family: "Segoe UI", sans-serif;
}

/* כותרת עליונה */
.app-header {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 8px;
}
.app-title {
    font-size: 28px;
    font-weight: 700;
    color: #1b2745;
}
.app-subtitle {
    font-size: 14px;
    color: #6b7280;
}

/* כרטיסי מדדים */
.metric-container {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 16px;
    margin-top: 12px;
}
.metric-card {
    background: linear-gradient(135deg, #0f766e, #22c55e);
    border-radius: 14px;
    padding: 14px 16px;
    color: white;
    box-shadow: 0 8px 18px rgba(15, 118, 110, 0.25);
}
.metric-title {
    font-size: 13px;
    font-weight: 500;
    opacity: 0.9;
}
.metric-value {
    margin-top: 6px;
    font-size: 20px;
    font-weight: 700;
}

/* כפתור הורדה */
.download-container {
    margin-top: 20px;
}

/* כרטיס קלט */
.input-card {
    background: white;
    border-radius: 16px;
    padding: 16px 18px;
    box-shadow: 0 4px 14px rgba(148, 163, 184, 0.3);
}
.section-title {
    font-size: 18px;
    font-weight: 600;
    color: #111827;
    margin-bottom: 8px;
}
.section-caption {
    font-size: 13px;
    color: #6b7280;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------- כותרת עם לוגו Altshare ----------

header_cols = st.columns([1, 3])
with header_cols[0]:
    # שים קובץ לוגו בשם altshare_logo.png באותה תיקייה
    st.image("altshare_logo.png", width=160)
with header_cols[1]:
    st.markdown(
        """
        <div class="app-header">
            <div>
                <div class="app-title">Altshare – Option Report Tool</div>
                <div class="app-subtitle">
                    חישוב אוטומטי של Weighted Averages ו-Intrinsic Value לפי קובץ grants
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------- פונקציות חישוב ----------

def calc_remaining_life(row, report_date: datetime) -> float:
    """משחזר את נוסחת העמודה O באקסל."""
    emp_term = row.get("Employment Termination Date")
    orig_exp = row.get("Original Expiry Date")
    upd_exp = row.get("Updated Expiry Date")

    # IF(AND(EmpTermDate>$B$1,OriginalExpDate>$B$1), YEARFRAC(report_date, OriginalExpDate),
    #    IF(UpdatedExpDate>$B$1, YEARFRAC(report_date, UpdatedExpDate), 0))
    if pd.notnull(emp_term) and pd.notnull(orig_exp) and emp_term > report_date and orig_exp > report_date:
        return (orig_exp - report_date).days / 365.0

    if pd.notnull(upd_exp) and upd_exp > report_date:
        return (upd_exp - report_date).days / 365.0

    return 0.0


def convert_exercise_price(price, currency, fx_rates):
    """
    המרה לדולר רק לעמודת Exercise Price לפי דרישות:
    - אם המטבע ריק או USD -> לא ממירים
    - EUR / GBP -> ממירים לפי שער שהוזן (אם קיים)
    - ILS מעל 0.1 -> ממירים לפי שער שהוזן
    - מטבע לא מוכר -> לא שואלים שער, מניחים דולר (לא ממירים)
    """
    if pd.isna(price):
        return price

    curr = str(currency).strip().upper() if isinstance(currency, str) else ""

    if curr in ("", "USD"):
        return price

    if curr == "EUR":
        rate = fx_rates.get("EUR")
        if rate:
            return price * rate
        return price

    if curr == "GBP":
        rate = fx_rates.get("GBP")
        if rate:
            return price * rate
        return price

    if curr in ("ILS", "NIS", "₪"):
        if price <= 0.1:
            return price
        rate = fx_rates.get("ILS")
        if rate:
            return price * rate
        return price

    # מטבע אחר – לא נוגעים (נחשב כאילו כבר בדולר)
    return price


def compute_metrics(df: pd.DataFrame, report_date: datetime, closing_price: float, fx_rates):
    # המרת עמודות תאריך
    for col in ["Employment Termination Date", "Original Expiry Date", "Updated Expiry Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # עמודת Exercise Price ל-USD
    if "Exercise Price" not in df.columns:
        raise ValueError("חסרה עמודה 'Exercise Price' בקובץ")

    curr_col = "Exercise Price Currency" if "Exercise Price Currency" in df.columns else None

    df["Exercise Price (USD)"] = df.apply(
        lambda row: convert_exercise_price(
            row["Exercise Price"],
            row[curr_col] if curr_col else "",
            fx_rates,
        ),
        axis=1,
    )

    # O – Remaining Contractual Life
    df["Remaining Life (O)"] = df.apply(lambda row: calc_remaining_life(row, report_date), axis=1)

    # X – Intrinsic per option
    df["Intrinsic (X)"] = df["Exercise Price (USD)"].apply(
        lambda ex: max(closing_price - ex, 0) if pd.notnull(ex) else 0
    )

    # עמודות כמות
    AE = df.get("Outstanding", pd.Series(dtype=float)).fillna(0)
    AH = df.get("Exercisable", pd.Series(dtype=float)).fillna(0)
    O = df["Remaining Life (O)"].fillna(0)
    X = df["Intrinsic (X)"].fillna(0)

    AE_sum = AE.sum()
    AH_sum = AH.sum()

    # Weighted averages – לפי נוסחאות SUMPRODUCT מהאקסל
    results = {}

    # Weighted Average Exercise Price
    if AE_sum > 0:
        results["Weighted Average Exercise Price - Outstanding"] = float(
            (AE * df["Exercise Price (USD)"]).sum() / AE_sum
        )
    else:
        results["Weighted Average Exercise Price - Outstanding"] = None

    if AH_sum > 0:
        results["Weighted Average Exercise Price - Exercisable"] = float(
            (AH * df["Exercise Price (USD)"]).sum() / AH_sum
        )
    else:
        results["Weighted Average Exercise Price - Exercisable"] = None

    # Weighted Average Remaining Contractual Life
    if AE_sum > 0:
        results["Weighted Average Remaining Contractual Life - Outstanding"] = float(
            (AE * O).sum() / AE_sum
        )
    else:
        results["Weighted Average Remaining Contractual Life - Outstanding"] = None

    if AH_sum > 0:
        results["Weighted Average Remaining Contractual Life - Exercisable"] = float(
            (AH * O).sum() / AH_sum
        )
    else:
        results["Weighted Average Remaining Contractual Life - Exercisable"] = None

    # Aggregate Intrinsic Value
    results["Aggregate Intrinsic Value - Outstanding"] = float((AE * X).sum())
    results["Aggregate Intrinsic Value - Exercisable"] = float((AH * X).sum())

    # טבלת Summary
    summary_df = (
        pd.DataFrame(
            [
                {"Metric": k, "Value": v}
                for k, v in results.items()
            ]
        )
        .set_index("Metric")
    )

    return df, results, summary_df


def to_excel_bytes(df_calc: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_calc.to_excel(writer, sheet_name="Data & Calculations", index=False)
        summary_df.reset_index().to_excel(writer, sheet_name="Summary", index=False)
    buffer.seek(0)
    return buffer.getvalue()


# ---------- UI – אזור קלט ----------

left, right = st.columns([1.1, 1.9])

with left:
    st.markdown('<div class="input-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">קלט</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">בחר קובץ grants, תאריך דוח ומחיר סגירה.</div>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("קובץ grants (Excel)", type=["xlsx", "xls"])

    report_date = st.date_input("תאריך הדוח")
    closing_price = st.number_input("מחיר סגירה (USD)", min_value=0.0, step=0.01)

    fx_rates = {}

    # נזהה אילו מטבעות בכלל קיימים בקובץ (אם הועלה)
    detected_currencies = set()
    if uploaded_file is not None:
        tmp_df = pd.read_excel(uploaded_file)
        if "Exercise Price Currency" in tmp_df.columns:
            detected_currencies = (
                tmp_df["Exercise Price Currency"]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
                .unique()
            )
        # צריך להחזיר את ה-pointer להתחלה לפני קריאה נוספת
        uploaded_file.seek(0)

    if "EUR" in detected_currencies:
        fx_rates["EUR"] = st.number_input(
            "שער EUR→USD (כמה דולר עבור 1€)",
            min_value=0.0,
            step=0.0001,
            format="%.4f",
        )

    if "GBP" in detected_currencies:
        fx_rates["GBP"] = st.number_input(
            "שער GBP→USD (כמה דולר עבור £1)",
            min_value=0.0,
            step=0.0001,
            format="%.4f",
        )

    if any(c in detected_currencies for c in ["ILS", "NIS", "₪"]):
        fx_rates["ILS"] = st.number_input(
            "שער ILS→USD (כמה דולר עבור ₪1)",
            min_value=0.0,
            step=0.0001,
            format="%.4f",
        )

    run_calc = st.button("🚀 בצע חישוב")

    st.markdown("</div>", unsafe_allow_html=True)


# ---------- UI – אזור תוצאות ----------

with right:
    if run_calc:
        if uploaded_file is None:
            st.error("אנא העלה קובץ Excel.")
        elif closing_price <= 0:
            st.error("אנא הזן מחיר סגירה גדול מ-0.")
        else:
            try:
                # קריאה מחדש של הקובץ (אחרי ה-seek)
                df_input = pd.read_excel(uploaded_file)

                # המרה לתאריך
                report_dt = datetime.combine(report_date, datetime.min.time())

                df_calc, results, summary_df = compute_metrics(
                    df_input.copy(),
                    report_dt,
                    closing_price,
                    fx_rates,
                )

                # כרטיסי מדדים
                st.markdown(
                    '<div class="section-title">תוצאות חישוב</div>',
                    unsafe_allow_html=True,
                )
                st.markdown(
                    '<div class="section-caption">המדדים מחושבים בהתאם לנוסחאות האקסל (O, X, AN).</div>',
                    unsafe_allow_html=True,
                )

                st.markdown('<div class="metric-container">', unsafe_allow_html=True)
                for metric_name, value in results.items():
                    if value is None:
                        txt_val = "N/A"
                    elif "Intrinsic" in metric_name:
                        txt_val = f"{value:,.2f}"
                    else:
                        txt_val = f"{value:,.6f}"

                    card_html = f"""
                    <div class="metric-card">
                        <div class="metric-title">{metric_name}</div>
                        <div class="metric-value">{txt_val}</div>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

                # טבלת summary
                st.write("")
                st.dataframe(
                    summary_df.style.format("{:,.6f}"),
                    use_container_width=True,
                )

                # הורדת קובץ אקסל
                excel_bytes = to_excel_bytes(df_calc, summary_df)

                st.markdown('<div class="download-container">', unsafe_allow_html=True)
                st.download_button(
                    label="⬇️ הורד קובץ Excel עם חישובים (Data & Summary)",
                    data=excel_bytes,
                    file_name="option_report_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.markdown("</div>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"שגיאה בחישוב: {e}")
    else:
        st.info("העלה קובץ, הזן נתונים ולחץ על ״בצע חישוב״ כדי לראות תוצאות.")
