import io
import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.set_page_config(page_title="Total Sales Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
h1, h2, h3 {direction: rtl; text-align: right;}
div[data-testid="stMetric"] {border: 1px solid #d9e2ec; padding: 14px; border-radius: 12px;}
</style>
""", unsafe_allow_html=True)

st.title("📊 لوحة تحليل المبيعات — Total")
st.caption("ارفع ملف Excel يحتوي على الشيتات المطلوبة لعرض البيانات والتحقق الأولي منها.")

expected_sheets = ["Total", "Target", "قصادير", "Sales Report", "Data", "Total price", "Statement"]
uploaded = st.file_uploader("ارفع ملف Excel (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("ابدأ برفع ملف Excel. يمكنك استخدام قالب البداية من المشروع: Total_Sales_Template.xlsx")
    st.markdown("### الشيتات المتوقعة")
    st.write(" • ".join(expected_sheets))
    st.stop()

try:
    excel = pd.ExcelFile(uploaded)
    available = excel.sheet_names
    missing = [s for s in expected_sheets if s not in available]
    st.success(f"تم قراءة الملف. عدد الشيتات: {len(available)}")
    if missing:
        st.warning("الشيتات التالية غير موجودة أو أسماؤها مختلفة: " + "، ".join(missing))
    selected = st.sidebar.selectbox("اختر الشيت", available)
    df = pd.read_excel(uploaded, sheet_name=selected, header=1 if selected != "دليل الاستخدام" else 1)
    st.subheader(f"عرض شيت: {selected}")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        "تنزيل بيانات الشيت بصيغة CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{selected.replace(' ', '_')}.csv",
        mime="text/csv"
    )

    st.divider()
    st.subheader("ملخص تلقائي مبدئي")
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        col_count = min(4, len(numeric_cols))
        metric_cols = st.columns(col_count)
        for i, col in enumerate(numeric_cols[:col_count]):
            metric_cols[i].metric(f"مجموع {col}", f"{df[col].sum(numeric_only=True):,.2f}")
    else:
        st.info("لا توجد أعمدة رقمية واضحة في هذا الشيت لحساب ملخص تلقائي.")

    if selected in ["Total", "Sales Report", "Data"] and not df.empty:
        st.subheader("رسم بياني سريع")
        candidate = [c for c in numeric_cols if c.lower() in ["actual", "total_actual", "value", "total", "total_actual", "total_price"]]
        if candidate:
            st.bar_chart(df[candidate[0]].fillna(0))
        else:
            st.caption("لم يتم العثور على عمود رقمي باسم متوقع للرسم؛ يمكن إضافة تحليل مخصص بعد معرفة بنية البيانات الفعلية.")
except Exception as e:
    st.error("تعذر قراءة الملف أو أحد الشيتات. تأكد أن الملف بصيغة .xlsx وأن رؤوس الأعمدة في الصف الثاني.")
    st.exception(e)
