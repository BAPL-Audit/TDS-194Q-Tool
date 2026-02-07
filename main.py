import pandas as pd
import streamlit as st

from utils import (
    compute_194q,
    export_to_excel_openpyxl,
    read_26q,
    read_books,
    read_purchase,
)

st.set_page_config(page_title="TDS 194Q Reconciliation", layout="wide")
st.title("📊 TDS 194Q Reconciliation Tool")
st.markdown("Upload required Excel files below:")

purchase_file = st.file_uploader("📥 Purchase Register (GST exclusive)", type=["xlsx"])
books_file = st.file_uploader("📥 Books TDS Ledger (194Q)", type=["xlsx"])
tds_file = st.file_uploader("📥 Deduction Register (26Q)", type=["xlsx"])

if purchase_file and books_file and tds_file:
    try:
        purchase_raw = pd.read_excel(purchase_file)
        books_raw = pd.read_excel(books_file)
        tds_raw = pd.read_excel(tds_file)

        purchase = read_purchase(purchase_raw)
        books = read_books(books_raw)
        tds26q = read_26q(tds_raw)

        purchase_summary = compute_194q(purchase)

        books_sum = books.groupby("party_norm", as_index=False).agg(
            tds_books=("tds_books", "sum")
        )

        tds26q_sum = tds26q.groupby("party_norm", as_index=False).agg(
            gross_26q=("gross_26q", "sum"),
            tds_26q=("tds_26q", "sum"),
        )

        final = (
            purchase_summary.merge(books_sum, on="party_norm", how="left")
            .merge(tds26q_sum, on="party_norm", how="left")
            .fillna(0)
        )

        final["short_books"] = final["tds_books"] - final["tds_required"]
        final["short_26q"] = final["tds_26q"] - final["tds_required"]

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            [
                "📘 Purchase Summary",
                "🧮 194Q Computation",
                "📒 Books vs Required",
                "🧾 26Q vs Required",
                "🚨 Final Reconciliation",
            ]
        )

        with tab1:
            st.dataframe(purchase_summary, use_container_width=True)

        with tab2:
            st.dataframe(
                purchase_summary[purchase_summary["liable_amount"] > 0],
                use_container_width=True,
            )

        with tab3:
            st.dataframe(
                final[["party_norm", "tds_required", "tds_books", "short_books"]],
                use_container_width=True,
            )

        with tab4:
            st.dataframe(
                final[["party_norm", "tds_required", "tds_26q", "short_26q"]],
                use_container_width=True,
            )

        with tab5:
            st.dataframe(final, use_container_width=True)

        excel_file = export_to_excel_openpyxl(
            {
                "Purchase Summary": purchase_summary,
                "194Q Computation": purchase_summary[
                    purchase_summary["liable_amount"] > 0
                ],
                "Books vs Required": final[
                    ["party_norm", "tds_required", "tds_books", "short_books"]
                ],
                "26Q vs Required": final[
                    ["party_norm", "tds_required", "tds_26q", "short_26q"]
                ],
                "Final Reconciliation": final,
            }
        )

        st.download_button(
            "⬇️ Download Formatted Excel Reconciliation",
            data=excel_file,
            file_name="194Q_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"Processing failed: {exc}")
else:
    st.info("⬆️ Upload all three files to proceed.")
