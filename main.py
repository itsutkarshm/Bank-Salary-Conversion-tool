import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import base64

st.set_page_config(page_title="Salary Bank Transfer Generator", page_icon="💰")

# === Modern Styling ===
st.markdown(
    """
<style>
    .stButton>button {
        background-color: #e7e7e7;
        color: black;
        border: none;
        border-radius: 20px;
        padding: 8px 16px;
        margin: 4px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #40916c;
        color: white;
    }
    .red-header {
        color: red;
        font-weight: bold;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.title("💰 Salary Bank Transfer Sheet Generator")

# === Upload File ===
uploaded_file = st.file_uploader("📥 Upload Salary Sheet (Excel)", type=["xlsx"])

# === Narration Settings ===
st.markdown("### ✏️ Narration Settings")
debit_narration_template = st.text_input(
    "Debit Narration Template", "Salary paid for the month of {month} {cfl}"
)
credit_narration_template = st.text_input(
    "Credit Narration Template", "Salary credited for the month of {month}"
)

# === Process ===
if uploaded_file:
    with st.spinner("⏳ Processing Salary Sheet..."):

        df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip()

        required_cols = [
            "Employee",
            "Employee Name",
            "Bank Name",
            "IFSC Code",
            "Bank A/C No.",
            "CFL",
            "Branch",
            "Start Date",
            "End Date",
            "Net Pay",
        ]

        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            st.error(f"❌ Missing columns in Salary Sheet: {missing}")
        else:
            df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
            df["Month"] = df["Start Date"].dt.strftime("%b-%Y")

            # All Headers
            all_headers = [
                "Client_Code",
                "Product_Code",
                "Payment_Type",
                "Payment_Ref_No.",
                "Payment_Date",
                "Instrument Date",
                "Dr_Ac_No",
                "Amount",
                "Bank_Code_Indicator",
                "Beneficiary_Code",
                "Beneficiary_Name",
                "Beneficiary_Bank",
                "Beneficiary_Branch / IFSC Code",
                "Beneficiary_Acc_No",
                "Location",
                "Print_Location",
                "Instrument_Number",
                "Ben_Add1",
                "Ben_Add2",
                "Ben_Add3",
                "Ben_Add4",
                "Beneficiary_Email",
                "Beneficiary_Mobile",
                "Debit_Narration",
                "Credit_Narration",
                "Payment Details 1",
                "Payment Details 2",
                "Payment Details 3",
                "Payment Details 4",
                "Enrichment_1",
                "Enrichment_2",
                "Enrichment_3",
                "Enrichment_4",
                "Enrichment_5",
                "Enrichment_6",
                "Enrichment_7",
                "Enrichment_8",
                "Enrichment_9",
                "Enrichment_10",
                "Enrichment_11",
                "Enrichment_12",
                "Enrichment_13",
                "Enrichment_14",
                "Enrichment_15",
                "Enrichment_16",
                "Enrichment_17",
                "Enrichment_18",
                "Enrichment_19",
                "Enrichment_20",
            ]

            populated_cols = [
                "Client_Code",
                "Product_Code",
                "Payment_Type",
                "Dr_Ac_No",
                "Amount",
                "Bank_Code_Indicator",
                "Beneficiary_Name",
                "Beneficiary_Bank",
                "Beneficiary_Branch / IFSC Code",
                "Beneficiary_Acc_No",
                "Debit_Narration",
                "Credit_Narration",
            ]

            # Convert Full Data for Immediate Table
            records = []
            for _, row in df.iterrows():
                bank_name = str(row["Bank Name"]).upper()
                payment_type = "IFT" if "KOTAK" in bank_name else "NEFT"
                dr_ac_no = (
                    "6550063533" if row["Branch"] == "UP Phase 3" else "6550063526"
                )
                month = row["Month"]
                cfl = row["CFL"]

                record = {header: "" for header in all_headers}
                record.update(
                    {
                        "Client_Code": "AWOKEIND",
                        "Product_Code": "SALARY",
                        "Payment_Type": payment_type,
                        "Dr_Ac_No": dr_ac_no,
                        "Amount": row["Net Pay"],
                        "Bank_Code_Indicator": "M",
                        "Beneficiary_Name": row["Employee Name"],
                        "Beneficiary_Bank": row["Bank Name"],
                        "Beneficiary_Branch / IFSC Code": row["IFSC Code"],
                        "Beneficiary_Acc_No": row["Bank A/C No."],
                        "Debit_Narration": debit_narration_template.format(
                            month=month, cfl=cfl
                        ),
                        "Credit_Narration": credit_narration_template.format(
                            month=month
                        ),
                    }
                )
                records.append(record)

            converted_df = pd.DataFrame(records, columns=all_headers)

            st.success(
                f"✅ Full Salary Conversion Completed ({len(converted_df)} Records)"
            )

            styled_df = converted_df.style.set_table_styles(
                [
                    {
                        "selector": f"th.col{converted_df.columns.get_loc(col)}",
                        "props": "color: red;",
                    }
                    for col in populated_cols
                ]
            )

            st.dataframe(styled_df, use_container_width=True)

            # === CFL-wise Filter and Download ===
            st.markdown("### 🎯 Generate CFL-wise Transfer Files")

            unique_cfls = sorted(df["CFL"].dropna().unique().tolist())
            col1, col2 = st.columns(2)
            with col1:
                select_all = st.button("Select All CFLs")
            with col2:
                clear_all = st.button("Clear All CFLs")

            if "selected_cfls" not in st.session_state:
                st.session_state.selected_cfls = []

            if select_all:
                st.session_state.selected_cfls = unique_cfls
            if clear_all:
                st.session_state.selected_cfls = []

            selected_cfls = st.multiselect(
                "Select CFL(s) to Download",
                options=unique_cfls,
                default=st.session_state.selected_cfls,
            )
            st.session_state.selected_cfls = selected_cfls

            if selected_cfls:
                if len(selected_cfls) == 1:
                    cfl = selected_cfls[0]
                    output_df = converted_df[df["CFL"] == cfl]

                    excel_buffer = BytesIO()
                    output_df.to_excel(excel_buffer, index=False, engine="openpyxl")
                    excel_buffer.seek(0)

                    st.download_button(
                        label=f"📥 Download {cfl} Transfer File",
                        data=excel_buffer,
                        file_name=f"{cfl}_Bank_Transfer_File.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zipf:
                        for cfl in selected_cfls:
                            output_df = converted_df[df["CFL"] == cfl]
                            if output_df.empty:
                                continue
                            excel_buffer = BytesIO()
                            output_df.to_excel(
                                excel_buffer, index=False, engine="openpyxl"
                            )
                            excel_buffer.seek(0)
                            zipf.writestr(
                                f"{cfl}_Bank_Transfer_File.xlsx", excel_buffer.read()
                            )

                    zip_buffer.seek(0)
                    b64 = base64.b64encode(zip_buffer.read()).decode()
                    href = f'<a href="data:application/zip;base64,{b64}" download="CFL_Wise_Bank_Transfer_Files.zip">📥 Download ZIP File</a>'
                    st.markdown(href, unsafe_allow_html=True)
