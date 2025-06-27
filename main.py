import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Salary Bank Transfer Generator", page_icon="💰")

st.title("💰 Salary Bank Transfer Sheet Generator")

# === Upload File ===
uploaded_file = st.file_uploader("📥 Upload Salary Sheet (Excel)", type=["xlsx"])

st.markdown("### ✏️ Narration Settings")

# Editable Narrations
default_debit_narr = "Salary paid for the month of {month} {cfl}"
default_credit_narr = "Salary credited for the month of {month}"

debit_narr_template = st.text_input("Debit Narration Template", value=default_debit_narr)
credit_narr_template = st.text_input("Credit Narration Template", value=default_credit_narr)

st.markdown("---")

# === Process Logic ===
if uploaded_file:
    with st.spinner("⏳ Generating Bank Transfer File..."):
        
        df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip()

        required_cols = [
            "Employee", "Employee Name", "Narration", "Bank Name", "IFSC Code", "Bank A/C No.",
            "Date of Joining", "Branch", "Region", "District", "CFL", "Block", "Department",
            "Designation", "Company", "Start Date", "End Date", "Leave Without Pay", "Absent Days",
            "Payment Days", "Basic", "Conveyance Allowance", "Dearness Allowance",
            "House Rent Allowance", "Leave Travel Allowance", "Medical Allowance", "Other Allowance",
            "Gross Pay", "ESI - Employee Contribution", "ESI - Employer Contribution",
            "PF (Employee's Contribution)", "PF (Employer's Contribution)", "Loan Repayment",
            "Total Deduction", "Net Pay"
        ]

        missing = [col for col in required_cols if col not in df.columns]

        if missing:
            st.error(f"❌ Missing columns in Salary Sheet: {missing}")
        else:
            records = []

            for _, row in df.iterrows():
                bank_name = str(row["Bank Name"]).upper()
                payment_type = "IFT" if "KOTAK" in bank_name else "NEFT"
                branch = str(row["Branch"]).strip()
                dr_ac_no = "6550063533" if branch == "UP Phase 3" else "6550063526"

                # Extract Month from Start Date
                start_date = pd.to_datetime(row["Start Date"], errors='coerce')
                month_year = start_date.strftime("%B-%Y") if not pd.isna(start_date) else "Unknown"

                # Narration with placeholders replaced
                debit_narr = debit_narr_template.format(month=month_year, cfl=row["CFL"])
                credit_narr = credit_narr_template.format(month=month_year)

                records.append({
                    "Client_Code": "AWOKEIND",
                    "Product_Code": "SALARY",
                    "Payment_Type": payment_type,
                    "Dr_Ac_No": dr_ac_no,
                    "Amount": row["Net Pay"],
                    "Bank_Code_Indicator": "M",
                    "Beneficiary_Name": row["Employee Name"],
                    "Beneficiary_Branch / IFSC Code": row["IFSC Code"],
                    "Beneficiary_Acc_No": row["Bank A/C No."],
                    "Debit_Narration": debit_narr,
                    "Credit_Narration": credit_narr
                })

            output_df = pd.DataFrame(records)

            # Download Excel
            output = BytesIO()
            output_df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)

            st.success(f"✅ Bank Transfer File generated with {len(output_df)} records.")
            st.dataframe(output_df.head())

            st.download_button(
                label="📤 Download Transfer File",
                data=output,
                file_name="Salary_Bank_Transfer_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
