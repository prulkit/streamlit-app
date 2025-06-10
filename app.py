import streamlit as st
import yfinance as yf
import pandas as pd
import requests
import logging
from datetime import datetime
from io import BytesIO
from typing import Tuple  # Import Tuple for type annotations

# ---------- Logging ----------
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ---------- Step 1: Find Ticker from Company Name ----------
def find_ticker(company_name: str) -> str:
    url = f"https://query2.finance.yahoo.com/v1/finance/search?q={company_name}"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        for item in data.get("quotes", []):
            if item.get("quoteType") == "EQUITY":
                return item.get("symbol")
    except Exception as e:
        st.error(f"Ticker lookup error for {company_name}: {e}")
        logging.error(f"Ticker lookup error for {company_name}: {e}")
    return None

# ---------- Step 2: Fetch Financial Data from yfinance ----------
def get_yfinance_data(ticker: str) -> dict:
    try:
        t = yf.Ticker(ticker)
        return {
            "info": t.info,
            "financials": t.financials,
            "balance_sheet": t.balance_sheet,
            "cashflow": t.cashflow,
            "quarterly_financials": t.quarterly_financials,
            "quarterly_balance_sheet": t.quarterly_balance_sheet,
            "quarterly_cashflow": t.quarterly_cashflow,
            "currency": t.info.get("financialCurrency", "Unknown")
        }
    except Exception as e:
        st.error(f"yfinance error for {ticker}: {e}")
        logging.error(f"yfinance error for {ticker}: {e}")
        return {}

# ---------- Step 3: Save to Excel and Return a Buffer ----------
def save_to_excel(company_name: str, ticker: str, data: dict) -> Tuple[BytesIO, str]:
    # Generate a filename based on the ticker and current date
    filename = f"{ticker}_public_diligence_{datetime.now():%Y%m%d}.xlsx"
    currency = data.get("currency", "Unknown")
    output = BytesIO()

    def convert_to_millions(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce') / 1e6
        return df

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Overview Sheet (transposed info from yfinance)
        if "info" in data and isinstance(data["info"], dict):
            overview_df = pd.DataFrame({
                "Attribute": list(data["info"].keys()),
                "Value": list(data["info"].values())
            })
            overview_df.to_excel(writer, sheet_name="Overview", index=False)

        # Helper function to write financial sheets with a currency note
        def write_financial_sheet(sheet_name: str, df: pd.DataFrame):
            df_millions = convert_to_millions(df)
            df_millions.reset_index(inplace=True)
            # Add currency info as the first row
            header_note = pd.DataFrame(
                [[f"Currency: {currency} (in millions)"] + [""] * (df_millions.shape[1] - 1)],
                columns=df_millions.columns
            )
            final_df = pd.concat([header_note, df_millions], ignore_index=True)
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Annual Statements Sheets
        if "financials" in data and not data["financials"].empty:
            write_financial_sheet("Annual_Income", data["financials"])
        if "balance_sheet" in data and not data["balance_sheet"].empty:
            write_financial_sheet("Annual_Balance", data["balance_sheet"])
        if "cashflow" in data and not data["cashflow"].empty:
            write_financial_sheet("Annual_Cashflow", data["cashflow"])

        # Quarterly Statements Sheets
        if "quarterly_financials" in data and not data["quarterly_financials"].empty:
            write_financial_sheet("Quarterly_Income", data["quarterly_financials"])
        if "quarterly_balance_sheet" in data and not data["quarterly_balance_sheet"].empty:
            write_financial_sheet("Quarterly_Balance", data["quarterly_balance_sheet"])
        if "quarterly_cashflow" in data and not data["quarterly_cashflow"].empty:
            write_financial_sheet("Quarterly_Cashflow", data["quarterly_cashflow"])

    output.seek(0)
    logging.info(f"Excel report generated: {filename}")
    return output, filename

# ---------- Step 4: Run the Public Diligence Pipeline ----------
def run_public_diligence(company_name: str):
    ticker = find_ticker(company_name)
    if not ticker:
        st.warning(f"Ticker not found for {company_name}")
        return None, None
    st.info(f"Found ticker for {company_name}: {ticker}")
    logging.info(f"Found ticker for {company_name}: {ticker}")
    data = get_yfinance_data(ticker)
    if not data:
        st.warning(f"Could not fetch financial data for {company_name}")
        return None, None
    excel_buffer, filename = save_to_excel(company_name, ticker, data)
    return excel_buffer, filename

# ---------- Streamlit App Interface ----------
st.title("Desktop Diligence Dashboard")
st.markdown("This tool automates desktop diligence by fetching financial data from online sources using yfinance and saving it as an Excel report.")

# Text input for comma-separated company names (e.g., "Apple, Microsoft")
company_input = st.text_input("Enter Public Company Names (comma-separated)", "Apple, Microsoft")

if st.button("Run Public Diligence"):
    companies = [name.strip() for name in company_input.split(',') if name.strip()]
    st.write(f"Running analysis for: {', '.join(companies)}")
    # Process each company separately
    for company in companies:
        with st.spinner(f"Processing {company}..."):
            excel_buffer, filename = run_public_diligence(company)
            if excel_buffer:
                st.success(f"Excel report generated for {company}")
                # Provide a download button for the Excel file
                st.download_button(
                    label=f"Download {filename}",
                    data=excel_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
