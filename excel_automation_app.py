import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import os
import logging
import sys
import io

# -------------------- Paths & Setup --------------------
OUTPUT_DIR = "output"
OUTPUT_CLEANED_FILE = os.path.join(OUTPUT_DIR, "cleaned_data.xlsx")
OUTPUT_REPORT_FILE = os.path.join(OUTPUT_DIR, "summary_report.xlsx")
LOG_FILE = os.path.join(OUTPUT_DIR, "automation.log")

os.makedirs(OUTPUT_DIR, exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# -------------------- Functions --------------------
def load_data(file):
    """Load Excel or CSV into pandas DataFrame"""
    try:
        logging.info(f"Loading file: {file}")
        if isinstance(file, str):
            if file.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
        else:  # Streamlit UploadedFile
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
        logging.info(f"Loaded {len(df)} rows successfully")
        return df
    except Exception as e:
        logging.error(f"Error loading file {file}: {e}")
        st.error(f"❌ Failed to load file: {e}")
        return None

def clean_data(df):
    """Clean and validate data"""
    try:
        logging.info("Cleaning data...")
        before_rows = len(df)

        # Drop duplicates
        df = df.drop_duplicates()

        # Fill missing numeric values with mean
        numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns
        df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())

        # Fill missing text values with 'Unknown'
        text_cols = df.select_dtypes(include=["object"]).columns
        df[text_cols] = df[text_cols].fillna("Unknown")

        after_rows = len(df)
        logging.info(f"Data cleaned: {before_rows - after_rows} duplicates removed")
        return df
    except Exception as e:
        logging.error(f"Error cleaning data: {e}")
        st.error(f"❌ Cleaning step failed: {e}")
        return df

def validate_with_sql(df):
    """Apply SQL validation rules"""
    try:
        logging.info("Applying SQL validation rules...")
        conn = sqlite3.connect(":memory:")
        df.to_sql("data", conn, index=False, if_exists="replace")

        query = """
        SELECT *
        FROM data
        WHERE Amount >= 0  -- remove negative amounts
        """
        validated_df = pd.read_sql_query(query, conn)
        conn.close()
        logging.info(f"Validation complete: {len(df) - len(validated_df)} invalid rows removed")
        return validated_df
    except Exception as e:
        logging.error(f"SQL validation failed: {e}")
        st.error(f"❌ SQL validation step failed: {e}")
        return df

def generate_report(df):
    """Generate summary statistics and charts"""
    try:
        logging.info("Generating summary report and charts...")
        summary = df.describe(include="all")
        summary.to_excel(OUTPUT_REPORT_FILE)

        # Generate histogram if "Amount" column exists
        if "Amount" in df.columns:
            plt.figure(figsize=(6, 4))
            df["Amount"].hist(bins=20)
            plt.title("Distribution of Amounts")
            plt.xlabel("Amount")
            plt.ylabel("Frequency")
            plt.tight_layout()
            plt.savefig(os.path.join(OUTPUT_DIR, "amount_distribution.png"))
            plt.close()
        logging.info("Report and visualization generated successfully")
    except Exception as e:
        logging.error(f"Error generating report: {e}")
        st.error(f"❌ Report generation failed: {e}")

# -------------------- Streamlit UI --------------------
st.title("📊 Excel Workflow Automation Tool")

uploaded_file = st.file_uploader("📂 Upload Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        st.success("✅ File loaded successfully!")
        st.dataframe(df.head())

        action = st.radio(
            "Select an Action", 
            ["Clean Data", "Validate Data", "Transform Data", "Generate Report"]
        )

        if action == "Clean Data":
            df = clean_data(df)
            st.write("✅ Cleaned Data Preview")
            st.dataframe(df.head())
            df.to_excel(OUTPUT_CLEANED_FILE, index=False)

        elif action == "Validate Data":
            df = validate_with_sql(df)
            st.write("✅ Validated Data Preview")
            st.dataframe(df.head())
            df.to_excel(OUTPUT_CLEANED_FILE, index=False)

        elif action == "Transform Data":
            df = clean_data(df)  # Using clean_data as transformation
            st.write("✅ Transformed Data Preview")
            st.dataframe(df.head())
            df.to_excel(OUTPUT_CLEANED_FILE, index=False)

        elif action == "Generate Report":
            generate_report(df)
            st.success(f"📄 Report generated successfully! Check the '{OUTPUT_DIR}' folder.")

        # Download processed data
        towrite = io.BytesIO()
        df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="📥 Download Processed File",
            data=towrite,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👆 Upload an Excel file to get started.")
