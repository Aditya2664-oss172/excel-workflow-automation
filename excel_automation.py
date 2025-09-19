import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import os
import logging
import sys

# Paths
INPUT_FILE = "/home/aditya-chauhan/Desktop/automation/sample_data.xlsx"
OUTPUT_CLEANED_FILE = "output/cleaned_data.xlsx"
OUTPUT_REPORT_FILE = "output/summary_report.xlsx"
LOG_FILE = "output/automation.log"

# Ensure output directory exists
os.makedirs("output", exist_ok=True)

# Configure logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def load_data(file_path):
    """Load Excel or CSV into pandas DataFrame"""
    try:
        logging.info(f"Loading file: {file_path}")
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        logging.info(f"Loaded {len(df)} rows successfully")
        return df
    except Exception as e:
        logging.error(f"Error loading file {file_path}: {e}")
        sys.exit(f"❌ Failed to load file: {e}")

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
        sys.exit(f"❌ Cleaning step failed: {e}")

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
        sys.exit(f"❌ SQL validation step failed: {e}")

def generate_report(df):
    """Generate summary statistics and charts"""
    try:
        logging.info("Generating summary report and charts...")
        summary = df.describe(include="all")
        summary.to_excel(OUTPUT_REPORT_FILE)

        if "Amount" in df.columns:
            plt.figure(figsize=(6, 4))
            df["Amount"].hist(bins=20)
            plt.title("Distribution of Amounts")
            plt.xlabel("Amount")
            plt.ylabel("Frequency")
            plt.savefig("output/amount_distribution.png")
            plt.close()
        logging.info("Report and visualization generated successfully")
    except Exception as e:
        logging.error(f"Error generating report: {e}")
        sys.exit(f"❌ Report generation failed: {e}")

def main():
    logging.info("Starting Excel Workflow Automation 🚀")
    try:
        df = load_data(INPUT_FILE)
        df_clean = clean_data(df)
        df_validated = validate_with_sql(df_clean)
        df_validated.to_excel(OUTPUT_CLEANED_FILE, index=False)
        logging.info(f"Cleaned data saved to {OUTPUT_CLEANED_FILE}")
        generate_report(df_validated)
        logging.info("Automation complete ✅")
    except Exception as e:
        logging.critical(f"Unexpected error: {e}")
        sys.exit(f"❌ Process failed: {e}")

if __name__ == "__main__":
    main()
