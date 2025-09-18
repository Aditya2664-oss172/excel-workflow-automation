import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import os

# Paths
INPUT_FILE = "data/sample_data.xlsx"
OUTPUT_CLEANED_FILE = "output/cleaned_data.xlsx"
OUTPUT_REPORT_FILE = "output/summary_report.xlsx"

# Ensure output directory exists
os.makedirs("output", exist_ok=True)

def load_data(file_path):
    """Load Excel or CSV into pandas DataFrame"""
    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    return df

def clean_data(df):
    """Clean and validate data"""
    print("Cleaning data...")

    # Drop duplicates
    df = df.drop_duplicates()

    # Fill missing numeric values with mean
    numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns
    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())

    # Fill missing text values with 'Unknown'
    text_cols = df.select_dtypes(include=["object"]).columns
    df[text_cols] = df[text_cols].fillna("Unknown")

    return df

def validate_with_sql(df):
    """Apply SQL validation rules"""
    print("Validating with SQL rules...")

    # Create in-memory SQLite DB
    conn = sqlite3.connect(":memory:")
    df.to_sql("data", conn, index=False, if_exists="replace")

    # Example SQL rules
    query = """
    SELECT *
    FROM data
    WHERE Amount >= 0  -- remove negative amounts
    """
    validated_df = pd.read_sql_query(query, conn)
    conn.close()
    return validated_df

def generate_report(df):
    """Generate summary statistics and charts"""
    print("Generating report...")

    # Save summary statistics
    summary = df.describe(include="all")
    summary.to_excel(OUTPUT_REPORT_FILE)

    # Example visualization
    if "Amount" in df.columns:
        plt.figure(figsize=(6, 4))
        df["Amount"].hist(bins=20)
        plt.title("Distribution of Amounts")
        plt.xlabel("Amount")
        plt.ylabel("Frequency")
        plt.savefig("output/amount_distribution.png")
        plt.close()

def main():
    print("Starting Excel Workflow Automation...")

    df = load_data(INPUT_FILE)
    print(f"Loaded {len(df)} rows")

    df_clean = clean_data(df)
    df_validated = validate_with_sql(df_clean)

    # Save cleaned data
    df_validated.to_excel(OUTPUT_CLEANED_FILE, index=False)

    generate_report(df_validated)

    print("Automation complete ✅")
    print(f"Cleaned data saved to: {OUTPUT_CLEANED_FILE}")
    print(f"Report saved to: {OUTPUT_REPORT_FILE}")

if __name__ == "__main__":
    main()
