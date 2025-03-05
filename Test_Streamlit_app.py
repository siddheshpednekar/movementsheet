import streamlit as st
import pandas as pd
from io import BytesIO
import re
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font

# Streamlit App
st.title("Movement Sheet Web App")

# File uploader for Excel files
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        package_file_path = r"Package File Movement sheet.xlsx"

        # Read the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)
        package_df = pd.read_excel(package_file_path)

        # Required columns
        required_columns = {"BookingStatus", "ServiceType", "ServiceTime", "ETA", "DepartureFlightNumber",
                            "ArrivalFlightNumber", "TransitFlightNumber", "ETD", "Origin", "Destination",
                            "PackageName", "Nationality", "TravelClass", "Remarks"}

        # Check for missing columns
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            st.error(f"Missing columns in uploaded file: {', '.join(missing_columns)}")
            st.stop()

        st.write("### Uploaded Data:")
        st.dataframe(df)

        # Package Mapping
        package_mapping = dict(zip(package_df["Package_Name"], package_df["Revised_Package_Name"]))

        # Step 2: Filter BookingStatus
        df = df[df["BookingStatus"].isin(["Completed", "PaymentCompleted"])]

        # Step 3: Adjust "ServiceType" for RoundTrip
        df.loc[(df["ServiceType"] == "RoundTrip") & (df["ServiceTime"] == df["ETA"]), "ServiceType"] = "Arrival"
        df.loc[(df["ServiceType"] == "RoundTrip") & (df["ServiceTime"] != df["ETA"]), "ServiceType"] = "Departure"

        # Step 4: Create "Flight No." column
        def get_flight_no(row):
            if row["ServiceType"] == "Departure" and row["DepartureFlightNumber"] not in ["NA", "", None]:
                return row["DepartureFlightNumber"]
            elif row["ServiceType"] == "Arrival" and row["ArrivalFlightNumber"] not in ["NA", "", None]:
                return row["ArrivalFlightNumber"]
            elif row["ServiceType"] == "Transit":
                return f"{row['ArrivalFlightNumber']} / {row['TransitFlightNumber']}"
            return ""

        df["Flight No."] = df.apply(get_flight_no, axis=1)

        # Step 5: Create "ETA/ETD" column
        df["ETA/ETD"] = df["ETD"].fillna(df["ETA"])

        # Step 6: Clean "Origin" and "Destination" columns
        def clean_location(value):
            return "" if pd.isna(value) else re.split(r"[,/]", value)[0].strip()

        df["Origin"] = df["Origin"].apply(clean_location)
        df["Destination"] = df["Destination"].apply(clean_location)

        # Step 7: Map "Package" column
        df["Package"] = df["PackageName"].apply(lambda x: package_mapping.get(x, x))

        # Step 8: Remove rows where "Remarks" is "Cancelled"
        df = df[df["Remarks"] != "Cancelled"]

        # Step 9: Save to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

        st.write("✅ Data processing completed successfully!")
        st.write("### Processed Data:")
        st.dataframe(df)

        # Step 10: Download Button
        st.download_button(
            label="Download Excel",
            data=output,
            file_name="modified_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
