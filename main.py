import streamlit as st
import pandas as pd
import io

def main():
    st.title("Excel Data Manipulation")

    # File uploads
    uploaded_file1 = st.file_uploader("Upload Excel1", type=["xlsx"])
    uploaded_file2 = st.file_uploader("Upload Excel2", type=["xlsx"])

    if uploaded_file1 and uploaded_file2:
        st.write("Files uploaded successfully!")

        # Load Excel files
        excel1_df = pd.read_excel(uploaded_file1)
        excel2_df = pd.read_excel(uploaded_file2, sheet_name=None)

        # Strip spaces from Excel1 "Email" column
        excel1_df["Email"] = excel1_df["Email"].str.strip()

        # Process Excel2 sheets and create Excel3
        excel3_writer = pd.ExcelWriter("Excel3.xlsx", engine="xlsxwriter")
        for sheet_name, sheet_data in excel2_df.items():
            # Add "Sessions" column and initialize with 0
            sheet_data["Sessions"] = 0

            # Strip spaces from Excel2 "Email" column
            if "Email" in sheet_data.columns:
                sheet_data["Email"] = sheet_data["Email"].str.strip()

            # Copy data to Excel3
            sheet_data.to_excel(excel3_writer, sheet_name=sheet_name, index=False)

            # Check if "Email" parameter exists in Excel2
            if "Email" in sheet_data.columns:
                for index, row in sheet_data.iterrows():
                    email = row["Email"]
                    matching_sessions = excel1_df.loc[excel1_df["Email"] == email, "Sessions"].values
                    if len(matching_sessions) > 0:
                        excel3_writer.sheets[sheet_name].write(index + 1, sheet_data.columns.get_loc("Sessions"), matching_sessions[0])

        excel3_writer.close()  # Close the writer

        # Provide download link for Excel3
        excel3_data = io.BytesIO()
        with open("Excel3.xlsx", "rb") as file:
            excel3_data.write(file.read())
        st.download_button("Download Excel3", data=excel3_data.getvalue(), file_name="Excel3.xlsx")

if __name__ == "__main__":
    main()
