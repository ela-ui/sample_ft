import pandas as pd
import streamlit as st

# Function to process and merge the sheets
def process_files(sheet1, sheet2, sheet3):
    # Step 2: Merge the first two sheets on 'Beneficiary Addr. Line 3' (sheet1) and 'REMARKS' (sheet2)
    merged_data = pd.merge(sheet2, sheet1, left_on='REMARKS', right_on='Beneficiary Addr. Line 3', how='inner')

    # Step 3: Remove duplicates based on 'REMARKS' and 'AMOUNT'
    merged_data = merged_data.drop_duplicates(subset=['REMARKS', 'AMOUNT'])

    # Step 4: Update 'UTR NUMBER' where its value is '- -' with 'REFERENCE NUMBER'
    merged_data.loc[merged_data['UTR NUMBER'].str.strip() == '- -', 'UTR NUMBER'] = merged_data['REFERENCE NUMBER']

    # Step 5: Select the first 16 columns (A to P) and the last column (AX)
    columns_to_save = list(merged_data.columns[:16]) + [merged_data.columns[-1]]

    # Step 6: Extract 'utr1' from 'Narrative' in the third sheet (sheet3)
    sheet3['utr1'] = sheet3['Narrative'].str.extract(r'/(?P<utr_value>[^/]+)/')['utr_value']
    sheet3['utr1'] = sheet3['utr1'].fillna(sheet3['Narrative'].str.extract(r'_(?P<after_underscore>.+)$')['after_underscore'])
    sheet3['utr1'] = sheet3['utr1'].where(sheet3['utr1'].notna(), None)

    # Step 7: Merge the updated intermediate data with the third sheet on 'utr1'
    final_merged_data = pd.merge(sheet3, merged_data, left_on='utr1', right_on='UTR NUMBER', how='inner')

    # Step 8: Select the first 9 columns
    columns_a_to_i = list(final_merged_data.columns[:9])

    # Generate intermediate file (16 columns + AX)
    intermediate_file = merged_data[columns_to_save]

    # Generate final file (first 9 columns)
    final_file = final_merged_data[columns_a_to_i]

    return intermediate_file, final_file

# Streamlit interface
def main():
    st.title("Mapping Tool")

    # Step 1: Upload files
    sheet1_file = st.file_uploader("Upload first sheet (bulk.xlsx)", type=["xlsx"])
    sheet2_file = st.file_uploader("Upload second sheet (payment.xlsx)", type=["xlsx"])
    sheet3_file = st.file_uploader("Upload third sheet (statement.xlsx)", type=["xlsx"])

    # If all files are uploaded
    if sheet1_file and sheet2_file and sheet3_file:
        # Load the sheets
        sheet1 = pd.read_excel(sheet1_file)
        sheet2 = pd.read_excel(sheet2_file)
        sheet3 = pd.read_excel(sheet3_file)

        # Process and merge the sheets
        intermediate_output, final_output = process_files(sheet1, sheet2, sheet3)

        # Display the intermediate and final outputs
        st.write("Intermediate Processed Data (16 columns + AX)", intermediate_output)
        st.write("Final Mapped Data (first 9 columns)", final_output)

        # Step 9: Download buttons for both intermediate and final processed files

        # Save intermediate file
        intermediate_file_path = "payment_status_output.xlsx"
        intermediate_output.to_excel(intermediate_file_path, index=False)
        with open(intermediate_file_path, "rb") as f:
            st.download_button(
                label="Download Intermediate Processed File",
                data=f,
                file_name=intermediate_file_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Save final file
        final_file_path = "final_mapped_output.xlsx"
        final_output.to_excel(final_file_path, index=False)
        with open(final_file_path, "rb") as f:
            st.download_button(
                label="Download Final Mapped File",
                data=f,
                file_name=final_file_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Run the Streamlit app
if __name__ == "__main__":
    main()
