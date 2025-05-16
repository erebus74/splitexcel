import pandas as pd
import streamlit as st
import io
from datetime import datetime

st.title("Excel File Processor")

# File upload section
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if uploaded_file:
    try:
        # Read the uploaded file
        df_dict = pd.read_excel(uploaded_file, sheet_name=None)
        main_df = pd.read_excel(uploaded_file)  # For the first feature
        
        st.success("File uploaded successfully!")
        
        # Show preview
        st.subheader("Data Preview")
        st.write(main_df.head())
        
        # Create two columns for buttons
        col1, col2 = st.columns(2)
        
        with col1:
            # Original feature: Split by unique values into sheets
            if st.button("Split by Unique Values into Sheets"):
                with st.spinner("Processing..."):
                    if 'eiin' not in main_df.columns:
                        st.error("The Excel file must contain a 'eiin' column")
                    else:
                        unique_values = main_df["eiin"].unique()
                        output = io.BytesIO()
                        
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            for value in unique_values:
                                filtered_df = main_df[main_df["eiin"] == value]
                                sheet_name = str(value)[:31]
                                for char in [':', '\\', '?', '/', '*', '[', ']']:
                                    sheet_name = sheet_name.replace(char, '')
                                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        output.seek(0)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"split_by_values_{timestamp}.xlsx"
                        
                        st.success("Processing complete!")
                        st.download_button(
                            label="Download Split-by-Values File",
                            data=output,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        
        with col2:
            # New feature: Split worksheets into separate workbooks
            if st.button("Split Worksheets to Separate Files"):
                with st.spinner("Processing..."):
                    if len(df_dict) == 1:
                        st.warning("File already has only one worksheet")
                    else:
                        # Create a zip file containing all sheets as separate files
                        from zipfile import ZipFile
                        zip_buffer = io.BytesIO()
                        
                        with ZipFile(zip_buffer, 'w') as zip_file:
                            for sheet_name, df in df_dict.items():
                                single_sheet_buffer = io.BytesIO()
                                with pd.ExcelWriter(single_sheet_buffer, engine='openpyxl') as writer:
                                    df.to_excel(writer, sheet_name="Sheet1", index=False)
                                single_sheet_buffer.seek(0)
                                zip_file.writestr(f"{sheet_name}.xlsx", single_sheet_buffer.read())
                        
                        zip_buffer.seek(0)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"separated_sheets_{timestamp}.zip"
                        
                        st.success("Processing complete!")
                        st.download_button(
                            label="Download All Sheets as Separate Files",
                            data=zip_buffer,
                            file_name=filename,
                            mime="application/zip"
                        )
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
