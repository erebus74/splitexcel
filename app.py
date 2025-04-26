import pandas as pd
import streamlit as st
import io
from datetime import datetime

st.title("Excel File Processor")

# File upload
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if uploaded_file:
    try:
        # Read the uploaded file
        df = pd.read_excel(uploaded_file)
        
        # Check if 'code' column exists
        if 'code' not in df.columns:
            st.error("The Excel file must contain a 'code' column")
        else:
            st.success("File uploaded successfully!")
            
            # Show preview
            st.subheader("Data Preview")
            st.write(df.head())
            
            # Process file when button is clicked
            if st.button("Process File"):
                with st.spinner("Processing..."):
                    # Get unique values
                    unique_values = df["code"].unique()
                    
                    # Create in-memory Excel file
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for value in unique_values:
                            # Filter data for current value
                            filtered_df = df[df["code"] == value]
                            
                            # Clean sheet name
                            sheet_name = str(value)[:31]
                            for char in [':', '\\', '?', '/', '*', '[', ']']:
                                sheet_name = sheet_name.replace(char, '')
                            
                            # Write to sheet
                            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Prepare download
                    output.seek(0)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"processed_file_{timestamp}.xlsx"
                    
                    st.success("Processing complete!")
                    st.download_button(
                        label="Download Processed File",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")