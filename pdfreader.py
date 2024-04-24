import streamlit as st
from tabula import read_pdf
import pandas as pd
from io import BytesIO

def save_df_to_excel(dfs):
    """Save a list of dataframes to an Excel file, each dataframe on a different sheet."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f'Table {i + 1}')
    return output

def main():
    st.title('PDF Table Extractor')
    
    # Introduction and user instructions
    st.write("""
    This application allows you to extract tables from a PDF file and download them in an Excel format. 
    Please upload your PDF and choose the appropriate extraction method based on the type of tables in your PDF:
    """)
    
    # Explanation of extraction methods
    st.markdown("""
    **Lattice:** Best for PDFs with tables that have clear, visible borders. This method tends to work well 
    with structured PDFs where tables are well-defined by grid lines.

    **Stream:** Ideal for PDFs where tables are defined by whitespace rather than visible lines. This method 
    is suitable for documents with less structured tables that do not have clear borders.
    """)
    
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    extraction_method = st.selectbox("Choose the extraction mode:", ["Lattice", "Stream"])

    if uploaded_file is not None:
        # Process the uploaded PDF and extract tables
        try:
            if extraction_method == "Lattice":
                dfs = read_pdf(uploaded_file, pages='all', multiple_tables=True, lattice=True)
            else:  # Stream method
                dfs = read_pdf(uploaded_file, pages='all', multiple_tables=True, stream=True)

            if dfs:
                st.success('Tables extracted successfully!')
                st.write(f"Extracted {len(dfs)} tables.")
                
                # Save the tables to an Excel file
                excel_file = save_df_to_excel(dfs)
                excel_file.seek(0)  # Move the pointer to the start of the file

                # Download link for Excel file
                st.download_button(label="Download Excel file with all tables",
                                   data=excel_file,
                                   file_name="extracted_tables.xlsx",
                                   mime="application/vnd.ms-excel")
            else:
                st.warning('No tables found in the PDF file.')
        except Exception as e:
            st.error(f'An error occurred: {e}')

    # Footer
    st.text("Powered by DA Business Information")

if __name__ == "__main__":
    main()
