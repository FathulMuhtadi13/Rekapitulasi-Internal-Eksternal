import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# Function to load the Excel file
def load_data(file):
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Main Streamlit application
def main():
    st.set_page_config(page_title="Rekapitulasi Audit ISO Internal & Eksternal RECARE", layout="wide")
    st.title("Rekapitulasi Audit ISO Internal & Eksternal RECARE")
    
    # Sidebar for file upload
    uploaded_file = st.sidebar.file_uploader("Upload file Excel", type=["xlsx"])

    if uploaded_file is None:
        st.info("Silahkan masukan file excel anda")
    else:
        df = load_data(uploaded_file)

        if df is not None:
            st.success("File anda telah terbaca")
            
            # Clean column names
            df.columns = df.columns.str.strip()

            # Date filters
            st.sidebar.subheader("Filter berdasarkan Tanggal")
            due_date_range = st.sidebar.date_input("Pilih Rentang Due Date", 
                                                    [df['Due Date'].min(), df['Due Date'].max()])
            closing_date_range = st.sidebar.date_input("Pilih Rentang Tgl Closing", 
                                                        [df['Tgl Closing'].min(), df['Tgl Closing'].max()])

            # Finding Category filter
            st.sidebar.subheader("Filter berdasarkan Finding Category")
            finding_category = st.sidebar.multiselect("Pilih Finding Category", 
                                                       options=df['Finding Category'].unique(),
                                                       default=df['Finding Category'].unique())

            # Finding Status filter
            st.sidebar.subheader("Filter berdasarkan Finding Status")
            finding_status = st.sidebar.multiselect("Pilih Finding Status", 
                                                     options=df['Finding Status'].unique(),
                                                     default=df['Finding Status'].unique())

            # Area filter
            st.sidebar.subheader("Filter berdasarkan Area")
            area = st.sidebar.multiselect("Pilih Area", 
                                           options=df['Area'].unique(),
                                           default=df['Area'].unique())

            # Apply filters
            filtered_df = df[
                (df['Due Date'].between(*due_date_range)) &
                (df['Tgl Closing'].between(*closing_date_range)) &
                (df['Finding Category'].isin(finding_category)) &
                (df['Finding Status'].isin(finding_status)) &
                (df['Area'].isin(area))
            ]

            # Show filtered data
            st.subheader("Data Audit yang Difilter")
            st.dataframe(filtered_df)

            # Download button for filtered data
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download Data yang Difilter", 
                               csv, 
                               "filtered_data.csv", 
                               "text/csv", 
                               key='download-csv')

            # Visualization: Example chart for Finding Status
            fig = px.histogram(filtered_df, x="Finding Status", title="Distribusi Status Temuan", 
                               color="Finding Category", barmode='group')
            st.plotly_chart(fig)

            # Editable Data Table (this requires the st-aggrid library)
            st.subheader("Edit Data")
            import streamlit_aggrid as st_aggrid

            grid_response = st_aggrid.aggrid(
                filtered_df,
                editable=True,
                height=300,
                fit_columns_on_grid_load=True,
            )

            updated_df = grid_response['data']  # Get updated data
            if st.button("Simpan Perubahan"):
                # Save updated data back to Excel
                with pd.ExcelWriter(uploaded_file, mode='a', engine='openpyxl') as writer:
                    updated_df.to_excel(writer, sheet_name='Updated Data', index=False)
                st.success("Perubahan berhasil disimpan!")

if __name__ == "__main__":
    main()
