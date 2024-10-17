import streamlit as st
import pandas as pd
import plotly.express as px
import streamlit.components.v1 as components
from st_aggrid import AgGrid, GridOptionsBuilder

# Fungsi untuk memuat data dari file Excel
def load_data(file):
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Fungsi untuk menyimpan data ke file Excel
def save_data_to_excel(df, filename):
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, index=False, sheet_name='Updated Data')

# Fungsi utama aplikasi Streamlit
def main():
    st.set_page_config(page_title="Rekapitulasi Audit ISO Internal & Eksternal RECARE", layout="wide")
    st.title("Rekapitulasi Audit ISO Internal & Eksternal RECARE")
    
    # Sidebar untuk upload file
    uploaded_file = st.sidebar.file_uploader("Upload file Excel", type=["xlsx"])

    if uploaded_file is None:
        st.info("Silahkan masukan file excel anda")
    else:
        st.success("File anda telah terbaca")
        df = load_data(uploaded_file)

        if df is not None:
            # Bersihkan nama kolom
            df.columns = df.columns.str.strip()

            # Filter berdasarkan tanggal
            st.sidebar.subheader("Filter berdasarkan Tanggal")
            due_date_range = st.sidebar.date_input("Pilih Rentang Due Date", 
                                                    [df['Due Date'].min(), df['Due Date'].max()])
            closing_date_range = st.sidebar.date_input("Pilih Rentang Tgl Closing", 
                                                        [df['Tgl Closing'].min(), df['Tgl Closing'].max()])

            # Filter berdasarkan Finding Category
            st.sidebar.subheader("Filter berdasarkan Finding Category")
            finding_category = st.sidebar.multiselect("Pilih Finding Category", 
                                                       options=df['Finding Category'].unique(),
                                                       default=df['Finding Category'].unique())

            # Filter berdasarkan Finding Status
            st.sidebar.subheader("Filter berdasarkan Finding Status")
            finding_status = st.sidebar.multiselect("Pilih Finding Status", 
                                                     options=df['Finding Status'].unique(),
                                                     default=df['Finding Status'].unique())

            # Filter berdasarkan Area
            st.sidebar.subheader("Filter berdasarkan Area")
            area = st.sidebar.multiselect("Pilih Area", 
                                           options=df['Area'].unique(),
                                           default=df['Area'].unique())

            # Terapkan filter
            filtered_df = df[
                (df['Finding Category'].isin(finding_category)) &
                (df['Finding Status'].isin(finding_status)) &
                (df['Area'].isin(area)) &
                # Tetap berikan data jika ada salah satu tanggal yang valid
                ((df['Due Date'].between(*due_date_range) | df['Due Date'].isna()) |
                 (df['Tgl Closing'].between(*closing_date_range) | df['Tgl Closing'].isna()))
            ]

            # Tampilkan data yang sudah difilter
            st.subheader("Data Audit yang Difilter")
            st.dataframe(filtered_df)

            # Menu download untuk data yang difilter
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download Data yang Difilter", 
                               csv, 
                               "filtered_data.csv", 
                               "text/csv", 
                               key='download-csv')

            # Visualisasi: Histogram status temuan
            fig = px.histogram(filtered_df, x="Finding Status", title="Distribusi Status Temuan", 
                               color="Finding Category", barmode='group')
            st.plotly_chart(fig, use_container_width=True)

            # Tabel yang dapat diedit
            st.subheader("Edit Data")
            grid_options = GridOptionsBuilder.from_dataframe(filtered_df)
            grid_options.configure_default_column(editable=True)
            grid_options.configure_pagination(paginationPageSize=10)
            grid_options.configure_side_bar()
            grid_options.configure_columns(list(filtered_df.columns))

            grid_response = AgGrid(filtered_df, gridOptions=grid_options.build(), enable_enterprise_modules=True)

            updated_df = pd.DataFrame(grid_response['data'])  # Dapatkan data yang diperbarui
            if st.button("Simpan Perubahan"):
                save_data_to_excel(updated_df, uploaded_file.name)
                st.success("Perubahan berhasil disimpan!")

if __name__ == "__main__":
    main()
