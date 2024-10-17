import streamlit as st
import pandas as pd
import plotly.express as px

# Judul Dashboard
st.title("Rekapitulasi Audit ISO Internal & Eksternal RECARE")

# Upload File di Sidebar
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=['xlsx'])

# Pesan jika file belum diupload atau sudah terbaca
if uploaded_file is None:
    st.info("Silahkan masukan file excel anda")
else:
    st.success("File anda telah terbaca")
    # Membaca file excel yang diupload
    df = pd.read_excel(uploaded_file)

    # Menampilkan struktur data agar sesuai dengan header yang diminta
    columns_required = ['No', 'Audit Description', 'Klausul', 'Standar ISO', 'Standard Description', 
                        'PPWI', 'Finding Descriptions', 'Finding Category', 'Root Cause', 'Correction', 
                        'Corrective Action and Evidence', 'Due Date', 'Tgl Closing', 
                        'Finding Status', 'Area', 'Evaluation']

    # Memeriksa apakah kolom dalam data sesuai dengan kolom yang diperlukan
    if set(columns_required).issubset(df.columns):
        # Filter untuk range date berdasarkan 'Due Date' dan 'Tgl Closing'
        date_range = st.sidebar.date_input("Pilih Rentang Tanggal (Due Date)", [])
        closing_range = st.sidebar.date_input("Pilih Rentang Tanggal (Tgl Closing)", [])

        # Filter untuk kategori lainnya
        finding_category = st.sidebar.multiselect("Filter Finding Category", options=df['Finding Category'].unique())
        finding_status = st.sidebar.multiselect("Filter Finding Status", options=df['Finding Status'].unique())
        area = st.sidebar.multiselect("Filter Area", options=df['Area'].unique())

        # Filter data berdasarkan input dari user
        filtered_data = df.copy()
        if date_range:
            filtered_data = filtered_data[(filtered_data['Due Date'] >= pd.to_datetime(date_range[0])) & 
                                          (filtered_data['Due Date'] <= pd.to_datetime(date_range[1]))]
        if closing_range:
            filtered_data = filtered_data[(filtered_data['Tgl Closing'] >= pd.to_datetime(closing_range[0])) & 
                                          (filtered_data['Tgl Closing'] <= pd.to_datetime(closing_range[1]))]
        if finding_category:
            filtered_data = filtered_data[filtered_data['Finding Category'].isin(finding_category)]
        if finding_status:
            filtered_data = filtered_data[filtered_data['Finding Status'].isin(finding_status)]
        if area:
            filtered_data = filtered_data[filtered_data['Area'].isin(area)]

        # Tampilkan tabel hasil filter
        st.dataframe(filtered_data)

        # Download filtered data
        def convert_df(df):
            return df.to_csv(index=False).encode('utf-8')
        
        csv = convert_df(filtered_data)
        st.download_button(label="Download Filtered Data", data=csv, file_name='filtered_data.csv', mime='text/csv')

        # Chart Visualisasi
        st.subheader("Visualisasi Chart")
        if not filtered_data.empty:
            fig = px.bar(filtered_data, x='Finding Category', color='Finding Status', barmode='group')
            st.plotly_chart(fig)

        # Menambahkan animasi dan CSS
        st.markdown(
            """
            <style>
            @keyframes fadeIn {
                from {opacity: 0;}
                to {opacity: 1;}
            }

            .element-container {
                animation: fadeIn 1.5s ease-in-out;
            }

            .stButton>button {
                background-color: #4CAF50;
                color: white;
                transition: 0.3s;
            }

            .stButton>button:hover {
                background-color: #45a049;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        # Menambahkan fitur edit data (gunakan st_aggrid)
        from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

        st.subheader("Edit Data Tabel")
        gb = GridOptionsBuilder.from_dataframe(filtered_data)
        gb.configure_default_column(editable=True)
        grid_options = gb.build()

        grid_response = AgGrid(
            filtered_data,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.AS_INPUT
        )

        # Memperbarui data setelah diedit
        edited_df = grid_response['data']

        # Fungsi untuk menyimpan perubahan ke file excel
        def save_changes_to_excel(df, file_path):
            with pd.ExcelWriter(file_path) as writer:
                df.to_excel(writer, index=False)

        # Tombol untuk menyimpan perubahan
        if st.button("Save Changes"):
            save_changes_to_excel(edited_df, 'updated_data.xlsx')
            st.success("Data telah disimpan")
    else:
        st.error("Kolom dalam file tidak sesuai dengan yang diharapkan.")
