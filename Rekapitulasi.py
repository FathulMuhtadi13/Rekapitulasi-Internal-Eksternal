import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from datetime import datetime
import io
import xlsxwriter
import os

st.set_page_config(
    page_title="Monitoring Tindak Lanjut Audit ISO Internal & Eksternal - Rekayasa Cakrawala Resources 2024",
    layout="wide"
)

image_path = "logo.png"
col1, col2 = st.columns([1, 5])

with col1:
    st.write("")  # Tambahkan teks kosong untuk memberi jarak
    if os.path.exists(image_path):
        st.image(image_path, use_column_width=True)
    else:
        st.error(f"File tidak ditemukan: {image_path}")

with col2:
    st.markdown(
        """
        <div style='display: flex; flex-direction: column; justify-content: center; height: 100%;'>
            <h2 style='font-size:24px; margin: 0;'>
                Monitoring Tindak Lanjut Audit ISO Internal & Eksternal
            </h2>
            <h2 style='font-size:24px; margin: 0;'>
                Rekayasa Cakrawala Resources 2024
            </h2>
        </div>
        """, 
        unsafe_allow_html=True
    )

# Function to load data from an Excel file
def load_data(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Function to create tornado chart data
def prepare_tornado_chart_data(df):
    findings_grouped = df.groupby(['Finding category', 'Finding Status']).size().reset_index(name='Count')
    return findings_grouped

# Function to create tornado chart
def create_tornado_chart(findings_grouped):
    total_findings = findings_grouped['Count'].sum()
    fig = go.Figure()

    for status in findings_grouped['Finding Status'].unique():
        filtered_df = findings_grouped[findings_grouped['Finding Status'] == status]
        fig.add_trace(go.Bar(
            y=filtered_df['Finding category'],
            x=filtered_df['Count'],
            name=status,
            orientation='h',
            text=filtered_df['Count'],
            textposition='inside',
            textfont=dict(
                size=24,  # Set the size of the text (value) on the chart
                color='black'  # Optional: Set text color for better visibility
            ),
            marker=dict(
                line=dict(
                    color='rgba(255, 255, 255, 0.5)',  # border color
                    width=1
                ),
                opacity=0.8,
                # Add rounded corners to bars (if supported)
                cornerradius=10  # specify the corner radius
            ),
            hoverinfo='text',
        ))

    x_max = findings_grouped['Count'].max() + 5
    for count_value in range(0, int(x_max) + 1, 2):
        fig.add_shape(
            type="line",
            x0=count_value,
            y0=-0.5,
            x1=count_value,
            y1=len(findings_grouped['Finding category']) - 0.5,
            line=dict(color="gray", width=0.5, dash="dash"),
            layer="below"
        )

    fig.update_layout(
        title=dict(text=f'Tornado Chart of Findings - Total Findings: {total_findings}', font=dict(size=25)),
        xaxis_title='Count',
        yaxis_title='Finding Category',
        barmode='stack',
        template='plotly_white',
        showlegend=True,
         legend=dict(
            font=dict(size=25)  # Set the size of the legend text
        ),
        height=600,
        margin=dict(l=30, r=30, t=30, b=30),  # Adjust margins if necessary
        yaxis=dict(
            title_font=dict(size=25),  # Increase font size for Y-axis title
            tickfont=dict(size=25)     # Increase font size for Y-axis labels
        ),
        xaxis=dict(
            tickmode='linear',  # Ensures that ticks are linear
            dtick=5  # Set ticks every 5
        )
    )

    return fig

# Main function for the Streamlit application
def main():
    

    # Custom CSS untuk mengatur gaya judul
    st.markdown(
        """
        <style>
        h1 {
            font-weight: 400 !important; /* Mengatur font-weight menjadi normal */
            font-family: Calibri, sans-serif; /* Mengatur jenis font */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.sidebar.file_uploader("Upload file Excel", type=["xlsx"])
    date_filter = st.sidebar.date_input("Filter by Date Range", [])

    if uploaded_file is None:
        st.info("Silahkan masukan file excel anda")
    else:
        df = load_data(uploaded_file)

        if df is None:
            return
        
        df.columns = df.columns.str.strip()
        required_columns = ['Due Date', 'Tgl Closing', 'Finding category', 'Finding Status', 'Area']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Kolom berikut tidak ditemukan dalam file Excel anda: {missing_columns}")
            return

        st.success("File anda telah terbaca dengan baik")

        df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
        df['Tgl Closing'] = pd.to_datetime(df['Tgl Closing'], errors='coerce')
        df.loc[~df['Tgl Closing'].isna(), 'Finding Status'] = 'Close'

        if len(date_filter) == 2:
            start_date, end_date = date_filter
            df = df[(df['Due Date'] >= pd.to_datetime(start_date)) & (df['Due Date'] <= pd.to_datetime(end_date))]

        finding_category_options = ['All'] + sorted(df['Finding category'].dropna().unique().tolist())
        finding_status_options = ['All'] + sorted(df['Finding Status'].dropna().unique().tolist())

        selected_category = st.sidebar.selectbox("Filter by Finding Category", finding_category_options)
        selected_status = st.sidebar.selectbox("Filter by Finding Status", finding_status_options)

        if selected_category != 'All':
            df = df[df['Finding category'] == selected_category]

        if selected_status != 'All':
            df = df[df['Finding Status'] == selected_status]

        search_input = st.sidebar.text_input("Search Table")

        if search_input:
            df = df[df.apply(lambda row: row.astype(str).str.contains(search_input, case=False).any(), axis=1)]

        findings_grouped = prepare_tornado_chart_data(df)
        fig_tornado = create_tornado_chart(findings_grouped)
        st.plotly_chart(fig_tornado, use_container_width=True)

        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_default_column(editable=True, resizable=True, wrapText=True, autoHeight=True)
        gb.configure_grid_options(pagination=True, paginationPageSize=10, enableFilter=True, rowHeight=40)

        custom_css = {
    ".ag-row": {
        "font-size": "20px !important",
        "font-family": "Calibri, sans-serif !important",
        "white-space": "normal !important",
        "word-wrap": "break-word !important"
    },
    ".ag-row:nth-child(even)": {
        "background-color": "#e8fbff !important"
    },
    ".ag-row:nth-child(odd)": {
        "background-color": "#ffffff !important"
    },
    ".ag-header-cell": {
        "font-size": "24px !important",  # Set the header font size
        "font-family": "Calibri, sans-serif !important",  # Set the header font family
        "white-space": "normal !important",  # Allow the text to wrap
        "word-wrap": "break-word !important",  # Enable word wrapping
        "overflow": "visible !important",  # Ensure the header text is not clipped
        "text-overflow": "ellipsis",  # Optional: use ellipsis for overflow
        "line-height": "1.2 !important",  # Adjust line height for better readability
        "height": "auto !important"  # Allow height to adjust based on content
    },
    ".ag-cell": {
        "white-space": "normal !important",
        "word-wrap": "break-word !important",
        "max-width": "250px",
        "overflow": "hidden",
        "text-overflow": "ellipsis"
    }
}


        grid_options = gb.build()
        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            data_return_mode=DataReturnMode.AS_INPUT,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            height=500,
            custom_css=custom_css,
        )

        updated_df = pd.DataFrame(grid_response['data'])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            updated_df.to_excel(writer, index=False, sheet_name='Updated Data')
            tornado_data = prepare_tornado_chart_data(updated_df)
            tornado_summary = tornado_data.pivot_table(index='Finding category', columns='Finding Status', values='Count', fill_value=0)
            tornado_summary.to_excel(writer, sheet_name='Dashboard', startrow=5)

            dashboard_ws = writer.sheets['Dashboard']
            chart = writer.book.add_chart({'type': 'bar', 'subtype': 'stacked'})
            for status in tornado_data['Finding Status'].unique():
                chart.add_series({
                    'name': status,
                    'categories': f"='Dashboard'!$A$6:$A${6 + len(tornado_summary)}",
                    'values': f"='Dashboard'!${chr(66 + list(tornado_data['Finding Status'].unique()).index(status))}$6:${chr(66 + list(tornado_data['Finding Status'].unique()).index(status))}${6 + len(tornado_summary)}",
                })
            chart.set_title({'name': 'Tornado Chart of Findings'})
            chart.set_x_axis({'name': 'Count'})
            chart.set_y_axis({'name': 'Finding Category'})
            dashboard_ws.insert_chart('E5', chart)

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="rekapitulasi_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
