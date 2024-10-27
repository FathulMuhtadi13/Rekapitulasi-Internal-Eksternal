import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from PIL import Image
import os
from st_aggrid import AgGrid, GridOptionsBuilder, ColumnsAutoSizeMode

# Set page configuration
st.set_page_config(
    page_title="Monitoring Tindak Lanjut Audit ISO Internal & Eksternal RECARE 2024",
    layout="wide"  
)

# Set file path for logo
image_path = os.path.join("logo.png")

# CSS Styling to mimic Shadcn UI inspired look
st.markdown(
    """
    <style>
        /* All existing CSS styles */
        /* Additional custom styles */
    </style>
    """,
    unsafe_allow_html=True
)

# Sidebar file uploader
st.sidebar.header("Upload Excel File")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx"])

def prepare_sankey_chart_data(df):
    areas = df['Area'].unique()
    categories = df['Finding category'].unique()
    statuses = df['Finding Status'].unique()

    labels = list(areas) + list(categories) + list(statuses)
    sources = []
    targets = []
    values = []

    for area in areas:
        for category in categories:
            count = df[(df['Area'] == area) & (df['Finding category'] == category)].shape[0]
            if count > 0:
                sources.append(labels.index(area))
                targets.append(labels.index(category))
                values.append(count)

    for category in categories:
        for status in statuses:
            count = df[(df['Finding category'] == category) & (df['Finding Status'] == status)].shape[0]
            if count > 0:
                sources.append(labels.index(category))
                targets.append(labels.index(status))
                values.append(count)

    return labels, sources, targets, values

def create_sankey_chart(labels, sources, targets, values):
    link_colors = []
    for target in targets:
        if labels[target] == "Open":
            link_colors.append("lightpink")
        elif labels[target] == "Close":
            link_colors.append("lightgreen")
        else:
            link_colors.append("gray")

    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=15,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=labels,
            color=['#66b3ff'] * len(labels)
        ),
        link=dict(
            source=sources,
            target=targets,
            value=values,
            color=link_colors,
            hovertemplate='Source: %{source.label}<br />Target: %{target.label}<br />Count: %{value}<extra></extra>'
        ))])

    fig.update_layout(
        title_text="Sankey Diagram of Findings",
        font_size=24,
        height=600,
        paper_bgcolor="white",
        plot_bgcolor="white"
    )

    return fig

# Display content if file is uploaded
if uploaded_file is not None:
    col1, col2 = st.columns([1, 5])
    with col1:
        try:
            image = Image.open(image_path)
            st.image(image, use_column_width=True, caption="")
        except FileNotFoundError:
            st.error(f"File not found: {image_path}")
    with col2:
        st.markdown("<h1 class='header'>Monitoring Tinlan Audit ISO Internal & Eksternal RECARE</h1>", unsafe_allow_html=True)

    try:
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")

        required_columns = ["Finding category", "Finding Status", "Area"]
        if not all(column in df.columns for column in required_columns):
            st.error(f"Excel file must contain columns: {', '.join(required_columns)}.")
        else:
            # Sidebar filters
            search_term = st.sidebar.text_input("Search Findings")
            selected_category = st.sidebar.multiselect("Filter by Category", options=df["Finding category"].unique())
            selected_status = st.sidebar.multiselect("Filter by Status", options=df["Finding Status"].unique())
            selected_area = st.sidebar.multiselect("Filter by Area", options=df["Area"].unique())

            # Apply filters
            filtered_df = df.copy()
            if search_term:
                filtered_df = filtered_df[filtered_df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
            if selected_category:
                filtered_df = filtered_df[filtered_df["Finding category"].isin(selected_category)]
            if selected_status:
                filtered_df = filtered_df[filtered_df["Finding Status"].isin(selected_status)]
            if selected_area:
                filtered_df = filtered_df[filtered_df["Area"].isin(selected_area)]

            # CSS Styling for modern cards
            st.markdown(
                """
                <style>
                    .card {
                        background: #fff;
                        border-radius: 16px;
                        padding: 20px;
                        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
                        transition: transform 0.2s, box-shadow 0.2s;
                    }
                    .card:hover {
                        transform: translateY(-8px);
                        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.15);
                    }
                    .card h3 {
                        font-size: 1.25rem;
                        font-weight: 600;
                        margin-bottom: 0.5rem;
                        color: #111827;
                    }
                    .card p {
                        font-size: 2rem;
                        font-weight: 700;
                        margin: 0.25rem 0;
                        color: #3B82F6;
                    }
                    .card span {
                        display: inline-block;
                        font-size: 0.9rem;
                        color: #6B7280;
                        margin-top: 0.5rem;
                    }
                </style>
                """,
                unsafe_allow_html=True
            )

            # Display card content using columns
            cols = st.columns(3)
            total_ofi = filtered_df[filtered_df['Finding category'] == 'OFI'].shape[0]
            total_ob = filtered_df[filtered_df['Finding category'] == 'OB'].shape[0]
            total_nc = filtered_df[filtered_df['Finding category'] == 'NC'].shape[0]

            with cols[0]:
                st.markdown(
                    f"""
                    <div class="card">
                        <h3 style="font-size: 30px;">Total NC</h3>
                        <p style="font-size: 50px;">{total_nc}</p>
                        <span></span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with cols[1]:
                st.markdown(
                    f"""
                    <div class="card">
                        <h3 style="font-size: 30px;">Total OB</h3>
                        <p style="font-size: 50px;">{total_ob}</p>
                        <span></span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with cols[2]:
                st.markdown(
                    f"""
                    <div class="card">
                        <h3 style="font-size: 30px;">Total OFI</h3>
                        <p style="font-size: 50px;">{total_ofi}</p>
                        <span></span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            labels, sources, targets, values = prepare_sankey_chart_data(filtered_df)
            sankey_chart = create_sankey_chart(labels, sources, targets, values)
            st.plotly_chart(sankey_chart, use_container_width=True)

            # Custom CSS for the AgGrid
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

            # Set up the Streamlit app and custom styles
            st.markdown(
                """
                <style>
                * {
                    box-sizing: border-box;
                    -webkit-box-sizing: border-box;
                    -moz-box-sizing: border-box;
                }
                body {
                    font-family: Helvetica;
                    -webkit-font-smoothing: antialiased;
                    background: rgba(71, 147, 227, 1);
                }
                h2 {
                    text-align: center;
                    font-size: 18px;
                    text-transform: uppercase;
                    letter-spacing: 2px;
                    color: white;
                    padding: 40px 0;
                }
                </style>
                """,
                unsafe_allow_html=True
            )

            
            # Configure AgGrid with the DataFrame
            grid_options = GridOptionsBuilder.from_dataframe(filtered_df)

            # Limit column width to 50px to achieve a condensed column width
            for col in filtered_df.columns:
                grid_options.configure_column(col, width=130, wrapText=True)

            # Set row height
            grid_options.configure_grid_options(rowHeight=100)  # Set a specific row height (in pixels)

            # Enable pagination and auto-size the columns to fit the contents
            grid_options.configure_pagination(paginationAutoPageSize=True)

            # Set grid options for displaying the DataFrame in AgGrid with the CSS applied
            AgGrid(
                filtered_df, 
                gridOptions=grid_options.build(), 
                columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
                theme='streamlit',  # Uses the CSS styles defined earlier
                custom_css=custom_css,  # Apply custom CSS styles here
            )

    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
else:
    st.info("Please upload an Excel file to display the content.")
