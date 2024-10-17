import streamlit as st
import pandas as pd
import plotly.express as px

# Function to load the Excel file
def load_data(file):
    try:
        df = pd.read_excel(file, sheet_name='your_sheet_name_here')
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Main Streamlit application
def main():
    st.title("Cost Analysis Dashboard")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        df = load_data(uploaded_file)

        if df is not None:
            # Debug: Print available columns
            st.write("Available columns in the DataFrame:", df.columns.tolist())

            # Strip spaces from column names
            df.columns = df.columns.str.strip()

            # Check if 'DATE' column exists
            if 'DATE' in df.columns:
                # Convert 'DATE' to datetime format
                df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')

                # Proceed with your analysis or visualization
                st.write(df)

                # Example: Plotting a line chart
                fig = px.line(df, x='DATE', y='AMOUNT', title='Amount over Time')
                st.plotly_chart(fig)
            else:
                st.error("The 'DATE' column does not exist in the uploaded file.")

if __name__ == "__main__":
    main()
