import streamlit as st
import pandas as pd
import io

# Page Configuration
st.set_page_config(
    page_title="Telecom Data Parser & Viewer",
    page_icon="📱",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        background-color: #4f46e5;
        color: white;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-weight: 600;
        border: none;
    }
    .stButton>button:hover {
        background-color: #4338ca;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    st.title("📱 Telecom Data Parser & Viewer")
    st.markdown("Upload your structured data file (Excel or pipe-delimited log matching your specific format) to parse and view records like **MSISDN**, **IMSI**, and **Date/Time**.")

    with st.sidebar:
        st.header("⚙️ Configuration & Reference")
        st.info("Supports Excel files (`.xlsx`, `.xls`) or raw text/CSV pipe-delimited logs.")
        
        st.markdown("---")
        st.markdown("### 📋 Expected Format Pattern")
        st.code("MSISDN|IMSI|MSISDN_abc|MSISDN_DATETIME|...", language="text")
        
        st.markdown("### 💡 Sample Record")
        st.code("3486672316|410060497914995|3486672316_abc|3486672316_20260827162355|...", language="text")

    # File Upload Section
    uploaded_file = st.file_uploader(
        "Upload your Telecom Data File",
        type=["xlsx", "xls", "csv", "txt"],
        help="Upload an Excel sheet or text file containing the pipe-separated schema."
    )

    if uploaded_file is not None:
        try:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            # Read file based on type
            if file_extension in ['xlsx', 'xls']:
                df = pd.read_excel(uploaded_file, header=None)
            elif file_extension in ['csv', 'txt']:
                string_data = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
                df = pd.read_csv(string_data, sep='|', header=None, low_memory=False)
            else:
                st.error("Unsupported file format.")
                return

            # Standard telecom columns mapping based on your format sample
            standard_cols = [
                "MSISDN", 
                "IMSI", 
                "MSISDN_abc", 
                "DATETIME_1", 
                "DATETIME_2", 
                "EMPTY_1", 
                "DATETIME_3", 
                "EMPTY_2", "EMPTY_3", "EMPTY_4", "EMPTY_5", 
                "EMPTY_6", "EMPTY_7", "EMPTY_8", "EMPTY_9", "EMPTY_10"
            ]
            
            # Pad or trim columns to match standard schema length
            if len(df.columns) <= len(standard_cols):
                df.columns = standard_cols[:len(df.columns)]
            else:
                extra_cols = [f"EXTRA_{i}" for i in range(len(df.columns) - len(standard_cols))]
                df.columns = standard_cols + extra_cols

            st.success(f"Successfully loaded and parsed file: **{uploaded_file.name}**")

            # Data Overview metrics
            st.markdown("### 📊 Dataset Overview")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", f"{df.shape[0]:,}")
            with col2:
                st.metric("Total Fields/Columns", f"{df.shape[1]}")
            with col3:
                missing_vals = df.isna().sum().sum()
                st.metric("Missing Values", f"{missing_vals:,}")

            # Data Preview & Filtering
            st.markdown("---")
            st.subheader("🔍 Data Explorer & Search")

            search_query = st.text_input("Search by MSISDN or IMSI", placeholder="e.g. 3486672316 or 410060497914995")
            
            filtered_df = df.copy()
            if search_query:
                mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
                filtered_df = filtered_df[mask]
                st.info(f"Found **{len(filtered_df):,}** matching records.")

            # Display the interactive dataframe highlighting primary target columns first
            st.dataframe(filtered_df, use_container_width=True, height=400)

            # Download processed data option
            st.markdown("---")
            st.subheader("📥 Export Processed Data")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='ParsedTelecomData')
            processed_data = output.getvalue()

            st.download_button(
                label="Download Parsed Data as Excel",
                data=processed_data,
                file_name="parsed_telecom_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing file: {e}")
    else:
        st.warning("Please upload an Excel or pipe-delimited file to begin.")

if __name__ == "__main__":
    main()