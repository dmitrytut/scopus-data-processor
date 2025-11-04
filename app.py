"""
Streamlit application for processing Scopus data
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

from config import (
    DEFAULT_FUZZY_MATCH_THRESHOLD,
    AFFILIATION_KEYWORDS,
    AFFILIATION_EXCLUDE_KEYWORDS,
    DEFAULT_TITLE_EXCLUDE_KEYWORDS,
    HIGHLIGHT_COLOR_MULTIPLE_DEPTS,
    DEFAULT_UNITED_SHEET_NAME
)
from utils import process_scopus_data, export_to_excel_with_highlighting


# Page configuration
st.set_page_config(
    page_title="Scopus Data Processor",
    page_icon="📊",
    layout="wide"
)

# Header
st.title("📊 Scopus Data Processor")
st.markdown("---")

# Initialize session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'stats' not in st.session_state:
    st.session_state.stats = None
if 'output_buffer' not in st.session_state:
    st.session_state.output_buffer = None

# ============== SIDEBAR ==============
with st.sidebar:
    # Process button
    process_button = st.button(
        "🚀 Process Data",
        type="primary",
        use_container_width=True
    )

    st.markdown("---")

    st.header("⚙️ Settings")

    st.subheader("📁 File Upload")

    # Scopus file upload
    scopus_file = st.file_uploader(
        "Scopus Export File",
        type=['xlsx', 'xls', 'csv'],
        help="Upload Scopus export file. Formats: xlsx, xls, csv"
    )

    # United file upload
    united_file = st.file_uploader(
        "United Database File",
        type=['xlsx', 'xls', 'csv'],
        help="Upload file with existing articles. Formats: xlsx, xls, csv"
    )

    # United sheet name
    united_sheet_name = st.text_input(
        "Sheet Name in United File",
        value=DEFAULT_UNITED_SHEET_NAME,
        help="Specify the sheet name to read from United file"
    )

    # Department mapping file
    dept_file = st.file_uploader(
        "Department Mapping File",
        type=['xlsx', 'xls'],
        help="Upload file with author-department mapping. Columns: 'Author Name' (short or full format), 'Departament'"
    )

    st.markdown("---")
    st.subheader("🔧 Processing Parameters")

    # Affiliation search settings
    affiliation_keywords = []
    affiliation_keywords_text = st.text_area(
        "Affiliation Keywords (one per line)",
        value='\n'.join(AFFILIATION_KEYWORDS),
        height=100,
        help="Only authors from these affiliations will be included in the `Authors` column"
    )

    if affiliation_keywords_text.strip():
        affiliation_keywords = [
            line.strip() for line in affiliation_keywords_text.split('\n')
            if line.strip()
        ]

    # Affiliation exclusion settings
    affiliation_exclude_keywords = []
    affiliation_exclude_keywords_text = st.text_area(
        "Affiliation Exclusion Keywords (one per line)",
        value='\n'.join(AFFILIATION_EXCLUDE_KEYWORDS),
        height=100,
        help="Authors from affiliations containing these keywords will be EXCLUDED from results"
    )

    if affiliation_exclude_keywords_text.strip():
        affiliation_exclude_keywords = [
            line.strip() for line in affiliation_exclude_keywords_text.split('\n')
            if line.strip()
        ]

    # Year filtering
    st.markdown("**Year Filtering**")
    year_filter_enabled = st.checkbox("Enable year filtering", value=True)

    selected_years = None
    if year_filter_enabled:
        # Get year range for selection
        current_year = datetime.now().year
        year_options = list(range(2015, current_year + 2))

        selected_years = st.multiselect(
            "Select Year(s)",
            options=year_options,
            default=[current_year],
            help="Select one or more years to filter"
        )

        # Convert to appropriate format
        if selected_years:
            if len(selected_years) == 1:
                selected_years = selected_years[0]
            # otherwise keep as list

    # Title filtering
    st.markdown("**Title Exclusion**")
    title_filter_enabled = st.checkbox(
        "Exclude articles with specific substrings",
        value=True
    )

    title_exclude_keywords = []
    if title_filter_enabled:
        title_exclude_text = st.text_area(
            "Substrings to Exclude (one per line)",
            value='\n'.join(DEFAULT_TITLE_EXCLUDE_KEYWORDS),
            height=100,
            help="Articles containing these substrings in Title will be excluded"
        )

        if title_exclude_text.strip():
            title_exclude_keywords = [
                line.strip() for line in title_exclude_text.split('\n')
                if line.strip()
            ]

    # Fuzzy matching threshold
    fuzzy_threshold = st.slider(
        "Fuzzy Matching Threshold (%)",
        min_value=80,
        max_value=100,
        value=DEFAULT_FUZZY_MATCH_THRESHOLD,
        help="Similarity threshold for duplicate detection (95%+ recommended)"
    )

# ============== MAIN AREA ==============

# Check if all files are uploaded
if process_button:
    if not scopus_file:
        st.error("❌ Please upload Scopus Export file")
    elif not united_file:
        st.error("❌ Please upload United database file")
    else:
        # Load data
        with st.spinner("📂 Loading files..."):
            try:
                # Load Scopus
                df_scopus = pd.read_excel(scopus_file)

                # Load United
                df_united = pd.read_excel(united_file, sheet_name=united_sheet_name)

                # Load departments (if empty, create empty dataframe)
                df_departments = pd.read_excel(dept_file) if dept_file else pd.DataFrame(columns=['Author Name', 'Departament'])

                st.success(f"✅ Files loaded successfully")
                st.info(f"📊 Scopus: {len(df_scopus)} articles | United: {len(df_united)} articles | Departments: {len(df_departments)} records")

            except Exception as e:
                st.error(f"❌ Error loading files: {str(e)}")
                st.stop()

        # Process data
        with st.spinner("⚙️ Processing data..."):
            try:
                result_df, stats = process_scopus_data(
                    df_scopus,
                    df_united,
                    df_departments,
                    threshold=fuzzy_threshold,
                    year=selected_years if year_filter_enabled else None,
                    title_exclude_keywords=title_exclude_keywords if title_filter_enabled else None,
                    affiliation_keywords=affiliation_keywords,
                    affiliation_exclude_keywords=affiliation_exclude_keywords
                )

                # Save to session state
                st.session_state.result_df = result_df
                st.session_state.stats = stats
                st.session_state.processed = True

            except Exception as e:
                st.error(f"❌ Error processing data: {str(e)}")
                st.stop()

# Display results
if st.session_state.processed and st.session_state.result_df is not None:
    st.markdown("---")
    st.header("📈 Processing Results")

    stats = st.session_state.stats
    result_df = st.session_state.result_df

    # Statistics in columns
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Source Articles (Scopus)", stats['original_scopus_count'])
        if stats.get('after_year_filter_scopus', 0) != stats['original_scopus_count']:
            st.caption(f"After year filter: {stats['after_year_filter_scopus']}")

    with col2:
        st.metric("New Articles Found", stats['new_articles'])
        st.caption(f"Duplicates: {stats['duplicates_found']}")

    with col3:
        st.metric("Articles with Affiliated Authors", stats['affiliated_articles'])
        if stats['no_affiliated_authors'] > 0:
            st.caption(f"Without affiliated authors: {stats['no_affiliated_authors']}")

    with col4:
        st.metric("Require Review", stats['highlighted_depts'])
        st.caption("Departments not found / multiple")

    # Additional statistics
    if stats.get('excluded_by_title', 0) > 0:
        st.info(f"📌 Excluded by Title: {stats['excluded_by_title']} articles")

    st.markdown("---")

    # Preview results
    st.subheader("👀 Preview")
    st.dataframe(
        result_df.drop(columns=['_highlight', '_highlight_reason'], errors='ignore').head(10),
        use_container_width=True
    )

    st.caption(f"Showing first 10 of {len(result_df)} rows")

    # Export to Excel
    st.markdown("---")
    st.subheader("💾 Download Results")

    # Generate filename
    current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

    if selected_years:
        if isinstance(selected_years, int):
            year_str = str(selected_years)
        else:
            year_str = '-'.join(map(str, sorted(selected_years)))
    else:
        year_str = 'all_years'

    output_filename = f"new_articles_{year_str}_{current_date}.xlsx"

    # Create Excel file in memory
    output_buffer = io.BytesIO()

    # Use our export function
    temp_path = f"temp_{output_filename}"
    success = export_to_excel_with_highlighting(
        result_df,
        temp_path,
        highlight_color=HIGHLIGHT_COLOR_MULTIPLE_DEPTS
    )

    if success:
        # Read file and load to buffer
        with open(temp_path, 'rb') as f:
            output_buffer.write(f.read())

        # Remove temporary file
        os.remove(temp_path)

        # Download button
        st.download_button(
            label="📥 Download Excel with Results",
            data=output_buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.success("✅ File ready for download!")
        st.info("💡 Yellow highlighted cells in 'Departament' column require manual review")
    else:
        st.error("❌ Error creating Excel file")

else:
    # Instructions if not yet processed
    st.info("👈 Upload files and configure settings in the sidebar, then click 'Process Data'")

    st.markdown("### 📖 Instructions")
    st.markdown("""
    1. **Upload files:**
       - Scopus Export (source data from Scopus)
       - United database (existing articles)
       - Department mapping file (Author Name → Department)

    2. **Configure parameters:**
       - Affiliation keywords (which institutions to search for)
       - Year filter (optional)
       - Title exclusion keywords (Correction, Erratum, etc.)
       - Fuzzy Matching threshold (95%+ recommended)

    3. **Click 'Process Data'**

    4. **Download results** in Excel format

    **Output:**
    - New articles (not found in United database)
    - Only articles with affiliated authors
    - Automatic department assignment
    - Highlighting for cells requiring manual review
    """)
