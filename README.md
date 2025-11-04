# Scopus Data Processor

A web application for processing Scopus export data for the given Affiliation. Automatically identifies new articles, extracts affiliated authors, and maps them to departments.

## Features

- ✓ **Automated duplicate detection** using fuzzy matching
- ✓ **Smart filtering** by year and title keywords
- ✓ **Author extraction** - identifies affiliated authors from co-authored papers
- ✓ **Department mapping** - auto-assigns departments based on author database
- ✓ **Excel export** with visual highlighting for manual review items
- ✓ **Cross-platform** - works on Windows, macOS, and Linux

## Quick Start

### Prerequisites

- Python 3.9 or higher

### Installation & Launch

1. **Install dependencies:**
   ```bash
   # Windows
   pip install -r requirements.txt

   # macOS/Linux
   pip3 install -r requirements.txt
   ```

2. **Run the application:**
   ```bash
   # Windows - double-click or run
   run_app.bat

   # macOS/Linux
   ./run_app.sh

   # Or directly with streamlit
   streamlit run app.py
   ```

3. **Access the app** - opens automatically at `http://localhost:8501`

## Usage

1. **Upload files** in the sidebar:
   - Scopus Export file (source data)
   - United database file (existing articles)
   - Department mapping file (author -> department)

2. **Configure settings:**
   - Year filter (single or multiple years)
   - Title exclusion keywords (e.g., "Correction", "Erratum")
   - Fuzzy matching threshold (recommended: 95-98%)

3. **Process data** - click "Process Data" button

4. **Download results** - Excel file with highlighted cells requiring manual review

## File Requirements

### Scopus Export File
- **Format:** Excel (.xlsx)
- **Required columns:** Authors, Author full names, Authors with affiliations, Title, Year, Source title

### United Database File
- **Format:** Excel (.xlsx)
- **Required columns:** Title, Year, Authors
- **Sheet name:** Configurable (default: "Last")

### Department Mapping File
- **Format:** Excel (.xlsx)
- **Columns:**
  - `Author Name` - format: "LastName, F."
  - `Departament` - department name

## Output

The application produces an Excel file with:
- Only NEW articles (not found in United database)
- Only articles with affiliated authors
- Automatic department assignment
- Highlighting for:
  - Authors not found in department mapping
  - Multiple departments detected

## Tech Stack

- **[Streamlit](https://streamlit.io/)** - Web interface
- **[pandas](https://pandas.pydata.org/)** - Data processing
- **[fuzzywuzzy](https://github.com/seatgeek/fuzzywuzzy)** - Fuzzy string matching
- **[openpyxl](https://openpyxl.readthedocs.io/)** - Excel file handling

## Project Structure

```
./
jupyter/scopus_processor.ipynb      # Original Jupyter notebook
app.py                              # Main Streamlit application
utils.py                            # Data processing functions
config.py                           # Default configuration
requirements.txt                    # Python dependencies
run_app.bat                         # Windows launcher
run_app.sh                          # macOS/Linux launcher
README.md                           # This file
```

## Alternative: Jupyter Notebook

For users familiar with Python, the original Jupyter notebook is available in `scopus_processor.ipynb`.

Run with:
```bash
jupyter notebook ./jupyter/scopus_processor.ipynb
```