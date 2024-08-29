

---

# Google Search Suggestions and Excel Update

This project includes two files to demonstrate how to perform Google searches, extract search suggestions, and update an Excel file with the longest and shortest suggestions. The two files are:

- **Jupyter Notebook (`assignment_1.ipynb`)**: An interactive notebook for exploring and executing the code.
- **Python Script (`assignment_1.py`)**: A standalone script for automated execution.

## Prerequisites

Ensure you have the following installed:

- Python 3.x
- Jupyter Notebook
- Selenium
- OpenPyXL
- ChromeDriver (compatible with your version of Google Chrome)

## Installation

1. **Install Dependencies**:
   Install the required Python packages using pip:

   ```bash
   pip3 install -U selenium openpyxl
   ```

## Usage

### Jupyter Notebook (`assignment_1.ipynb`)

1. **Open the Notebook**:
   - Start Jupyter Notebook and open `assignment_1.ipynb`.

2. **Run the Cells**:
   - Execute the cells sequentially to perform the search and update the Excel file interactively.

### Python Script (`assignment_1.py`)

1. **Prepare Your Excel File**:
   - Place your Excel file at the specified path (e.g., `/Users/shanto/Desktop/4beats/4BeatsQ1.xlsx`).
   - Ensure it has sheets named after the days of the week (e.g., "Monday", "Tuesday", etc.).
   - The sheet should contain columns where the third column has search keywords, and the fourth and fifth columns will be updated with the longest and shortest search suggestions.

2. **Run the Script**:
   - Execute the script in your Python environment:

     ```bash
     python assignment_1.py
     ```

   - Update the `file_path` variable in the script to match the location of your Excel file if necessary.

## Script Overview

1. **Google Search**:
   - Uses Selenium to open Google and perform a search.
   - Collects search suggestions from the dropdown list.

2. **Update Excel File**:
   - Opens the specified Excel file and selects the sheet corresponding to the current day.
   - Searches for keywords in the Excel file, retrieves search suggestions, and updates the longest and shortest suggestions in the appropriate columns.

## Example Output

Both the notebook and the script will print the longest and shortest search suggestions. The Excel file will be updated with these suggestions for the specified day.

```plaintext
Longest suggestion: [Longest suggestion text]
Shortest suggestion: [Shortest suggestion text]
File overwritten successfully!
```

