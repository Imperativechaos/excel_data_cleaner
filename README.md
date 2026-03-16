# Python Automation Demos

Collection of practical Python scripts that automate boring repetitive tasks — especially Excel data cleaning, report generation, email sending, and workflow hacks.

These are live demos of what I build for clients on Fiverr. Each script is clean, commented, and ready to run or customize.

## Scripts Included

### 1. Excel Data Cleaner & Professional Formatter  
**File:** [excel_data_cleaner.py](excel_data_cleaner.py)

Cleans messy Excel files and turns them into polished, business-ready spreadsheets.

**Features:**
- Removes fully empty rows and columns
- Strips and title-cases column headers
- Fills missing values with "N/A" (customizable)
- Applies professional formatting:
  - Bold + white text on blue header row
  - Centered headers
  - Auto-adjusts column widths
- Uses **pandas** for data wrangling + **openpyxl** for styling

**Before / After Example**

(Add screenshots here — drag & drop images into repo or use imgur links)

Before: messy spreadsheet with junk headers, blanks, uneven columns  
After: clean table, colored headers, perfect widths

**Quick Run Example**

```python
from excel_data_cleaner import clean_and_format_excel

clean_and_format_excel("your_messy_file.xlsx")
# → Outputs cleaned_your_messy_file.xlsx

#install requirements

pip install pandas openpyxl
