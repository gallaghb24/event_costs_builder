# Superdrug ITG Invoice Generator v3.0

A Streamlit application for generating Superdrug ITG invoices from production files and timesheets with full Excel template formatting preservation.

## 🚀 New in Version 3.0

- **Timesheet Integration**: Automatically populate studio hours from timesheet CSV
- **Studio QC Exclusion**: Non-chargeable QC hours automatically excluded
- **Smart Hour Rounding**: All hours rounded up to nearest 0.25
- **Enhanced Workflow**: Seamless progression from data import to invoice generation
- **Improved UI**: Dedicated tabs for each step of the process

## 📋 Features

### Core Functionality
- **Template-based Processing**: Works with your existing Excel invoice template
- **Multiple File Handling**: Combine multiple Production Lines INTERNAL files
- **Automatic Deduplication**: Removes duplicates based on Brief Ref
- **Smart Filtering**: Excludes draft/unapproved items automatically
- **Formula Preservation**: Maintains all Excel formulas and formatting
- **Event Code Generation**: Converts "Event 10 2025" to "E1025" for filenames

### Timesheet Processing
- Excludes Studio QC hours (non-chargeable)
- Rounds all hours up to nearest 0.25
- Matches projects by Project Ref
- Determines work type from charge codes
- Sets Core/OAB based on ROI designation

### Data Mappings
- **Studio Tab**: Job-level aggregation with hours from timesheet
- **Print Tab**: Line-level items with correct field mappings
- **Cost Calculations**: Automatic Core/OAB cost separation

## Installation

1. Install required packages:
```bash
pip install streamlit pandas openpyxl
```

2. Install LibreOffice (for formula recalculation):
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS
brew install libreoffice
```

## Running the Application

```bash
streamlit run invoice_app.py
```

The app will open in your browser at http://localhost:8501

## How to Use

### 1. Upload Template
- Click on the sidebar
- Upload your Superdrug ITG Invoice Template Excel file
- Click "Load Template" to analyze the file

### 2. Enter Event Information
- Enter the event name (e.g., "Event 10 2025")
- The app automatically generates the event code (e.g., "E1025")
- This code is used in the output filename

### 3. Upload Production Files
- Go to the "Production Files" tab
- Upload one or more Production Lines INTERNAL Excel files
- Click "Process Production Files"
- The app will:
  - Combine all uploaded files
  - Remove duplicates based on Brief Ref
  - Filter out draft and unapproved items
  - Prepare data for Studio and Print tabs

### 4. Add Studio Hours
- Go to the "Studio Hours" tab
- For each project, enter:
  - **Studio Hours**: Time spent on the project
  - **Type**: Artwork, Creative Artwork, or Digital
  - **Core/OAB**: Classification as CORE or OAB
- You can either:
  - Edit directly in the table
  - Upload a CSV file with hours data

### 5. Review Data
- Check the "Data Review" tab to verify all information
- Download CSV exports if needed for records

### 6. Preview Costs
- View the "Cost Preview" tab for:
  - Core vs OAB breakdown
  - Studio vs Print costs
  - Project-level breakdown
  - Grand total

### 7. Generate Invoice
- Go to "Generate Invoice" tab
- Click "Generate Invoice" button
- Download the completed Excel file
- Filename will include the event code (e.g., E1025_Superdrug_ITG_Invoice_20250117.xlsx)

## Data Processing Rules

### Production File Requirements
The Production Lines INTERNAL files should have headers in row 2 with these columns:
- Project Ref
- Event Name
- Project Description
- Project Owner
- Brief Ref
- POS Code
- Brief Description
- Part URN
- Part
- Height/Width
- Allocated Qty / Total including Spares
- Material
- Content Brief Status
- Production Supplier Brief Status
- Production Buy Price
- Production Sell Price

### Filtering Rules

**Excluded Production Statuses:**
- draft
- saved
- awaiting rfq
- rfq responses
- estimates awaiting approval
- client approved estimates

**Studio Tab Processing:**
- Aggregated at job level (by Project Ref)
- Includes items where Production Supplier Brief Status = 'not applicable'
- Excludes items where Content Brief Status = 'not applicable'
- Lines count = number of brief refs per project

**Print Tab Processing:**
- Line-level data from production files
- Excludes items where Production Supplier Brief Status = 'not applicable'
- Maps Production Buy Price → Print Cost
- Maps Production Sell Price → Production Cost

### Event Code Generation
- "Event 10 2025" → "E1025"
- "Event 4 2026" → "E0426"
- Format: E + two-digit event number + last two digits of year

## Studio Hours CSV Format

If uploading studio hours via CSV, use this format:

```csv
Project Ref,Studio Hours,Type,Core/OAB
SDG2161,8.5,Artwork,CORE
SDG2210,12,Creative Artwork,OAB
SDG2244,6,Digital,CORE
```

The CSV will be matched to projects by Project Ref.

## Template Structure

The Excel template should contain these worksheets:
- **Event Summary - Core**: Summary of core costs with formulas
- **Event Summary - OAB**: Summary of OAB costs with formulas
- **Studio**: Job-level data with artwork hours
- **Print**: Line-level print items
- **Tags**: Manual entry items
- **Stock, Collate, Pack & Delivery**: Delivery costs
- **Late C&P breakdown**: Manual breakdown items

## Formula Recalculation

The app uses the included `recalc.py` script to recalculate Excel formulas using LibreOffice. This ensures all formulas in the generated invoice are properly calculated.

## Tips

1. **Prepare your data**: Format your data in a spreadsheet first, then copy as CSV
2. **Check Core/OAB assignments**: Ensure each job is correctly classified
3. **Verify costs**: Review the preview before generating the final invoice
4. **Save templates**: Keep your template file updated with current client names
5. **Batch processing**: Process multiple jobs at once by pasting all data

## Troubleshooting

**Template won't load:**
- Ensure the Excel file is not corrupted
- Check that all required sheets are present

**Formula errors:**
- Install LibreOffice for formula recalculation
- Check that all referenced cells exist

**Mis
## 🛠️ Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup Steps

1. **Clone or download the files**
```bash
mkdir superdrug-invoice-app
cd superdrug-invoice-app
```

2. **Create a virtual environment (recommended)**
```bash
python -m venv venv

# On Windows:
venv\Scripts\activate

# On Mac/Linux:
source venv/bin/activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run the application**
```bash
streamlit run invoice_app_v3.py
```

The app will open in your browser at http://localhost:8501

## 📁 File Structure

```
superdrug-invoice-app/
│
├── invoice_app_v3.py       # Main application file
├── requirements.txt        # Python dependencies
├── README.md              # This file
└── templates/             # Folder for Excel templates (optional)
    └── Exxxx_Superdrug_ITG_Invoice_Template.xlsx
```

## 📊 Usage Guide

### Step 1: Upload Template
1. Click on the sidebar
2. Upload your Superdrug ITG Invoice Template Excel file
3. Click "Load Template"
4. Verify the template info shows Core/OAB clients

### Step 2: Process Production Files
1. Go to "Production Files" tab
2. Upload one or more Production Lines INTERNAL Excel files
3. Click "Process Production Files"
4. The app will:
   - Combine all files
   - Remove duplicates based on Brief Ref
   - Filter out draft/unapproved items
   - Create Studio and Print data

### Step 3: Import Timesheet Data
1. Go to "Timesheet Data" tab
2. Upload your timesheet CSV export
3. Click "Process Timesheet"
4. The app will:
   - Exclude Studio QC hours
   - Round hours up to nearest 0.25
   - Match projects and populate hours
   - Set types and Core/OAB classifications

### Step 4: Review Studio Hours
1. Go to "Studio Hours Review" tab
2. Review auto-populated data
3. Edit any values as needed:
   - Adjust hours (in 0.25 increments)
   - Change Type (Artwork/Creative Artwork/Digital)
   - Modify Core/OAB assignments
4. Changes are saved automatically

### Step 5: Preview Costs
1. Go to "Cost Preview" tab
2. Review breakdown:
   - Core vs OAB costs
   - Studio vs Production costs
   - Project-level breakdown
3. Verify totals before proceeding

### Step 6: Generate Invoice
1. Go to "Generate Invoice" tab
2. Click "Generate Invoice" button
3. Download the completed Excel file
4. Filename format: `E1025_Superdrug_ITG_Invoice_20250922.xlsx`

## 📝 Data Processing Rules

### Production File Filtering
**Excluded Production Statuses:**
- draft
- saved
- awaiting rfq
- rfq responses
- estimates awaiting approval
- client approved estimates

### Studio Tab Processing
- Aggregated at job level (by Project Ref)
- Includes items where Production Supplier Brief Status = 'not applicable'
- Excludes items where Content Brief Status = 'not applicable'
- Lines count = number of brief refs per project

### Print Tab Processing
- Line-level data from production files
- Excludes items where Production Supplier Brief Status = 'not applicable'
- Maps Production Sell Price to Production Cost column

### Timesheet Processing
- **Excluded**: Any rows with Charge Code containing 'QC' or 'Studio QC'
- **Rounding**: All aggregated hours rounded UP to nearest 0.25
  - 1.1 hours → 1.25 hours
  - 2.26 hours → 2.5 hours
  - 3.0 hours → 3.0 hours
- **Type Determination**:
  - Contains 'creative' → Creative Artwork
  - Contains 'digital' or 'tec' → Digital
  - Default → Artwork
- **Core/OAB Assignment**:
  - Description contains 'ROI' → OAB
  - Default → CORE

## 📑 Input File Requirements

### Production Lines INTERNAL Files
- Excel format (.xlsx)
- Headers in row 2
- Required columns:
  - Project Ref
  - Event Name
  - Project Description
  - Project Owner
  - Brief Ref
  - Production Supplier Brief Status
  - Content Brief Status
  - Production Sell Price
  - Total including Spares

### Timesheet CSV
- CSV format
- Required columns:
  - Job Number (format: 1/SDGxxxx)
  - Job Description
  - Total (hours)
  - Charge Code
  - Resource Name

### Template Excel File
- Must contain these sheets:
  - Event Summary - Core
  - Event Summary - OAB
  - Studio
  - Print
  - Tags
  - Stock, Collate, Pack & Delivery
  - Late C&P breakdown

## 🎯 Field Mappings

### Print Tab Output Columns
| Column | Field | Source |
|--------|-------|--------|
| A | Project Ref | Project Ref |
| B | Event Name | Event Name |
| C | Project Description | Project Description |
| D | Project Owner | Project Owner |
| E | Brief Ref | Brief Ref |
| F | POS Code | POS Code |
| G | Brief Description | Brief Description |
| H | Part URN | Part URN |
| I | Part | Part |
| J | Height | Height |
| K | Width | Width |
| L | Colours Front | Colours Front |
| M | Colours Back | Colours Back |
| N | Material | Material |
| O | No of Pages | No of Pages |
| P | Production Finishing Notes | Production Finishing Notes |
| Q | Production Supplier Comments | Production Supplier Comments |
| R | Allocated Qty | Allocated Qty |
| S | Spares | Spares |
| T | Total including Spares | Total including Spares |
| U | No of Stores | No of Stores |
| V | In Store Deadline | In Store Deadline |
| W | Content Brief Status | Content Brief Status |
| X | Production Supplier Brief Status | Production Supplier Brief Status |
| Y | Production Cost | Production Sell Price |
| Z | Core/OAB | Formula (VLOOKUP) |

## 🔧 Troubleshooting

### Template won't load
- Ensure the Excel file is not corrupted
- Check that all required sheets are present
- Try opening and re-saving the template

### No hours appearing from timesheet
- Check Job Number format (should be 1/SDGxxxx)
- Verify projects match between production files and timesheet
- Ensure Charge Code doesn't contain 'QC'

### Formula errors in output
- Verify all Project Refs match between Studio and Print tabs
- Check that Core/OAB assignments are complete
- Ensure template formulas are intact

### High production costs
- Check quantities in Total including Spares column
- Verify Production Sell Price values are reasonable
- Some items may have very large quantities

## 💡 Tips

1. **Prepare your data**: Clean your production files and timesheet before uploading
2. **Check project matches**: Ensure Project Refs are consistent across all files
3. **Review ROI projects**: Verify OAB assignments for ROI-designated projects
4. **Save your work**: Download CSV exports at each stage for records
5. **Template updates**: Keep your template file updated with current client names and rates

## 🔄 Version History

- **v3.0** (Current): Added timesheet integration, QC exclusion, hour rounding
- **v2.0**: Production file processing, format preservation
- **v1.0**: Basic invoice generation with manual data entry

---

Built with ❤️ using Streamlit for efficient invoice management
