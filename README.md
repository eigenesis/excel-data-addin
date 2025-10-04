# Excel Data Extract & Insert Add-in

An Office Excel add-in that allows you to extract data from sheets/files and insert data automatically.

## Features

### 📤 Extract Data
- **Current Sheet**: Extract all data from the active worksheet
- **Specific Sheet**: Extract data from a named sheet
- **Range**: Extract data from a specific range (e.g., A1:C10)
- **File**: Import data from external CSV files

### 📥 Insert Data
- **Current Selection**: Insert data at the current selected cell
- **New Sheet**: Create a new sheet and insert data
- **Specific Range**: Insert at a specific range address
- **Auto-format**: Optionally format inserted data as a table

## VBA Installation (Recommended for Mac)

Since Office add-in sideloading has limited support on Excel for Mac, use the VBA macro version instead:

1. **Open Excel**
2. **Enable Developer Tab** (if not visible):
   - Go to **Excel** > **Preferences** > **Ribbon & Toolbar**
   - Check "Developer" in the right column
3. **Open Visual Basic Editor**:
   - Click **Developer** tab > **Visual Basic** (or press Alt+F11)
4. **Import the Module**:
   - In VBA Editor: **File** > **Import File**
   - Navigate to and select `DataExtractInsert.bas`
5. **Add a Button (Optional)**:
   - Go back to Excel
   - **Developer** > **Insert** > **Button (Form Control)**
   - Draw the button on your sheet
   - Assign macro: `ShowDataToolsMenu`
   - Right-click button > **Edit Text** and rename to "Data Tools"

### Usage

Run the main menu: **Developer** > **Macros** > **ShowDataToolsMenu** > **Run**

Or click the button if you created one.

---

## Web Add-in Setup (macOS - Advanced)

1. Install dependencies:
```bash
npm install
```

2. Start the development server:
```bash
npm run serve
```

3. Manually sideload the add-in in Excel:
   - Open Excel for Mac
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **+ Add a Custom Add-in** > **Add from File**
   - Navigate to and select `manifest.xml` from this folder
   - Click **OK**

4. The "Data Tools" button should appear in the Home tab ribbon

## Usage

1. Click the "Show Data Tools" button in the Home tab
2. The task pane will open on the right side
3. **To Extract:**
   - Select your data source (sheet, range, or file)
   - Click "Extract Data"
   - The data will be shown in the output area and auto-populated in the insert section

4. **To Insert:**
   - Paste or edit JSON data in the textarea
   - Choose where to insert the data
   - Optionally enable auto-formatting
   - Click "Insert Data"

## Data Format

Data should be in JSON array format:
```json
[
  ["Header1", "Header2", "Header3"],
  ["Value1", "Value2", "Value3"],
  ["Value4", "Value5", "Value6"]
]
```

## File Structure

```
excel-data-addin/
├── manifest.xml           # Add-in configuration
├── package.json          # NPM dependencies
├── README.md            # This file
└── src/
    ├── taskpane.html    # UI layout
    └── taskpane.js      # Main functionality
```

## Notes

- CSV file import is supported with basic parsing
- Excel file import (.xlsx) requires additional libraries
- The add-in requires Office.js API
- Auto-formatting creates styled tables automatically