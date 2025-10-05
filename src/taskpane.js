/* global Excel, Office */

let extractedData = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Extract section event handlers
    document.querySelectorAll('input[name="extractSource"]').forEach(radio => {
      radio.addEventListener('change', handleExtractSourceChange);
    });

    document.getElementById('extractBtn').addEventListener('click', extractData);

    // Insert section event handlers
    document.getElementById('insertLocation').addEventListener('change', handleInsertLocationChange);
    document.getElementById('insertBtn').addEventListener('click', insertData);

    showOutput('Ready! Select an action above.');
  }
});

function handleExtractSourceChange(event) {
  // Hide all option divs
  document.getElementById('sheetOptions').classList.add('hidden');
  document.getElementById('rangeOptions').classList.add('hidden');
  document.getElementById('fileOptions').classList.add('hidden');

  // Show relevant option div
  const source = event.target.value;
  if (source === 'specific') {
    document.getElementById('sheetOptions').classList.remove('hidden');
  } else if (source === 'range') {
    document.getElementById('rangeOptions').classList.remove('hidden');
  } else if (source === 'file') {
    document.getElementById('fileOptions').classList.remove('hidden');
  }
}

function handleInsertLocationChange(event) {
  const location = event.target.value;
  if (location === 'specific') {
    document.getElementById('insertRangeOptions').classList.remove('hidden');
  } else {
    document.getElementById('insertRangeOptions').classList.add('hidden');
  }
}

async function extractData() {
  try {
    const source = document.querySelector('input[name="extractSource"]:checked').value;

    await Excel.run(async (context) => {
      let range;

      switch(source) {
        case 'sheet':
          // Extract from current sheet
          const currentSheet = context.workbook.worksheets.getActiveWorksheet();
          range = currentSheet.getUsedRange();
          break;

        case 'specific':
          // Extract from specific sheet
          const sheetName = document.getElementById('sheetName').value;
          if (!sheetName) {
            showOutput('Error: Please enter a sheet name');
            return;
          }
          const sheet = context.workbook.worksheets.getItem(sheetName);
          range = sheet.getUsedRange();
          break;

        case 'range':
          // Extract from specific range
          const rangeAddress = document.getElementById('rangeAddress').value;
          if (!rangeAddress) {
            showOutput('Error: Please enter a range address');
            return;
          }
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          range = activeSheet.getRange(rangeAddress);
          break;

        case 'file':
          // Handle file upload
          const fileInput = document.getElementById('fileInput');
          if (fileInput.files.length === 0) {
            showOutput('Error: Please select a file');
            return;
          }
          await handleFileExtraction(fileInput.files[0]);
          return;
      }

      range.load('values, address, rowCount, columnCount');
      await context.sync();

      extractedData = range.values;

      // Convert to contextual JSON with headers as keys
      const contextualJSON = convertToContextualJSON(extractedData);

      const summary = `Extracted ${range.rowCount} rows x ${range.columnCount} columns from ${range.address}\n\n` +
                     `Data preview:\n${JSON.stringify(contextualJSON.slice(0, 3), null, 2)}` +
                     (contextualJSON.length > 3 ? '\n...(showing first 3 records)' : '');

      showOutput(summary);

      // Auto-populate insert textarea with contextual JSON
      document.getElementById('insertData').value = JSON.stringify(contextualJSON, null, 2);
    });

  } catch (error) {
    showOutput(`Error: ${error.message}`);
  }
}

async function handleFileExtraction(file) {
  const reader = new FileReader();

  reader.onload = async (e) => {
    try {
      if (file.name.endsWith('.csv')) {
        // Parse CSV
        const text = e.target.result;
        extractedData = parseCSV(text);
        showOutput(`Extracted ${extractedData.length} rows from ${file.name}\n\n` +
                  `Data preview:\n${JSON.stringify(extractedData.slice(0, 5), null, 2)}`);
        document.getElementById('insertData').value = JSON.stringify(extractedData, null, 2);
      } else {
        // For Excel files, we'd need a library like xlsx.js
        showOutput('Excel file parsing requires xlsx library. For now, please use CSV files or copy data directly from Excel.');
      }
    } catch (error) {
      showOutput(`Error parsing file: ${error.message}`);
    }
  };

  if (file.name.endsWith('.csv')) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function parseCSV(text) {
  const lines = text.split('\n');
  return lines.map(line => {
    // Simple CSV parser - handles basic cases
    const values = [];
    let currentValue = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        values.push(currentValue.trim());
        currentValue = '';
      } else {
        currentValue += char;
      }
    }
    values.push(currentValue.trim());

    return values;
  }).filter(row => row.some(cell => cell !== ''));
}

// Convert array-of-arrays to array-of-objects with headers as keys
function convertToContextualJSON(arrayData) {
  if (!arrayData || arrayData.length === 0) return [];

  const headers = arrayData[0];
  const rows = arrayData.slice(1);

  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

// Convert array-of-objects back to array-of-arrays for Excel insertion
function convertFromContextualJSON(objectArray) {
  if (!objectArray || objectArray.length === 0) return [];

  const headers = Object.keys(objectArray[0]);
  const rows = objectArray.map(obj => headers.map(header => obj[header]));

  return [headers, ...rows];
}

async function insertData() {
  try {
    const dataText = document.getElementById('insertData').value;
    if (!dataText) {
      showOutput('Error: No data to insert');
      return;
    }

    let data;
    try {
      data = JSON.parse(dataText);
    } catch (e) {
      showOutput('Error: Invalid JSON format');
      return;
    }

    if (!Array.isArray(data) || data.length === 0) {
      showOutput('Error: Data must be a non-empty array');
      return;
    }

    // Check if data is array of objects (contextual JSON) and convert if needed
    let arrayData;
    if (data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0])) {
      // Convert from contextual JSON to array format
      arrayData = convertFromContextualJSON(data);
    } else {
      // Already in array format
      arrayData = data;
    }

    await Excel.run(async (context) => {
      const location = document.getElementById('insertLocation').value;
      const autoFormat = document.getElementById('autoFormat').checked;
      let targetRange;

      switch(location) {
        case 'current':
          // Insert at current selection
          targetRange = context.workbook.getSelectedRange();
          break;

        case 'newSheet':
          // Create new sheet and insert
          const newSheet = context.workbook.worksheets.add();
          newSheet.activate();
          targetRange = newSheet.getRange('A1').getResizedRange(arrayData.length - 1, arrayData[0].length - 1);
          break;

        case 'specific':
          // Insert at specific range
          const rangeAddr = document.getElementById('insertRange').value;
          if (!rangeAddr) {
            showOutput('Error: Please enter a target range');
            return;
          }
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          targetRange = activeSheet.getRange(rangeAddr).getResizedRange(arrayData.length - 1, arrayData[0].length - 1);
          break;
      }

      targetRange.values = arrayData;

      if (autoFormat) {
        // Format as table
        const table = context.workbook.tables.add(targetRange, true);
        table.name = `DataTable_${Date.now()}`;
        table.style = 'TableStyleMedium2';
      }

      targetRange.format.autofitColumns();
      targetRange.format.autofitRows();

      await context.sync();

      showOutput(`Successfully inserted ${data.length} rows x ${data[0].length} columns` +
                (autoFormat ? ' (formatted as table)' : ''));
    });

  } catch (error) {
    showOutput(`Error: ${error.message}`);
  }
}

function showOutput(message) {
  document.getElementById('output').textContent = message;
}