/* global Excel, Office */

let extractedData = null;
let extractedJSON = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Tab switching
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', switchTab);
    });

    // Extract section event handlers
    document.querySelectorAll('input[name="extractSource"]').forEach(radio => {
      radio.addEventListener('change', handleExtractSourceChange);
    });

    document.getElementById('extractBtn').addEventListener('click', extractData);
    document.getElementById('copyBtn').addEventListener('click', copyToClipboard);
    document.getElementById('goToInsertBtn').addEventListener('click', () => switchTab({target: {dataset: {tab: 'insert'}}}));

    // Insert section event handlers
    document.getElementById('insertLocation').addEventListener('change', handleInsertLocationChange);
    document.getElementById('insertBtn').addEventListener('click', insertData);

    // Auto-detect range when "Range" option is selected
    setupRangeAutoDetection();

    // Fraud detection event handlers
    document.getElementById('fraudEnv').addEventListener('change', handleEnvChange);
    document.getElementById('runFraudCheckBtn').addEventListener('click', runFraudDetection);
    document.getElementById('clearSettingsBtn').addEventListener('click', clearFraudSettings);

    // Load saved API settings
    loadFraudSettings();

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
      extractedJSON = convertToContextualJSON(extractedData);

      // Show result in extract tab
      document.getElementById('extractedDataBox').textContent = JSON.stringify(extractedJSON, null, 2);
      document.getElementById('extractResult').classList.remove('hidden');

      // Auto-populate insert textarea with contextual JSON
      document.getElementById('insertData').value = JSON.stringify(extractedJSON, null, 2);

      showOutput(`✓ Extracted ${range.rowCount} rows x ${range.columnCount} columns from ${range.address}`);
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

// Tab switching
function switchTab(event) {
  const tabName = event.target.dataset.tab;

  // Update tab buttons
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  event.target.classList.add('active');

  // Update tab content
  document.querySelectorAll('.tab-content').forEach(content => {
    content.classList.remove('active');
  });
  document.getElementById(`${tabName}Tab`).classList.add('active');
}

// Copy extracted data to clipboard
function copyToClipboard() {
  const dataText = document.getElementById('insertData').value;

  navigator.clipboard.writeText(dataText).then(() => {
    const btn = document.getElementById('copyBtn');
    const originalText = btn.textContent;
    btn.textContent = '✓ Copied!';
    btn.style.background = '#38a169';

    setTimeout(() => {
      btn.textContent = originalText;
      btn.style.background = '';
    }, 2000);
  }).catch(err => {
    showOutput(`Failed to copy: ${err.message}`);
  });
}

// Auto-detect selected range
function setupRangeAutoDetection() {
  document.getElementById('opt-range').addEventListener('change', async () => {
    if (document.getElementById('opt-range').checked) {
      await detectSelectedRange();
    }
  });
}

async function detectSelectedRange() {
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load('address');
      await context.sync();

      // Auto-fill the range input
      document.getElementById('rangeAddress').value = selectedRange.address.split('!')[1] || selectedRange.address;

      // Show auto-detected info
      document.getElementById('detectedRange').textContent = selectedRange.address;
      document.getElementById('autoRangeInfo').classList.remove('hidden');
    });
  } catch (error) {
    console.error('Could not detect range:', error);
  }
}
// Fraud Detection Functions
function handleEnvChange() {
  const env = document.getElementById('fraudEnv').value;
  const customEnvGroup = document.getElementById('customEnvGroup');
  
  if (env === 'custom') {
    customEnvGroup.classList.remove('hidden');
  } else {
    customEnvGroup.classList.add('hidden');
  }
}

function loadFraudSettings() {
  const apiKey = localStorage.getItem('fraud_api_key');
  const env = localStorage.getItem('fraud_env');
  const customUrl = localStorage.getItem('fraud_custom_url');
  const proxyUrl = localStorage.getItem('fraud_proxy_url');

  if (apiKey) document.getElementById('apiKey').value = apiKey;
  if (env) document.getElementById('fraudEnv').value = env;
  if (customUrl) document.getElementById('customEnvUrl').value = customUrl;
  if (proxyUrl) document.getElementById('proxyUrl').value = proxyUrl;

  handleEnvChange();
}

function clearFraudSettings() {
  localStorage.removeItem('fraud_api_key');
  localStorage.removeItem('fraud_env');
  localStorage.removeItem('fraud_custom_url');
  localStorage.removeItem('fraud_proxy_url');

  document.getElementById('apiKey').value = '';
  document.getElementById('fraudEnv').value = 'production';
  document.getElementById('customEnvUrl').value = '';
  document.getElementById('proxyUrl').value = '';
  handleEnvChange();

  showOutput('✓ Settings cleared');
}

async function runFraudDetection() {
  try {
    const apiKey = document.getElementById('apiKey').value;
    if (!apiKey) {
      showOutput('Error: Please enter an API key');
      return;
    }

    // Save settings
    localStorage.setItem('fraud_api_key', apiKey);
    localStorage.setItem('fraud_env', document.getElementById('fraudEnv').value);
    localStorage.setItem('fraud_custom_url', document.getElementById('customEnvUrl').value);
    localStorage.setItem('fraud_proxy_url', document.getElementById('proxyUrl').value);

    showOutput('🔄 Extracting current sheet data...');

    //Extract current sheet data
    await Excel.run(async (context) => {
      const currentSheet = context.workbook.worksheets.getActiveWorksheet();
      const range = currentSheet.getUsedRange();
      range.load('values, address, rowCount, columnCount');
      await context.sync();

      const data = range.values;
      const jsonData = convertToContextualJSON(data);

      showOutput(`🔄 Sending ${jsonData.length} records to fraud detection API...\n⏳ This may take 30-60 seconds...`);

      // Check if proxy URL is provided
      const proxyUrl = document.getElementById('proxyUrl').value.trim();
      const env = document.getElementById('fraudEnv').value;

      // Make API call with timeout
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 120000); // 2 minute timeout

      let response;

      try {
        if (proxyUrl) {
          // Use Supabase Edge Function proxy
          showOutput(`🔄 Using proxy: ${proxyUrl}\n⏳ Please wait...`);

          let environment = env;
          if (env === 'custom') {
            environment = document.getElementById('customEnvUrl').value.trim();
            if (!environment) {
              showOutput('Error: Please enter a custom environment name');
              return;
            }
          }

          response = await fetch(proxyUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              environment: environment,
              apiKey: apiKey,
              data: jsonData
            }),
            signal: controller.signal
          });

        } else {
          // Call API directly
          let apiUrl;

          if (env === 'production') {
            apiUrl = 'https://api.airia.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01';
          } else if (env === 'dev') {
            apiUrl = 'https://dev.api.airiadev.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01';
          } else {
            const customEnv = document.getElementById('customEnvUrl').value.trim();
            if (!customEnv) {
              showOutput('Error: Please enter a custom environment name');
              return;
            }
            // Format: demo -> demo.api.airia.ai
            apiUrl = `https://${customEnv}.api.airia.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01`;
          }

          showOutput(`🔄 Calling API: ${apiUrl}\n⏳ Please wait...`);

          response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
              'X-API-KEY': apiKey,
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              userInput: JSON.stringify(jsonData),
              asyncOutput: false
            }),
            signal: controller.signal,
            mode: 'cors'
          });
        }

        clearTimeout(timeoutId);

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`API Error ${response.status}: ${errorText || response.statusText}`);
        }

        showOutput('🔄 Processing response...');

        const result = await response.json();
        let fraudData;

        // Parse the API response - handle various response formats
        if (typeof result === 'string') {
          fraudData = JSON.parse(result);
        } else if (result.output) {
          fraudData = typeof result.output === 'string' ? JSON.parse(result.output) : result.output;
        } else if (result.result) {
          fraudData = typeof result.result === 'string' ? JSON.parse(result.result) : result.result;
        } else {
          fraudData = result;
        }

        if (!Array.isArray(fraudData) || fraudData.length === 0) {
          throw new Error('API returned invalid or empty data');
        }

        showOutput('✓ Received fraud detection results. Updating sheet...');

        // Convert back to array format and insert
        const arrayData = convertFromContextualJSON(fraudData);

        // Clear existing content and insert new data with fraud scores
        range.clear();
        const newRange = currentSheet.getRange('A1').getResizedRange(arrayData.length - 1, arrayData[0].length - 1);
        newRange.values = arrayData;

        // Apply conditional formatting based on risk level
        await applyRiskFormatting(context, currentSheet, fraudData);

        newRange.format.autofitColumns();
        await context.sync();

        showOutput(`✓ Fraud detection complete! Updated ${fraudData.length} records with risk scores.`);

      } catch (fetchError) {
        clearTimeout(timeoutId);
        if (fetchError.name === 'AbortError') {
          throw new Error('Request timeout - API took too long to respond (>2 minutes)');
        }
        throw fetchError;
      }
    });

  } catch (error) {
    let errorMsg = error.message;

    // Provide helpful error messages
    if (errorMsg.includes('Failed to fetch') || errorMsg.includes('NetworkError')) {
      errorMsg = 'Network Error: Unable to reach API. Possible causes:\n' +
                 '1. CORS restrictions (API must allow requests from GitHub Pages)\n' +
                 '2. Invalid API URL\n' +
                 '3. Network connectivity issues\n' +
                 '4. API server is down\n\n' +
                 'Original error: ' + errorMsg;
    }

    showOutput(`❌ Error: ${errorMsg}`);
  }
}

async function applyRiskFormatting(context, sheet, fraudData) {
  // Find the Risk Level and Fraud Score column indices
  const headers = Object.keys(fraudData[0]);
  const riskLevelIndex = headers.indexOf('Risk Level');
  const fraudScoreIndex = headers.indexOf('Fraud Score');
  
  if (riskLevelIndex === -1) return;
  
  // Apply formatting row by row
  for (let i = 0; i < fraudData.length; i++) {
    const rowIndex = i + 2; // +2 because: +1 for header, +1 for 1-based indexing
    const riskLevel = fraudData[i]['Risk Level'];
    const fraudScore = fraudData[i]['Fraud Score'];
    
    const rowRange = sheet.getRange(`A${rowIndex}:ZZ${rowIndex}`);
    
    // Color code based on risk level and fraud score
    if (riskLevel === 'HIGH' || fraudScore > 0.7) {
      rowRange.format.fill.color = '#ffcccc'; // Light red
    } else if (riskLevel === 'MEDIUM' || fraudScore > 0.3) {
      rowRange.format.fill.color = '#ffe6cc'; // Light orange
    } else if (riskLevel === 'LOW' && fraudScore === 0) {
      rowRange.format.fill.color = '#e6ffe6'; // Light green
    }
  }
}
