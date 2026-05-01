import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';
import EmployeeWeekConsolidator from './TransactionDuplicateRemover';

function App() {
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState('');
  const [processedData, setProcessedData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [stats, setStats] = useState(null);
  const [activeTab, setActiveTab] = useState('gljournal'); // Default to GL Journal tab
  const [comparisonData, setComparisonData] = useState(null);
  const [verificationData, setVerificationData] = useState(null);

  // Define expected staffing headers
  const STAFFING_HEADERS = [
    'Staffing Entity',
    'Worksite Customer',
    'EmployeeID',
    'Employee Name',
    'TransactionGUID',
    'WeekWorked',
    'Reg Hours',
    'Reg Wages',
    'OT Hours',
    'OT Wages',
    'Double Hours',
    'Double Wages',
    'Total Hours',
    'Total Wages',
    'WC Wages',
    'WC Amount',
    'WC Code',
    'Bill Amount',
    'Final Billed',
    'InvoiceNumber'
  ];

  // Handle file selection
  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setFileName(selectedFile.name);
      setError('');
      setSuccess('');
      setProcessedData(null);
      setStats(null);
      setComparisonData(null);
      setVerificationData(null);
    }
  };

  // Process the selected Excel file for GL Journal formatting
  const processGLJournalFile = async () => {
    if (!file) {
      setError('Please select an Excel file first.');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      // Read the file
      const data = await readFileAsync(file);
      
      // Parse the Excel data
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true
      });
      
      // Get the first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert to array of arrays (preserving all data exactly as is)
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // Calculate totals for verification
      let totalDebit = 0;
      let totalCredit = 0;
      
      rawData.forEach(row => {
        // Assuming debit is in column 4 (index 3) and credit is in column 5 (index 4)
        if (row[3] && typeof row[3] === 'number') {
          totalDebit += row[3];
        }
        if (row[4] && typeof row[4] === 'number') {
          totalCredit += row[4];
        }
      });
      
      setStats({
        totalRows: rawData.length,
        totalDebit: totalDebit,
        totalCredit: totalCredit,
        balanced: Math.abs(totalDebit - totalCredit) < 0.01 // Account for small floating point differences
      });
      
      // Store the data for preview
      setProcessedData(rawData);
      
      setSuccess('File processed successfully! You can now export the formatted file.');
    } catch (err) {
      console.error('Error processing file:', err);
      setError(`Error processing file: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // Helper function to compare two rows for exact duplication
  const compareRows = (row1, row2, headers) => {
    if (row1.length !== row2.length) {
      return { 
        exactMatch: false, 
        differences: [`Row lengths differ: ${row1.length} vs ${row2.length}`] 
      };
    }
    
    const differences = [];
    for (let i = 0; i < row1.length; i++) {
      // Skip empty cells or cells that are undefined/null in both rows
      if ((!row1[i] && !row2[i]) || (row1[i] === undefined && row2[i] === undefined)) {
        continue;
      }
      
      // Compare values, converting to string to handle different types
      const val1 = row1[i] !== undefined && row1[i] !== null ? row1[i].toString() : '';
      const val2 = row2[i] !== undefined && row2[i] !== null ? row2[i].toString() : '';
      
      if (val1 !== val2) {
        // Add the column name if headers are available for this index
        const columnName = headers && headers[i] ? headers[i] : `Column ${i+1}`;
        differences.push(`${columnName}: "${val1}" vs "${val2}"`);
      }
    }
    
    return {
      exactMatch: differences.length === 0,
      differences
    };
  };

  // Helper function to detect header format
  const detectHeaderFormat = (headers) => {
    const headerStr = headers.map(h => h?.toString().trim().toLowerCase()).join(',');
    
    // Check for format 1: Original format
    const format1Headers = ['staffing entity', 'worksite customer', 'employeeid', 'employee name', 'transactionguid', 'weekworked', 'reg hours'];
    const isFormat1 = format1Headers.every(h => headerStr.includes(h));
    
    // Check for other expected columns to validate format
    const hasExpectedColumns = headerStr.includes('transactionguid') && 
                               (headerStr.includes('reg hours') || headerStr.includes('reghours'));
    
    if (isFormat1) {
      return { format: 'staffing', valid: true };
    } else if (hasExpectedColumns) {
      return { format: 'alternative', valid: true };
    } else {
      return { format: 'unknown', valid: false };
    }
  };

  // Process file for removing transaction duplicates
  const processTransactionDuplicates = async () => {
    if (!file) {
      setError('Please select an Excel file first.');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      // Read the file
      const data = await readFileAsync(file);
      
      // Parse the Excel data
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true
      });
      
      // Get the first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert to JSON with headers
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (jsonData.length < 2) {
        throw new Error('The file does not contain enough data.');
      }
      
      // Get headers
      const headers = jsonData[0];

      // Detect header format
      const headerFormat = detectHeaderFormat(headers);
      if (!headerFormat.valid) {
        throw new Error('Invalid file format. Expected headers including TransactionGUID and Reg Hours columns.');
      }
      
      console.log(`Detected header format: ${headerFormat.format}`);

      // Find the index of TransactionGUID column
      const transactionGuidIndex = headers.findIndex(
        header => header && header.toString().trim().toLowerCase() === 'transactionguid'
      );
      
      if (transactionGuidIndex === -1) {
        throw new Error('TransactionGUID column not found in the file.');
      }
      
      // Find the index of Reg Hours column - handles both "Reg Hours" and "RegHours"
      const regHoursIndex = headers.findIndex(
        header => {
          const normalized = header && header.toString().trim().toLowerCase().replace(/\s+/g, '');
          return normalized === 'reghours' || normalized === 'regularhours';
        }
      );
      
      // Process data to remove duplicates
      const uniqueTransactions = new Map();
      const duplicatesRemoved = [];
      const uniqueRows = [];
      const duplicateSets = {};
      const skippedZeroHours = [];
      
      // Skip header row, process all data rows
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row.length === 0) continue; // Skip empty rows
        
        const transactionId = row[transactionGuidIndex]?.toString().trim();
        
        if (!transactionId) {
          // Skip rows with empty transaction IDs
          continue;
        }
        
        // Check if Reg Hours is 0 (if column exists)
        if (regHoursIndex !== -1) {
          const regHours = parseFloat(row[regHoursIndex]) || 0;
          if (regHours === 0) {
            // Skip rows with 0 regular hours and track them
            skippedZeroHours.push({
              rowIndex: i,
              row,
              transactionId
            });
            continue;
          }
        }
        
        if (!uniqueTransactions.has(transactionId)) {
          // This is a new unique transaction ID
          uniqueTransactions.set(transactionId, { rowIndex: i, row });
          uniqueRows.push(row);
        } else {
          // This is a duplicate transaction ID
          const originalRow = uniqueTransactions.get(transactionId);
          
          // Add to duplicates removed
          duplicatesRemoved.push({
            rowIndex: i,
            row,
            originalRowIndex: originalRow.rowIndex
          });
          
          // Create or update the duplicate set for this transaction ID
          if (!duplicateSets[transactionId]) {
            duplicateSets[transactionId] = {
              transactionId,
              originalRow: originalRow.row,
              originalRowIndex: originalRow.rowIndex,
              duplicates: []
            };
          }
          
          duplicateSets[transactionId].duplicates.push({
            rowIndex: i,
            row
          });
        }
      }
      
      // Create a new array with header and unique rows
      const processedRows = [headers, ...uniqueRows];
      
      // Create comparison report data
      const comparisonReportData = [];
      
      // Add headers for the comparison report: Status + Original headers
      comparisonReportData.push(['Status', ...headers]);
      
      // For each transaction ID with duplicates
      for (const transactionId in duplicateSets) {
        const set = duplicateSets[transactionId];
        
        // Add the original (kept) row
        comparisonReportData.push(['KEPT', ...set.originalRow]);
        
        // Add each duplicate (removed) row
        set.duplicates.forEach(duplicate => {
          comparisonReportData.push(['REMOVED', ...duplicate.row]);
        });
        
        // Add a blank row as separator
        comparisonReportData.push(Array(headers.length + 1).fill(''));
      }
      
      // Create verification sheet data
      const verificationHeaders = [
        'Transaction ID',
        'Kept Row #',
        'Removed Row #',
        'Is Exact Duplicate?',
        'Differences (if any)'
      ];
      
      const verificationRows = [];
      for (const transactionId in duplicateSets) {
        const set = duplicateSets[transactionId];
        
        set.duplicates.forEach(duplicate => {
          // Compare rows to check if they're exact duplicates
          const isExactDuplicate = compareRows(set.originalRow, duplicate.row, headers);
          
          // Find differences if not an exact duplicate
          let differences = '';
          if (!isExactDuplicate.exactMatch) {
            differences = isExactDuplicate.differences.join(', ');
          }
          
          verificationRows.push([
            transactionId,
            set.originalRowIndex + 1, // +1 for Excel row number (1-based)
            duplicate.rowIndex + 1,   // +1 for Excel row number (1-based)
            isExactDuplicate.exactMatch ? 'Yes' : 'No',
            differences
          ]);
        });
      }
      
      // Save comparison and verification data for export
      setComparisonData(comparisonReportData);
      setVerificationData([verificationHeaders, ...verificationRows]);
      
      // Update stats with verification info
      setStats({
        totalRows: jsonData.length - 1, // Excluding header row
        uniqueRows: uniqueRows.length,
        duplicatesRemoved: duplicatesRemoved.length,
        skippedZeroHours: skippedZeroHours.length,
        duplicateSets: Object.keys(duplicateSets).length,
        exactDuplicatesCount: verificationRows.filter(row => row[3] === 'Yes').length,
        partialDuplicatesCount: verificationRows.filter(row => row[3] === 'No').length,
        headers
      });
      
      // Store the data for preview
      setProcessedData(processedRows);
      
      const formatName = headerFormat.format === 'staffing' ? 'Staffing format' : 'Alternative format';
      const successMessage = skippedZeroHours.length > 0 
        ? `File processed successfully (${formatName})! Duplicates have been removed and ${skippedZeroHours.length} rows with 0 regular hours were excluded. You can now export the results.`
        : `File processed successfully (${formatName})! Duplicates have been removed. You can now export the results.`;
      setSuccess(successMessage);
    } catch (err) {
      console.error('Error processing file:', err);
      setError(`Error processing file: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // Process file for removing customer transaction duplicates
  const processCustomerTransactionDuplicates = async () => {
    if (!file) {
      setError('Please select an Excel file first.');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      // Read the file
      const data = await readFileAsync(file);
      
      // Parse the Excel data
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true
      });
      
      // Get the first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert to JSON with headers
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (jsonData.length < 2) {
        throw new Error('The file does not contain enough data.');
      }
      
      // Get headers
      const headers = jsonData[0];

      // Validate customer format headers
      const requiredHeaders = ['customername', 'employeeid', 'transactionguid', 'reghours'];
      const headerStr = headers.map(h => h?.toString().trim().toLowerCase().replace(/\s+/g, '')).join(',');
      
      const hasRequiredHeaders = requiredHeaders.every(h => headerStr.includes(h));
      if (!hasRequiredHeaders) {
        throw new Error('Invalid file format. Expected headers: CustomerName, EmployeeID, TransactionGUID, Reg Hours, etc.');
      }

      // Find the index of TransactionGUID column
      const transactionGuidIndex = headers.findIndex(
        header => header && header.toString().trim().toLowerCase() === 'transactionguid'
      );
      
      if (transactionGuidIndex === -1) {
        throw new Error('TransactionGUID column not found in the file.');
      }
      
      // Find the index of Reg Hours column
      const regHoursIndex = headers.findIndex(
        header => {
          const normalized = header && header.toString().trim().toLowerCase().replace(/\s+/g, '');
          return normalized === 'reghours' || normalized === 'regularhours';
        }
      );
      
      // Process data to remove duplicates
      const uniqueTransactions = new Map();
      const duplicatesRemoved = [];
      const uniqueRows = [];
      const duplicateSets = {};
      const skippedZeroHours = [];
      
      // Skip header row, process all data rows
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row.length === 0) continue; // Skip empty rows
        
        const transactionId = row[transactionGuidIndex]?.toString().trim();
        
        if (!transactionId) {
          // Skip rows with empty transaction IDs
          continue;
        }
        
        // Check if Reg Hours is 0 (if column exists)
        if (regHoursIndex !== -1) {
          const regHours = parseFloat(row[regHoursIndex]) || 0;
          if (regHours === 0) {
            // Skip rows with 0 regular hours and track them
            skippedZeroHours.push({
              rowIndex: i,
              row,
              transactionId
            });
            continue;
          }
        }
        
        if (!uniqueTransactions.has(transactionId)) {
          // This is a new unique transaction ID
          uniqueTransactions.set(transactionId, { rowIndex: i, row });
          uniqueRows.push(row);
        } else {
          // This is a duplicate transaction ID
          const originalRow = uniqueTransactions.get(transactionId);
          
          // Add to duplicates removed
          duplicatesRemoved.push({
            rowIndex: i,
            row,
            originalRowIndex: originalRow.rowIndex
          });
          
          // Create or update the duplicate set for this transaction ID
          if (!duplicateSets[transactionId]) {
            duplicateSets[transactionId] = {
              transactionId,
              originalRow: originalRow.row,
              originalRowIndex: originalRow.rowIndex,
              duplicates: []
            };
          }
          
          duplicateSets[transactionId].duplicates.push({
            rowIndex: i,
            row
          });
        }
      }
      
      // Create a new array with header and unique rows
      const processedRows = [headers, ...uniqueRows];
      
      // Create comparison report data
      const comparisonReportData = [];
      
      // Add headers for the comparison report: Status + Original headers
      comparisonReportData.push(['Status', ...headers]);
      
      // For each transaction ID with duplicates
      for (const transactionId in duplicateSets) {
        const set = duplicateSets[transactionId];
        
        // Add the original (kept) row
        comparisonReportData.push(['KEPT', ...set.originalRow]);
        
        // Add each duplicate (removed) row
        set.duplicates.forEach(duplicate => {
          comparisonReportData.push(['REMOVED', ...duplicate.row]);
        });
        
        // Add a blank row as separator
        comparisonReportData.push(Array(headers.length + 1).fill(''));
      }
      
      // Create verification sheet data
      const verificationHeaders = [
        'Transaction ID',
        'Kept Row #',
        'Removed Row #',
        'Is Exact Duplicate?',
        'Differences (if any)'
      ];
      
      const verificationRows = [];
      for (const transactionId in duplicateSets) {
        const set = duplicateSets[transactionId];
        
        set.duplicates.forEach(duplicate => {
          // Compare rows to check if they're exact duplicates
          const isExactDuplicate = compareRows(set.originalRow, duplicate.row, headers);
          
          // Find differences if not an exact duplicate
          let differences = '';
          if (!isExactDuplicate.exactMatch) {
            differences = isExactDuplicate.differences.join(', ');
          }
          
          verificationRows.push([
            transactionId,
            set.originalRowIndex + 1, // +1 for Excel row number (1-based)
            duplicate.rowIndex + 1,   // +1 for Excel row number (1-based)
            isExactDuplicate.exactMatch ? 'Yes' : 'No',
            differences
          ]);
        });
      }
      
      // Save comparison and verification data for export
      setComparisonData(comparisonReportData);
      setVerificationData([verificationHeaders, ...verificationRows]);
      
      // Update stats with verification info
      setStats({
        totalRows: jsonData.length - 1, // Excluding header row
        uniqueRows: uniqueRows.length,
        duplicatesRemoved: duplicatesRemoved.length,
        skippedZeroHours: skippedZeroHours.length,
        duplicateSets: Object.keys(duplicateSets).length,
        exactDuplicatesCount: verificationRows.filter(row => row[3] === 'Yes').length,
        partialDuplicatesCount: verificationRows.filter(row => row[3] === 'No').length,
        headers
      });
      
      // Store the data for preview
      setProcessedData(processedRows);
      
      const successMessage = skippedZeroHours.length > 0 
        ? `File processed successfully (Customer format)! Duplicates have been removed and ${skippedZeroHours.length} rows with 0 regular hours were excluded. You can now export the results.`
        : `File processed successfully (Customer format)! Duplicates have been removed. You can now export the results.`;
      setSuccess(successMessage);
    } catch (err) {
      console.error('Error processing file:', err);
      setError(`Error processing file: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // Process the file based on active tab
  const processFile = () => {
    if (activeTab === 'gljournal') {
      processGLJournalFile();
    } else if (activeTab === 'duplicateremover') {
      processTransactionDuplicates();
    } else if (activeTab === 'customerduplicateremover') {
      processCustomerTransactionDuplicates();
    }
  };

  // Helper function to read file asynchronously
  const readFileAsync = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('Error reading file'));
      reader.readAsArrayBuffer(file);
    });
  };

  // Export the processed data to Excel
  const handleExport = () => {
    if (!processedData) {
      setError('No processed data to export.');
      return;
    }

    try {
      // Create a new workbook with the processed data
      const ws = XLSX.utils.aoa_to_sheet(processedData);
      const wb = XLSX.utils.book_new();
      
      // Set sheet name based on the active tab
      const sheetName = activeTab === 'gljournal' ? "GL Journal" : 
                       (activeTab === 'customerduplicateremover' ? "Customer Transactions" : "Unique Transactions");
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      
      // Generate a filename
      const prefix = activeTab === 'gljournal' ? 'formatted' : 
                    (activeTab === 'customerduplicateremover' ? 'customer_unique' : 'unique');
      const outputFileName = `${prefix}_${fileName}`;
      
      // Export to Excel
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`File exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting file:', err);
      setError(`Error exporting file: ${err.message}`);
    }
  };

  // Export verification data to a separate Excel file
  const exportVerificationData = () => {
    if (!verificationData) {
      setError('No verification data available.');
      return;
    }

    try {
      // Create a new workbook with the verification data
      const ws = XLSX.utils.aoa_to_sheet(verificationData);
      
      // Apply header style
      const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4472C4" } }
      };
      
      // Set column widths for better readability
      ws['!cols'] = [
        { wch: 40 }, // Transaction ID
        { wch: 10 }, // Kept Row #
        { wch: 10 }, // Removed Row #
        { wch: 15 }, // Is Exact Duplicate?
        { wch: 50 }  // Differences
      ];
      
      // Apply styles to all cells in the first row (header)
      for (let c = 0; c < verificationData[0].length; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (!ws[cellRef]) ws[cellRef] = { v: verificationData[0][c] };
        ws[cellRef].s = headerStyle;
      }
      
      // Conditional formatting for Yes/No in "Is Exact Duplicate?" column
      for (let r = 1; r < verificationData.length; r++) {
        const exactMatchCellRef = XLSX.utils.encode_cell({ r, c: 3 });
        if (ws[exactMatchCellRef] && ws[exactMatchCellRef].v === 'Yes') {
          ws[exactMatchCellRef].s = { fill: { fgColor: { rgb: "E2EFDA" } } }; // Light green
        } else if (ws[exactMatchCellRef] && ws[exactMatchCellRef].v === 'No') {
          ws[exactMatchCellRef].s = { fill: { fgColor: { rgb: "FCE4D6" } } }; // Light orange
        }
      }
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Duplicate Verification");
      
      // Generate a filename
      const outputFileName = `verification_${fileName}`;
      
      // Export to Excel
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`Verification data exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting verification file:', err);
      setError(`Error exporting verification file: ${err.message}`);
    }
  };

  // Export comparison report to Excel
  const exportComparisonReport = () => {
    if (!comparisonData) {
      setError('No comparison data available.');
      return;
    }

    try {
      // Create a new workbook with the comparison data
      const ws = XLSX.utils.aoa_to_sheet(comparisonData);
      
      // Apply conditional formatting for KEPT vs REMOVED rows
      const keptStyle = { 
        fill: { fgColor: { rgb: "E2EFDA" } }, // Light green
        font: { bold: true }
      };
      
      const removedStyle = {
        fill: { fgColor: { rgb: "FCE4D6" } } // Light orange/peach
      };
      
      // Apply header style
      const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4472C4" } }
      };
      
      // Apply styles to all cells in the first row (header)
      for (let c = 0; c < comparisonData[0].length; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (!ws[cellRef]) ws[cellRef] = { v: comparisonData[0][c] };
        ws[cellRef].s = headerStyle;
      }
      
      // Apply styles to all other rows
      for (let r = 1; r < comparisonData.length; r++) {
        if (comparisonData[r].length === 0 || !comparisonData[r][0]) {
          // Skip empty separator rows
          continue;
        }
        
        const status = comparisonData[r][0];
        const style = status === 'KEPT' ? keptStyle : removedStyle;
        
        // Apply the appropriate style to each cell in this row
        for (let c = 0; c < comparisonData[r].length; c++) {
          const cellRef = XLSX.utils.encode_cell({ r, c });
          if (!ws[cellRef]) ws[cellRef] = { v: comparisonData[r][c] };
          ws[cellRef].s = style;
        }
      }
      
      // Set column widths for better readability
      const colWidths = [
        { wch: 10 } // Status column
      ];
      
      // Add reasonable column widths for each data column
      stats.headers.forEach(() => {
        colWidths.push({ wch: 25 });
      });
      
      ws['!cols'] = colWidths;
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Side-by-Side Comparison");
      
      // Generate a filename
      const outputFileName = `comparison_${fileName}`;
      
      // Export to Excel
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`Side-by-side comparison report exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting comparison file:', err);
      setError(`Error exporting comparison file: ${err.message}`);
    }
  };

  // Render tab content based on active tab
  const renderStats = () => {
    if (!stats) return null;

    if (activeTab === 'gljournal') {
      return (
        <div className="stats-container">
          <h3>File Statistics</h3>
          <p><strong>Total Rows:</strong> {stats.totalRows}</p>
          <p><strong>Total Debit:</strong> {stats.totalDebit.toFixed(2)}</p>
          <p><strong>Total Credit:</strong> {stats.totalCredit.toFixed(2)}</p>
          <p><strong>Balance Status:</strong> {stats.balanced ? 'Balanced ✓' : 'Unbalanced ✗'}</p>
        </div>
      );
    } else if (activeTab === 'duplicateremover' || activeTab === 'customerduplicateremover') {
      return (
        <div className="stats-container">
          <h3>File Statistics</h3>
          <p><strong>Total Rows:</strong> {stats.totalRows}</p>
          <p><strong>Unique Transactions:</strong> {stats.uniqueRows}</p>
          <p><strong>Duplicates Removed:</strong> {stats.duplicatesRemoved}</p>
          {stats.skippedZeroHours > 0 && (
            <p><strong>Skipped (0 Reg Hours):</strong> {stats.skippedZeroHours}</p>
          )}
          
          {verificationData && (
            <>
              <p><strong>Transaction IDs with Duplicates:</strong> {stats.duplicateSets}</p>
              <p><strong>Exact Duplicates:</strong> {stats.exactDuplicatesCount}</p>
              <p><strong>Partial Duplicates:</strong> {stats.partialDuplicatesCount}</p>
              <div className="report-buttons">
                <button 
                  onClick={exportVerificationData} 
                  className="verify-button"
                >
                  Export Verification Report
                </button>
                {comparisonData && (
                  <button 
                    onClick={exportComparisonReport} 
                    className="compare-button"
                  >
                    Export Side-by-Side Comparison
                  </button>
                )}
              </div>
            </>
          )}
        </div>
      );
    }
  };

  // Render tab header text based on active tab
  const getTabHeaderText = () => {
    if (activeTab === 'gljournal') {
      return (
        <>
          <h1>GL Journal Excel Formatter</h1>
          <p>Format your GL Journal files while maintaining data integrity</p>
        </>
      );
    } else if (activeTab === 'duplicateremover') {
      return (
        <>
          <h1>Staffing Transaction Duplicate Remover</h1>
          <p>Remove duplicate staffing transactions based on TransactionGUID</p>
          <p className="format-note">
            <strong>Note:</strong> Rows with 0 regular hours will be automatically excluded.
            <br />
            <strong>Supported formats:</strong> Files with TransactionGUID and Reg Hours columns
          </p>
        </>
      );
    } else if (activeTab === 'employeeconsolidator') {
      return (
        <>
          <h1>Employee Week Consolidator</h1>
          <p>Consolidate employee data by client and sum total weeks worked</p>
        </>
      );
    } else if (activeTab === 'customerduplicateremover') {
      return (
        <>
          <h1>Customer Transaction Duplicate Remover</h1>
          <p>Remove duplicate customer transactions based on TransactionGUID</p>
          <p className="format-note">
            <strong>Note:</strong> Rows with 0 regular hours will be automatically excluded.
            <br />
            <strong>Expected headers:</strong> CustomerName, EmployeeID, Employee Name, WeekWorked, Reg Hours, 
            Reg Wages, OT Hours, OT Wages, Double Hours, Double Wages, Total Hours, Total Wages, 
            WC Wages, TransactionGUID, WC Amount, WC Code, Bill Amount, Final Billed
          </p>
        </>
      );
    }
  };

  // Get process button text based on active tab
  const getProcessButtonText = () => {
    if (loading) return 'Processing...';
    
    if (activeTab === 'gljournal') {
      return 'Process File';
    } else if (activeTab === 'duplicateremover' || activeTab === 'customerduplicateremover') {
      return 'Remove Duplicates';
    }
    
    return 'Process File';
  };

  // Get helper text based on active tab
  const getHelperText = () => {
    if (activeTab === 'gljournal') {
      return "Upload your GL Journal Excel file for processing.";
    } else if (activeTab === 'duplicateremover') {
      return (
        <span>
          Expected columns include: Staffing Entity, Worksite Customer, EmployeeID, Employee Name, 
          TransactionGUID, WeekWorked, Reg Hours, Reg Wages, OT Hours, OT Wages, Double Hours, 
          Double Wages, Total Hours, Total Wages, WC Wages, WC Amount, WC Code, Bill Amount, Final Billed, InvoiceNumber
        </span>
      );
    } else if (activeTab === 'customerduplicateremover') {
      return (
        <span>
          Upload an Excel file with customer transaction data. The tool will remove duplicate transactions 
          based on TransactionGUID and exclude rows with 0 regular hours.
        </span>
      );
    }
    
    return "";
  };

  return (
    <div className="app-container">
      <header className="app-header">
        {getTabHeaderText()}
      </header>

      <div className="tabs-container">
        <div className="tabs">
          <button 
            className={`tab-button ${activeTab === 'gljournal' ? 'active' : ''}`}
            onClick={() => setActiveTab('gljournal')}
          >
            GL Journal Formatter
          </button>
          <button 
            className={`tab-button ${activeTab === 'duplicateremover' ? 'active' : ''}`}
            onClick={() => setActiveTab('duplicateremover')}
          >
            Staffing Duplicate Remover
          </button>
          <button 
            className={`tab-button ${activeTab === 'employeeconsolidator' ? 'active' : ''}`}
            onClick={() => setActiveTab('employeeconsolidator')}
          >
            Employee Consolidator
          </button>
          <button 
            className={`tab-button ${activeTab === 'customerduplicateremover' ? 'active' : ''}`}
            onClick={() => setActiveTab('customerduplicateremover')}
          >
            Customer Duplicate Remover
          </button>
        </div>
      </div>

      <main className="app-main">
        {activeTab === 'employeeconsolidator' ? (
          <EmployeeWeekConsolidator />
        ) : (
          <>
            <div className="file-upload-container">
              <h2>Upload Excel File</h2>
              <p className="helper-text">{getHelperText()}</p>
              <div className="file-input-wrapper">
                <input 
                  type="file" 
                  accept=".xls,.xlsx" 
                  onChange={handleFileChange} 
                  className="file-input"
                  id="file-upload"
                />
                <label htmlFor="file-upload" className="file-input-label">
                  {fileName ? fileName : 'Choose file...'}
                </label>
              </div>
              
              <button 
                onClick={processFile} 
                disabled={!file || loading}
                className="process-button"
              >
                {getProcessButtonText()}
              </button>
            </div>

            {error && <div className="error-message">{error}</div>}
            {success && <div className="success-message">{success}</div>}

            {renderStats()}

            {processedData && (
              <div className="results-container">
                <h2>Processing Complete</h2>
                <p>
                  {activeTab === 'gljournal' 
                    ? 'The file has been processed with full data integrity preserved.' 
                    : activeTab === 'customerduplicateremover'
                    ? 'Customer duplicate transactions have been removed based on TransactionGUID.'
                    : 'Duplicate transactions have been removed based on TransactionGUID.'}
                </p>
                <div className="button-group">
                  <button onClick={handleExport} className="export-button">
                    Download Excel File
                  </button>
                  
                  {(activeTab === 'duplicateremover' || activeTab === 'customerduplicateremover') && verificationData && (
                    <>
                      <button onClick={exportVerificationData} className="verify-button">
                        Export Verification Report
                      </button>
                      
                      {comparisonData && (
                        <button onClick={exportComparisonReport} className="compare-button">
                          Export Side-by-Side Comparison
                        </button>
                      )}
                    </>
                  )}
                </div>
                
                <div className="preview-container">
                  <h3>Data Preview</h3>
                  <div className="table-wrapper">
                    <table className="data-preview">
                      <tbody>
                        {processedData.slice(0, 5).map((row, rowIndex) => (
                          <tr key={rowIndex} className={rowIndex === 0 ? 'header-row' : ''}>
                            {row.slice(0, 8).map((cell, cellIndex) => (
                              <td key={cellIndex}>{cell !== null && cell !== undefined ? cell.toString() : ''}</td>
                            ))}
                            {row.length > 8 && <td>...</td>}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {processedData.length > 5 && (
                      <div className="preview-note">
                        Showing first 5 rows of {processedData.length} total rows
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </>
        )}
      </main>

      <footer className="app-footer">
        <p>Excel Processing Tools &copy; {new Date().getFullYear()}</p>
      </footer>
    </div>
  );
}

export default App;