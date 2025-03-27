/**
 * excelformat.js
 * Utility functions for formatting GL Journal Proof Excel files in React
 */

import * as XLSX from 'xlsx';

// Define the column headers for the formatted output
export const COLUMN_HEADERS = [
  'Company',
  'Account',
  'Account Description',
  'Debit',
  'Credit',
  'Batch Number',
  'Journal Number',
  'Entity',
  'Reference',
  'Date',
  'Source',
  'Key',
  'Period'
];

/**
 * Map raw data row to properly formatted row with correct column names
 * @param {Array} row - Raw data row
 * @returns {Object} - Formatted row with proper column names
 */
export function mapRow(row) {
  return {
    'Company': row[0],
    'Account': row[1],
    'Account Description': row[2],
    'Debit': row[3] || 0,
    'Credit': row[4] || 0,
    'Batch Number': row[5],
    'Journal Number': row[6],
    'Entity': row[7],
    'Reference': row[8],
    'Date': row[9],
    'Source': row[10],
    'Key': row[11],
    'Period': row[12]
  };
}

/**
 * Format a date as MM/DD/YYYY
 * @param {Date} date - Date to format
 * @returns {string} - Formatted date string
 */
export function formatDate(date) {
  if (!date) return '';
  
  // If it's already a string, return it
  if (typeof date === 'string') return date;
  
  const d = new Date(date);
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const year = d.getFullYear();
  
  return `${month}/${day}/${year}`;
}

/**
 * Format currency values with 2 decimal places
 * @param {number} value - Number to format
 * @returns {number} - Formatted number with 2 decimal places
 */
export function formatCurrency(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'string') {
    value = parseFloat(value.replace(/,/g, ''));
    if (isNaN(value)) return 0;
  }
  return Math.round(value * 100) / 100; // Round to 2 decimal places
}

/**
 * Calculate journal balances and add summary rows
 * @param {Array} data - Array of journal entries
 * @returns {Array} - Data with summary rows added
 */
export function addSummaryRows(data) {
  const result = [...data];
  
  // Group by batch and journal
  const journalGroups = {};
  
  data.forEach(row => {
    const key = `${row['Batch Number']}-${row['Journal Number']}`;
    if (!journalGroups[key]) {
      journalGroups[key] = {
        entries: [],
        totalDebit: 0,
        totalCredit: 0
      };
    }
    
    journalGroups[key].entries.push(row);
    journalGroups[key].totalDebit += formatCurrency(row['Debit']);
    journalGroups[key].totalCredit += formatCurrency(row['Credit']);
  });
  
  // Add summary rows
  Object.keys(journalGroups).forEach(key => {
    const group = journalGroups[key];
    const lastEntryIndex = result.indexOf(group.entries[group.entries.length - 1]);
    
    const summaryRow = {
      'Company': group.entries[0]['Company'],
      'Account': '',
      'Account Description': 'JOURNAL TOTAL',
      'Debit': group.totalDebit,
      'Credit': group.totalCredit,
      'Batch Number': group.entries[0]['Batch Number'],
      'Journal Number': group.entries[0]['Journal Number'],
      'Entity': group.entries[0]['Entity'],
      'Reference': '',
      'Date': '',
      'Source': group.entries[0]['Source'],
      'Key': '',
      'Period': ''
    };
    
    // Insert the summary row after the last entry of this journal
    result.splice(lastEntryIndex + 1, 0, summaryRow);
  });
  
  return result;
}

/**
 * Format GL Journal data
 * @param {Array} rawData - Raw data from Excel file
 * @returns {Array} - Formatted data
 */
export function formatGLJournal(rawData) {
  // Filter out empty rows
  const filteredData = rawData.filter(row => row.length > 0);
  
  // Map rows to proper column structure
  const mappedData = filteredData.map(row => mapRow(row));
  
  // Apply formatting to values
  const formattedData = mappedData.map(row => {
    return {
      ...row,
      'Debit': formatCurrency(row['Debit']),
      'Credit': formatCurrency(row['Credit']),
      'Date': formatDate(row['Date'])
    };
  });
  
  // Add summary rows with totals
  const dataWithSummary = addSummaryRows(formattedData);
  
  // Add a grand total row
  const grandTotalDebit = dataWithSummary.reduce((sum, row) => sum + formatCurrency(row['Debit']), 0);
  const grandTotalCredit = dataWithSummary.reduce((sum, row) => sum + formatCurrency(row['Credit']), 0);
  
  const grandTotalRow = {
    'Company': '',
    'Account': '',
    'Account Description': 'GRAND TOTAL',
    'Debit': grandTotalDebit,
    'Credit': grandTotalCredit,
    'Batch Number': '',
    'Journal Number': '',
    'Entity': '',
    'Reference': '',
    'Date': '',
    'Source': '',
    'Key': '',
    'Period': ''
  };
  
  dataWithSummary.push(grandTotalRow);
  
  // Convert objects to arrays in the same order as COLUMN_HEADERS
  const formattedRows = dataWithSummary.map(row => 
    COLUMN_HEADERS.map(header => row[header])
  );
  
  // Return formatted data with headers as the first row
  return [COLUMN_HEADERS, ...formattedRows];
}

/**
 * Export data to an Excel file for browser download
 * @param {Array} data - Data to export
 * @param {string} fileName - Filename for the exported Excel file
 */
export function exportToExcel(data, fileName) {
  // Convert data to worksheet
  const ws = XLSX.utils.aoa_to_sheet(data);
  
  // Set column widths
  ws['!cols'] = [
    { wch: 10 },  // Company
    { wch: 10 },  // Account
    { wch: 35 },  // Account Description
    { wch: 15 },  // Debit
    { wch: 15 },  // Credit
    { wch: 12 },  // Batch Number
    { wch: 12 },  // Journal Number
    { wch: 25 },  // Entity
    { wch: 12 },  // Reference
    { wch: 12 },  // Date
    { wch: 20 },  // Source
    { wch: 25 },  // Key
    { wch: 40 }   // Period
  ];
  
  // Create a new workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "GL Journal");
  
  // Apply styles
  applyStyles(ws, data.length);
  
  // Save to file
  XLSX.writeFile(wb, fileName);
}

/**
 * Apply styles to the worksheet
 * @param {Object} ws - Worksheet
 * @param {number} rowCount - Number of rows
 */
function applyStyles(ws, rowCount) {
  // Apply header styles
  const headerStyle = {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "4F81BD" } },
    alignment: { horizontal: "center" }
  };
  
  // Apply currency styles
  const currencyStyle = {
    numFmt: '#,##0.00',
    alignment: { horizontal: "right" }
  };
  
  // Apply summary row styles
  const summaryStyle = {
    font: { bold: true },
    fill: { fgColor: { rgb: "DDEBF7" } }
  };
  
  // Apply grand total row styles
  const grandTotalStyle = {
    font: { bold: true },
    fill: { fgColor: { rgb: "BDD7EE" } },
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "double", color: { rgb: "000000" } }
    }
  };
  
  // Apply styles to header row
  COLUMN_HEADERS.forEach((header, colIndex) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
    if (!ws[cellRef]) ws[cellRef] = { v: header };
    ws[cellRef].s = headerStyle;
  });
  
  // Apply styles to data rows
  for (let rowIndex = 1; rowIndex < rowCount; rowIndex++) {
    // Get cell values for this row
    const accountDescCell = ws[XLSX.utils.encode_cell({ r: rowIndex, c: 2 })];
    const isJournalTotal = accountDescCell && accountDescCell.v === 'JOURNAL TOTAL';
    const isGrandTotal = accountDescCell && accountDescCell.v === 'GRAND TOTAL';
    
    // Apply appropriate styles based on row type
    for (let colIndex = 0; colIndex < COLUMN_HEADERS.length; colIndex++) {
      const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      if (!ws[cellRef]) continue;
      
      // Apply number format to Debit and Credit columns
      if (colIndex === 3 || colIndex === 4) {
        ws[cellRef].s = currencyStyle;
      }
      
      // Apply summary styles to journal total rows
      if (isJournalTotal) {
        ws[cellRef].s = { ...ws[cellRef].s, ...summaryStyle };
      }
      
      // Apply grand total styles to the grand total row
      if (isGrandTotal) {
        ws[cellRef].s = { ...ws[cellRef].s, ...grandTotalStyle };
      }
    }
  }
}