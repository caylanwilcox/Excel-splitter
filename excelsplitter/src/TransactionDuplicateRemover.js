import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './DuplicateRemover.css';

function EmployeeWeekConsolidator() {
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState('');
  const [processedData, setProcessedData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [stats, setStats] = useState(null);

  // Define expected headers for employee data
  const EXPECTED_HEADERS = [
    'EmployeeID',
    'FirstName', 
    'LastName',
    'SSN',
    'WeekWorked',
    'Client Name'
    // Job Title is optional (7th column)
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
    }
  };

  // Process the selected Excel file
  const processFile = async () => {
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
      
      // Get headers from the file
      const headers = jsonData[0];

      // Check if first row looks like headers or data
      const firstRow = jsonData[0];
      let hasHeaders = false;
      let dataStartIndex = 0;
      
             // Check if first row contains header-like text (non-numeric, descriptive)
       if (firstRow && firstRow.length >= 6) {
         const firstCellStr = firstRow[0]?.toString().toLowerCase() || '';
         const thirdCellStr = firstRow[2]?.toString() || '';
         
         // If first cell contains words like "first", "name", "employee" or third cell is not a number (SSN), likely headers
         if (firstCellStr.includes('first') || firstCellStr.includes('name') || firstCellStr.includes('employee') || 
             (isNaN(parseFloat(thirdCellStr)) && thirdCellStr.length < 15)) {
           hasHeaders = true;
           dataStartIndex = 1;
         }
       }
       
       // Use positional mapping: EmployeeID, FirstName, LastName, SSN, WeekWorked, Client Name, Job Title
       const employeeIDIndex = 0;
       const firstNameIndex = 1;
       const lastNameIndex = 2;
       const ssnIndex = 3;
       const weekWorkedIndex = 4;
       const clientNameIndex = 5;
       const jobTitleIndex = 6;

       // Validate we have enough columns in the data
       const sampleRow = jsonData[dataStartIndex] || [];
       if (sampleRow.length < 6) {
         throw new Error(`File must have at least 6 columns: EmployeeID, FirstName, LastName, SSN, WeekWorked, Client Name. Found only ${sampleRow.length} columns.`);
       }
      
      console.log('Processing file:', {
        hasHeaders,
        totalRows: jsonData.length,
        dataStartIndex,
        sampleRow: jsonData[dataStartIndex]
      });

      // Process data to consolidate by employee and client
      const consolidatedData = new Map();
      const originalRecords = [];
      const skippedRecords = [];
      
      // Process data rows starting from the correct index
      for (let i = dataStartIndex; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue; // Skip empty rows
        
        const employeeID = row[employeeIDIndex]?.toString().trim() || '';
        const firstName = row[firstNameIndex]?.toString().trim() || '';
        const lastName = row[lastNameIndex]?.toString().trim() || '';
        const ssn = row[ssnIndex]?.toString().trim() || '';
        const weekWorked = row[weekWorkedIndex]?.toString().trim() || '';
        const clientName = row[clientNameIndex]?.toString().trim() || '';
        const jobTitle = (row.length > jobTitleIndex && row[jobTitleIndex]) ? row[jobTitleIndex].toString().trim() : '';
        
        console.log(`Row ${i + 1}:`, { employeeID, firstName, lastName, ssn, weekWorked, clientName, jobTitle });
        
        // Parse the week worked date
        let weekWorkedDate = null;
        let weeksWorked = 1; // Default to 1 week per record
        
        const weekValue = row[weekWorkedIndex];
        if (weekValue !== undefined && weekValue !== null && weekValue !== '') {
          // Try to parse as a date first
          if (weekValue.toString().includes('/')) {
            try {
              // Handle dates like "2/18/24 0:00" or "6/9/24 0:00"
              const dateStr = weekValue.toString().split(' ')[0]; // Remove time portion
              weekWorkedDate = new Date(dateStr);
              if (isNaN(weekWorkedDate.getTime())) {
                weekWorkedDate = null;
              }
              weeksWorked = 1; // Each date record represents 1 week
            } catch (e) {
              console.log(`Could not parse date: ${weekValue}`);
              weekWorkedDate = null;
            }
          } else {
            // Try to parse as a number
            const parsed = parseFloat(weekValue.toString().replace(/[^\d.-]/g, ''));
            if (!isNaN(parsed) && parsed > 0) {
              weeksWorked = parsed;
            }
          }
        }
        
        // More lenient validation - require employeeID, firstName and clientName as minimum
        if (!employeeID || !firstName || !clientName) {
          const missingFields = [];
          if (!employeeID) missingFields.push('EmployeeID');
          if (!firstName) missingFields.push('FirstName');
          if (!clientName) missingFields.push('Client Name');
          
          skippedRecords.push({
            rowIndex: i + 1,
            reason: `Missing required data: ${missingFields.join(', ')}`,
            data: row
          });
          console.log(`Skipping row ${i + 1}: missing ${missingFields.join(', ')}`);
          continue;
        }
        
        // Normalize data for consistent key generation
        const normalizedEmployeeID = employeeID.trim();
        const normalizedFirstName = firstName.trim().toLowerCase().replace(/\s+/g, ' ');
        const normalizedLastName = lastName.trim().toLowerCase().replace(/\s+/g, ' ');
        const normalizedClientName = clientName.trim().toLowerCase().replace(/\s+/g, ' ');
        
        // Normalize SSN by removing all non-digits and formatting consistently
        const normalizedSSN = ssn ? ssn.replace(/\D/g, '') : '';
        
        // Create unique key for employee + client combination using EmployeeID and SSN for better identification
        const employeeKey = `${normalizedEmployeeID}${normalizedSSN ? `_${normalizedSSN}` : ''}_${normalizedFirstName}_${normalizedLastName}`;
        const key = `${employeeKey}|${normalizedClientName}`;
        
        originalRecords.push({
          rowIndex: i + 1,
          employeeID,
          firstName,
          lastName,
          ssn,
          clientName,
          jobTitle,
          weekWorked,
          weekWorkedDate,
          weeksWorked,
          consolidationKey: key
        });
        
        if (consolidatedData.has(key)) {
          // Add to existing consolidation
          const existing = consolidatedData.get(key);
          existing.totalWeeks += weeksWorked;
          existing.recordCount += 1;
          
          // Keep the earliest week worked date
          if (weekWorkedDate && (!existing.earliestWeekWorkedDate || weekWorkedDate < existing.earliestWeekWorkedDate)) {
            existing.earliestWeekWorkedDate = weekWorkedDate;
            existing.displayWeekWorked = weekWorked; // Store the original date string
          }
          
          // Keep the most recent job title if it's not empty
          if (jobTitle && jobTitle.trim() !== '') {
            existing.jobTitle = jobTitle;
          }
          // Update display names with the most complete version (non-empty values preferred)
          if (employeeID && employeeID.trim()) existing.displayEmployeeID = employeeID.trim();
          if (firstName && firstName.trim()) existing.displayFirstName = firstName.trim();
          if (lastName && lastName.trim()) existing.displayLastName = lastName.trim();
          if (ssn && ssn.trim()) existing.displaySSN = ssn.trim();
          if (clientName && clientName.trim()) existing.displayClientName = clientName.trim();
          
          console.log(`Consolidating: ${key} - Total weeks now: ${existing.totalWeeks}`);
        } else {
          // Create new consolidation entry
          consolidatedData.set(key, {
            employeeID: normalizedEmployeeID,
            firstName: normalizedFirstName,
            lastName: normalizedLastName,
            ssn: normalizedSSN,
            clientName: normalizedClientName,
            jobTitle,
            totalWeeks: weeksWorked,
            recordCount: 1,
            earliestWeekWorkedDate: weekWorkedDate,
            // Store display versions (original case) for output
            displayEmployeeID: employeeID.trim(),
            displayFirstName: firstName.trim(),
            displayLastName: lastName.trim(),
            displaySSN: ssn.trim(),
            displayClientName: clientName.trim(),
            displayWeekWorked: weekWorked
          });
          console.log(`New entry: ${key} - Weeks: ${weeksWorked}`);
        }
      }
      
      console.log('Consolidation summary:', {
        originalRecordsCount: originalRecords.length,
        skippedRecordsCount: skippedRecords.length,
        consolidatedEntriesCount: consolidatedData.size
      });
      
      // Convert consolidated data to array format
      const consolidatedArray = Array.from(consolidatedData.values());
      
      // Sort by Last Name, then by First Name, then by Client Name using display versions
      consolidatedArray.sort((a, b) => {
        const lastNameA = a.displayLastName || a.lastName;
        const lastNameB = b.displayLastName || b.lastName;
        const firstNameA = a.displayFirstName || a.firstName;
        const firstNameB = b.displayFirstName || b.firstName;
        const clientNameA = a.displayClientName || a.clientName;
        const clientNameB = b.displayClientName || b.clientName;
        
        if (lastNameA !== lastNameB) {
          return lastNameA.localeCompare(lastNameB);
        }
        if (firstNameA !== firstNameB) {
          return firstNameA.localeCompare(firstNameB);
        }
        return clientNameA.localeCompare(clientNameB);
      });
      
      // Create output headers
      const outputHeaders = ['EmployeeID', 'FirstName', 'LastName', 'SSN', 'WeekWorked', 'Client Name', 'Job Title'];
      
      // Create consolidated rows using display versions for proper formatting
      const consolidatedRows = consolidatedArray.map(item => [
        item.displayEmployeeID || item.employeeID,
        item.displayFirstName || item.firstName,
        item.displayLastName || item.lastName,
        item.displaySSN || item.ssn,
        item.displayWeekWorked || 'N/A',
        item.displayClientName || item.clientName,
        item.jobTitle
      ]);
      
      // Create final data array with headers
      const finalData = [outputHeaders, ...consolidatedRows];
      
      // Create detailed report for review
      const detailReportHeaders = [
        'Original Row #',
        'EmployeeID',
        'FirstName', 
        'LastName',
        'SSN',
        'Week Worked',
        'Client Name',
        'Job Title',
        'Weeks Calculated',
        'Consolidation Key'
      ];
      
      const detailReportRows = originalRecords.map(record => [
        record.rowIndex,
        record.employeeID,
        record.firstName,
        record.lastName,
        record.ssn,
        record.weekWorked,
        record.clientName,
        record.jobTitle,
        record.weeksWorked,
        record.consolidationKey
      ]);
      
      // Calculate statistics
      const totalOriginalRecords = originalRecords.length;
      const totalConsolidatedRecords = consolidatedArray.length;
      const totalWeeksAcrossAll = consolidatedArray.reduce((sum, item) => sum + item.totalWeeks, 0);
      const uniqueEmployees = new Set(consolidatedArray.map(item => `${item.displayFirstName || item.firstName} ${item.displayLastName || item.lastName}`)).size;
      const uniqueClients = new Set(consolidatedArray.map(item => item.displayClientName || item.clientName)).size;
      
      setStats({
        totalFileRows: jsonData.length,
        totalOriginalRecords,
        totalConsolidatedRecords,
        recordsConsolidated: totalOriginalRecords - totalConsolidatedRecords,
        totalWeeksAcrossAll,
        uniqueEmployees,
        uniqueClients,
        skippedRecords: skippedRecords.length,
        detailReportData: [detailReportHeaders, ...detailReportRows],
        skippedRecordsData: skippedRecords,
        headers: outputHeaders
      });
      
      // Store the data for preview and export
      setProcessedData(finalData);
      
      setSuccess('File processed successfully! Duplicate employee records have been consolidated and earliest week worked dates preserved.');
    } catch (err) {
      console.error('Error processing file:', err);
      setError(`Error processing file: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };
  
  // Helper function to find column index with flexible matching
  const findColumnIndex = (headers, possibleNames) => {
    return headers.findIndex(header => {
      if (!header) return false;
      const headerLower = header.toString().trim().toLowerCase();
      return possibleNames.some(name => headerLower === name.toLowerCase());
    });
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

  // Export the consolidated data to Excel
  const handleExport = () => {
    if (!processedData) {
      setError('No processed data to export.');
      return;
    }

    try {
      // Create a new workbook with the consolidated data
      const ws = XLSX.utils.aoa_to_sheet(processedData);
      
      // Apply header style
      const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4472C4" } }
      };
      
      // Apply styles to header row
      for (let c = 0; c < processedData[0].length; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (!ws[cellRef]) ws[cellRef] = { v: processedData[0][c] };
        ws[cellRef].s = headerStyle;
      }
      
      // Set column widths
      ws['!cols'] = [
        { wch: 12 }, // EmployeeID
        { wch: 15 }, // FirstName
        { wch: 15 }, // LastName
        { wch: 12 }, // SSN
        { wch: 15 }, // WeekWorked
        { wch: 20 }, // Client Name
        { wch: 25 }  // Job Title
      ];
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Consolidated Employee Data");
      
      // Generate a filename
      const outputFileName = `consolidated_${fileName}`;
      
      // Export to Excel
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`Consolidated file exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting file:', err);
      setError(`Error exporting file: ${err.message}`);
    }
  };

  // Export detailed report to Excel
  const exportDetailReport = () => {
    if (!stats || !stats.detailReportData) {
      setError('No detail report data available.');
      return;
    }

    try {
      // Create workbook with multiple sheets
      const wb = XLSX.utils.book_new();
      
      // Detail report sheet
      const detailWs = XLSX.utils.aoa_to_sheet(stats.detailReportData);
      
      // Apply header style
      const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4472C4" } }
      };
      
      // Apply styles to header row
      for (let c = 0; c < stats.detailReportData[0].length; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (!detailWs[cellRef]) detailWs[cellRef] = { v: stats.detailReportData[0][c] };
        detailWs[cellRef].s = headerStyle;
      }
      
      detailWs['!cols'] = [
        { wch: 12 }, // Original Row #
        { wch: 12 }, // EmployeeID
        { wch: 15 }, // FirstName
        { wch: 15 }, // LastName
        { wch: 12 }, // SSN
        { wch: 20 }, // Week Worked
        { wch: 20 }, // Client Name
        { wch: 25 }, // Job Title
        { wch: 15 }, // Weeks Calculated
        { wch: 35 }  // Consolidation Key
      ];
      
      XLSX.utils.book_append_sheet(wb, detailWs, "Original Records Detail");
      
      // Skipped records sheet (if any)
      if (stats.skippedRecordsData && stats.skippedRecordsData.length > 0) {
        const skippedHeaders = ['Original Row #', 'Reason Skipped', 'Raw Data'];
        const skippedRows = stats.skippedRecordsData.map(record => [
          record.rowIndex,
          record.reason,
          record.data.join(' | ')
        ]);
        
        const skippedWs = XLSX.utils.aoa_to_sheet([skippedHeaders, ...skippedRows]);
        
        // Apply header style to skipped records sheet
        for (let c = 0; c < skippedHeaders.length; c++) {
          const cellRef = XLSX.utils.encode_cell({ r: 0, c });
          if (!skippedWs[cellRef]) skippedWs[cellRef] = { v: skippedHeaders[c] };
          skippedWs[cellRef].s = headerStyle;
        }
        
        skippedWs['!cols'] = [
          { wch: 15 }, // Original Row #
          { wch: 20 }, // Reason Skipped
          { wch: 50 }  // Raw Data
        ];
        
        XLSX.utils.book_append_sheet(wb, skippedWs, "Skipped Records");
      }
      
      // Generate a filename
      const outputFileName = `detail_report_${fileName}`;
      
      // Export to Excel
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`Detail report exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting detail report:', err);
      setError(`Error exporting detail report: ${err.message}`);
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>Employee Duplicate Remover</h1>
        <p>Remove duplicate employee records and preserve earliest week worked dates</p>
      </header>

      <main className="app-main">
        <div className="file-upload-container">
          <h2>Upload Excel File</h2>
          <p className="helper-text">
            Expected columns (in order): EmployeeID, FirstName, LastName, SSN, WeekWorked, Client Name (Job Title is optional)
            <br />
            The tool will consolidate duplicate records by Employee + Client combination and keep the first/earliest week worked date.
            <br />
            Note: Duplicate employees (same EmployeeID/SSN) working for the same client will be consolidated into a single record.
          </p>
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
            {loading ? 'Processing...' : 'Consolidate Data'}
          </button>
        </div>

        {error && <div className="error-message">{error}</div>}
        {success && <div className="success-message">{success}</div>}

        {stats && (
          <div className="stats-container">
            <h3>Consolidation Statistics</h3>
            <p><strong>Total Rows in File:</strong> {stats.totalFileRows}</p>
            <p><strong>Original Records Processed:</strong> {stats.totalOriginalRecords}</p>
            <p><strong>Consolidated Records:</strong> {stats.totalConsolidatedRecords}</p>
            <p><strong>Records Consolidated:</strong> {stats.recordsConsolidated}</p>
            <p><strong>Total Weeks Across All Records:</strong> {stats.totalWeeksAcrossAll.toFixed(2)}</p>
            <p><strong>Unique Employees:</strong> {stats.uniqueEmployees}</p>
            <p><strong>Unique Clients:</strong> {stats.uniqueClients}</p>
            {stats.skippedRecords > 0 && (
              <p style={{color: 'orange'}}><strong>Skipped Records:</strong> {stats.skippedRecords} (missing required data)</p>
            )}
            
            <div className="report-buttons">
              <button 
                onClick={exportDetailReport} 
                className="verify-button"
              >
                Export Detail Report
              </button>
            </div>
          </div>
        )}

        {processedData && (
          <div className="results-container">
            <h2>Consolidation Complete</h2>
            <p>Duplicate employee records have been consolidated by client with earliest week worked dates preserved.</p>
            <div className="button-group">
              <button onClick={handleExport} className="export-button">
                Download Consolidated Excel File
              </button>
              
              <button onClick={exportDetailReport} className="verify-button">
                Export Detail Report
              </button>
            </div>
            
            <div className="preview-container">
              <h3>Consolidated Data Preview</h3>
              <div className="table-wrapper">
                <table className="data-preview">
                  <tbody>
                    {processedData.slice(0, 10).map((row, rowIndex) => (
                      <tr key={rowIndex} className={rowIndex === 0 ? 'header-row' : ''}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex}>{cell !== null && cell !== undefined ? cell.toString() : ''}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {processedData.length > 10 && (
                  <div className="preview-note">
                    Showing first 10 rows of {processedData.length} total rows
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      <footer className="app-footer">
        <p>Employee Week Consolidator &copy; {new Date().getFullYear()}</p>
      </footer>
    </div>
  );
}

export default EmployeeWeekConsolidator;