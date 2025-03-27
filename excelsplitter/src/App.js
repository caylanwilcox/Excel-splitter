import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState('');
  const [processedData, setProcessedData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [stats, setStats] = useState(null);

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
      // Create a new workbook preserving the exact same data
      const ws = XLSX.utils.aoa_to_sheet(processedData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      
      // Generate a filename
      const outputFileName = `formatted_${fileName}`;
      
      // Export to Excel (exact same data)
      XLSX.writeFile(wb, outputFileName);
      
      setSuccess(`File exported as "${outputFileName}"`);
    } catch (err) {
      console.error('Error exporting file:', err);
      setError(`Error exporting file: ${err.message}`);
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>GL Journal Excel Formatter</h1>
        <p>Format your GL Journal files while maintaining data integrity</p>
      </header>

      <main className="app-main">
        <div className="file-upload-container">
          <h2>Upload Excel File</h2>
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
            {loading ? 'Processing...' : 'Process File'}
          </button>
        </div>

        {error && <div className="error-message">{error}</div>}
        {success && <div className="success-message">{success}</div>}

        {stats && (
          <div className="stats-container">
            <h3>File Statistics</h3>
            <p><strong>Total Rows:</strong> {stats.totalRows}</p>
            <p><strong>Total Debit:</strong> {stats.totalDebit.toFixed(2)}</p>
            <p><strong>Total Credit:</strong> {stats.totalCredit.toFixed(2)}</p>
            <p><strong>Balance Status:</strong> {stats.balanced ? 'Balanced ✓' : 'Unbalanced ✗'}</p>
          </div>
        )}

        {processedData && (
          <div className="results-container">
            <h2>Processing Complete</h2>
            <p>The file has been processed with full data integrity preserved.</p>
            <button onClick={handleExport} className="export-button">
              Download Excel File
            </button>
            
            <div className="preview-container">
              <h3>Data Preview</h3>
              <div className="table-wrapper">
                <table className="data-preview">
                  <tbody>
                    {processedData.slice(0, 5).map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {row.slice(0, 8).map((cell, cellIndex) => (
                          <td key={cellIndex}>{cell !== null && cell !== undefined ? cell.toString() : ''}</td>
                        ))}
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
      </main>

      <footer className="app-footer">
        <p>GL Journal Excel Formatter &copy; {new Date().getFullYear()}</p>
      </footer>
    </div>
  );
}

export default App;