import React, { useState } from 'react';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import './CSVDateSplitter.css'; // Ensure this CSS file exists

const CSVDateSplitter = () => {
  const [file, setFile] = useState(null);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [dateColumn, setDateColumn] = useState('Date of Visit');
  const [error, setError] = useState('');
  const [processing, setProcessing] = useState(false);
  const [processedCount, setProcessedCount] = useState(0);
  const [skippedCount, setSkippedCount] = useState(0);
  const [skippedRows, setSkippedRows] = useState([]);
  const [totalRows, setTotalRows] = useState(0);

  const handleFileChange = (e) => {
    setError('');
    setFile(e.target.files[0]);
    resetCountsAndData();
  };

  const handleDateColumnChange = (e) => {
    setDateColumn(e.target.value.trim());
  };

  const resetCountsAndData = () => {
    setProcessedCount(0);
    setSkippedCount(0);
    setSkippedRows([]);
    setTotalRows(0);
  };

  const validateDates = () => {
    if (!startDate || !endDate) {
      setError('Please select both start and end dates.');
      return false;
    }
    if (new Date(endDate) < new Date(startDate)) {
      setError('End Date cannot be earlier than Start Date.');
      return false;
    }
    return true;
  };

  const processFile = () => {
    setError('');
    resetCountsAndData();

    if (!validateDates()) {
      return;
    }

    if (file) {
      const fileType = file.name.split('.').pop().toLowerCase();
      if (fileType === 'csv') {
        processCSVFile(file);
      } else if (fileType === 'xlsx' || fileType === 'xls') {
        processExcelFile(file);
      } else {
        setError('Unsupported file type. Please upload a CSV or Excel file.');
      }
    } else {
      setError('Please upload a file.');
    }
  };

  const processCSVFile = (file) => {
    setProcessing(true);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (result) => {
        const jsonData = result.data;
        const headers = result.meta.fields;
        setTotalRows(jsonData.length);

        const normalizedDateColumn = dateColumn.trim().toLowerCase();
        const actualDateColumn = headers.find(
          (header) => header.trim().toLowerCase() === normalizedDateColumn
        );

        if (!actualDateColumn) {
          setError(`The CSV does not contain a "${dateColumn}" column. Please check the column name.`);
          setProcessing(false);
          return;
        }

        transferToExcel(jsonData, actualDateColumn);
        setProcessing(false);
      },
      error: (err) => {
        setError('Failed to parse CSV file. Please ensure it is a valid CSV.');
        console.error('Error parsing CSV:', err);
        setProcessing(false);
      },
    });
  };

  const processExcelFile = (file) => {
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

      const headers = worksheet[0].map((header) => header.trim().toLowerCase());
      const jsonData = worksheet.slice(1).map((row) => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      });

      setTotalRows(jsonData.length);

      const normalizedDateColumn = dateColumn.trim().toLowerCase();
      const actualDateColumn = headers.find(
        (header) => header.trim().toLowerCase() === normalizedDateColumn
      );

      if (!actualDateColumn) {
        setError(`The Excel file does not contain a "${dateColumn}" column. Please check the column name.`);
        setProcessing(false);
        return;
      }

      transferToExcel(jsonData, actualDateColumn);
      setProcessing(false);
    };
    reader.onerror = (err) => {
      setError('Failed to read Excel file.');
      console.error('Error reading Excel file:', err);
      setProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // This function now collects all filtered rows and writes them to a single Excel file
  const transferToExcel = (data, actualDateColumn) => {
    let processed = 0;
    let skipped = 0;
    const skippedRowsLocal = [];

    const start = new Date(startDate);
    const end = new Date(endDate);
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    const validDateRows = data.filter((row, index) => {
      const dateString = row[actualDateColumn];
      const rowDate = parseDate(dateString);

      if (!rowDate || rowDate < start || rowDate > end) {
        skipped += 1;
        skippedRowsLocal.push({ rowNumber: index + 2, reason: 'Date out of range or invalid format' });
        return false;
      }

      processed += 1;
      return true;
    });

    setProcessedCount(processed);
    setSkippedCount(skipped);
    setSkippedRows(skippedRowsLocal);

    if (validDateRows.length === 0) {
      setError('No valid data found within the specified date range.');
      return;
    }

    // Transfer all processed rows to a single Excel sheet
    const worksheet = XLSX.utils.json_to_sheet(validDateRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Data');
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, `filtered_data_${startDate}_to_${endDate}.xlsx`);

    alert(
      `Processed Rows: ${processed}\nSkipped Rows: ${skipped}`
    );
  };

  const isValidDate = (date) => {
    return date instanceof Date && !isNaN(date);
  };

  const parseDate = (dateString) => {
    if (!dateString) return null;
    let parsedDate = new Date(dateString);
    if (isValidDate(parsedDate)) return parsedDate;

    const parts = dateString.split('/');
    if (parts.length === 3) {
      let [month, day, year] = parts.map(Number);
      if (year < 100) year += 2000;
      parsedDate = new Date(year, month - 1, day);
      if (isValidDate(parsedDate)) return parsedDate;
    }
    return null;
  };

  return (
    <div className="container">
      <h1>CSV/Excel Date Splitter Dashboard</h1>
      <p>
        Upload a CSV or Excel file, specify the date column, select a date range, and filter data based on the date range.
        The processed rows will be transferred to a single Excel sheet.
      </p>

      {error && <div className="error">{error}</div>}

      <div className="section">
        <label htmlFor="file-upload" className="custom-file-upload">
          Upload CSV or Excel File
          <input id="file-upload" type="file" accept=".csv,.xlsx,.xls" onChange={handleFileChange} />
        </label>
        {file && <div className="file-name">Selected File: {file.name}</div>}
      </div>

      <div className="section">
        <label htmlFor="date-column">Date Column Name:</label>
        <input
          type="text"
          id="date-column"
          value={dateColumn}
          onChange={handleDateColumnChange}
          placeholder="Date of Visit"
        />
      </div>

      <div className="section date-inputs">
        <div className="date-input">
          <label htmlFor="start-date">Start Date:</label>
          <input
            type="date"
            id="start-date"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
          />
        </div>
        <div className="date-input">
          <label htmlFor="end-date">End Date:</label>
          <input
            type="date"
            id="end-date"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
          />
        </div>
      </div>

      {totalRows > 0 && (
        <div className="summary">
          <p>Total Rows in File: {totalRows}</p>
          <p>Processed Rows: {processedCount}</p>
          <p>Skipped Rows: {skippedCount}</p>
        </div>
      )}

      {skippedRows.length > 0 && (
        <div className="skipped-rows">
          <h3>Skipped Rows:</h3>
          <ul>
            {skippedRows.map((row) => (
              <li key={row.rowNumber}>
                Row {row.rowNumber}: {row.reason}
              </li>
            ))}
          </ul>
        </div>
      )}

      <div className="section">
        <button
          onClick={processFile}
          disabled={!file || !startDate || !endDate || processing}
          className="split-button"
        >
          {processing ? 'Processing...' : 'Transfer to Excel'}
        </button>
      </div>
    </div>
  );
};

export default CSVDateSplitter;
