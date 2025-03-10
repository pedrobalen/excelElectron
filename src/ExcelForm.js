import React, { useState } from 'react';
import * as XLSX from 'xlsx';


function ExcelForm() {
  const [tableData, setTableData] = useState([]);
  const [selectedResearch, setSelectedResearch] = useState('');
  const [columnName, setColumnName] = useState('');
  const [cellValue, setCellValue] = useState('');
  const [fileName, setFileName] = useState('');
  const [originalWorkbook, setOriginalWorkbook] = useState(null);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      const reader = new FileReader();

      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const wb = XLSX.read(binaryStr, { type: 'binary' });
        setOriginalWorkbook(wb);
        
        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: ''
        });

        if (jsonData.length > 1) {
          const headers = jsonData[1];
          const rows = jsonData.slice(2);
          const formattedData = rows.map((row) => {
            const rowObject = {};
            row.forEach((cell, index) => {
              rowObject[headers[index]] = cell;
            });
            return rowObject;
          });

          setTableData(formattedData);
        }
      };

      reader.readAsBinaryString(file);
    }
  };

  const handleSubmit = (e) => {
    e.preventDefault();

    if (!selectedResearch || !columnName || !cellValue) {
      alert('Please fill in all fields');
      return;
    }

    const rowIdx = tableData.findIndex(row => Object.values(row)[0] === selectedResearch);
    
    if (rowIdx === -1) {
      alert('Research not found');
      return;
    }

    if (!tableData[0] || !tableData[0].hasOwnProperty(columnName)) {
      alert('Invalid column name');
      return;
    }

    const updatedData = [...tableData];
    updatedData[rowIdx] = {
      ...updatedData[rowIdx],
      [columnName]: cellValue
    };

    setTableData(updatedData);

    if (originalWorkbook && fileName) {
      const ws = originalWorkbook.Sheets[originalWorkbook.SheetNames[0]];
      
      const headers = Object.keys(updatedData[0]);
      const colIdx = headers.indexOf(columnName);
      
      const cellAddress = XLSX.utils.encode_cell({
        r: rowIdx + 2,
        c: colIdx
      });

      const existingCell = ws[cellAddress] || {};
      ws[cellAddress] = {
        ...existingCell,
        v: cellValue,
        w: cellValue.toString()
      };

      XLSX.writeFile(originalWorkbook, fileName);
    }

    setSelectedResearch('');
    setColumnName('');
    setCellValue('');
  };

  return (
    <div>
      <div className="file-upload">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          className="file-input"
        />
      </div>

      {tableData.length > 0 && (
        <form onSubmit={handleSubmit} className="excel-form">
          <div className="form-group">
            <label htmlFor="selectedResearch">Select Research:</label>
            <select
              id="selectedResearch"
              value={selectedResearch}
              onChange={(e) => setSelectedResearch(e.target.value)}
            >
              <option value="">Select a research</option>
              {tableData.map((row) => (
                <option key={Object.values(row)[0]} value={Object.values(row)[0]}>
                  {Object.values(row)[0]}
                </option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label htmlFor="columnName">Column Name:</label>
            <select
              id="columnName"
              value={columnName}
              onChange={(e) => setColumnName(e.target.value)}
            >
              <option value="">Select a column</option>
              {tableData[0] &&
                Object.keys(tableData[0]).map((col) => (
                  <option key={col} value={col}>
                    {col}
                  </option>
                ))}
            </select>
          </div>

          <div className="form-group">
            <label htmlFor="cellValue">Cell Value:</label>
            <input
              type="text"
              id="cellValue"
              value={cellValue}
              onChange={(e) => setCellValue(e.target.value)}
            />
          </div>

          <button type="submit">Save Changes</button>
        </form>
      )}
    </div>
  );
}

export default ExcelForm;
