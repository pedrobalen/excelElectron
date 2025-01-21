import React, { useState } from "react";
import * as XLSX from "xlsx";
import "./styles.css";

function ExcelImporter() {
  const [tableData, setTableData] = useState([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [newColumnName, setNewColumnName] = useState("");
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [fileName, setFileName] = useState("");
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [originalWorkbook, setOriginalWorkbook] = useState(null);

  const saveToExcel = () => {
    if (tableData.length === 0 || !fileName || !originalWorkbook) return;

    // Get the original worksheet
    const ws = originalWorkbook.Sheets[originalWorkbook.SheetNames[0]];

    // Convert our current data to the format expected by xlsx
    const headers = Object.keys(tableData[0]);

    // Update the cells in the worksheet while preserving formatting
    tableData.forEach((row, rowIndex) => {
      headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({
          r: rowIndex + 2,
          c: colIndex,
        }); // +2 because of header rows
        const existingCell = ws[cellAddress] || {};

        // Preserve the cell's style and other properties while updating its value
        ws[cellAddress] = {
          ...existingCell,
          v: row[header], // new value
          w: row[header].toString(), // formatted text
        };
      });
    });

    // Save the file with preserved formatting
    XLSX.writeFile(originalWorkbook, fileName);
    setHasUnsavedChanges(false);
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      const reader = new FileReader();

      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const wb = XLSX.read(binaryStr, { type: "binary" });
        setOriginalWorkbook(wb); // Store the original workbook

        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: "",
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
          setHasUnsavedChanges(false);
        }
      };

      reader.readAsBinaryString(file);
    }
  };

  const updateTableData = (newData) => {
    setTableData(newData);
    setHasUnsavedChanges(true);
  };

  const handleAddColumn = () => {
    if (newColumnName.trim() === "") return;

    const updatedData = tableData.map((row) => ({
      ...row,
      [newColumnName]: "",
    }));

    updateTableData(updatedData);
    setNewColumnName("");
  };

  const handleAddRow = () => {
    const newRow = {};
    if (tableData.length > 0) {
      Object.keys(tableData[0]).forEach((key) => {
        newRow[key] = "";
      });
    }
    updateTableData([...tableData, newRow]);
  };

  const startEditing = (rowIndex, column, value) => {
    setEditingCell({ rowIndex, column });
    setEditValue(value.toString());
  };

  const handleCellEdit = (e) => {
    setEditValue(e.target.value);
  };

  const finishEditing = () => {
    if (editingCell) {
      const updatedData = [...tableData];
      updatedData[editingCell.rowIndex][editingCell.column] = editValue;
      updateTableData(updatedData);
      setEditingCell(null);
      setEditValue("");
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      finishEditing();
    }
  };

  const handleDeleteRow = (rowIndex) => {
    const updatedData = tableData.filter((_, index) => index !== rowIndex);
    updateTableData(updatedData);
  };

  const filteredData = tableData.filter((row) => {
    return Object.values(row).some((value) =>
      value.toString().toLowerCase().includes(searchTerm.toLowerCase())
    );
  });

  // Rest of the component remains the same...
  return (
    <div className="excel-container">
      <div className="controls-container">
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
        <input
          type="text"
          placeholder="Search..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="search-input"
        />
        {fileName && (
          <button
            onClick={saveToExcel}
            className={`action-button ${
              hasUnsavedChanges ? "action-button-highlight" : ""
            }`}
            disabled={!hasUnsavedChanges}
          >
            {hasUnsavedChanges ? "Save Changes" : "Saved"}
          </button>
        )}
      </div>

      <div className="controls-container">
        <input
          type="text"
          placeholder="New Column Name"
          value={newColumnName}
          onChange={(e) => setNewColumnName(e.target.value)}
          className="column-input"
        />
        <button onClick={handleAddColumn} className="action-button">
          Add Column
        </button>
        <button onClick={handleAddRow} className="action-button">
          Add Row
        </button>
      </div>

      {tableData.length > 0 && (
        <div className="table-container">
          <table className="excel-table">
            <thead>
              <tr>
                {Object.keys(tableData[0]).map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {Object.entries(row).map(([column, cell], cellIndex) => (
                    <td
                      key={cellIndex}
                      onClick={() =>
                        !editingCell && startEditing(rowIndex, column, cell)
                      }
                    >
                      {editingCell?.rowIndex === rowIndex &&
                      editingCell?.column === column ? (
                        <input
                          type="text"
                          value={editValue}
                          onChange={handleCellEdit}
                          onBlur={finishEditing}
                          onKeyDown={handleKeyDown}
                          autoFocus
                          className="cell-input"
                        />
                      ) : (
                        cell
                      )}
                    </td>
                  ))}
                  <td>
                    <button
                      onClick={() => handleDeleteRow(rowIndex)}
                      className="delete-button"
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default ExcelImporter;
