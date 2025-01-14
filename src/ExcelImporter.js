import React, { useState } from "react";
import * as XLSX from "xlsx";

function ExcelImporter() {
  const [tableData, setTableData] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();

      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const wb = XLSX.read(binaryStr, { type: "binary" });
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
        }
      };

      reader.readAsBinaryString(file);
    }
  };

  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      {tableData.length > 0 && (
        <table>
          <thead>
            <tr>
              {Object.keys(tableData[0]).map((header, index) => (
                <th key={index}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tableData.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {Object.values(row).map((cell, cellIndex) => (
                  <td key={cellIndex}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default ExcelImporter;
