import React, { useState } from "react";
import * as XLSX from "xlsx";

function ExcelImporter() {
  const [tableData, setTableData] = useState([]);
  const [searchTerm, setSearchTerm] = useState("");

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

  // Filter rows based on search term
  const filteredData = tableData.filter((row) => {
    return Object.values(row).some((value) =>
      value.toString().toLowerCase().includes(searchTerm.toLowerCase())
    );
  });

  return (
    <div>
      <div style={{ marginBottom: '20px' }}>
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          style={{ marginRight: '20px' }}
        />
        <input
          type="text"
          placeholder="Search..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          style={{
            padding: '8px',
            borderRadius: '4px',
            border: '1px solid #ccc',
            width: '200px'
          }}
        />
      </div>
      {tableData.length > 0 && (
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              {Object.keys(tableData[0]).map((header, index) => (
                <th 
                  key={index}
                  style={{
                    backgroundColor: '#f4f4f4',
                    padding: '12px',
                    borderBottom: '2px solid #ddd',
                    textAlign: 'left'
                  }}
                >
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredData.map((row, rowIndex) => (
              <tr 
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? '#ffffff' : '#f9f9f9'
                }}
              >
                {Object.values(row).map((cell, cellIndex) => (
                  <td 
                    key={cellIndex}
                    style={{
                      padding: '8px',
                      borderBottom: '1px solid #ddd'
                    }}
                  >
                    {cell}
                  </td>
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