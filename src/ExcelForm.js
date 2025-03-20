import React, { useState } from "react";
import * as XLSX from "xlsx";
import "./ExcelForm.css";

const ExcelForm = () => {
  const [excelData, setExcelData] = useState(null);
  const [selectedResearch, setSelectedResearch] = useState("");
  const [formFields, setFormFields] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      setExcelData(data);
    };

    reader.readAsBinaryString(file);
  };

  const determineFieldType = (value) => {
    if (!value || typeof value !== "string") return "text";

    const lowerValue = value.toLowerCase();
    if (lowerValue === "yes" || lowerValue === "no") return "radio";
    if (value.includes("\n")) return "textarea";
    return "text";
  };

  const handleResearchSelect = (research) => {
    setSelectedResearch(research);

    if (!research || !excelData) {
      setFormFields([]);
      return;
    }

    // Get headers from the second row (index 1) instead of first row
    const headers = excelData[1];
    const selectedRow = excelData.find((row) => row[0] === research);

    console.log("Headers from second row:", headers);
    console.log("Selected Row:", selectedRow);

    if (headers && selectedRow) {
      let fields = [];

      // Start from index 1 to skip the research name column
      for (let i = 1; i < headers.length; i++) {
        if (headers[i]) {
          // Only create fields for non-empty headers
          fields.push({
            header: String(headers[i]),
            value: String(selectedRow[i] || ""),
            columnIndex: i,
          });
        }
      }

      console.log("Generated Fields:", fields);
      setFormFields(fields);
    }
  };

  const handleFieldChange = (columnIndex, value) => {
    setFormFields((prev) =>
      prev.map((field) =>
        field.columnIndex === columnIndex ? { ...field, value } : field
      )
    );
  };

  const handleSubmit = (e) => {
    e.preventDefault();

    const updatedRow = [
      selectedResearch,
      ...formFields.map((field) => field.value),
    ];

    const newExcelData = excelData.map((row) =>
      row[0] === selectedResearch ? updatedRow : row
    );

    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "updated_research.xlsx");
  };

  return (
    <div className="google-form-container">
      <div className="form-header">
        <h1>Research Data Form</h1>
        <p className="form-description">
          Upload Excel file and edit research data
        </p>
      </div>

      <div className="form-card">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="file-input"
        />
      </div>

      {excelData && (
        <div className="form-card">
          <label>Select Research:</label>
          <select
            value={selectedResearch}
            onChange={(e) => handleResearchSelect(e.target.value)}
          >
            <option value="">Choose a research</option>
            {excelData.slice(1).map((row, index) => (
              <option key={index} value={row[0]}>
                {row[0]}
              </option>
            ))}
          </select>
        </div>
      )}

      {/* Debug information */}
      <div style={{ display: "none" }}>
        <p>Selected Research: {selectedResearch}</p>
        <p>Form Fields Count: {formFields.length}</p>
      </div>

      {/* Form fields */}
      {formFields.length > 0 && (
        <form onSubmit={handleSubmit}>
          {formFields.map((field) => (
            <div key={field.columnIndex} className="form-card">
              <h3>{field.header}</h3>
              <input
                type="text"
                value={field.value}
                onChange={(e) =>
                  handleFieldChange(field.columnIndex, e.target.value)
                }
                className="form-input"
              />
            </div>
          ))}
          <button type="submit" className="submit-button">
            Save Changes
          </button>
        </form>
      )}
    </div>
  );
};

export default ExcelForm;
