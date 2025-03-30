import React, { useState } from "react";
import * as XLSX from "xlsx";
import "./ExcelForm.css";

const ExcelForm = () => {
  const [excelData, setExcelData] = useState(null);
  const [selectedRow, setSelectedRow] = useState("");
  const [formFields, setFormFields] = useState([]);
  const [isAddingRow, setIsAddingRow] = useState(false);
  const [newRowName, setNewRowName] = useState("");
  const [isAddingColumn, setIsAddingColumn] = useState(false);
  const [newColumnName, setNewColumnName] = useState("");
  const [originalWorkbook, setOriginalWorkbook] = useState(null);
  const [fileName, setFileName] = useState("");
  const [activeSheet, setActiveSheet] = useState("");
  const [availableSheets, setAvailableSheets] = useState([]);
  const [successMessage, setSuccessMessage] = useState("");
  const [isDeletingRow, setIsDeletingRow] = useState(false);
  const [isDeletingColumn, setIsDeletingColumn] = useState(false);
  const [selectedColumnToDelete, setSelectedColumnToDelete] = useState("");

  const showSuccessMessage = (message) => {
    setSuccessMessage(message);
    setTimeout(() => setSuccessMessage(""), 3000);
  };

  const loadSheetData = (workbook, sheetName) => {
    const worksheet = workbook.Sheets[sheetName];

    for (let cell in worksheet) {
      if (cell[0] === "!") continue;

      if (worksheet[cell].w) {
        worksheet[cell].v = worksheet[cell].w;
      }
    }

    const data = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
      defval: "",
    });

    setExcelData(data);
    setActiveSheet(sheetName);
    setFormFields([]);
    setSelectedRow("");
  };

  const handleSheetChange = (sheetName) => {
    if (!originalWorkbook) return;
    loadSheetData(originalWorkbook, sheetName);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "binary" });
      setOriginalWorkbook(workbook);
      setAvailableSheets(workbook.SheetNames);

      loadSheetData(workbook, workbook.SheetNames[0]);
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

  const handleRowSelect = (rowName) => {
    setSelectedRow(rowName);

    if (!rowName || !excelData) {
      setFormFields([]);
      return;
    }

    const headers = excelData[1];
    const selectedRowData = excelData.find((row) => row[0] === rowName);

    console.log("Selected row:", selectedRowData);

    if (headers && selectedRowData) {
      let fields = [];

      for (let i = 1; i < headers.length; i++) {
        if (headers[i]) {
          let value = selectedRowData[i];

          if (typeof value === "number") {
            const originalValue = XLSX.SSF.format("General", value);
            if (originalValue.includes("/")) {
              value = originalValue;
            }
          }

          fields.push({
            header: String(headers[i]),
            value: String(value || ""),
            columnIndex: i,
          });
        }
      }

      console.log("Generated fields:", fields);
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

    if (!originalWorkbook || !activeSheet) return;

    const updatedRow = [selectedRow, ...formFields.map((field) => field.value)];
    const newExcelData = excelData.map((row) =>
      row[0] === selectedRow ? updatedRow : row
    );

    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    originalWorkbook.Sheets[activeSheet] = ws;

    XLSX.writeFile(originalWorkbook, fileName || "updated_research.xlsx");
  };

  const handleAddRow = () => {
    if (!newRowName || !excelData || !originalWorkbook || !activeSheet) return;

    const headers = excelData[1] || [];

    let nextEmptyRowIndex = excelData.findIndex((row) => !row[0]);
    if (nextEmptyRowIndex === -1) {
      nextEmptyRowIndex = excelData.length;
    }

    const newRow = Array(headers.length).fill("");
    newRow[0] = newRowName;

    const newExcelData = [...excelData];
    newExcelData[nextEmptyRowIndex] = newRow;
    setExcelData(newExcelData);

    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    originalWorkbook.Sheets[activeSheet] = ws;

    setNewRowName("");
    setIsAddingRow(false);
    showSuccessMessage("Linha adicionada com sucesso!");
  };

  const handleAddColumn = () => {
    if (!newColumnName || !excelData || !originalWorkbook || !activeSheet)
      return;

    const headers = excelData[1] || [];

    let nextEmptyColIndex = headers.findIndex((header) => !header);
    if (nextEmptyColIndex === -1) {
      nextEmptyColIndex = headers.length;
    }

    const newExcelData = excelData.map((row, index) => {
      const newRow = [...(row || [])];

      if (index === 1) {
        newRow[nextEmptyColIndex] = newColumnName;
      } else {
        newRow[nextEmptyColIndex] = "";
      }
      return newRow;
    });

    setExcelData(newExcelData);
    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    originalWorkbook.Sheets[activeSheet] = ws;

    setNewColumnName("");
    setIsAddingColumn(false);
    showSuccessMessage("Coluna adicionada com sucesso!");
  };

  const handleDeleteRow = () => {
    if (!selectedRow || !excelData || !originalWorkbook || !activeSheet) return;

    const newExcelData = excelData.filter((row) => row[0] !== selectedRow);
    setExcelData(newExcelData);

    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    originalWorkbook.Sheets[activeSheet] = ws;

    setSelectedRow("");
    setFormFields([]);
    setIsDeletingRow(false);
    showSuccessMessage("Linha excluída com sucesso!");
  };

  const handleDeleteColumn = () => {
    if (
      !selectedColumnToDelete ||
      !excelData ||
      !originalWorkbook ||
      !activeSheet
    )
      return;

    const headers = excelData[1] || [];
    const columnIndex = headers.findIndex(
      (header) => header === selectedColumnToDelete
    );

    if (columnIndex === -1) return;

    const newExcelData = excelData.map((row) => {
      const newRow = [...row];
      newRow.splice(columnIndex, 1);
      return newRow;
    });

    if (selectedRow && formFields.length > 0) {
      const newFormFields = formFields
        .filter((field) => field.header !== selectedColumnToDelete)
        .map((field) => ({
          ...field,
          columnIndex:
            field.columnIndex > columnIndex
              ? field.columnIndex - 1
              : field.columnIndex,
        }));
      setFormFields(newFormFields);
    }

    setExcelData(newExcelData);
    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    originalWorkbook.Sheets[activeSheet] = ws;

    setSelectedColumnToDelete("");
    setIsDeletingColumn(false);
    showSuccessMessage("Coluna excluída com sucesso!");
  };

  return (
    <div className="google-form-container">
      <div className="form-header">
        <div className="header-actions">
          <button
            onClick={() => window.history.back()}
            className="return-button"
          >
            ← Voltar
          </button>
          <h1>Formulário de Pesquisa</h1>
        </div>
        <p className="form-description">
          Carregue arquivo Excel e edite os dados da pesquisa
        </p>
      </div>

      {successMessage && (
        <div className="success-message">{successMessage}</div>
      )}

      <div className="form-card">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="file-input"
        />
      </div>

      {availableSheets.length > 0 && (
        <div className="sheet-tabs">
          {availableSheets.map((sheetName) => (
            <button
              key={sheetName}
              onClick={() => handleSheetChange(sheetName)}
              className={`sheet-tab ${
                activeSheet === sheetName ? "active" : ""
              }`}
            >
              {sheetName}
            </button>
          ))}
        </div>
      )}

      {excelData && (
        <>
          <div className="form-card">
            <div className="actions-container">
              <button
                onClick={() => setIsAddingRow(true)}
                className="action-button"
              >
                Adicionar Nova Linha
              </button>
              <button
                onClick={() => setIsAddingColumn(true)}
                className="action-button"
              >
                Adicionar Nova Coluna
              </button>
              <button
                onClick={() => setIsDeletingRow(true)}
                className="action-button delete"
              >
                Excluir Linha
              </button>
              <button
                onClick={() => setIsDeletingColumn(true)}
                className="action-button delete"
              >
                Excluir Coluna
              </button>
            </div>

            {isAddingRow && (
              <div className="add-form">
                <input
                  type="text"
                  value={newRowName}
                  onChange={(e) => setNewRowName(e.target.value)}
                  placeholder="Digite o nome da linha"
                  className="form-input"
                />
                <div className="button-group">
                  <button onClick={handleAddRow} className="submit-button">
                    Adicionar
                  </button>
                  <button
                    onClick={() => {
                      setIsAddingRow(false);
                      setNewRowName("");
                    }}
                    className="cancel-button"
                  >
                    Cancelar
                  </button>
                </div>
              </div>
            )}

            {isAddingColumn && (
              <div className="add-form">
                <input
                  type="text"
                  value={newColumnName}
                  onChange={(e) => setNewColumnName(e.target.value)}
                  placeholder="Digite o nome da coluna"
                  className="form-input"
                />
                <div className="button-group">
                  <button onClick={handleAddColumn} className="submit-button">
                    Adicionar
                  </button>
                  <button
                    onClick={() => {
                      setIsAddingColumn(false);
                      setNewColumnName("");
                    }}
                    className="cancel-button"
                  >
                    Cancelar
                  </button>
                </div>
              </div>
            )}

            {isDeletingRow && (
              <div className="add-form">
                <select
                  value={selectedRow}
                  onChange={(e) => setSelectedRow(e.target.value)}
                  className="form-input"
                >
                  <option value="">Selecione a linha para excluir</option>
                  {excelData.slice(2).map((row, index) => (
                    <option key={index} value={row[0]}>
                      {row[0]}
                    </option>
                  ))}
                </select>
                <div className="button-group">
                  <button
                    onClick={handleDeleteRow}
                    className="submit-button delete"
                  >
                    Excluir
                  </button>
                  <button
                    onClick={() => {
                      setIsDeletingRow(false);
                      setSelectedRow("");
                    }}
                    className="cancel-button"
                  >
                    Cancelar
                  </button>
                </div>
              </div>
            )}

            {isDeletingColumn && (
              <div className="add-form">
                <select
                  value={selectedColumnToDelete}
                  onChange={(e) => setSelectedColumnToDelete(e.target.value)}
                  className="form-input"
                >
                  <option value="">Selecione a coluna para excluir</option>
                  {(excelData[1] || []).slice(1).map((header, index) => (
                    <option key={index} value={header}>
                      {header}
                    </option>
                  ))}
                </select>
                <div className="button-group">
                  <button
                    onClick={handleDeleteColumn}
                    className="submit-button delete"
                  >
                    Excluir
                  </button>
                  <button
                    onClick={() => {
                      setIsDeletingColumn(false);
                      setSelectedColumnToDelete("");
                    }}
                    className="cancel-button"
                  >
                    Cancelar
                  </button>
                </div>
              </div>
            )}

            <label>Selecione a Linha:</label>
            <select
              value={selectedRow}
              onChange={(e) => handleRowSelect(e.target.value)}
            >
              <option value="">Escolha uma linha</option>
              {excelData.slice(1).map((row, index) => (
                <option key={index} value={row[0]}>
                  {row[0]}
                </option>
              ))}
            </select>
          </div>
        </>
      )}

      <div style={{ display: "none" }}>
        <p>Selected Row: {selectedRow}</p>
        <p>Form Fields Count: {formFields.length}</p>
      </div>

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
            Salvar Alterações
          </button>
        </form>
      )}
    </div>
  );
};

export default ExcelForm;
