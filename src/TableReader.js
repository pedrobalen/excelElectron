import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import "./TableReader.css";

function ExcelImporter() {
  const [tableData, setTableData] = useState([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [correlationTerms, setCorrelationTerms] = useState([""]);
  const [isCorrelationMode, setIsCorrelationMode] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [fileName, setFileName] = useState("");
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [originalWorkbook, setOriginalWorkbook] = useState(null);
  const [currentSheet, setCurrentSheet] = useState("");
  const [availableSheets, setAvailableSheets] = useState([]);
  const [isColumnModalOpen, setIsColumnModalOpen] = useState(false);
  const [selectedColumn, setSelectedColumn] = useState("");
  const [isColumnDropdownOpen, setIsColumnDropdownOpen] = useState(false);

  const saveToExcel = () => {
    if (tableData.length === 0 || !fileName || !originalWorkbook) return;
    const ws = originalWorkbook.Sheets[originalWorkbook.SheetNames[0]];
    const headers = Object.keys(tableData[0]);

    tableData.forEach((row, rowIndex) => {
      headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({
          r: rowIndex + 2,
          c: colIndex,
        });
        const existingCell = ws[cellAddress] || {};

        ws[cellAddress] = {
          ...existingCell,
          v: row[header],
          w: row[header].toString(),
        };
      });
    });

    XLSX.writeFile(originalWorkbook, fileName);
    setHasUnsavedChanges(false);
  };

  const loadSheetData = (wb, sheetName) => {
    const ws = wb.Sheets[sheetName];

    // Ensure cells with "/" are treated as text
    for (let cell in ws) {
      if (cell[0] === "!") continue;

      // If the cell has a value and contains a slash
      if (
        ws[cell].v &&
        typeof ws[cell].v === "string" &&
        ws[cell].v.includes("/")
      ) {
        ws[cell].t = "s"; // Force cell type to be string
        ws[cell].w = ws[cell].v; // Ensure the formatted text matches the value
      }

      // For any cell with a formatted value, use that instead
      if (ws[cell].w) {
        ws[cell].v = ws[cell].w;
      }
    }

    const jsonData = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: false,
      defval: "",
    });

    if (jsonData.length > 1) {
      const headers = jsonData[1];
      const rows = jsonData.slice(2);
      const formattedData = rows.map((row) => {
        const rowObject = {};
        row.forEach((cell, index) => {
          // Ensure the cell value is treated as a string
          rowObject[headers[index]] = String(cell || "");
        });
        return rowObject;
      });

      setTableData(formattedData);
      setCurrentSheet(sheetName);
      setHasUnsavedChanges(false);
    }
  };

  const handleSheetChange = (sheetName) => {
    if (hasUnsavedChanges) {
      const confirmChange = window.confirm(
        "Voce tem mudancas não salvas, deseja continuar?"
      );
      if (!confirmChange) return;
    }
    loadSheetData(originalWorkbook, sheetName);
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      const reader = new FileReader();

      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const wb = XLSX.read(binaryStr, { type: "binary" });
        setOriginalWorkbook(wb);
        setAvailableSheets(wb.SheetNames);
        loadSheetData(wb, wb.SheetNames[0]);
      };

      reader.readAsBinaryString(file);
    }
  };

  const updateTableData = (newData) => {
    setTableData(newData);
    setHasUnsavedChanges(true);
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

  const handleReturn = () => {
    window.history.back();
  };

  const highlightText = (text, terms) => {
    if (!text || !terms || terms.length === 0) return text;

    let parts = [text.toString()];
    terms.forEach((term, index) => {
      if (!term.trim()) return;

      const newParts = [];
      parts.forEach((part) => {
        if (typeof part === "string") {
          const splitPart = part.split(new RegExp(`(${term})`, "gi"));
          splitPart.forEach((subPart, i) => {
            if (subPart.toLowerCase() === term.toLowerCase()) {
              newParts.push(
                <span key={`${index}-${i}`} className={`term${index + 1}-text`}>
                  {subPart}
                </span>
              );
            } else if (subPart) {
              newParts.push(subPart);
            }
          });
        } else {
          newParts.push(part);
        }
      });
      parts = newParts;
    });

    return parts;
  };

  const calculateCorrelation = (terms) => {
    const validTerms = terms.filter((term) => term.trim());
    if (!validTerms.length) return null;

    const termsLower = validTerms.map((term) => term.toLowerCase());

    const correlationData = tableData.map((row, index) => {
      const rowStr = Object.values(row).join(" ").toLowerCase();
      const termPresence = termsLower.map((term) => ({
        hasTerm: rowStr.includes(term),
        columns: Object.entries(row)
          .filter(([_, value]) => value.toString().toLowerCase().includes(term))
          .map(([key]) => key),
      }));

      return {
        termPresence,
        rowIndex: index,
        rowData: row,
      };
    });

    const stats = correlationData.reduce(
      (acc, curr) => {
        const allTermsPresent = curr.termPresence.every((t) => t.hasTerm);

        if (allTermsPresent) {
          acc.allTerms++;
          if (acc.examples.length < 3) {
            acc.examples.push({
              rowIndex: curr.rowIndex,
              rowData: curr.rowData,
              termColumns: curr.termPresence.map((t) => t.columns),
            });
          }
        }

        curr.termPresence.forEach((presence, idx) => {
          if (presence.hasTerm) {
            acc.termCounts[idx] = (acc.termCounts[idx] || 0) + 1;
          }
        });

        return acc;
      },
      {
        allTerms: 0,
        termCounts: Array(termsLower.length).fill(0),
        examples: [],
      }
    );

    const totalRows = tableData.length;
    const termRates = stats.termCounts.map((count) => count / totalRows);
    const allTermsRate = stats.allTerms / totalRows;
    const expectedRate = termRates.reduce((acc, rate) => acc * rate, 1);

    return {
      allTermsOccurrence: stats.allTerms,
      termTotals: stats.termCounts,
      termRates: termRates.map((rate) => rate * 100),
      allTermsRate: allTermsRate * 100,
      correlationStrength: allTermsRate / (expectedRate || 1),
      examples: stats.examples,
      validTerms,
    };
  };

  const correlationResults = useMemo(() => {
    return isCorrelationMode ? calculateCorrelation(correlationTerms) : null;
  }, [isCorrelationMode, correlationTerms, tableData]);

  const filteredData = useMemo(() => {
    if (!tableData.length) return [];

    if (isCorrelationMode) {
      const validTerms = correlationTerms.filter((term) => term.trim());
      if (!validTerms.length) return tableData;

      return tableData.filter((row) => {
        const rowStr = Object.values(row).join(" ").toLowerCase();
        return validTerms.every((term) => rowStr.includes(term.toLowerCase()));
      });
    }

    return tableData.filter((row) => {
      return Object.values(row).some((value) =>
        value.toString().toLowerCase().includes(searchTerm.toLowerCase())
      );
    });
  }, [tableData, searchTerm, correlationTerms, isCorrelationMode]);

  const handleAddTerm = () => {
    setCorrelationTerms([...correlationTerms, ""]);
  };

  const handleRemoveTerm = (index) => {
    if (correlationTerms.length > 1) {
      setCorrelationTerms(correlationTerms.filter((_, i) => i !== index));
    }
  };

  const handleTermChange = (index, value) => {
    const newTerms = [...correlationTerms];
    newTerms[index] = value;
    setCorrelationTerms(newTerms);
  };

  const getColumnData = (columnName) => {
    if (!tableData.length || !columnName) return [];
    return tableData.map((row) => row[columnName]);
  };

  const getColumnNames = () => {
    if (!tableData.length) return [];
    return Object.keys(tableData[0]);
  };

  const ColumnDataModal = () => {
    if (!isColumnModalOpen || !selectedColumn) return null;

    const columnData = getColumnData(selectedColumn);

    return (
      <div className="column-modal-overlay">
        <div className="column-modal">
          <div className="column-modal-header">
            <h3>{selectedColumn}</h3>
            <button
              onClick={() => {
                setIsColumnModalOpen(false);
                setSelectedColumn("");
              }}
              className="close-modal-button"
            >
              ×
            </button>
          </div>
          <div className="column-modal-content">
            <div className="column-data-list">
              <table className="column-data-table">
                <tbody>
                  {columnData.map((value, index) => (
                    <tr key={index}>
                      <td>{value}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="excel-container">
      <div className="controls-container">
        <button onClick={handleReturn} className="return-button">
          ← Voltar
        </button>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />

        {tableData.length > 0 && (
          <div className="column-consultation">
            <button
              onClick={() => setIsColumnDropdownOpen(!isColumnDropdownOpen)}
              className="consult-column-button"
            >
              Consultar Coluna
            </button>
            {isColumnDropdownOpen && (
              <div className="column-dropdown">
                {getColumnNames().map((columnName, index) => (
                  <button
                    key={index}
                    className="column-option"
                    onClick={() => {
                      setSelectedColumn(columnName);
                      setIsColumnModalOpen(true);
                      setIsColumnDropdownOpen(false);
                    }}
                  >
                    {columnName}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        <div className="search-container">
          <div className="search-mode-toggle">
            <button
              onClick={() => {
                setIsCorrelationMode(!isCorrelationMode);
                if (!isCorrelationMode) {
                  setCorrelationTerms([""]);
                }
              }}
              className={`mode-button ${isCorrelationMode ? "active" : ""}`}
            >
              {isCorrelationMode ? "Modo Correlação" : "Modo Busca Simples"}
            </button>
          </div>
          <div className="search-inputs">
            {isCorrelationMode ? (
              <div className="correlation-terms-container">
                {correlationTerms.map((term, index) => (
                  <div key={index} className="correlation-term-input">
                    <input
                      type="text"
                      placeholder={`Termo ${index + 1}...`}
                      value={term}
                      onChange={(e) => handleTermChange(index, e.target.value)}
                      className="search-input"
                    />
                    {correlationTerms.length > 1 && (
                      <button
                        onClick={() => handleRemoveTerm(index)}
                        className="remove-term-button"
                        title="Remover termo"
                      >
                        ×
                      </button>
                    )}
                  </div>
                ))}
                <button
                  onClick={handleAddTerm}
                  className="add-term-button"
                  title="Adicionar novo termo"
                >
                  + Adicionar termo
                </button>
              </div>
            ) : (
              <input
                type="text"
                placeholder="Procurar..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="search-input"
              />
            )}
          </div>
          {isCorrelationMode && correlationResults && (
            <div className="correlation-stats">
              <div className="correlation-header">
                Estatísticas de Correlação
              </div>
              <div className="correlation-details">
                <div className="stat-group">
                  {correlationResults.validTerms.map((term, index) => (
                    <div key={index} className="stat-item">
                      <span className="stat-label">
                        Termo {index + 1} ("{term}"):
                      </span>
                      <span className="stat-value">
                        {correlationResults.termTotals[index]} ocorrências (
                        {correlationResults.termRates[index].toFixed(1)}%)
                      </span>
                    </div>
                  ))}
                  {correlationResults.validTerms.length > 0 && (
                    <div className="stat-item">
                      <span className="stat-label">Co-ocorrências:</span>
                      <span className="stat-value">
                        {correlationResults.allTermsOccurrence} (
                        {correlationResults.allTermsRate.toFixed(1)}% do total)
                      </span>
                    </div>
                  )}
                </div>

                {correlationResults.validTerms.length > 0 && (
                  <div className="correlation-strength-container">
                    <div className="correlation-strength">
                      <span className="strength-label">
                        Força da Correlação:
                      </span>
                      <span
                        className={`strength-value ${
                          correlationResults.correlationStrength > 1.2
                            ? "strong-positive"
                            : correlationResults.correlationStrength < 0.8
                            ? "strong-negative"
                            : "neutral"
                        }`}
                      >
                        {correlationResults.correlationStrength.toFixed(2)}x
                      </span>
                      <span className="correlation-hint">
                        {correlationResults.correlationStrength > 1.2
                          ? " (Correlação positiva forte)"
                          : correlationResults.correlationStrength > 1
                          ? " (Correlação positiva leve)"
                          : correlationResults.correlationStrength < 0.8
                          ? " (Correlação negativa forte)"
                          : correlationResults.correlationStrength < 1
                          ? " (Correlação negativa leve)"
                          : " (Sem correlação)"}
                      </span>
                    </div>
                  </div>
                )}

                {correlationResults.examples.length > 0 && (
                  <div className="correlation-examples">
                    <h4>Exemplos de Co-ocorrência:</h4>
                    {correlationResults.examples.map((example, index) => (
                      <div key={index} className="example-item">
                        <div className="example-header">
                          Exemplo {index + 1}:
                        </div>
                        <div className="example-content">
                          {Object.entries(example.rowData).map(
                            ([column, value], i) => (
                              <div
                                key={i}
                                className={`example-field ${
                                  example.termColumns.some((termCols) =>
                                    termCols.includes(column)
                                  )
                                    ? `term${
                                        example.termColumns.findIndex(
                                          (termCols) =>
                                            termCols.includes(column)
                                        ) + 1
                                      }-highlight`
                                    : ""
                                }`}
                              >
                                <strong>{column}:</strong> {value}
                              </div>
                            )
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}
          <span className="row-counter">
            {filteredData.length} resultados em {tableData.length} linhas
          </span>
        </div>
        {fileName && (
          <button
            onClick={saveToExcel}
            className={`action-button ${
              hasUnsavedChanges ? "action-button-highlight" : ""
            }`}
            disabled={!hasUnsavedChanges}
          >
            {hasUnsavedChanges ? "Salvar Alterações" : "Salvo"}
          </button>
        )}
      </div>

      {availableSheets.length > 0 && (
        <div className="sheet-tabs">
          {availableSheets.map((sheetName) => (
            <button
              key={sheetName}
              onClick={() => handleSheetChange(sheetName)}
              className={`sheet-tab ${
                currentSheet === sheetName ? "active" : ""
              }`}
            >
              {sheetName}
            </button>
          ))}
        </div>
      )}

      {tableData.length > 0 && (
        <div className="table-container">
          <table className="excel-table">
            <thead>
              <tr>
                {Object.keys(tableData[0]).map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredData.map((row, rowIndex) => (
                <tr
                  key={rowIndex}
                  className={isCorrelationMode ? "correlation-row" : ""}
                >
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
                      ) : isCorrelationMode ? (
                        highlightText(
                          cell,
                          correlationResults?.validTerms || []
                        )
                      ) : (
                        highlightText(cell, [searchTerm])
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {isColumnModalOpen && <ColumnDataModal />}
    </div>
  );
}

export default ExcelImporter;
