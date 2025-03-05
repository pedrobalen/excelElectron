import React from "react";
import ReactDOM from "react-dom/client";
import "./styles.css";
import ExcelImporter from "./ExcelImporter";
import ExcelForm from "./ExcelForm";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
    <ExcelForm />
  </React.StrictMode>
);
