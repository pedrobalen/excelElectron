import React from "react";
import ReactDOM from "react-dom/client";
import "./styles.css";
import ExcelImporter from "./ExcelImporter";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
    <ExcelImporter />
  </React.StrictMode>
);
