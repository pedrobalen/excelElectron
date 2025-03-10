import React from "react";
import { HashRouter, Routes, Route } from 'react-router-dom';
import ReactDOM from "react-dom/client";
import ExcelImporter from "./ExcelImporter";
import './index.css';
import ExcelForm from "./ExcelForm";
import Menu from "./Menu";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
    <HashRouter>
      <Routes>
        <Route path="/" element={<Menu />} />
        <Route path="/excel-importer" element={<ExcelImporter />} />
        <Route path="/excel-form" element={<ExcelForm />} />
      </Routes>
    </HashRouter>
  </React.StrictMode>
);
