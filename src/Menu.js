import React from 'react';
import './menu.css';

function Menu() {
  return (
    <div className="menu-container">
      <nav>
        <ul className="menu-list">
          <li>
            <a href="#/excel-importer" className="menu-item">
              Visualizar tabela
            </a>
          </li>
          <li>
            <a href="#/excel-form" className="menu-item">
              Formulario
            </a>
          </li>
        </ul>
      </nav>
    </div>
  );
}

export default Menu;
