import React from 'react';
import ExcelImporter from './ExcelImporter';
import './App.css';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>DECA CUTTING OPTIMIZER</h1>
      </header>
      <main>
        <ExcelImporter />
      </main>
    </div>
  );
}

export default App;
