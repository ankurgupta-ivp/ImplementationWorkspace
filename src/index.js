import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

// Global CSS reset matching the prototype
const style = document.createElement('style');
style.textContent = `
  @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Roboto', sans-serif; }
  .btn {
    border: none; border-radius: 4px; padding: 6px 14px;
    font-family: 'Roboto', sans-serif; font-size: 12px; font-weight: 500;
    cursor: pointer; display: inline-flex; align-items: center; gap: 5px; white-space: nowrap;
  }
  .btn:disabled { opacity: 0.6; cursor: not-allowed; }
  .btn-primary { background: #404789; color: #fff; }
  .btn-primary:hover:not(:disabled) { background: #5259b3; }
  .btn-outlined { background: transparent; border: 1px solid #404789; color: #404789; }
  .btn-outlined:hover { background: #f0f0ff; }
  .btn-secondary { background: #da9b38; color: #fff; }
  .btn-secondary:hover { background: #c4882f; }
  .btn-ghost { background: transparent; border: 1px solid #c4c4c4; color: #444; }
  .btn-ghost:hover { border-color: #404789; color: #404789; }
  .btn-danger { background: #d32f2f; color: #fff; }
  .btn-danger:hover { background: #b71c1c; }
  .btn-sm { padding: 4px 10px; font-size: 11px; }
  .btn-xs { padding: 2px 7px; font-size: 11px; }
  ::-webkit-scrollbar { width: 7px; height: 7px; }
  ::-webkit-scrollbar-track { background: #f0f4f9; }
  ::-webkit-scrollbar-thumb { background: #c4c4c4; border-radius: 4px; }
`;
document.head.appendChild(style);

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<React.StrictMode><App /></React.StrictMode>);
