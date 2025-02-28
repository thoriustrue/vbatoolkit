import { Buffer } from 'buffer';
globalThis.Buffer = Buffer;
window.Buffer = Buffer;

import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.js';
import { ErrorBoundary } from './components/ErrorLogger';
import './index.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
