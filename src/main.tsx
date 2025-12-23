import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';
import type { OfficeExtension } from '@microsoft/office-js';

const rootElement = document.getElementById('root');

const renderApp = () => {
  if (!rootElement) {
    return;
  }

  ReactDOM.createRoot(rootElement).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};

if ((window as any).Office) {
  Office.onReady((info: OfficeExtension.Info) => {
    if (info.host === Office.HostType.Word) {
      renderApp();
    }
  });
} else {
  renderApp();
}
