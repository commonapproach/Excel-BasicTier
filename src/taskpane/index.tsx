/* global document, Office, module, require */
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import App from './components/App';
import DialogContextProvider from './context/DialogContext';

const rootElement: HTMLElement | null = document.getElementById('container');
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  const language = Office.context.displayLanguage;
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <DialogContextProvider>
        <App />
      </DialogContextProvider>
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept('./components/App', () => {
    const NextApp = require('./components/App').default;
    root?.render(NextApp);
  });
}
